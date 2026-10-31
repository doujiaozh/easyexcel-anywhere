/**
 * Streaming xlsx writer.
 *
 * Data never accumulates: rows are consumed from an async iterable and
 * immediately encoded into sheet XML, which is pushed through a raw
 * DEFLATE encoder straight into the ZIP stream. With a path target the
 * output lands on disk via a Node fs stream; with a WritableStream
 * target the bytes go wherever the caller wants (browser download,
 * in-memory buffer, service worker ...).
 */

import {
  cellRef,
  contentTypesXml,
  rootRelsXml,
  SHEET_XML_EPILOGUE,
  SHEET_XML_PROLOGUE,
  stylesXml,
  workbookRelsXml,
  workbookXml,
  xmlEscape,
} from '../core/xml-fragments.js';
import { ZipWriter } from '../core/zip-writer.js';
import { buildWriteBindings, toCellValue } from '../read/mapper.js';

export interface WriteOptions {
  /** Sheet name. Default: 'Sheet1'. */
  sheet?: string;
  /** Row class decorated with @ExcelProperty (headers come from the class). */
  class?: new (...args: unknown[]) => unknown;
  /** Explicit headers when writing plain arrays. */
  headers?: string[];
  /** DEFLATE level 0-9. Default 6. */
  compressionLevel?: number;
}

function isNode(): boolean {
  const g = globalThis as { process?: { versions?: { node?: string } } };
  return typeof g.process?.versions?.node === 'string';
}

async function* once(str: string): AsyncGenerator<Uint8Array> {
  yield new TextEncoder().encode(str);
}

async function* asRows(rows: AsyncIterable<unknown> | Iterable<unknown>): AsyncGenerator<unknown> {
  if (Symbol.asyncIterator in (rows as object)) {
    yield* rows as AsyncIterable<unknown>;
    return;
  }
  for (const r of rows as Iterable<unknown>) yield r;
}

function buildRowXml(values: Array<string | number | boolean | null>, rowIndex: number): string {
  let end = values.length;
  while (end > 0 && (values[end - 1] === null || values[end - 1] === undefined)) end--;
  let s = `<row r="${rowIndex}">`;
  for (let i = 0; i < end; i++) {
    const v = values[i];
    const ref = cellRef(rowIndex, i);
    if (v === null || v === undefined) {
      s += `<c r="${ref}"/>`;
      continue;
    }
    if (typeof v === 'number' || typeof v === 'bigint') {
      s += `<c r="${ref}"><v>${v}</v></c>`;
      continue;
    }
    if (typeof v === 'boolean') {
      s += `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`;
      continue;
    }
    s += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(String(v))}</t></is></c>`;
  }
  return s + '</row>\n';
}

export class ExcelWriter {
  private zw: ZipWriter | null = null;
  private nodeStream: ReturnType<typeof import('node:fs').createWriteStream> | null = null;
  private nodeDone: Promise<void> | null = null;

  constructor(
    private readonly target: string | WritableStream<Uint8Array>,
    private readonly opts: WriteOptions = {}
  ) {}

  private async open(): Promise<ZipWriter> {
    if (this.zw) return this.zw;
    let out: WritableStream<Uint8Array>;
    if (typeof this.target === 'string') {
      if (!isNode()) {
        throw new Error('A filesystem path target requires Node.js; pass a WritableStream in browsers');
      }
      const fs = await import('node:fs');
      const ws = fs.createWriteStream(this.target);
      this.nodeStream = ws;
      // capture completion eagerly: 'close' may fire before we would attach later
      this.nodeDone = new Promise<void>((resolve, reject) => {
        ws.on('close', () => resolve());
        ws.on('error', (e) => reject(e));
      });
      const stream = await import('node:stream');
      out = stream.Writable.toWeb(ws) as WritableStream<Uint8Array>;
    } else {
      out = this.target;
    }
    this.zw = new ZipWriter(out);
    return this.zw;
  }

  /**
   * Stream `rows` into the workbook. One call per file; rows may be an
   * (async) iterable of class instances or of plain arrays.
   */
  async writeRows(rows: AsyncIterable<unknown> | Iterable<unknown>): Promise<void> {
    const zw = await this.open();
    const sheetName = this.opts.sheet ?? 'Sheet1';
    const level = this.opts.compressionLevel ?? 6;

    await zw.addFile('[Content_Types].xml', once(contentTypesXml([sheetName])));
    await zw.addFile('_rels/.rels', once(rootRelsXml()));
    await zw.addFile('xl/workbook.xml', once(workbookXml([sheetName])));
    await zw.addFile('xl/_rels/workbook.xml.rels', once(workbookRelsXml([sheetName])));
    await zw.addFile('xl/styles.xml', once(stylesXml()));

    const schema = this.opts.class ? buildWriteBindings(this.opts.class) : null;
    const headers = this.opts.headers ?? (schema ? schema.map((m) => m.header) : null);

    zw.level = level;
    await zw.addFile('xl/worksheets/sheet1.xml', this.sheetContent(rows, schema, headers));
  }

  private async *sheetContent(
    rows: AsyncIterable<unknown> | Iterable<unknown>,
    schema: ReturnType<typeof buildWriteBindings> | null,
    headers: string[] | null
  ): AsyncGenerator<Uint8Array> {
    const enc = new TextEncoder();
    let batch = SHEET_XML_PROLOGUE;
    let rowIndex = 1;

    if (headers) {
      batch += buildRowXml(headers, rowIndex++);
    }

    for await (const row of asRows(rows)) {
      const values = schema
        ? schema.map((m) => toCellValue((row as Record<string, unknown>)[m.key], m))
        : (row as Array<string | number | boolean | null>);
      batch += buildRowXml(values, rowIndex++);
      if (batch.length > 0x10000) {
        yield enc.encode(batch);
        batch = '';
      }
    }
    yield enc.encode(batch + SHEET_XML_EPILOGUE);
  }

  /** Finalize the archive and close the output. */
  async close(): Promise<void> {
    if (!this.zw) {
      // nothing written: still produce a valid (empty) workbook
      await this.writeRows([]);
    }
    await this.zw!.finish();
    if (this.nodeDone) await this.nodeDone;
  }
}
