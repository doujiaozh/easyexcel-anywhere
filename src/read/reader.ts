/**
 * Reading pipeline: xlsx byte source -> streaming ZIP scan -> sheet XML rows.
 *
 * Entries are scanned sequentially (worksheets are assumed to follow
 * workbook.xml / its rels / sharedStrings in the archive, which is what
 * Excel, LibreOffice and this library's writer produce). Backpressure is
 * end-to-end: the file stream is only pulled while the caller is pulling rows.
 */

import type { ByteSource } from '../core/bytes.js';
import { ZipReader, type ZipEntry } from '../core/zip-reader.js';
import {
  parseRels,
  parseSharedStrings,
  parseWorkbook,
  SheetXmlReader,
  type RawRow,
} from '../core/xml-reader.js';
import { buildReadBindings, instantiate, type ReadBinding } from './mapper.js';

export type { RawRow };

export interface ReadOptions<T = RawRow> {
  /** Sheet name or 0-based index. Default: first sheet. */
  sheet?: string | number;
  /** Map rows onto instances of this class (decorated with @ExcelProperty). */
  class?: new (...args: unknown[]) => T;
  /** Row callback (registered handlers receive rows before the generator yields them). */
  onRow?: (row: T) => void | Promise<void>;
  /** Stop after this many data rows. */
  max?: number;
}

const WORKSHEET_PREFIX = 'xl/worksheets/';

/** rels targets are package-absolute ("/xl/...") or relative to the rels part ("xl/..."). */
function normalizeWorksheetPath(target: string): string {
  const t = target.replace(/\\/g, '/');
  if (t.startsWith('/')) return t.slice(1);
  return t.startsWith('xl/') ? t : 'xl/' + t;
}

export class ExcelReader<T = RawRow> {
  constructor(private readonly source: ByteSource, private readonly opts: ReadOptions<T> = {}) {}

  /** Rows as mapped class instances (when `class` given) or RawRow objects. */
  async *rows(): AsyncGenerator<T> {
    const zr = new ZipReader(this.source);
    const sheetXml = new SheetXmlReader();

    let sheets: Array<{ name: string; rId: string }> | null = null;
    let rels: Map<string, string> | null = null;
    let sharedStrings: string[] | null = null;
    let targetPath: string | null = null;
    let worksheetCounter = 0;

    const select = (): string | null => {
      const sel = this.opts.sheet ?? 0;
      if (typeof sel === 'number') {
        if (sheets && rels) {
          const s = sheets[sel];
          if (s) {
            const t = rels.get(s.rId);
            if (t) return normalizeWorksheetPath(t);
          }
          return null;
        }
        return null; // resolved positionally below
      }
      if (sheets && rels) {
        const s = sheets.find((x) => x.name === sel);
        if (s) {
          const t = rels.get(s.rId);
          if (t) return normalizeWorksheetPath(t);
          throw new Error(`Relationship ${s.rId} of sheet "${sel}" not found in workbook rels`);
        }
        throw new Error(`Sheet "${sel}" not found in workbook; available: ${sheets.map((s) => s.name).join(', ')}`);
      }
      return null;
    };

    let bindings: ReadBinding[] | null = null;
    let yielded = 0;

    while (true) {
      const entry = await zr.next();
      if (!entry) break;
      const name = entry.name.replace(/\\/g, '/');

      if (name === 'xl/workbook.xml') {
        sheets = await parseWorkbook(entry.open());
        continue;
      }
      if (name === 'xl/_rels/workbook.xml.rels') {
        rels = await parseRels(entry.open());
        continue;
      }
      if (name === 'xl/sharedStrings.xml') {
        sharedStrings = await parseSharedStrings(entry.open());
        continue;
      }
      if (name.startsWith(WORKSHEET_PREFIX)) {
        worksheetCounter++;
        if (targetPath === null) {
          targetPath = select(); // may stay null: fall back to positional match
        }
        const wanted =
          targetPath === name ||
          (targetPath === null && worksheetCounter - 1 === (this.opts.sheet ?? 0));
        if (!wanted) {
          await entry.skip();
          continue;
        }
        // ---- stream the selected sheet ----
        const rows = sheetXml.rows(entry.open(), sharedStrings);
        if (this.opts.class) {
          for await (const row of rows) {
            const raw = row as RawRow;
            if (bindings === null) {
              bindings = buildReadBindings(this.opts.class, raw.cells);
              continue; // header row consumed
            }
            const obj = instantiate(this.opts.class);
            for (const b of bindings) {
              obj[b.key] = b.convert(raw.cells[b.col] ?? null);
            }
            if (this.opts.onRow) await this.opts.onRow(obj as T);
            yield obj as T;
            if (this.opts.max && ++yielded >= this.opts.max) return;
          }
          return;
        }
        for await (const row of rows) {
          if (this.opts.onRow) await this.opts.onRow(row as T);
          yield row as T;
          if (this.opts.max && ++yielded >= this.opts.max) return;
        }
        return;
      }
      await entry.skip();
    }
    throw new Error(
      `Sheet ${typeof this.opts.sheet === 'string' ? `"${this.opts.sheet}"` : this.opts.sheet ?? 0} not found in archive`
    );
  }

  /** Convenience: await every row with a callback (event-style API). */
  async forEach(onRow: (row: T) => void | Promise<void>): Promise<void> {
    const opts = { ...this.opts, onRow };
    const reader = new ExcelReader<T>(this.source, opts);
    for await (const _ of reader.rows()) {
      void _;
    }
  }

  /** Convenience: collect every row into an array. Beware of memory. */
  async toArray(): Promise<T[]> {
    const out: T[] = [];
    for await (const row of this.rows()) out.push(row);
    return out;
  }
}
