/**
 * SAX-driven streaming XML readers for the parts of an .xlsx package we care about.
 *
 * `SheetXmlReader.rows()` is the hot path: it is a *pull-based* async generator.
 * The XML source is only fed one chunk at a time, and only until the next row
 * is complete - so parsing pauses while the consumer processes the row,
 * which is exactly the backpressure needed to keep memory flat.
 */

import sax from 'sax';

/** A raw worksheet row: 1-based row index + cell values (null for gaps/empties). */
export interface RawRow {
  rowIndex: number;
  cells: (string | null)[];
}

function colToIndex(ref: string): number {
  let i = 0;
  for (const ch of ref) {
    const c = ch.charCodeAt(0);
    if (c >= 65 && c <= 90) i = i * 26 + (c - 64);
    else if (c >= 97 && c <= 122) i = i * 26 + (c - 96);
    else break;
  }
  return i - 1;
}

function rowNumberOf(ref: string): number {
  const m = /(\d+)\s*$/.exec(ref);
  return m ? parseInt(m[1], 10) : 0;
}

const SAX_OPTIONS: sax.SAXOptions = {
  trim: false,
  normalize: false,
  lowercase: true,
  xmlns: false,
  position: false,
};

/** Worksheet XML -> async stream of rows. */
export class SheetXmlReader {
  /**
   * @param stream inflated bytes of the sheet XML part
   * @param sharedStrings shared-strings table (used when cells carry t="s");
   *                       pass null when the sheet only uses inline strings
   */
  async *rows(
    stream: ReadableStream<Uint8Array>,
    sharedStrings: string[] | null
  ): AsyncGenerator<RawRow> {
    const parser = sax.parser(false, SAX_OPTIONS);

    // Rows completed inside the current parser chunk. A single parser.write()
    // can finish many rows at once (a 64KB chunk easily holds hundreds), so a
    // single slot would silently drop all but the last one. The queue is drained
    // by the consumer before the next chunk is pulled - that IS the backpressure.
    const pendingRows: RawRow[] = [];
    let cells: (string | null)[] = [];
    let nextCellIdx = 0;
    let cellType = '';
    let cellRef = '';
    let text = '';
    let inValue = false; // inside <v> or <is>
    let rowNum = 0;
    let fatal: Error | null = null;

    parser.onopentag = (node: sax.Tag) => {
      switch (node.name) {
        case 'row': {
          cells = [];
          nextCellIdx = 0;
          const r = node.attributes['r'];
          rowNum = r ? rowNumberOf(r) : rowNum + 1;
          break;
        }
        case 'c': {
          cellRef = node.attributes['r'] ?? '';
          cellType = node.attributes['t'] ?? '';
          text = '';
          break;
        }
        case 'v':
        case 'is':
          inValue = true;
          break;
        default:
          break;
      }
    };

    parser.ontext = (t: string) => {
      if (inValue) text += t;
    };

    parser.oncdata = (c: string) => {
      if (inValue) text += c;
    };

    parser.onclosetag = (name: string) => {
      switch (name) {
        case 'v':
        case 'is':
          inValue = false;
          break;
        case 'c': {
          const idx = cellRef ? colToIndex(cellRef) : nextCellIdx;
          while (cells.length < idx) cells.push(null);
          cells[idx] = resolveCell(cellType, text, sharedStrings);
          nextCellIdx = idx + 1;
          break;
        }
        case 'row':
          pendingRows.push({ rowIndex: rowNum, cells });
          break;
        default:
          break;
      }
    };

    parser.onerror = (e: Error) => {
      fatal = fatal ?? e;
    };

    const reader = stream.getReader();
    const decoder = new TextDecoder('utf-8');
    try {
      while (true) {
        if (pendingRows.length > 0) {
          yield pendingRows.shift()!;
          continue;
        }
        if (fatal) throw fatal;
        const { value, done } = await reader.read();
        if (done) {
          const tail = decoder.decode();
          if (tail) parser.write(tail);
          parser.end();
          if (fatal) throw fatal;
          if (pendingRows.length > 0) continue; // rows flushed by parser.end()
          break;
        }
        const piece = decoder.decode(value, { stream: true });
        if (piece) parser.write(piece);
      }
    } finally {
      reader.releaseLock();
    }
  }
}

function resolveCell(type: string, text: string, sharedStrings: string[] | null): string | null {
  switch (type) {
    case 'inlineStr':
    case 'str':
      return text;
    case 's': {
      if (!sharedStrings) return null;
      const idx = parseInt(text, 10);
      return Number.isFinite(idx) && idx >= 0 && idx < sharedStrings.length
        ? sharedStrings[idx]
        : null;
    }
    case 'b':
      return text === '1' || text === 'true' ? 'TRUE' : 'FALSE';
    case 'e':
      return text ? `#${text}` : null;
    default:
      return text === '' ? null : text;
  }
}

/** Parse sharedStrings.xml from an inflated byte stream. Memory: unique strings only. */
export async function parseSharedStrings(stream: ReadableStream<Uint8Array>): Promise<string[]> {
  const parser = sax.parser(false, SAX_OPTIONS);

  const out: string[] = [];
  let current = '';
  let inSi = false;
  let inT = false;
  let fatal: Error | null = null;

  parser.onopentag = (node: sax.Tag) => {
    if (node.name === 'si') {
      inSi = true;
      current = '';
    } else if (node.name === 't' && inSi) {
      inT = true;
    }
  };
  parser.ontext = (t: string) => {
    if (inT) current += t;
  };
  parser.onclosetag = (name: string) => {
    if (name === 't') {
      inT = false;
    } else if (name === 'si') {
      out.push(current);
      inSi = false;
    }
  };
  parser.onerror = (e: Error) => {
    fatal = fatal ?? e;
  };

  for await (const piece of decodeUtf8(stream)) {
    parser.write(piece);
    if (fatal) throw fatal;
  }
  parser.end();
  if (fatal) throw fatal;
  return out;
}

/** Parse xl/workbook.xml -> ordered sheet definitions (name + relationship id). */
export async function parseWorkbook(
  stream: ReadableStream<Uint8Array>
): Promise<Array<{ name: string; rId: string }>> {
  const parser = sax.parser(false, SAX_OPTIONS);

  const sheets: Array<{ name: string; rId: string }> = [];
  let fatal: Error | null = null;
  parser.onopentag = (node: sax.Tag) => {
    if (node.name === 'sheet') {
      sheets.push({
        name: node.attributes['name'] ?? '',
        rId: node.attributes['r:id'] ?? '',
      });
    }
  };
  parser.onerror = (e: Error) => {
    fatal = fatal ?? e;
  };

  for await (const piece of decodeUtf8(stream)) {
    parser.write(piece);
    if (fatal) throw fatal;
  }
  parser.end();
  if (fatal) throw fatal;
  return sheets;
}

/** Parse an OOXML .rels part -> Map<rId, target>. */
export async function parseRels(stream: ReadableStream<Uint8Array>): Promise<Map<string, string>> {
  const parser = sax.parser(false, SAX_OPTIONS);

  const rels = new Map<string, string>();
  let fatal: Error | null = null;
  parser.onopentag = (node: sax.Tag) => {
    if (node.name === 'relationship') {
      const id = node.attributes['id'] ?? '';
      const target = node.attributes['target'] ?? '';
      if (id) rels.set(id, target);
    }
  };
  parser.onerror = (e: Error) => {
    fatal = fatal ?? e;
  };

  for await (const piece of decodeUtf8(stream)) {
    parser.write(piece);
    if (fatal) throw fatal;
  }
  parser.end();
  if (fatal) throw fatal;
  return rels;
}

/** Streaming UTF-8 decode of a byte stream into string chunks (multi-byte safe). */
export async function* decodeUtf8(stream: ReadableStream<Uint8Array>): AsyncGenerator<string> {
  const reader = stream.getReader();
  const decoder = new TextDecoder('utf-8');
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        const tail = decoder.decode();
        if (tail) yield tail;
        return;
      }
      const piece = decoder.decode(value, { stream: true });
      if (piece) yield piece;
    }
  } finally {
    reader.releaseLock();
  }
}
