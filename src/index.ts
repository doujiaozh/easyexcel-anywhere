/**
 * EasyExcel-Anywhere
 *
 * Stream-based, OOM-free xlsx processing for Node.js and browsers:
 * pure TypeScript, no JVM, no native modules.
 *
 *   "用 Web Streams 降维打击 Java，让百万行表格处理跑在每一个浏览器里。"
 *   "Beating Java at its own game with Web Streams - million-row
 *    spreadsheet processing that runs in every browser."
 */

// Side-effect import: records the TS `design:type` metadata emitted for
// decorated classes so @ExcelProperty can infer string/number/boolean/date
// without an explicit `type` option. Safe on Node >= 18 and in browsers.
import 'reflect-metadata';

export { ExcelIgnore, ExcelProperty } from './annotations/ExcelProperty.js';
export type {
  CellType,
  ExcelPropertyOptions,
  PropertyMeta,
  ValueConverter,
} from './annotations/ExcelProperty.js';

export { ExcelReader } from './read/reader.js';
export type { ReadOptions, RawRow } from './read/reader.js';

export { ExcelWriter } from './write/writer.js';
export type { WriteOptions } from './write/writer.js';

export type { ByteSource } from './core/bytes.js';
export { ZipReader, type ZipEntry, type ZipEntryInfo } from './core/zip-reader.js';
export { ZipWriter, type ZipEntryContent, type ZipWrittenEntry } from './core/zip-writer.js';
export { SheetXmlReader } from './core/xml-reader.js';

import { ExcelReader } from './read/reader.js';
import { ExcelWriter } from './write/writer.js';
import type { ReadOptions, RawRow } from './read/reader.js';
import type { WriteOptions } from './write/writer.js';
import type { ByteSource } from './core/bytes.js';

/**
 * Facade entry point, mirroring the feel of alibaba/easyexcel.
 *
 * @example
 * ```ts
 * // read: stream a 1M-row file row by row
 * for await (const user of EasyExcel.read(file, { class: User })) {
 *   // memory stays flat: the file is only pulled while you pull rows
 * }
 *
 * // write: rows flow straight into the .xlsx byte stream
 * const w = EasyExcel.write('/tmp/out.xlsx', { class: User });
 * await w.writeRows(users);
 * await w.close();
 * ```
 */
export class EasyExcel {
  static read<T = RawRow>(source: ByteSource, options?: ReadOptions<T>): ExcelReader<T> {
    return new ExcelReader<T>(source, options);
  }

  static write(
    target: string | WritableStream<Uint8Array>,
    options?: WriteOptions
  ): ExcelWriter {
    return new ExcelWriter(target, options);
  }
}
