/**
 * Streaming ZIP writer: emits a valid .zip archive byte-by-byte without
 * ever buffering a whole entry (or the whole archive) in memory.
 *
 * Each entry is written as:
 *   [local file header (sizes zeroed, general purpose bit 3 set)]
 *   [DEFLATE bytes, produced incrementally by pako]
 *   [16-byte data descriptor: signature + crc32 + compressed + uncompressed]
 * followed at the end by a central directory + EOCD.
 *
 * The data-descriptor layout is what makes true streaming possible: we
 * cannot know crc/sizes before the data exists, and we cannot rewind an
 * output stream we are appending to. The layout is part of the ZIP spec
 * (APPNOTE 4.4.8 / 4.4.13) and is accepted by Excel, LibreOffice,
 * 7-Zip, the JDK, JSZip, fflate and this library's own reader.
 */

import { Deflate } from 'pako';
import { CRC32_INIT, crc32Final, crc32Update } from './crc32.js';

const CHUNK = 64 * 1024;
const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_END = 0x06054b50;
const SIG_DATA_DESC = 0x08074b50;

const ZIP_METHOD_DEFLATE = 8;
const ZIP_FLAG_DATA_DESCRIPTOR = 0x0008;

// Fixed DOS time/date (2026-10-31 12:00) - jars/tools dislike zeros.
const DOS_TIME = (12 << 11) | (0 << 5) | 0;
const DOS_DATE = ((2026 - 1980) << 9) | (10 << 5) | 31;

/** Content of a zip entry: a string, a byte array, or an async stream of byte arrays. */
export type ZipEntryContent = string | Uint8Array | AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>;

export interface ZipWrittenEntry {
  name: string;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
}

interface CentralEntry extends ZipWrittenEntry {
  nameBytes: Uint8Array;
  headerOffset: number;
}

function u16(view: DataView, off: number, v: number): void {
  view.setUint16(off, v, true);
}
function u32(view: DataView, off: number, v: number): void {
  view.setUint32(off, v >>> 0, true);
}

async function* normalizeContent(content: ZipEntryContent): AsyncGenerator<Uint8Array> {
  if (typeof content === 'string') {
    yield new TextEncoder().encode(content);
    return;
  }
  if (content instanceof Uint8Array) {
    yield content;
    return;
  }
  if (content instanceof ReadableStream) {
    const reader = content.getReader();
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        if (value) yield value;
      }
    } finally {
      reader.releaseLock();
    }
    return;
  }
  yield* content as AsyncIterable<Uint8Array>;
}

export class ZipWriter {
  private readonly writer: WritableStreamDefaultWriter<Uint8Array>;
  private readonly entries: CentralEntry[] = [];
  private offset = 0;
  private closed = false;
  /** DEFLATE level (0-9) used for subsequent entries. Default 6. */
  level = 6;

  constructor(out: WritableStream<Uint8Array>) {
    this.writer = out.getWriter();
  }

  private async write(bytes: Uint8Array): Promise<void> {
    await this.writer.write(bytes);
    this.offset += bytes.length;
  }

  /** Append one entry, streaming its content through a raw DEFLATE encoder. */
  async addFile(name: string, content: ZipEntryContent): Promise<ZipWrittenEntry> {
    if (this.closed) throw new Error('ZipWriter already finished');
    const nameBytes = new TextEncoder().encode(name);

    const def = new Deflate({ raw: true, level: this.level, chunkSize: CHUNK });
    const pending: Uint8Array[] = [];
    def.onData = (chunk: Uint8Array) => {
      pending.push(chunk);
    };

    const headerOffset = this.offset;
    const local = new DataView(new ArrayBuffer(30));
    u32(local, 0, SIG_LOCAL);
    u16(local, 4, 20); // version needed
    u16(local, 6, ZIP_FLAG_DATA_DESCRIPTOR);
    u16(local, 8, ZIP_METHOD_DEFLATE);
    u16(local, 10, DOS_TIME);
    u16(local, 12, DOS_DATE);
    u32(local, 14, 0); // crc (in descriptor)
    u32(local, 18, 0); // compressed size (in descriptor)
    u32(local, 22, 0); // uncompressed size (in descriptor)
    u16(local, 26, nameBytes.length);
    u16(local, 28, 0); // extra length
    await this.write(new Uint8Array(local.buffer));
    await this.write(nameBytes);

    let crcState = CRC32_INIT;
    let uncompressed = 0;
    let compressed = 0;

    const flush = async (): Promise<void> => {
      for (const c of pending.splice(0, pending.length)) {
        compressed += c.length;
        await this.write(c); // awaits output-stream backpressure
      }
    };

    for await (const chunk of normalizeContent(content)) {
      crcState = crc32Update(crcState, chunk);
      uncompressed += chunk.length;
      def.push(chunk, false);
      await flush();
    }
    def.push(new Uint8Array(0), true);
    await flush();

    if (def.err !== 0) {
      throw new Error(`ZipWriter: deflate error ${def.err} (${def.msg}) for ${name}`);
    }

    const crc = crc32Final(crcState);
    const desc = new DataView(new ArrayBuffer(16));
    u32(desc, 0, SIG_DATA_DESC);
    u32(desc, 4, crc);
    u32(desc, 8, compressed);
    u32(desc, 12, uncompressed);
    await this.write(new Uint8Array(desc.buffer));

    const entry: CentralEntry = {
      name,
      nameBytes,
      crc,
      compressedSize: compressed,
      uncompressedSize: uncompressed,
      headerOffset,
    };
    this.entries.push(entry);
    return { name, crc, compressedSize: compressed, uncompressedSize: uncompressed };
  }

  /** Write the central directory + EOCD and close the output stream. */
  async finish(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    const cdStart = this.offset;
    let cdSize = 0;

    for (const e of this.entries) {
      const rec = new DataView(new ArrayBuffer(46 + e.nameBytes.length));
      u32(rec, 0, SIG_CENTRAL);
      u16(rec, 4, 20); // version made by
      u16(rec, 6, 20); // version needed
      u16(rec, 8, ZIP_FLAG_DATA_DESCRIPTOR);
      u16(rec, 10, ZIP_METHOD_DEFLATE);
      u16(rec, 12, DOS_TIME);
      u16(rec, 14, DOS_DATE);
      u32(rec, 16, e.crc);
      u32(rec, 20, e.compressedSize);
      u32(rec, 24, e.uncompressedSize);
      u16(rec, 28, e.nameBytes.length);
      u16(rec, 30, 0); // extra length
      u16(rec, 32, 0); // comment length
      u16(rec, 34, 0); // disk number start
      u16(rec, 36, 0); // internal attributes
      u32(rec, 38, 0); // external attributes
      u32(rec, 42, e.headerOffset);
      new Uint8Array(rec.buffer, 46).set(e.nameBytes);
      const bytes = new Uint8Array(rec.buffer);
      await this.write(bytes);
      cdSize += bytes.length;
    }

    const eocd = new DataView(new ArrayBuffer(22));
    u32(eocd, 0, SIG_END);
    u16(eocd, 4, 0); // disk number
    u16(eocd, 6, 0); // disk with central directory
    u16(eocd, 8, this.entries.length);
    u16(eocd, 10, this.entries.length);
    u32(eocd, 12, cdSize);
    u32(eocd, 16, cdStart);
    u16(eocd, 20, 0); // comment length
    await this.write(new Uint8Array(eocd.buffer));

    await this.writer.close();
  }
}
