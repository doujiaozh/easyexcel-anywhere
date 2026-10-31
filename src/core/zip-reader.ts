/**
 * Streaming ZIP (RFC 1950/ZIP "stored-deflate") reader for .xlsx containers.
 *
 * The reader walks local file headers sequentially from a pure byte stream
 * (no random access, no full-file buffering). Two entry layouts are handled:
 *
 *  1. classic entries: local header carries crc/sizes -> the entry's
 *     `compressedSize` bytes are sliced and inflated.
 *  2. stream-written entries (data descriptor, general purpose bit 3): the
 *     compressed payload is fed incrementally to the inflater; when the
 *     DEFLATE stream terminates, `strm.next_in` tells us exactly how many
 *     bytes of the last pushed chunk were consumed, so the trailing
 *     16-byte descriptor can be parsed and leftover bytes pushed back.
 *
 * This makes archives produced by our own streaming writer, by Excel and
 * by every other mainstream zip producer readable.
 */

import { Inflate } from 'pako';
import { StreamReader, toWebStream, type ByteSource } from './bytes.js';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_END = 0x06054b50;
const SIG_DATA_DESC = 0x08074b50;

const CHUNK = 64 * 1024;

export interface ZipEntryInfo {
  name: string;
  method: number;
  flag: number;
  /** crc32; filled in only when known up-front or after a descriptor was parsed */
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  /** true when the entry stream ends with a 16-byte data descriptor */
  dataDescriptor: boolean;
}

export interface ZipEntry extends ZipEntryInfo {
  /** Open the (inflated) content of this entry as a Web ReadableStream. */
  open(): ReadableStream<Uint8Array>;
  /** Discard this entry's bytes. Must be called for entries you don't open. */
  skip(): Promise<void>;
}

/** Turn an async generator into a pull-based Web ReadableStream. */
function asyncGenToStream(gen: AsyncGenerator<Uint8Array>): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      const { value, done } = await gen.next();
      if (done) {
        controller.close();
        return;
      }
      controller.enqueue(value);
    },
    cancel() {
      gen.return?.(undefined);
    },
  });
}

interface PakoInflateInternal {
  strm: { next_in: number };
}

export class ZipReader {
  private readonly br: StreamReader;
  private consumed = true;

  constructor(source: ByteSource) {
    this.br = new StreamReader(toWebStream(source));
  }

  private markConsumed(): void {
    this.consumed = true;
  }

  /**
   * Yield the next local-file entry, or null when the end of the local
   * entry area is reached (central directory / EOCD / plain EOF).
   * The previous entry must have been fully consumed or skipped first.
   */
  async next(): Promise<ZipEntry | null> {
    if (!this.consumed) {
      throw new Error('ZipReader: previous entry must be fully consumed or skipped first');
    }
    const sig = await this.readU32();
    if (sig === null) return null;
    if (sig === SIG_END || sig === SIG_CENTRAL) return null;
    if (sig !== SIG_LOCAL) {
      throw new Error(`ZipReader: invalid local header signature 0x${sig.toString(16)}`);
    }
    if (!(await this.br.need(26))) {
      throw new Error('ZipReader: truncated local file header');
    }
    const hdr = this.br.read(26)!;
    const dv = new DataView(hdr.buffer, hdr.byteOffset, hdr.byteLength);
    const flag = dv.getUint16(2, true);
    const method = dv.getUint16(4, true);
    const crc = dv.getUint32(10, true);
    const compressedSize = dv.getUint32(14, true);
    const uncompressedSize = dv.getUint32(18, true);
    const nameLen = dv.getUint16(22, true);
    const extraLen = dv.getUint16(24, true);
    if (!(await this.br.need(nameLen))) {
      throw new Error('ZipReader: truncated entry name');
    }
    const nameBytes = this.br.read(nameLen)!;
    const name = new TextDecoder('utf-8').decode(nameBytes);
    await this.br.skip(extraLen);

    const dataDescriptor = (flag & 0x0008) !== 0 && compressedSize === 0;
    if (method !== 0 && method !== 8) {
      throw new Error(`ZipReader: unsupported compression method ${method} for ${name}`);
    }
    if (method === 0 && dataDescriptor) {
      throw new Error(`ZipReader: stored entry with data descriptor is not supported: ${name}`);
    }

    this.consumed = false;
    return {
      name,
      method,
      flag,
      crc,
      compressedSize,
      uncompressedSize,
      dataDescriptor,
      open: () => asyncGenToStream(this.iterEntry(name, method, compressedSize, dataDescriptor)),
      skip: () => this.skipEntry(name, method, compressedSize, dataDescriptor),
    };
  }

  private async readU32(): Promise<number | null> {
    if (!(await this.br.need(4))) return null;
    const b = this.br.read(4)!;
    return ((b[0] | (b[1] << 8) | (b[2] << 16) | (b[3] * 0x1000000)) >>> 0);
  }

  private async skipEntry(
    name: string,
    method: number,
    compressedSize: number,
    dataDescriptor: boolean
  ): Promise<void> {
    try {
      if (!dataDescriptor) {
        await this.br.skip(compressedSize);
        return;
      }
      // descriptor mode: drain (and discard) the deflated stream
      for await (const _ of this.iterEntry(name, method, compressedSize, dataDescriptor)) {
        void _;
      }
    } finally {
      this.markConsumed();
    }
  }

  private async *iterEntry(
    name: string,
    method: number,
    compressedSize: number,
    dataDescriptor: boolean
  ): AsyncGenerator<Uint8Array> {
    try {
      if (method === 0) {
        let rem = compressedSize;
        while (rem > 0) {
          const n = Math.min(rem, CHUNK);
          if (!(await this.br.need(n))) throw new Error(`ZipReader: truncated entry ${name}`);
          const piece = this.br.read(n)!;
          rem -= piece.length;
          yield piece;
        }
        return;
      }

      if (!dataDescriptor && compressedSize === 0) {
        return; // empty entry
      }

      const infl = new Inflate({ raw: true, chunkSize: CHUNK });
      const queue: Uint8Array[] = [];
      infl.onData = (chunk: Uint8Array) => {
        queue.push(chunk);
      };

      if (!dataDescriptor) {
        let rem = compressedSize;
        let finished = false;
        while (!(finished && queue.length === 0)) {
          if (queue.length > 0) {
            yield queue.shift()!;
            continue;
          }
          if (finished) break;
          if (rem <= 0) throw new Error(`ZipReader: deflate stream not terminated in ${name}`);
          const n = Math.min(rem, CHUNK);
          if (!(await this.br.need(n))) throw new Error(`ZipReader: truncated entry ${name}`);
          const piece = this.br.read(n)!;
          rem -= piece.length;
          infl.push(piece, rem === 0);
          if (infl.err !== 0 && !infl.ended) {
            throw new Error(`ZipReader: inflate error ${infl.err} (${infl.msg}) in ${name}`);
          }
          if (infl.ended) {
            // safety net for malformed archives: give back unconsumed bytes
            const consumed = (infl as unknown as PakoInflateInternal).strm.next_in;
            if (consumed < piece.length) this.br.unread(piece.subarray(consumed));
            finished = true;
          }
        }
        if (!infl.ended) throw new Error(`ZipReader: truncated deflate stream in ${name}`);
        return;
      }

      // ---- data descriptor mode ----
      let lastPiece: Uint8Array | null = null;
      while (!infl.ended) {
        while (queue.length > 0) yield queue.shift()!;
        if (!(await this.br.need(1))) {
          throw new Error(`ZipReader: unexpected EOF while inflating ${name}`);
        }
        const piece = this.br.read(Math.min(CHUNK, this.br.bufferedLength))!;
        lastPiece = piece;
        infl.push(piece, false);
        if (infl.err !== 0 && infl.ended) {
          throw new Error(`ZipReader: inflate error ${infl.err} (${infl.msg}) in ${name}`);
        }
      }
      // the terminating push fires its final onData chunk(s) *after* the loop
      // condition has already flipped; drain whatever is still queued
      while (queue.length > 0) yield queue.shift()!;
      const consumed = lastPiece
        ? (infl as unknown as PakoInflateInternal).strm.next_in
        : 0;
      let leftover = lastPiece ? lastPiece.subarray(consumed) : new Uint8Array(0);
      if (leftover.length < 16) {
        const needMore = 16 - leftover.length;
        if (!(await this.br.need(needMore))) {
          throw new Error(`ZipReader: truncated data descriptor in ${name}`);
        }
        const more = this.br.read(needMore)!;
        const merged = new Uint8Array(leftover.length + more.length);
        merged.set(leftover);
        merged.set(more, leftover.length);
        leftover = merged;
      }
      const dv = new DataView(leftover.buffer, leftover.byteOffset, 16);
      const sig = dv.getUint32(0, true);
      if (sig !== SIG_DATA_DESC) {
        throw new Error(`ZipReader: invalid data descriptor signature 0x${sig.toString(16)} in ${name}`);
      }
      // sizes stay informational on the returned entry
      if (leftover.length > 16) this.br.unread(leftover.subarray(16));
    } finally {
      this.markConsumed();
    }
  }
}
