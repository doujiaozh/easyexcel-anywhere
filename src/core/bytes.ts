/**
 * Universal byte-source normalization + a buffered *sequential* reader
 * built on top of the Web Streams API.
 *
 * The reader pulls chunks on demand only (demand-driven backpressure),
 * so memory stays bounded by the current chunk no matter how large the
 * input is. It runs unchanged on Node.js (>= 18.14) and in browsers.
 */

export type ByteSource =
  | ReadableStream<Uint8Array>
  | AsyncIterable<Uint8Array>
  | Iterable<Uint8Array>
  | Uint8Array
  | ArrayBuffer
  | Blob
  | { stream?: () => ReadableStream<Uint8Array> };

/** Normalize any supported byte source into a Web ReadableStream<Uint8Array>. */
export function toWebStream(source: ByteSource): ReadableStream<Uint8Array> {
  if (source instanceof ReadableStream) {
    return source;
  }
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) {
    const u = source instanceof Uint8Array ? source : new Uint8Array(source);
    return new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(u);
        controller.close();
      },
    });
  }
  const blobLike = source as { stream?: () => ReadableStream<Uint8Array> };
  if (typeof blobLike.stream === 'function') {
    return blobLike.stream();
  }
  const any = source as unknown as object;
  if (Symbol.asyncIterator in any) {
    const it = (source as AsyncIterable<Uint8Array>)[Symbol.asyncIterator]();
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        const { value, done } = await it.next();
        if (done) {
          controller.close();
          return;
        }
        if (value) controller.enqueue(value);
      },
      cancel() {
        it.return?.();
      },
    });
  }
  if (Symbol.iterator in any) {
    return new ReadableStream<Uint8Array>({
      start(controller) {
        for (const c of source as Iterable<Uint8Array>) controller.enqueue(c);
        controller.close();
      },
    });
  }
  throw new TypeError('Unsupported byte source for EasyExcel-Anywhere');
}

/**
 * Buffered sequential reader over a Web ReadableStream.
 * All methods are async and pull-based: bytes are only read from the
 * underlying stream when actually needed.
 */
export class StreamReader {
  private readonly reader: ReadableStreamDefaultReader<Uint8Array>;
  private chunks: Uint8Array[] = [];
  private len = 0;
  private eof = false;

  constructor(stream: ReadableStream<Uint8Array>) {
    this.reader = stream.getReader();
  }

  get bufferedLength(): number {
    return this.len;
  }

  get atEof(): boolean {
    return this.eof;
  }

  private async fill(): Promise<boolean> {
    if (this.eof) return false;
    const { value, done } = await this.reader.read();
    if (done) {
      this.eof = true;
      return false;
    }
    if (value && value.length > 0) {
      this.chunks.push(value);
      this.len += value.length;
    }
    return true;
  }

  /** Ensure at least n bytes are buffered (stop at EOF). True when n bytes are available. */
  async need(n: number): Promise<boolean> {
    while (this.len < n) {
      if (!(await this.fill())) break;
    }
    return this.len >= n;
  }

  /** Consume exactly n bytes. Call `need(n)` first. Returns null if fewer bytes remain. */
  read(n: number): Uint8Array | null {
    if (n <= 0) return new Uint8Array(0);
    if (this.len < n) return null;
    const out = new Uint8Array(n);
    let off = 0;
    while (off < n) {
      const c = this.chunks[0];
      const take = Math.min(c.length, n - off);
      out.set(take === c.length ? c : c.subarray(0, take), off);
      this.len -= take;
      off += take;
      if (take === c.length) this.chunks.shift();
      else this.chunks[0] = c.subarray(take);
    }
    return out;
  }

  /** Push bytes back to the front of the buffer (un-read). */
  unread(data: Uint8Array): void {
    if (data.length > 0) {
      this.chunks.unshift(data);
      this.len += data.length;
    }
  }

  /** Skip n bytes without materializing them (avoids large memcopies). */
  async skip(n: number): Promise<void> {
    if (n <= 0) return;
    let rem = n;
    while (rem > 0 && this.chunks.length > 0) {
      const c = this.chunks[0];
      if (c.length <= rem) {
        this.chunks.shift();
        this.len -= c.length;
        rem -= c.length;
      } else {
        this.chunks[0] = c.subarray(rem);
        this.len -= rem;
        rem = 0;
      }
    }
    while (rem > 0 && !this.eof) {
      const { value, done } = await this.reader.read();
      if (done) {
        this.eof = true;
        break;
      }
      if (value.length > rem) {
        this.unread(value.subarray(rem));
        rem = 0;
      } else {
        rem -= value.length;
      }
    }
  }

  /** Read everything that remains. Only for bounded/small payloads. */
  async drain(): Promise<Uint8Array> {
    while (await this.fill()) {
      /* keep pulling until EOF */
    }
    const out = new Uint8Array(this.len);
    let off = 0;
    for (const c of this.chunks) {
      out.set(c, off);
      off += c.length;
    }
    this.chunks = [];
    this.len = 0;
    return out;
  }
}
