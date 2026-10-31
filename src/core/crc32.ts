/**
 * Pure-JS CRC-32 (IEEE 802.3), table driven.
 * Streaming-friendly: update over chunks, finalize once.
 */

const TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    t[n] = c >>> 0;
  }
  return t;
})();

/** Initial raw state for a new CRC computation. */
export const CRC32_INIT = 0xffffffff;

/** Feed bytes into the raw CRC state. */
export function crc32Update(state: number, data: Uint8Array): number {
  let c = state;
  for (let i = 0; i < data.length; i++) {
    c = TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  }
  return c >>> 0;
}

/** Convert the raw state to the final ZIP CRC-32 checksum. */
export function crc32Final(state: number): number {
  return (state ^ 0xffffffff) >>> 0;
}

/** One-shot CRC-32. */
export function crc32(data: Uint8Array): number {
  return crc32Final(crc32Update(CRC32_INIT, data));
}
