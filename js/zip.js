/**
 * Minimal ZIP reader/writer using browser CompressionStream / DecompressionStream.
 * Supports store (0) and deflate (8) — enough for .xlsx (OOXML).
 */

function u16(view, offset) {
  return view.getUint16(offset, true);
}

function u32(view, offset) {
  return view.getUint32(offset, true);
}

function utf8Decode(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

function utf8Encode(str) {
  return new TextEncoder().encode(str);
}

async function inflateRaw(data) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("This browser cannot decompress ZIP/XLSX (needs DecompressionStream)");
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function deflateRaw(data) {
  if (typeof CompressionStream === "undefined") {
    // Fall back to store (no compression)
    return { method: 0, bytes: data };
  }
  const stream = new Blob([data]).stream().pipeThrough(new CompressionStream("deflate-raw"));
  const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
  return { method: 8, bytes };
}

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let k = 0; k < 8; k++) c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
  }
  return ~c >>> 0;
}

/**
 * @param {ArrayBuffer|Uint8Array} buffer
 * @returns {Promise<Map<string, Uint8Array>>}
 */
export async function unzip(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const files = new Map();

  // Find End of Central Directory
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 0xffff - 22); i--) {
    if (u32(view, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("Invalid ZIP: missing central directory");

  const count = u16(view, eocd + 10);
  let offset = u32(view, eocd + 16);

  for (let n = 0; n < count; n++) {
    if (u32(view, offset) !== 0x02014b50) throw new Error("Invalid ZIP central directory");
    const method = u16(view, offset + 10);
    const compSize = u32(view, offset + 20);
    const nameLen = u16(view, offset + 28);
    const extraLen = u16(view, offset + 30);
    const commentLen = u16(view, offset + 32);
    const localOffset = u32(view, offset + 42);
    const name = utf8Decode(bytes.subarray(offset + 46, offset + 46 + nameLen));
    offset += 46 + nameLen + extraLen + commentLen;

    if (name.endsWith("/")) continue;

    if (u32(view, localOffset) !== 0x04034b50) throw new Error(`Invalid local header for ${name}`);
    const localNameLen = u16(view, localOffset + 26);
    const localExtraLen = u16(view, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const compressed = bytes.subarray(dataStart, dataStart + compSize);

    let raw;
    if (method === 0) raw = compressed.slice();
    else if (method === 8) raw = await inflateRaw(compressed);
    else throw new Error(`Unsupported ZIP compression ${method} in ${name}`);

    files.set(name, raw);
  }

  return files;
}

/**
 * @param {Record<string, string|Uint8Array>} entries
 * @returns {Promise<Uint8Array>}
 */
export async function zip(entries) {
  const locals = [];
  const central = [];
  let offset = 0;

  for (const [name, content] of Object.entries(entries)) {
    const data = typeof content === "string" ? utf8Encode(content) : content;
    const nameBytes = utf8Encode(name);
    const { method, bytes: compressed } = await deflateRaw(data);
    const crc = crc32(data);

    const local = new Uint8Array(30 + nameBytes.length + compressed.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0, true); // flags
    lv.setUint16(8, method, true);
    lv.setUint16(10, 0, true); // time
    lv.setUint16(12, 0, true); // date
    lv.setUint32(14, crc, true);
    lv.setUint32(18, compressed.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    local.set(nameBytes, 30);
    local.set(compressed, 30 + nameBytes.length);

    const cen = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(cen.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0, true);
    cv.setUint16(10, method, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, compressed.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    cen.set(nameBytes, 46);

    locals.push(local);
    central.push(cen);
    offset += local.length;
  }

  const centralSize = central.reduce((n, b) => n + b.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, locals.length, true);
  ev.setUint16(10, locals.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);

  const total = offset + centralSize + 22;
  const out = new Uint8Array(total);
  let p = 0;
  for (const b of locals) {
    out.set(b, p);
    p += b.length;
  }
  for (const b of central) {
    out.set(b, p);
    p += b.length;
  }
  out.set(eocd, p);
  return out;
}
