// zip.js — minimal zip reader/writer, enough for .docx round-trips.
// Reading uses the browser's built-in DecompressionStream (Chrome 103+,
// Firefox 113+). Writing stores entries uncompressed (valid zip; Word is fine
// with it and docx files are small). No libraries.
"use strict";

const ZIP_CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    c = ZIP_CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream()
    .pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// ArrayBuffer -> Map(name -> Uint8Array)
async function unzip(buffer) {
  const b = new Uint8Array(buffer);
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);

  let eocd = -1; // End Of Central Directory: scan back past any zip comment
  for (let i = b.length - 22; i >= Math.max(0, b.length - 22 - 65535); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Not a valid .docx (zip) file");

  const count = dv.getUint16(eocd + 10, true);
  let off = dv.getUint32(eocd + 16, true);
  const td = new TextDecoder();
  const files = new Map();
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(off, true) !== 0x02014b50) throw new Error("Corrupt zip directory");
    const method = dv.getUint16(off + 10, true);
    const csize = dv.getUint32(off + 20, true);
    const nameLen = dv.getUint16(off + 28, true);
    const extraLen = dv.getUint16(off + 30, true);
    const commentLen = dv.getUint16(off + 32, true);
    const lho = dv.getUint32(off + 42, true);
    const name = td.decode(b.subarray(off + 46, off + 46 + nameLen));
    off += 46 + nameLen + extraLen + commentLen;

    const nl = dv.getUint16(lho + 26, true);
    const el = dv.getUint16(lho + 28, true);
    const dataStart = lho + 30 + nl + el;
    const raw = b.subarray(dataStart, dataStart + csize);
    let data;
    if (method === 0) data = raw.slice();
    else if (method === 8) data = await inflateRaw(raw);
    else throw new Error("Unsupported zip compression method: " + method);
    files.set(name, data);
  }
  return files;
}

// Map(name -> Uint8Array) -> Uint8Array (entries stored uncompressed)
function zipStore(files) {
  const te = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameB = te.encode(name);
    const crc = crc32(data);

    const lh = new Uint8Array(30 + nameB.length);
    const lv = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);          // version needed
    lv.setUint16(8, 0, true);           // method: store
    lv.setUint16(12, 0x21, true);       // date: 1980-01-01
    lv.setUint32(14, crc, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameB.length, true);
    lh.set(nameB, 30);
    localParts.push(lh, data);

    const ch = new Uint8Array(46 + nameB.length);
    const cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);          // version made by
    cv.setUint16(6, 20, true);          // version needed
    cv.setUint16(14, 0x21, true);       // date
    cv.setUint32(16, crc, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameB.length, true);
    cv.setUint32(42, offset, true);     // local header offset
    ch.set(nameB, 46);
    centralParts.push(ch);

    offset += lh.length + data.length;
  }
  const centralSize = centralParts.reduce((a, p) => a + p.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.size, true);
  ev.setUint16(10, files.size, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);

  const all = [...localParts, ...centralParts, eocd];
  const out = new Uint8Array(all.reduce((a, p) => a + p.length, 0));
  let pos = 0;
  for (const p of all) { out.set(p, pos); pos += p.length; }
  return out;
}

// Node (for tests)
if (typeof module !== "undefined") {
  module.exports = { crc32, unzip, zipStore };
}
