/**
 * Minimal, dependency-free pixel-dimension reader for PNG/JPEG buffers.
 * Used to compute a crop-to-cover srcRect so replacement images fill their
 * existing PPT box without distorting — no image-processing package needed.
 * Returns null for unsupported formats or malformed buffers instead of throwing.
 */
function getImageDimensions(buffer, ext) {
  try {
    if (!buffer || buffer.length < 24) return null;
    const e = String(ext || '').toLowerCase();

    if (e === 'png' || (buffer[0] === 0x89 && buffer[1] === 0x50)) {
      return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
    }

    if (e === 'jpg' || e === 'jpeg' || (buffer[0] === 0xff && buffer[1] === 0xd8)) {
      let offset = 2;
      while (offset + 4 <= buffer.length) {
        if (buffer[offset] !== 0xff) {
          offset += 1;
          continue;
        }
        const marker = buffer[offset + 1];
        if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
          offset += 2;
          continue;
        }
        if (marker === 0xd9 || marker === 0xda) break; // EOI / start of scan
        const segLength = buffer.readUInt16BE(offset + 2);
        const isSOF = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
        if (isSOF) {
          return { height: buffer.readUInt16BE(offset + 5), width: buffer.readUInt16BE(offset + 7) };
        }
        offset += 2 + segLength;
      }
      return null;
    }

    return null;
  } catch {
    return null;
  }
}

module.exports = { getImageDimensions };
