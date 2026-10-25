import { Inflate, zipSync } from 'fflate';

export const LIMITS = Object.freeze({
  fileBytes: 100 * 1024 * 1024,
  expandedBytes: 200 * 1024 * 1024,
  entryBytes: 100 * 1024 * 1024,
  xmlPartBytes: 12 * 1024 * 1024,
  xmlTotalBytes: 40 * 1024 * 1024,
  entries: 5000,
  expansionRatio: 300,
  deflateNoProgressBytes: 128 * 1024,
});

const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  crcTable[n] = c >>> 0;
}
export function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function safePath(name) {
  const candidates = [name];
  try { candidates.push(decodeURIComponent(name)); } catch { /* literal % is legal in ZIP */ }
  return candidates.every(p => p && !p.startsWith('/') && !/[\\\x00-\x1f:]/.test(p)
    && !p.split('/').some(s => s === '.' || s === '..'));
}

function inflateBounded(compressed, expectedSize) {
  if (!compressed.length) throw new Error('缺少 DEFLATE 压缩流。');
  const output = new Uint8Array(expectedSize);
  let written = 0;
  let finalized = false;
  let noProgressBytes = 0;
  const inflater = new Inflate((chunk, final) => {
    // Stop synchronously before accepting any excess output or feeding the next
    // compressed chunk. A false small ZIP size must not make us decode the whole
    // stream first. The decoder still needs a bounded transient chunk buffer.
    if (chunk.length > expectedSize - written) throw new Error('解压内容超过声明长度。');
    output.set(chunk, written);
    written += chunk.length;
    finalized = final;
  });
  const chunkBytes = 4 * 1024;
  for (let offset = 0; offset < compressed.length; offset += chunkBytes) {
    const end = Math.min(offset + chunkBytes, compressed.length);
    const before = written;
    inflater.push(compressed.subarray(offset, end), end === compressed.length);
    noProgressBytes = written === before ? noProgressBytes + end - offset : 0;
    // Public-API guard against a decoder retaining and repeatedly concatenating
    // trailing input after BFINAL, or excessive empty blocks. 128 KiB exceeds a
    // DEFLATE stored block's 65,535-byte payload plus its header.
    if (noProgressBytes > LIMITS.deflateNoProgressBytes) throw new Error('压缩流长时间没有输出，已达到安全处理上限。');
  }
  if (!finalized || written !== expectedSize) throw new Error('解压长度与声明不一致。');
  return output;
}

export function readArchive(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes.length > LIMITS.fileBytes) throw new Error('文件超过 100 MB。请先减小嵌入媒体或拆分文件。');
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new Error('这是旧版 Office 或加密文件。请在 Office 中另存为未加密的 PPTX / XLSX 副本。');
  if (bytes.length < 22) throw new Error('文件不是完整的 PPTX / XLSX 文档。');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = n => view.getUint16(n, true);
  const u32 = n => view.getUint32(n, true);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (u32(i) === 0x06054b50 && i + 22 + u16(i + 20) === bytes.length) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('压缩包目录不完整，无法安全检查。');
  const count = u16(eocd + 10), size = u32(eocd + 12), offset = u32(eocd + 16);
  if (u16(eocd + 4) || u16(eocd + 6) || u16(eocd + 8) !== count) throw new Error('不支持分卷压缩文档。');
  if (count === 65535 || offset === 0xffffffff || size === 0xffffffff) throw new Error('当前不支持 ZIP64 大型文档。');
  if (!count || count > LIMITS.entries || offset + size !== eocd) throw new Error('压缩包目录异常或内部文件数量过多。');
  const metadata = new Map(), canonicalNames = new Set(), ranges = [];
  let cursor = offset, total = 0, xmlTotal = 0;
  for (let i = 0; i < count; i++) {
    if (cursor + 46 > eocd || u32(cursor) !== 0x02014b50) throw new Error('压缩包目录损坏。');
    const flags = u16(cursor + 8), method = u16(cursor + 10), crc = u32(cursor + 16);
    const compressed = u32(cursor + 20), expanded = u32(cursor + 24);
    const nameLength = u16(cursor + 28), extraLength = u16(cursor + 30), commentLength = u16(cursor + 32);
    const local = u32(cursor + 42), next = cursor + 46 + nameLength + extraLength + commentLength;
    if (next > eocd || u16(cursor + 34)) throw new Error('压缩包条目损坏。');
    if (flags & 0x41) throw new Error('不支持加密的文档。请使用未加密副本。');
    if (![0, 8].includes(method)) throw new Error('文档使用了不支持的压缩方式。');
    let name;
    try { name = new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength)); }
    catch { throw new Error('内部文件名称编码无法安全识别。'); }
    if (!safePath(name)) throw new Error('文档包含不安全的内部路径，已停止处理。');
    let canonical = name;
    try { canonical = decodeURIComponent(name); } catch { /* preserve literal percent */ }
    canonical = canonical.replace(/[A-Z]/g, c => c.toLowerCase());
    if (canonicalNames.has(canonical)) throw new Error('文档包含重复或大小写冲突的内部文件，无法确定应使用哪一份。');
    canonicalNames.add(canonical);
    if (expanded === 0xffffffff || compressed === 0xffffffff || expanded > LIMITS.entryBytes) throw new Error('内部文件过大或使用 ZIP64。');
    total += expanded;
    if (/\.(xml|rels)$/i.test(name) || name === '[Content_Types].xml') {
      xmlTotal += expanded;
      if (expanded > LIMITS.xmlPartBytes) throw new Error('单个 XML 部件超过 12 MB 的安全处理上限。');
    }
    if (total > LIMITS.expandedBytes || xmlTotal > LIMITS.xmlTotalBytes
      || (expanded > 1024 * 1024 && expanded / Math.max(1, compressed) > LIMITS.expansionRatio)) {
      throw new Error('文档解压后过大或压缩比例异常，已停止处理以保护浏览器。');
    }
    if (local + 30 > offset || u32(local) !== 0x04034b50) throw new Error('压缩包文件头损坏。');
    const localNameLen = u16(local + 26), localExtraLen = u16(local + 28);
    const dataStart = local + 30 + localNameLen + localExtraLen;
    if (dataStart + compressed > offset || u16(local + 6) !== flags || u16(local + 8) !== method) throw new Error('压缩包文件头与目录不一致。');
    const localName = bytes.subarray(local + 30, local + 30 + localNameLen);
    const centralName = bytes.subarray(cursor + 46, cursor + 46 + nameLength);
    if (localName.length !== centralName.length || localName.some((b, j) => b !== centralName[j])) throw new Error('压缩包内外文件名称不一致。');
    if (!(flags & 8) && (u32(local + 14) !== crc || u32(local + 18) !== compressed || u32(local + 22) !== expanded)) throw new Error('压缩包大小或校验信息不一致。');
    ranges.push([local, dataStart + compressed]);
    if (method === 0 && expanded !== compressed) throw new Error('未压缩部件的大小信息不一致。');
    metadata.set(name, { compressed, expanded, crc, method, dataStart });
    cursor = next;
  }
  if (cursor !== eocd) throw new Error('压缩包目录长度不一致。');
  ranges.sort((a, b) => a[0] - b[0]);
  for (let i = 1; i < ranges.length; i++) if (ranges[i][0] < ranges[i - 1][1]) throw new Error('内部压缩数据相互重叠，已停止处理。');
  const entries = new Map();
  for (const [name, meta] of metadata) {
    // Decompress from the directory WE validated. A second ZIP parser might
    // choose a forged EOCD inside the archive comment and bypass size limits.
    const compressedData = bytes.subarray(meta.dataStart, meta.dataStart + meta.compressed);
    let data;
    try {
      data = meta.method === 0 ? compressedData.slice()
        : inflateBounded(compressedData, meta.expanded);
    } catch { throw new Error(`文档压缩数据损坏或解压超限：${name}。`); }
    if (!(data instanceof Uint8Array) || data.length !== meta.expanded || crc32(data) !== meta.crc) throw new Error(`内部文件 ${name} 校验失败，原文档可能已损坏。`);
    entries.set(name, data);
  }
  return { bytes, entries, metadata, expandedBytes: total };
}

export function writeArchive(entries) {
  const files = Object.create(null);
  for (const [name, data] of entries) {
    files[name] = [data, { level: /\.(png|jpg|jpeg|mp4|m4v|mov|mp3|zip)$/i.test(name) ? 0 : 6,
      mtime: new Date('2020-01-01T00:00:00Z') }];
  }
  return zipSync(files);
}
