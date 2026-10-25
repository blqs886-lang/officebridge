import { DOMParser } from '@xmldom/xmldom';

export const NS = Object.freeze({
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
});

export function elements(node, namespace, localName) {
  return node ? Array.from(node.getElementsByTagNameNS(namespace, localName)) : [];
}

export function attribute(node, name, fallback = '') {
  return node?.hasAttribute(name) ? node.getAttribute(name) : fallback;
}

export function parseXml(text, part = 'XML') {
  if (/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/.test(text)
    || /[\uD800-\uDFFF]/u.test(text)) {
    throw new Error(`${part} 含有 XML 不允许的字符，已停止处理。`);
  }
  if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(text)) throw new Error(`${part} 含有不支持的 XML 实体或 DTD。`);
  // OOXML requires XML, never HTML's forgiving recovery. Treat parser warnings as
  // invalid too: xmldom can otherwise repair malformed input silently.
  const problems = [];
  let doc;
  try {
    doc = new DOMParser({ onError: (level, message) => {
      // U+FFFD is a valid XML character. Preserve it for the scanner to report;
      // rejecting it would prevent diagnosing an already garbled document.
      if (level === 'warning' && String(message).includes('Unicode replacement character detected')) return;
      problems.push(`${level}: ${message}`);
    } })
      .parseFromString(text, 'application/xml');
  } catch {
    throw new Error(`${part} 的 XML 结构损坏，已停止处理。`);
  }
  if (problems.length || !doc?.documentElement || doc.doctype) {
    throw new Error(`${part} 的 XML 结构不完整或不合法，已停止处理。`);
  }
  return doc;
}

export function decodeXml(bytes, part) {
  let encoding = 'utf-8';
  let bom = false;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) { encoding = 'utf-16le'; bom = true; }
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) { encoding = 'utf-16be'; bom = true; }
  else if (bytes[0] === 0x3c && bytes[1] === 0) encoding = 'utf-16le';
  else if (bytes[0] === 0 && bytes[1] === 0x3c) encoding = 'utf-16be';
  else bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  let text;
  try { text = new TextDecoder(encoding, { fatal: true }).decode(bytes); }
  catch { throw new Error(`${part} 的文本编码无法安全读取。`); }
  const declared = /^\s*<\?xml\b[^?]*\bencoding\s*=\s*["']([^"']+)/i.exec(text)?.[1]?.toLowerCase();
  if (declared && !['utf-8', 'utf-16', 'utf-16le', 'utf-16be'].includes(declared)) {
    throw new Error(`${part} 使用 ${declared} 编码；当前仅支持 UTF-8 / UTF-16。`);
  }
  if (declared === 'utf-8' && encoding !== 'utf-8') throw new Error(`${part} 的编码声明与内容不一致。`);
  if (declared?.startsWith('utf-16') && encoding === 'utf-8') throw new Error(`${part} 的编码声明与内容不一致。`);
  if (['utf-16le', 'utf-16be'].includes(declared) && declared !== encoding) throw new Error(`${part} 的 UTF-16 字节顺序与编码声明不一致。`);
  return { text, encoding, bom };
}

export function encodeXml(text, { encoding, bom }) {
  if (encoding === 'utf-8') {
    const bytes = new TextEncoder().encode(text);
    if (!bom) return bytes;
    const out = new Uint8Array(bytes.length + 3);
    out.set([0xef, 0xbb, 0xbf]); out.set(bytes, 3); return out;
  }
  const offset = bom ? 2 : 0;
  const bytes = new Uint8Array(text.length * 2 + offset);
  const view = new DataView(bytes.buffer);
  const little = encoding === 'utf-16le';
  if (bom) view.setUint16(0, 0xfeff, little);
  for (let i = 0; i < text.length; i++) view.setUint16(offset + i * 2, text.charCodeAt(i), little);
  return bytes;
}

export function escapeAttribute(value, quote = '"') {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll(quote, quote === '"' ? '&quot;' : '&apos;');
}

// Token offsets and DOM namespaces together allow changing ONLY the chosen
// attribute values. Unknown namespaces, whitespace, comments, CDATA, Alternate-
// Content, formulas and text stay byte-for-byte intact after re-encoding.
export function replaceFontAttributes(text, doc, fixes) {
  const nodes = Array.from(doc.getElementsByTagName('*'));
  const tokens = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/[\s\S]*?>|<([A-Za-z_][\w.:-]*)(?:[^>"']|"[^"]*"|'[^']*')*>/g;
  const replacements = [];
  const counts = new Map();
  let index = 0;
  for (const match of text.matchAll(tokens)) {
    if (!match[1]) continue;
    const node = nodes[index++];
    if (!node || node.nodeName !== match[1]) throw new Error('XML 位置核对失败，未写入修复。');
    for (const fix of fixes) {
      if (node.namespaceURI !== fix.scope) continue;
      let key = null;
      if (node.namespaceURI === NS.a || node.namespaceURI === 'http://purl.oclc.org/ooxml/drawingml/main') {
        if (node.hasAttribute('typeface')) key = 'typeface';
      } else if (node.namespaceURI === NS.s && (node.localName === 'rFont'
        || (node.localName === 'name' && node.parentNode?.namespaceURI === NS.s && node.parentNode.localName === 'font'))) key = 'val';
      if (!key || node.getAttribute(key) !== fix.from) continue;
      const attrPattern = /\s([A-Za-z_][\w.:-]*)\s*=\s*(["'])([\s\S]*?)\2/g;
      for (const attr of match[0].matchAll(attrPattern)) {
        if (attr[1] !== key) continue;
        const quoteIndex = attr[0].indexOf(attr[2]);
        const start = match.index + attr.index + quoteIndex + 1;
        replacements.push({ start, end: start + attr[3].length, value: escapeAttribute(fix.to, attr[2]) });
        counts.set(fix.issueId, (counts.get(fix.issueId) || 0) + 1);
      }
    }
  }
  if (index !== nodes.length) throw new Error('XML 节点核对失败，未写入修复。');
  replacements.sort((a, b) => b.start - a.start);
  let last = Infinity;
  for (const r of replacements) {
    if (r.end > last) throw new Error('修复规则冲突，未写入文件。');
    text = text.slice(0, r.start) + r.value + text.slice(r.end);
    last = r.start;
  }
  return { text, counts };
}
