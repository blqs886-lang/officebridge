// Match family names, never font files or glyph coverage. Keep the full list
// in memory for this document. Reports disclose only names used by the file.
const indices = new WeakMap();
const familyKey = name => name.trim().normalize('NFC').toLowerCase();
export const FONT_LIST_MAX_BYTES = 512 * 1024;

export function normalizeTargetFonts(names) {
  if (!Array.isArray(names) || !names.length || names.length > 4096) {
    throw new Error('字体名单需要包含 1 到 4096 个名称，每行一个。');
  }
  const unique = new Map();
  for (const name of names) {
    if (typeof name !== 'string' || !name.trim() || name.length > 128
      || /[\x00-\x1f\x7f<>\uFFFD\uFFFE\uFFFF]/.test(name) || /[\uD800-\uDFFF]/u.test(name)) {
      throw new Error('字体名单含有无效名称。请使用每行一个字体家族名称的文本文件。');
    }
    const key = familyKey(name);
    if (!unique.has(key)) unique.set(key, name.trim());
  }
  const list = Object.freeze([...unique.values()]);
  indices.set(list, new Set(unique.keys()));
  return list;
}

export function parseFontList(bytes) {
  if (bytes.length > FONT_LIST_MAX_BYTES) throw new Error('字体名单上限为 512 KiB。');
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le'
    : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8';
  let text;
  try { text = new TextDecoder(encoding, { fatal: true }).decode(bytes); }
  catch { throw new Error('请将名单保存为 UTF-8 或带字节顺序标记的 UTF-16 文本后重试。'); }
  return normalizeTargetFonts(text.split(/\r\n|[\r\n]/).map(line => line.trim()).filter(Boolean));
}

export function listedFont(names, name) {
  const key = familyKey(name);
  return indices.get(names)?.has(key) ?? names.some(item => familyKey(item) === key);
}

export function compareFontList(names, declarations) {
  const listed = [], unlisted = [];
  for (const { name } of declarations) {
    if (/^\+(?:mj|mn)-/i.test(name.trim())) continue;
    (listedFont(names, name) ? listed : unlisted).push(name);
  }
  return { nameCount: names.length, listed, unlisted };
}
