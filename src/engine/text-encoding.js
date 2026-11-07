// A saved string can contain the original UTF-8 bytes rendered as Latin-1 or
// Windows-1252 characters. Offer one short, read-only example when at least
// two consecutive complete byte triples decode to basic/Extension A Han text.
const windows1252 = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008dŽ\u008f\u0090‘’“”•–—˜™š›œ\u009džŸ';
const byteByCharacter = new Map(Array.from(windows1252, (character, i) => [character, 0x80 + i]));

function continuation(text, index) {
  const code = text.charCodeAt(index);
  if (code >= 0x80 && code <= 0xbf) return code;
  return byteByCharacter.get(text[index]) ?? -1;
}

function hanAt(text, offset) {
  const lead = text.charCodeAt(offset);
  if (!(lead >= 0xe3 && lead <= 0xe9)) return '';
  const second = continuation(text, offset + 1), third = continuation(text, offset + 2);
  if (second < 0 || third < 0) return '';
  const code = ((lead & 15) << 12) | ((second & 63) << 6) | (third & 63);
  if ((code >= 0x3400 && code <= 0x4dbf) || (code >= 0x4e00 && code <= 0x9fff)) {
    return String.fromCharCode(code);
  }
  return '';
}

export function encodingExample(text) {
  if (!text || !/[ã-é]/.test(text)) return null;
  for (let i = 0; i < text.length;) {
    const start = i;
    let preview = '';
    while (preview.length < 12) {
      const character = hanAt(text, i);
      if (!character) break;
      preview += character;
      i += 3;
    }
    if (preview.length >= 2) return { stored: text.slice(start, i), preview };
    if (i === start) i++;
  }
  return null;
}

export const encodingFinding = {
  severity: 'warning',
  title: '文字可能在导入时使用了错误的编码',
  description: '文件中连续的字符可按 UTF-8 字节重新读成汉字。报告列出的短预览用于和原始资料核对，原文仍原样保留。',
  recommendation: '如果预览与原始资料一致，用正确编码重新导入文字，或在原应用中手动改正。更换字体不会改变已保存的字符。',
};
