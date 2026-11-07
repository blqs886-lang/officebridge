// Inspect formula tokens without evaluating expressions, resolving names or
// opening links. Excel allows paths, punctuation and doubled quotes inside
// string literals, quoted sheet names and structured table column names.
function endOfQuote(text, start) {
  const quote = text[start];
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] !== quote) continue;
    if (text[i + 1] === quote) { i++; continue; }
    return i + 1;
  }
  return text.length;
}

function endOfBrackets(text, start) {
  let depth = 1, nested = false;
  for (let i = start + 1; i < text.length; i++) {
    // In a structured reference a single quote escapes these characters.
    if (text[i] === "'" && text[i + 1] && "'[]#@".includes(text[i + 1])) { i++; continue; }
    if (text[i] === '[') { depth++; nested = true; }
    else if (text[i] === ']' && --depth === 0) return { end: i + 1, nested };
  }
  return { end: text.length, nested: true };
}

function skipSpace(text, i) {
  while (i < text.length && /\s/.test(text[i])) i++;
  return i;
}

function literalLink(text, open) {
  const start = skipSpace(text, open + 1);
  if (text[start] !== '"') return null;
  const end = endOfQuote(text, start);
  // Do not guess links constructed with &, concatenation or cell references.
  if (text[end - 1] !== '"' || ![',', ')', ';'].includes(text[skipSpace(text, end)])) return null;
  return text.slice(start + 1, end - 1).replaceAll('""', '"');
}

export function formulaTokens(text) {
  const calls = new Set(), literalLinks = new Set();
  const word = /[\p{L}\p{N}\p{M}_.$\\]+/uy;
  const sheetPrefix = /[^'"\[\]!+\-*/^&=<>(),;{}\s]*!/y;
  let external = false;
  for (let i = 0; i < text.length;) {
    const c = text[i];
    if (c === '"' || c === "'") {
      const end = endOfQuote(text, i);
      if (c === "'" && text[end] === '!') {
        const path = text.slice(i + 1, end - 1);
        const open = path.indexOf('[');
        if (open >= 0 && path.indexOf(']', open + 1) > open + 1) external = true;
      }
      i = end;
    } else if (c === '[') {
      const { end, nested } = endOfBrackets(text, i);
      // [1]Sheet!A1 and [1]!DefinedName use external workbooks. Tables
      // (including nested column specifiers) may contain function-like text.
      if (!nested && end > i + 2 && (i === 0 || !/[\p{L}\p{N}\p{M}_.$\\\]]/u.test(text[i - 1]))) {
        sheetPrefix.lastIndex = end;
        if (sheetPrefix.test(text)) external = true;
      }
      i = end;
    } else {
      word.lastIndex = i;
      const match = word.exec(text);
      if (!match) { i++; continue; }
      i = word.lastIndex;
      const open = skipSpace(text, i);
      if (text[open] !== '(') continue;
      const name = match[0].toUpperCase();
      calls.add(name);
      if (name === 'HYPERLINK') {
        const link = literalLink(text, open);
        if (link) literalLinks.add(link);
      }
    }
  }
  return { calls, literalLinks, external };
}
