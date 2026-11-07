import { NS, elements, attribute, fontAttribute } from './xml.js';
import { fontAdvice, isKnownFont, isProtectedFont } from './fonts.js';
import { formulaTokens } from './xlsx-formula.js';

const DYNAMIC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const RECENT_FUNCTIONS = new Set([
  'XLOOKUP', 'XMATCH', 'FILTER', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE', 'RANDARRAY',
  'LET', 'LAMBDA', 'BYCOL', 'BYROW', 'MAP', 'REDUCE', 'SCAN', 'MAKEARRAY', 'TAKE',
  'DROP', 'HSTACK', 'VSTACK', 'CHOOSECOLS', 'CHOOSEROWS', 'TEXTSPLIT', 'TEXTBEFORE',
  'TEXTAFTER', 'TOROW', 'TOCOL', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'GROUPBY',
  'PIVOTBY', 'TRIMRANGE', 'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'ISOMITTED',
]);
const DEPENDENT_FUNCTIONS = new Set([
  'RTD', 'WEBSERVICE', 'FILTERXML', 'IMAGE', 'STOCKHISTORY', 'PY', 'SQL.REQUEST',
  'CUBEVALUE', 'CUBEMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBERANKEDMEMBER',
  'CUBEMEMBERPROPERTY', 'CUBEKPIMEMBER',
]);
const TRUE = new Set(['1', 'true']);
const PRIVATE_USE = /[\uE000-\uF8FF\u{F0000}-\u{FFFFD}\u{100000}-\u{10FFFD}]/u;

function children(node, ns, localName) {
  const result = [];
  for (let child = node?.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && child.namespaceURI === ns && child.localName === localName) result.push(child);
  }
  return result;
}

function first(node, ns, localName) { return children(node, ns, localName)[0]; }
function isTrue(value) { return TRUE.has(value); }
function sample(items, count = items.length) { return items.slice(0, 12).join('、') + (count > 12 ? ` 等 ${count} 处` : ''); }
function compact(text, length = 160) { return text.length > length ? `${text.slice(0, length)}…` : text; }

function isWebHyperlink(target) {
  let url;
  try { url = new URL(target); } catch { return false; }
  if (['mailto:', 'tel:'].includes(url.protocol)) return true;
  if (!['http:', 'https:'].includes(url.protocol)) return false;
  const host = url.hostname.replace(/\.$/, '').toLowerCase();
  // A loopback URL names the recipient's own machine, not the sender's service.
  return host && host !== 'localhost' && !host.endsWith('.localhost')
    && !/^127\./.test(host) && host !== '[::1]' && host !== '0.0.0.0';
}

function resolvePart(base, target) {
  const pieces = target.startsWith('/') ? [] : base.split('/').slice(0, -1);
  for (const piece of target.replace(/\\/g, '/').split('/')) {
    if (piece === '..') pieces.pop();
    else if (piece && piece !== '.') pieces.push(piece);
  }
  return pieces.join('/');
}

function worksheetNames(xml) {
  const byId = new Map();
  const relationships = xml.get('xl/_rels/workbook.xml.rels');
  if (relationships) for (const rel of elements(relationships, NS.rel, 'Relationship')) {
    if (attribute(rel, 'TargetMode') !== 'External') byId.set(attribute(rel, 'Id'), resolvePart('xl/workbook.xml', attribute(rel, 'Target')));
  }
  const names = new Map();
  const workbook = xml.get('xl/workbook.xml');
  if (workbook) for (const sheet of elements(workbook, NS.s, 'sheet')) {
    const id = sheet.getAttributeNS(NS.r, 'id');
    if (byId.has(id)) names.set(byId.get(id), attribute(sheet, 'name'));
  }
  return names;
}

function readStyles(xml) {
  const styles = xml.get('xl/styles.xml');
  if (!styles) return { xfs: [], baseXfs: [], numberFormats: new Map(), fontSizes: [], fontNames: [] };
  const numberFormats = new Map(elements(styles, NS.s, 'numFmt').map(node => [attribute(node, 'numFmtId'), attribute(node, 'formatCode')]));
  const xfs = children(elements(styles, NS.s, 'cellXfs')[0], NS.s, 'xf');
  const baseXfs = children(elements(styles, NS.s, 'cellStyleXfs')[0], NS.s, 'xf');
  const fonts = children(elements(styles, NS.s, 'fonts')[0], NS.s, 'font');
  const fontSizes = fonts.map(font => Number(attribute(first(font, NS.s, 'sz'), 'val', '11')) || 11);
  const fontNames = fonts.map(font => attribute(first(font, NS.s, 'name'), 'val'));
  return { xfs, baseXfs, numberFormats, fontSizes, fontNames };
}

function richText(node) {
  // Phonetic guides (rPh) are not part of the actual cell value.
  let text = '';
  for (let child = node?.firstChild; child; child = child.nextSibling) {
    if (child.namespaceURI !== NS.s) continue;
    if (child.localName === 't') text += decodeSpreadsheetText(child.textContent);
    else if (child.localName === 'r') text += children(child, NS.s, 't').map(item => decodeSpreadsheetText(item.textContent)).join('');
  }
  return text;
}

function decodeSpreadsheetText(text) {
  // A single pass preserves escaped literal sequences such as _x005F_xFFFD_.
  return text.replace(/_x([0-9a-f]{4})_/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function textValue(cell, strings) {
  const type = attribute(cell, 't');
  if (type === 's') {
    const raw = first(cell, NS.s, 'v')?.textContent ?? '';
    return /^\d+$/.test(raw) ? strings[Number(raw)] ?? null : null;
  }
  if (type === 'inlineStr') return richText(first(cell, NS.s, 'is'));
  if (type === 'str') return decodeSpreadsheetText(first(cell, NS.s, 'v')?.textContent ?? '');
  return null;
}

function columnIndex(ref) {
  const letters = /^[A-Z]+/i.exec(ref)?.[0];
  if (!letters) return null;
  return [...letters.toUpperCase()].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
}

function cellPosition(ref) {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(ref);
  return match ? { col: columnIndex(match[1]), row: Number(match[2]) } : null;
}

function mergedRanges(doc) {
  return elements(doc, NS.s, 'mergeCell').map(node => {
    const [start, end = start] = attribute(node, 'ref').split(':').map(cellPosition);
    return start && end ? { start, end } : null;
  }).filter(Boolean);
}

function effectiveAlignment(cell, styles) {
  const index = Number(attribute(cell, 's', '0'));
  const xf = styles.xfs[index];
  const alignment = first(xf, NS.s, 'alignment');
  return { alignment, xf, fontSize: styles.fontSizes[Number(attribute(xf, 'fontId', '0'))] ?? 11 };
}

function numericPrecisionRisk(raw) {
  const match = /^([+-]?)(\d+)(?:\.(\d*))?(?:[Ee]([+-]?\d+))?$/.exec(raw);
  if (!match) return false;
  const integer = match[2].replace(/^0+/, '');
  const decimals = match[3] ?? '';
  const significant = (match[2] + decimals).replace(/^0+/, '').replace(/0+$/, '');
  if (!significant) return false;
  const exponent = Number(match[4] ?? '0');
  const integerLength = integer ? integer.length + exponent : exponent - (decimals.match(/^0*/)?.[0].length ?? 0);
  return significant.length > 15 || integerLength > 15;
}

function usesZeroPadding(raw, format) {
  if (!raw || !format.includes('00') || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(raw)) return false;
  // Recognize ordinary positive/negative/zero sections only. Conditions,
  // quoted strings, escapes and locale-specific masks need a real renderer.
  if (/[\[\]"\\]/.test(format)) return false;
  const value = Number(raw), sections = format.split(';');
  const section = value < 0 ? (sections[1] ?? sections[0]) : value === 0 ? (sections[2] ?? sections[0]) : sections[0];
  const mask = /^-?(0{2,})$/.exec(section);
  return !!mask && Number.isFinite(value) && Math.round(Math.abs(value)) < 10 ** (mask[1].length - 1);
}

function formulaFeatures(formula) {
  const { calls, literalLinks, external } = formulaTokens(formula);
  const names = [...calls];
  const recent = names.filter(name => /^_XLFN\.|^_XLWS\./.test(name) || RECENT_FUNCTIONS.has(name));
  const dependent = names.filter(name => /^_XLL\./.test(name) || DEPENDENT_FUNCTIONS.has(name.replace(/^_XLFN\.(?:_XLWS\.)?/, '')));
  const localLink = [...literalLinks].some(target => {
    const path = target.trim();
    // A leading # links within this workbook. [Book.xlsx]Sheet may also refer
    // to this workbook: without a path, do not claim it must be distributed.
    if (!path || path.startsWith('#') || isWebHyperlink(path)) return false;
    if (/^https?:/i.test(path)) {
      try { return !!new URL(path).hostname; } catch { return false; }
    }
    return /^(?:file:|smb:|[A-Za-z]:[\\/])/i.test(path)
      || (!/^[a-z][a-z\d+.-]*:/i.test(path) && (/[\\/]/.test(path) || /^[^:[\]]+\.[a-z\d]{1,12}(?:#.*)?$/i.test(path)));
  });
  return { recent, dependent, external, localLink };
}

function scanOtherFormulas(ctx, add) {
  const formulaLocals = new Set(['formula', 'formula1', 'formula2', 'calculatedColumnFormula', 'totalsRowFormula']);
  for (const [part, doc] of ctx.xml) {
    if (!part.startsWith('xl/')) continue;
    const grouped = { recent: [], dependent: [], external: [], localLink: [] };
    for (const node of elements(doc, NS.s, '*')) {
      if (!formulaLocals.has(node.localName) && node.localName !== 'definedName') continue;
      const features = formulaFeatures(node.textContent ?? '');
      const location = node.localName === 'definedName'
        ? `名称 ${compact(attribute(node, 'name'))}`
        : `${node.localName} ${attribute(node.parentNode, 'sqref') || attribute(node.parentNode?.parentNode, 'sqref') || attribute(node.parentNode, 'name') || ''}`.trim();
      // Named-formula version findings are emitted with workbook metadata above.
      if (features.recent.length && node.localName !== 'definedName') grouped.recent.push(`${location}（${[...new Set(features.recent)].join('、')}）`);
      if (features.dependent.length) grouped.dependent.push(`${location}（${[...new Set(features.dependent)].join('、')}）`);
      if (features.external) grouped.external.push(location);
      if (features.localLink) grouped.localLink.push(location);
    }
    if (grouped.recent.length) add('OTHER_FORMULA_VERSION', part, {
      severity: 'warning', title: '条件、验证或表公式需要核对软件版本',
      description: '单元格以外的公式使用未来函数标记或较新的函数。现代 Windows Excel 可能支持，老版本或其他软件仍需验证；不能据此认定公式错误。',
      evidence: compact(sample(grouped.recent), 800), recommendation: '在目标软件中检查条件格式、数据验证和表公式是否仍按预期工作。',
    });
    if (grouped.dependent.length) add('OTHER_FORMULA_DEPENDENCY', part, {
      severity: 'warning', title: '名称或规则公式依赖外部运行环境',
      description: '单元格以外的公式引用在线服务、数据源或加载项。接收方需要相应运行环境才能计算。',
      evidence: compact(sample(grouped.dependent), 800), recommendation: '在目标电脑核对相关数据源、许可、网络和加载项。',
    });
    if (grouped.external.length) add('OTHER_FORMULA_EXTERNAL_LINK', part, {
      severity: 'warning', title: '名称或规则公式引用其他工作簿',
      description: '这些公式依赖文件外的工作簿，换电脑后需重新确认源文件位置和访问权限。',
      evidence: compact(sample(grouped.external), 800), recommendation: '保留所需源文件，并在目标 Excel 中检查名称、规则和链接的结果。',
    });
    if (grouped.localLink.length) add('OTHER_FORMULA_FILE_HYPERLINK', part, {
      severity: 'warning', title: '名称或规则公式中的链接指向外部文件或本机地址',
      description: 'HYPERLINK 的固定目标是文件路径或本机地址。它可能依赖发送方的目录、共享文件夹或服务。',
      evidence: compact(sample(grouped.localLink), 800),
      recommendation: '核对文件随附、相对目录和接收方权限，再在目标电脑打开链接。工具不会访问或改写地址。',
    });
  }
}

function scanFonts(ctx, add, inventory, privateStyleFonts) {
  const fonts = new Map();
  for (const [part, doc] of ctx.xml) {
    if (!part.startsWith('xl/')) continue;
    const groups = new Map();
    for (const node of elements(doc, '*', '*')) {
      const keyName = fontAttribute(node);
      const name = keyName && attribute(node, keyName);
      if (!name || /^\+(?:mj|mn)-/.test(name)) continue;
      const key = `${node.namespaceURI}\0${name}`;
      const group = groups.get(key) ?? { name, scope: node.namespaceURI, count: 0, supplemental: 0, symbolSensitive: false };
      group.count += 1;
      const richRun = node.localName === 'rFont' ? node.parentNode?.parentNode : null;
      const privateRichText = richRun?.namespaceURI === NS.s && richRun.localName === 'r'
        && children(richRun, NS.s, 't').some(t => PRIVATE_USE.test(decodeSpreadsheetText(t.textContent)));
      if (node.localName === 'sym' || Number(attribute(node, 'charset')) === 2
        || Number(attribute(first(node.parentNode, NS.s, 'charset'), 'val')) === 2
        || privateRichText || (part === 'xl/styles.xml' && node.namespaceURI === NS.s && privateStyleFonts.has(name))
        || isProtectedFont(name)) group.symbolSensitive = true;
      if (node.namespaceURI === NS.a && node.localName === 'font' && attribute(node, 'script')
        && ['majorFont', 'minorFont'].includes(node.parentNode?.localName)) group.supplemental += 1;
      groups.set(key, group);
      const font = fonts.get(name) ?? { name, parts: new Set(), occurrences: 0 };
      font.parts.add(part);
      font.occurrences += 1;
      fonts.set(name, font);
    }
    const supplementalUnknown = [];
    for (const { name, scope, count, supplemental, symbolSensitive } of groups.values()) {
      const advice = fontAdvice(name, ctx.options ?? {});
      if (symbolSensitive && (advice || !isKnownFont(name, ctx.options ?? {}))) {
        add('SYMBOL_FONT_REVIEW', part, {
          severity: 'warning', title: `符号或数学字体「${name}」需要保留并核对`,
          description: '此字体与符号、数学、特殊字符集或私用区字符有关，普通字体替换可能让图标变成方框或其他符号。当前保留这项字体声明。',
          evidence: `在此部件中发现 ${count} 处同名字体声明。`,
          recommendation: '在目标电脑确认原字体与字符可用，再在原始应用中检查符号的含义和显示。',
        });
      } else if (advice && advice.to !== name && !symbolSensitive) {
        add('FONT_REPLACEMENT_AVAILABLE', part, {
          severity: 'warning', title: `字体「${name}」可选择替换`,
          description: `${advice.reason} 替换字体可能改变字宽、换行与图表布局，需要检查预览。`,
          evidence: `在此部件中发现 ${count} 处字体声明。`,
          recommendation: `确认目标电脑有「${advice.to}」后，再勾选替换；修复不会调整列宽、行高或内容。`,
          fixable: true, fix: { kind: 'font-replace', from: name, to: advice.to, scope },
        });
      } else if (!isKnownFont(name, ctx.options ?? {}) && supplemental === count) {
        supplementalUnknown.push(name);
      } else if (!isKnownFont(name, ctx.options ?? {})) {
        add('FONT_UNVERIFIED', part, {
          severity: 'info', title: `字体「${name}」尚未验证`,
          description: ctx.options?.targetFonts?.length
            ? '导入的目标电脑字体名单未列出此名称。其他语言的名称和 Office 云字体可能另外可用；请核对实际显示。'
            : '无法从文件确认目标 Windows 电脑是否安装此字体。缺失时可能发生替换、换行或表格文字截断。',
          evidence: `在此部件中发现 ${count} 处字体声明。`,
          recommendation: '在目标电脑安装获得授权的字体，或在原始 Excel 中更换为接收方确认可用的字体。',
        });
      }
    }
    if (supplementalUnknown.length) add('THEME_FONTS_UNVERIFIED', part, {
      severity: 'info', title: `${supplementalUnknown.length} 种主题备用字体尚未验证`,
      description: '主题为多种文字预设了备用字体。这些声明不代表可见文字实际用到了它们；仅在使用对应文字且目标电脑缺少字体时，才可能影响显示。',
      evidence: compact(supplementalUnknown.join('、'), 800),
      recommendation: '重点核对实际出现的文字与图表；不必仅因主题包含备用字体就全部替换。',
    });
  }
  inventory.fonts = [...fonts.values()].map(font => ({ ...font, parts: [...font.parts].sort() })).sort((a, b) => a.name.localeCompare(b.name));
}

function scanWorkbook(ctx, add, inventory) {
  const part = 'xl/workbook.xml';
  const workbook = ctx.xml.get(part);
  if (!workbook) return;
  const properties = elements(workbook, NS.s, 'workbookPr')[0];
  inventory.dateSystem = isTrue(attribute(properties, 'date1904')) ? '1904' : '1900';
  if (inventory.dateSystem === '1904') add('DATE_SYSTEM_1904', part, {
    severity: 'info', title: '工作簿使用 1904 日期系统',
    description: 'Windows 和 Mac Excel 都支持此设置，直接打开不代表日期有错。与 1900 日期系统的工作簿复制或外部链接日期时，需核对 1462 天的基准差异。',
    evidence: 'workbookPr date1904 为 true / 1。',
    recommendation: '保持原日期系统，并核对跨工作簿传递的日期；本工具不转换日期或改写序列值。',
  });
  if (isTrue(attribute(first(workbook.documentElement, NS.s, 'workbookProtection'), 'lockStructure'))) add('WORKBOOK_STRUCTURE_PROTECTED', part, {
    severity: 'info', title: '工作簿已保护工作表结构',
    description: '原文件限制新增、删除、移动或重命名工作表；这并不表示所有单元格都不能编辑。字体修复副本会保留此设置。',
    evidence: 'workbookProtection 的 lockStructure 已开启。',
    recommendation: '需要调整工作表结构时，请原作者提供允许相应编辑的副本。',
  });
  const calculation = elements(workbook, NS.s, 'calcPr')[0];
  if (attribute(calculation, 'calcMode') === 'manual') add('MANUAL_CALCULATION', part, {
    severity: 'warning', title: '工作簿使用手动计算',
    description: '打开或编辑后，公式结果可能仍是上次保存的缓存。跨电脑查看时应确认是否需要重新计算。',
    evidence: 'calcPr calcMode="manual"。',
    recommendation: '在目标 Excel 中确认计算模式并主动重新计算，核对结果后保存。',
  });
  for (const definedName of elements(workbook, NS.s, 'definedName')) {
    const name = attribute(definedName, 'name');
    if (isTrue(attribute(definedName, 'function')) || isTrue(attribute(definedName, 'vbProcedure')) || isTrue(attribute(definedName, 'xlm'))) {
      add('MACRO_DEFINED_NAME', part, {
        severity: 'warning', title: '名称中带有宏或自定义函数标记',
        description: '此名称依赖宏或自定义函数行为，不能仅靠文件扩展名判断是否能在目标电脑执行。',
        evidence: `名称：${compact(name)}。`, recommendation: '用原应用和目标 Excel 检查其来源与依赖；不要将文件改名为 XLSX 来去除宏。',
      });
    }
    if (/^_xlnm\.Print_(Area|Titles)$/i.test(name)) {
      inventory.hasPrintSettings = true;
    }
    const features = formulaFeatures(definedName.textContent ?? '');
    if (features.recent.length) add('NAMED_FORMULA_VERSION', part, {
      severity: 'warning', title: '名称公式需要核对目标软件版本',
      description: '此名称使用未来函数标记或较新的函数。这不是公式损坏的证明，当前 Windows Excel 可能支持；老版本或其他表格软件需另行验证。',
      evidence: `名称：${compact(name)}；函数：${[...new Set(features.recent)].join('、')}。`,
      recommendation: '在接收方实际使用的 Excel 版本中计算并核对结果，保留原公式。',
    });
  }
}

function scanSheet(ctx, part, doc, name, styles, strings, add, inventory, privateStyleFonts) {
  const cells = elements(doc, NS.s, 'c');
  const sheet = { name, part, cellCount: cells.length, formulaCount: 0 };
  if (isTrue(attribute(first(doc.documentElement, NS.s, 'sheetProtection'), 'sheet'))) add('WORKSHEET_EDIT_PROTECTED', part, {
    severity: 'info', title: '此工作表已启用单元格保护', location: name,
    description: '原作者已启用工作表保护。锁定单元格的编辑会受到限制，未锁定的单元格仍可能可编辑。此设置不是打开文件所需的加密密码。',
    evidence: 'sheetProtection 的 sheet 已开启。', recommendation: '请在目标 Excel 中核对需要填写的区域；需要更改保护范围时，联系原作者。',
  });
  const groups = new Map();
  const mark = (code, ref, extra = '') => {
    const group = groups.get(code) ?? { count: 0, refs: [], extras: new Set() };
    group.count++;
    if (group.refs.length < 12) group.refs.push(ref);
    if (extra && group.extras.size < 12) group.extras.add(compact(extra, 160));
    groups.set(code, group);
  };
  const columns = elements(doc, NS.s, 'col').map(node => ({ min: Number(attribute(node, 'min')), max: Number(attribute(node, 'max')), width: Number(attribute(node, 'width')), hidden: isTrue(attribute(node, 'hidden')), style: attribute(node, 'style', '0') }));
  const defaultWidth = Number(attribute(elements(doc, NS.s, 'sheetFormatPr')[0], 'defaultColWidth', '8.43')) || 8.43;
  const merges = mergedRanges(doc);
  const layoutBudget = cells.length * (columns.length + merges.length);
  const layoutLimited = layoutBudget > 2_000_000;
  let protectColumnStyles = false;
  const rememberStyle = styleIndex => {
    const xf = styles.xfs[Number(styleIndex)];
    const base = styles.baseXfs[Number(attribute(xf, 'xfId', '0'))];
    const family = styles.fontNames[Number(attribute(xf, 'fontId', attribute(base, 'fontId', '0')))];
    if (family) privateStyleFonts.add(family);
    if (base && ['0', 'false'].includes(attribute(xf, 'applyFont'))) {
      const inherited = styles.fontNames[Number(attribute(base, 'fontId', '0'))];
      if (inherited) privateStyleFonts.add(inherited);
    }
  };
  const occupied = new Set(cells.filter(cell => first(cell, NS.s, 'f') || first(cell, NS.s, 'v') || first(cell, NS.s, 'is')).map(cell => {
    const pos = cellPosition(attribute(cell, 'r'));
    return pos ? `${pos.row}:${pos.col}` : '';
  }));
  for (const cell of cells) {
    const ref = attribute(cell, 'r', '?');
    const formula = first(cell, NS.s, 'f');
    const value = first(cell, NS.s, 'v');
    const raw = value?.textContent?.trim() ?? '';
    const type = attribute(cell, 't');
    if (formula) {
      sheet.formulaCount += 1;
      const features = formulaFeatures(formula.textContent ?? '');
      if (features.recent.length) mark('FORMULA_VERSION_REQUIREMENT', ref, features.recent.join('、'));
      if (features.dependent.length) mark('FORMULA_ENVIRONMENT_DEPENDENCY', ref, features.dependent.join('、'));
      if (features.external) mark('EXTERNAL_WORKBOOK_FORMULA', ref);
      if (features.localLink) mark('FORMULA_FILE_HYPERLINK', ref);
      // Empty string results are valid. Lack of a numeric/boolean/error
      // cache is common in generated files; it does not invalidate a formula.
      const hasCachedValue = value ? !!raw || type === 'str' || type === 'inlineStr'
        : type === 'inlineStr' && first(cell, NS.s, 'is');
      if (!hasCachedValue) mark('FORMULA_CACHE_MISSING', ref);
      else if (type === 'e' && raw) mark('FORMULA_CACHE_ERROR', ref, compact(raw, 40));
    }
    // A formula result is not a source identifier. Do not tell users to
    // re-import their formulas as text based on last-calculated values.
    if (!formula && (!type || type === 'n') && raw) {
      if (numericPrecisionRisk(raw)) mark('NUMERIC_PRECISION_RISK', ref);
      if (/^[+-]?0\d+/.test(raw)) mark('NUMERIC_LEADING_ZEROS', ref);
    }
    const text = textValue(cell, strings);
    if (text && PRIVATE_USE.test(text)) {
      mark('PRIVATE_USE_TEXT', ref);
      if (cell.hasAttribute('s')) rememberStyle(attribute(cell, 's'));
      else if (cell.parentNode?.localName === 'row' && cell.parentNode.hasAttribute('s')
        && isTrue(attribute(cell.parentNode, 'customFormat'))) rememberStyle(attribute(cell.parentNode, 's'));
      else if (layoutLimited) protectColumnStyles = true;
      else {
        const idx = columnIndex(ref);
        rememberStyle(columns.find(col => idx >= col.min && idx <= col.max)?.style ?? '0');
      }
    }
    if (text?.includes('\uFFFD')) mark('TEXT_REPLACEMENT_CHARACTER', ref);
    if (text !== null && !formula && /^\d+$/.test(text) && (/^0\d/.test(text) || text.length > 15)) mark('TEXT_IDENTIFIER_PRESERVED', ref);
    const { alignment, xf, fontSize } = effectiveAlignment(cell, styles);
    const format = styles.numberFormats.get(attribute(xf, 'numFmtId')) ?? '';
    if ((!type || type === 'n') && usesZeroPadding(raw, format)) mark('ZERO_PADDED_DISPLAY', ref);

    // A deliberately conservative hint, not a renderer or an autofit operation.
    if (layoutLimited || text === null || !text || formula || isTrue(attribute(alignment, 'shrinkToFit')) || Number(attribute(alignment, 'textRotation', '0')) !== 0) continue;
    const pos = cellPosition(ref);
    if (!pos || merges.some(range => pos.col >= range.start.col && pos.col <= range.end.col && pos.row >= range.start.row && pos.row <= range.end.row)) continue;
    const col = columns.find(item => pos.col >= item.min && pos.col <= item.max);
    if (col?.hidden || isTrue(attribute(cell.parentNode, 'hidden'))) continue;
    const width = Number.isFinite(col?.width) && col.width > 0 ? col.width : defaultWidth;
    const units = [...text].reduce((total, character) => total + (character.codePointAt(0) > 0xff ? 2 : 1), 0) * fontSize / 11;
    if (!isTrue(attribute(alignment, 'wrapText'))) {
      if (units > width * 1.35 && occupied.has(`${pos.row}:${pos.col + 1}`)) mark('TEXT_CLIPPING_HEURISTIC', ref);
    } else if (isTrue(attribute(cell.parentNode, 'customHeight'))) {
      const height = Number(attribute(cell.parentNode, 'ht'));
      const lines = Math.max(text.split(/\r?\n/).length, Math.ceil(units / width));
      if (Number.isFinite(height) && lines * fontSize * 1.2 > height * 1.2) mark('WRAPPED_TEXT_HEIGHT_HEURISTIC', ref);
    }
  }
  if (protectColumnStyles) {
    // Don't multiply thousands of columns by thousands of private-use cells.
    // When resolving those intersections is too costly, protect all candidate
    // column/default families used on this sheet.
    rememberStyle('0');
    for (const col of columns) rememberStyle(col.style);
  }
  inventory.formulaCount += sheet.formulaCount;
  inventory.worksheets.push(sheet);
  if (layoutLimited) add('LAYOUT_HEURISTIC_LIMIT', part, {
    severity: 'info', title: '复杂工作表未进行文字截断估算', location: name,
    description: '这张工作表有较多单元格和布局规则，为避免长时间占用浏览器，已跳过文字长度与列宽、合并区域的交叉估算。字体、公式和数据类型检查仍已执行。',
    recommendation: '在目标 Excel 中人工检查长文本、固定行高和合并单元格。',
  });

  const definitions = {
    FORMULA_VERSION_REQUIREMENT: ['warning', '公式需要核对目标软件版本', '发现未来函数标记或较新的函数；这是版本依赖提示，不代表公式损坏或 Windows 不支持。当前 Excel、旧版 Excel、WPS 等支持范围可能不同。', '在接收方实际使用的软件版本中重算并核对结果，保留原公式。'],
    FORMULA_ENVIRONMENT_DEPENDENCY: ['warning', '公式依赖在线服务、数据源或加载项', '部分公式可能需要网络、许可、数据源或本机加载项。移动文件本身不会携带所有运行条件。', '在目标电脑配置所需的数据源、许可或加载项，并核对刷新后的结果。'],
    EXTERNAL_WORKBOOK_FORMULA: ['warning', '公式引用外部工作簿', '公式依赖另一个工作簿。换电脑后，路径、访问权限或源文件可能不同；缓存结果不保证仍然最新。', '一同提供所需源文件，在目标 Excel 中检查链接并重新计算。'],
    FORMULA_FILE_HYPERLINK: ['warning', '公式中的附件或本机链接可能打不开', 'HYPERLINK 公式写入了固定文件路径或本机服务地址。工作簿不会随附链接指向的文件；点击链接是否可用，取决于接收方的目录、权限或服务。', '一同发送附件并保留相对目录，或在原始 Excel 中更新链接，再到目标电脑打开确认。工具不访问链接、不改写公式。'],
    FORMULA_CACHE_MISSING: ['info', '部分公式没有保存上次计算结果', '公式仍在文件中。只能预览而不能计算公式的软件，可能显示空白或不同的占位值；缺少缓存不能证明公式损坏。', '在可信的原工作簿中重新计算、保存后再发送。如果接收方需要继续编辑，保留公式，不要粘贴为数值。'],
    FORMULA_CACHE_ERROR: ['warning', '部分公式上次保存的结果是错误值', '文件中保存了公式的错误结果。它可能来自当时的数据、依赖或软件版本，也可能是公式有意返回的值；当前工具没有执行公式。', '对照原工作簿，在接收方 Excel 中核对输入、函数和外部依赖后重新计算。'],
    NUMERIC_PRECISION_RISK: ['warning', '长数值需要核对精度', '这些单元格以数值保存，长度或有效数字超过 Excel 的 15 位精度范围。不能仅凭当前文件判断是否已经丢失数字，也不能恢复未知原值。', '如果它们是编号，请与原始数据核对，并在原应用中以文本重新导入；若用于计算，请核对所需精度。'],
    NUMERIC_LEADING_ZEROS: ['warning', '数值存储中出现前导零', 'XML 中带前导零的内容被标记为数值。目标软件可能按数值显示，不能假定这些零属于可丢弃格式或必须保留的编号。', '对照原始数据确认含义；本工具不会猜测并改写数据类型或补零。'],
    TEXT_IDENTIFIER_PRESERVED: ['info', '部分长编号或前导零已以文本保存', '这些内容当前以文本保存，前导零或长数字可保留。此项是保护提示，不是兼容性错误。', '保持文本类型，避免另存 CSV 后直接按数值重新导入。'],
    TEXT_REPLACEMENT_CHARACTER: ['warning', '文字中出现 Unicode 替换字符', '这些单元格保存了替换字符（�）。这可能来自较早的编码转换，也可能是有意输入；单凭当前文件不能确定或还原原字。', '与原始文稿核对；如属乱码，从原应用重新导出或从正确编码的数据源重新导入。'],
    PRIVATE_USE_TEXT: ['warning', '部分文字使用字体专用的字符', '这些单元格包含 Unicode 私用区字符。它们可能是图标、徽标或自定义字，换一套字体可能显示成方框或不同的图形。', '请在目标电脑安装获得授权的原字体，并逐个核对这些单元格。工具保留字符本身，不猜测它们代表的文字。'],
    ZERO_PADDED_DISPLAY: ['info', '部分前导零来自数字显示格式', '这些单元格用补零格式显示数值，前导零不是底层数值的一部分。XLSX 会保留格式，但复制到其他软件或导出纯文本时需核对。', '保留当前数字格式；如需作为外部编号使用，先核对导出的实际文本。'],
    TEXT_CLIPPING_HEURISTIC: ['info', '部分文字可能被相邻单元格挡住（估算）', '根据文本长度、列宽及相邻内容估计，部分未换行文字可能显示不全。此检查没有进行字体排版，也不能证明存在截断。', '在目标 Excel 中检查这些单元格，再决定是否调整列宽或换行；本工具不会自动改变布局。'],
    WRAPPED_TEXT_HEIGHT_HEURISTIC: ['info', '固定行高可能容纳不了换行文字（估算）', '这些单元格设置了换行且行高固定。估算的行数可能超出可见高度；实际显示取决于字体和渲染环境。', '在目标 Excel 中检查这些单元格，必要时手动调整行高。'],
  };
  for (const [code, group] of groups) {
    const [severity, title, description, recommendation] = definitions[code];
    const detail = group.extras.size ? `${code === 'FORMULA_CACHE_ERROR' ? '保存的错误值' : '函数'}：${compact([...group.extras].join('；'), 500)}。` : '';
    const refs = sample(group.refs, group.count);
    add(code, part, { severity, title, description, recommendation, location: `${name} · ${refs}`, evidence: `${group.count} 个单元格：${refs}。${detail}` });
  }
  const pageSetup = elements(doc, NS.s, 'pageSetup')[0];
  const marginDefaults = [
    { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  ];
  const margins = elements(doc, NS.s, 'pageMargins')[0];
  const customMargins = margins && !marginDefaults.some(preset => Object.entries(preset).every(([key, expected]) => !margins.hasAttribute(key) || Math.abs(Number(attribute(margins, key)) - expected) <= 0.001));
  const customPage = pageSetup && [...pageSetup.attributes].some(attr => !['scale', 'orientation', 'usePrinterDefaults', 'firstPageNumber'].includes(attr.localName)
    || (attr.localName === 'scale' && attr.value !== '100')
    || (attr.localName === 'orientation' && !['default', 'portrait'].includes(attr.value))
    || (attr.localName === 'firstPageNumber' && attr.value !== '1'));
  const printSettings = customMargins || customPage
    || elements(doc, NS.s, 'brk').length > 0
    || elements(doc, NS.s, 'headerFooter').some(node => node.textContent.trim())
    || elements(doc, NS.s, 'pageSetUpPr').some(node => isTrue(attribute(node, 'fitToPage')));
  if (printSettings) {
    inventory.hasPrintSettings = true;
    add('PRINT_LAYOUT_REVIEW', part, {
      severity: 'info', title: '打印分页需要在目标电脑复核', location: name,
      description: '工作表包含打印、分页或页眉页脚设置。目标电脑的字体、纸张和打印机驱动可能改变分页；当前检查无法验证实际打印效果。',
      recommendation: '在目标 Excel 的打印预览中确认纸张、缩放和分页，再输出 PDF 或打印。',
    });
  }
}

function scanPackageFeatures(ctx, add, inventory) {
  const names = worksheetNames(ctx.xml);
  for (const [part, doc] of ctx.xml) {
    if (!part.startsWith('xl/')) continue;
    if (part.endsWith('.rels')) {
      const source = part.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
      const cellRefs = new Map();
      for (const link of elements(ctx.xml.get(source), NS.s, 'hyperlink')) {
        const id = link.getAttributeNS(NS.r, 'id');
        if (!id) continue;
        const refs = cellRefs.get(id) ?? [];
        refs.push(attribute(link, 'ref'));
        cellRefs.set(id, refs);
      }
      for (const relationship of elements(doc, NS.rel, 'Relationship')) {
        if (attribute(relationship, 'TargetMode') !== 'External') continue;
        const target = attribute(relationship, 'Target');
        const id = attribute(relationship, 'Id');
        const isOrdinaryLink = attribute(relationship, 'Type').endsWith('/hyperlink') && isWebHyperlink(target);
        const refs = cellRefs.get(id) ?? [];
        const location = names.has(source) ? `${names.get(source)}${refs.length ? ` · ${sample(refs)}` : ''}` : source;
        add(isOrdinaryLink ? 'EXTERNAL_WEB_HYPERLINK' : 'EXTERNAL_RESOURCE_RELATIONSHIP', part, {
          severity: isOrdinaryLink ? 'info' : 'warning',
          title: isOrdinaryLink ? '包含需要外部访问的超链接' : '包含工作簿之外的文件或资源依赖',
          description: isOrdinaryLink
            ? '普通网页、邮件或电话链接本身不是兼容性错误。使用时仍取决于目标电脑的网络、权限及可用应用。'
            : '本机或相对文件路径、局域网位置、本机服务及外链图片等资源不会随工作簿自动搬到另一台电脑。换电脑后可能无法访问，当前没有验证链接是否有效。',
          location, evidence: `${id} → ${compact(target, 800)}`,
          recommendation: isOrdinaryLink
            ? '需要使用链接时，在目标环境确认可以访问。工具不会打开或改写链接。'
            : '在原始 Excel 中确认资源来源，与接收方一起准备所需文件或资源，再在目标电脑测试。工具保留链接，不会下载、嵌入或修改它。',
        });
      }
      continue;
    }
    if (/^xl\/externalLinks\/externalLink\d+\.xml$/i.test(part)) {
      inventory.externalLinks += 1;
      add('EXTERNAL_LINK_PART', part, {
        severity: 'warning', title: '工作簿包含外部链接',
        description: '此文件保存了外部工作簿或 DDE / OLE 链接。目标电脑可能无法访问来源，现有缓存也可能过期。',
        recommendation: '在目标 Excel 的链接管理中核对来源、访问权限与更新结果。',
      });
    }
    if (part === 'xl/connections.xml') for (const connection of elements(doc, NS.s, 'connection')) {
      inventory.connectionCount += 1;
      add('DATA_CONNECTION', part, {
        severity: 'warning', title: '数据连接需要目标电脑环境',
        description: '刷新可能依赖驱动、数据源、网络或登录状态。复制文件不能保证这些依赖可用，当前缓存不等同于刷新成功。',
        evidence: `连接名称：${compact(attribute(connection, 'name', '未命名'))}。`,
        recommendation: '在目标 Windows Excel 中验证连接、驱动与刷新权限。',
      });
    }
    const dynamic = elements(doc, DYNAMIC_NS, 'dynamicArrayProperties').filter(node => isTrue(attribute(node, 'fDynamic')));
    if (dynamic.length) add('DYNAMIC_ARRAY_METADATA', part, {
      severity: 'warning', title: '工作簿含动态数组标记',
      description: '支持动态数组的 Excel 可以处理此功能；旧版软件可能按传统数组处理，溢出范围和调整行为可能不同。此标记不是错误。',
      evidence: `${dynamic.length} 处 fDynamic 标记。`, recommendation: '使用支持所需动态数组功能的目标版本，并验证溢出区域和结果。',
    });
    const objects = elements(doc, NS.s, 'oleObject');
    if (objects.length) add('EMBEDDED_OBJECT', part, {
      severity: 'warning', title: '嵌入对象需要对应应用',
      description: '工作表中的 OLE 嵌入或链接对象可能依赖目标电脑安装的应用。预览可见不代表能够打开或继续编辑。',
      evidence: `${objects.length} 个对象。`, recommendation: '在目标电脑测试对象的打开与编辑；保留原嵌入内容。',
    });
    const controls = elements(doc, NS.s, 'control');
    if (controls.length) add('WORKBOOK_CONTROL', part, {
      severity: 'warning', title: '工作表控件需要兼容性检查',
      description: '表单或 ActiveX 控件可能受平台、Office 安全设置和相关代码影响，不能仅靠静态扫描确认可操作。',
      evidence: `${controls.length} 个控件。`, recommendation: '用目标 Excel 验证控件；本工具不会移除控件或修改安全设置。',
    });
  }
  for (const part of ctx.entries.keys()) {
    if (/^xl\/charts\/chart\d+\.xml$/i.test(part)) inventory.chartCount += 1;
    if (/^xl\/embeddings\//i.test(part) && !part.endsWith('/')) inventory.embeddedObjects += 1;
    if (/^xl\/(?:vbaProject|vbaData)\b|^xl\/(?:macrosheets|intlMacrosheets)\//i.test(part)) add('MACRO_CONTENT', part, {
      severity: 'warning', title: '文件内存在宏相关部件',
      description: '部件名称表明存在 VBA 或宏工作表内容。这类文件不提供自动修复；将扩展名改成 XLSX 不能安全移除或保留宏。',
      recommendation: '回到原应用确认真实格式和宏用途，再按宏工作簿流程处理。',
    });
  }
}

/** Inspect SpreadsheetML without changing formulas, values, styles or dimensions. */
export function scanXlsx(ctx) {
  const issues = [], fontIssues = [];
  const counters = new Map();
  const add = (code, part, details, list = issues) => {
    const key = `${code}:${part}`;
    const index = counters.get(key) ?? 0;
    counters.set(key, index + 1);
    list.push({ id: `${code}:${part}:${index}`, code, severity: 'info', part, location: part, fixable: false, defaultSelected: false, ...details });
  };
  const inventory = { worksheets: [], formulaCount: 0, fonts: [], chartCount: 0, externalLinks: 0, connectionCount: 0, embeddedObjects: 0, dateSystem: '1900', hasPrintSettings: false };
  const privateStyleFonts = new Set();
  scanWorkbook(ctx, add, inventory);
  const styles = readStyles(ctx.xml);
  const strings = elements(ctx.xml.get('xl/sharedStrings.xml'), NS.s, 'si').map(richText);
  const names = worksheetNames(ctx.xml);
  for (const [part, doc] of ctx.xml) if (/^xl\/worksheets\/[^/]+\.xml$/i.test(part)) {
    scanSheet(ctx, part, doc, names.get(part) ?? part.split('/').at(-1), styles, strings, add, inventory, privateStyleFonts);
  }
  scanFonts(ctx, (code, part, details) => add(code, part, details, fontIssues), inventory, privateStyleFonts);
  scanPackageFeatures(ctx, add, inventory);
  scanOtherFormulas(ctx, add);
  return { issues: [...fontIssues, ...issues], inventory };
}
