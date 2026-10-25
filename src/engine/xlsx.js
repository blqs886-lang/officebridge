import { NS, elements, attribute } from './xml.js';
import { fontAdvice, isKnownFont, isProtectedFont } from './fonts.js';

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
const FONT_ELEMENTS = new Set(['latin', 'ea', 'cs', 'sym', 'buFont', 'font']);
const TRUE = new Set(['1', 'true']);

function children(node, ns, localName) {
  const result = [];
  for (let child = node?.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && child.namespaceURI === ns && child.localName === localName) result.push(child);
  }
  return result;
}

function first(node, ns, localName) { return children(node, ns, localName)[0]; }
function isTrue(value) { return TRUE.has(value); }
function sample(items) { return items.slice(0, 12).join('、') + (items.length > 12 ? ` 等 ${items.length} 处` : ''); }
function compact(text, length = 160) { return text.length > length ? `${text.slice(0, length)}…` : text; }

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
  if (!styles) return { xfs: [], numberFormats: new Map(), fontSizes: [] };
  const numberFormats = new Map(elements(styles, NS.s, 'numFmt').map(node => [attribute(node, 'numFmtId'), attribute(node, 'formatCode')]));
  const xfs = children(elements(styles, NS.s, 'cellXfs')[0], NS.s, 'xf');
  const fontSizes = children(elements(styles, NS.s, 'fonts')[0], NS.s, 'font').map(font => Number(attribute(first(font, NS.s, 'sz'), 'val', '11')) || 11);
  return { xfs, numberFormats, fontSizes };
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
  const exponent = Number(match[4] ?? '0');
  const integerLength = integer ? integer.length + exponent : exponent - (decimals.match(/^0*/)?.[0].length ?? 0);
  return significant.length > 15 || integerLength > 15;
}

function formulaFeatures(formula) {
  // Quoted text and quoted sheet names must not be treated as function calls.
  const code = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'/g, "''");
  const names = [...code.matchAll(/(?:^|[^A-Z0-9_.])([A-Z_][A-Z0-9_.]*)\s*\(/gi)].map(match => match[1].toUpperCase());
  const recent = names.filter(name => /^_XLFN\.|^_XLWS\./.test(name) || RECENT_FUNCTIONS.has(name));
  const dependent = names.filter(name => /^_XLL\./.test(name) || DEPENDENT_FUNCTIONS.has(name.replace(/^_XLFN\.(?:_XLWS\.)?/, '')));
  // External references use a workbook-index/name followed by a sheet and !.
  // Table structured references, e.g. Table1[Amount], are not external links.
  const withoutStrings = formula.replace(/"(?:[^"]|"")*"/g, '""');
  const external = /'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'!|(?:^|[^A-Z0-9_])\[[^\]]+\][^'!+\-*/^&=<>(),;\s]+!/i.test(withoutStrings);
  return { recent, dependent, external };
}

function scanOtherFormulas(ctx, add) {
  const formulaLocals = new Set(['formula', 'formula1', 'formula2', 'calculatedColumnFormula', 'totalsRowFormula']);
  for (const [part, doc] of ctx.xml) {
    if (!part.startsWith('xl/')) continue;
    const grouped = { recent: [], dependent: [], external: [] };
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
  }
}

function scanFonts(ctx, add, inventory) {
  const fonts = new Map();
  for (const [part, doc] of ctx.xml) {
    if (!part.startsWith('xl/')) continue;
    const groups = new Map();
    for (const node of elements(doc, '*', '*')) {
      let name;
      if (node.namespaceURI === NS.a && FONT_ELEMENTS.has(node.localName)) name = attribute(node, 'typeface');
      if (node.namespaceURI === NS.s && ((node.localName === 'name' && node.parentNode?.localName === 'font') || node.localName === 'rFont')) name = attribute(node, 'val');
      if (!name || /^\+(?:mj|mn)-/.test(name)) continue;
      const key = `${node.namespaceURI}\0${name}`;
      const group = groups.get(key) ?? { name, scope: node.namespaceURI, count: 0, supplemental: 0, symbolSensitive: false };
      group.count += 1;
      if (node.localName === 'sym' || attribute(node, 'charset') === '2'
        || attribute(first(node.parentNode, NS.s, 'charset'), 'val') === '2'
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
          description: '此字体声明带有符号、数学或特殊字符集线索，普通字体替换可能把符号变成其他字符。当前不提供自动替换。',
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
          description: '无法从文件确认目标 Windows 电脑是否安装此字体。缺失时可能发生替换、换行或表格文字截断。',
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

function scanSheet(ctx, part, doc, name, styles, strings, add, inventory) {
  const cells = elements(doc, NS.s, 'c');
  const sheet = { name, part, cellCount: cells.length, formulaCount: 0 };
  const groups = new Map();
  const mark = (code, ref, extra = '') => {
    const group = groups.get(code) ?? { refs: [], extras: new Set() };
    group.refs.push(ref);
    if (extra) group.extras.add(extra);
    groups.set(code, group);
  };
  const columns = elements(doc, NS.s, 'col').map(node => ({ min: Number(attribute(node, 'min')), max: Number(attribute(node, 'max')), width: Number(attribute(node, 'width')), hidden: isTrue(attribute(node, 'hidden')) }));
  const defaultWidth = Number(attribute(elements(doc, NS.s, 'sheetFormatPr')[0], 'defaultColWidth', '8.43')) || 8.43;
  const merges = mergedRanges(doc);
  const layoutBudget = cells.length * (columns.length + merges.length);
  const layoutLimited = layoutBudget > 2_000_000;
  const occupied = new Set(cells.filter(cell => first(cell, NS.s, 'f') || first(cell, NS.s, 'v') || first(cell, NS.s, 'is')).map(cell => {
    const pos = cellPosition(attribute(cell, 'r'));
    return pos ? `${pos.row}:${pos.col}` : '';
  }));
  for (const cell of cells) {
    const ref = attribute(cell, 'r', '?');
    const formula = first(cell, NS.s, 'f');
    if (formula) {
      sheet.formulaCount += 1;
      const features = formulaFeatures(formula.textContent ?? '');
      if (features.recent.length) mark('FORMULA_VERSION_REQUIREMENT', ref, features.recent.join('、'));
      if (features.dependent.length) mark('FORMULA_ENVIRONMENT_DEPENDENCY', ref, features.dependent.join('、'));
      if (features.external) mark('EXTERNAL_WORKBOOK_FORMULA', ref);
    }
    const raw = first(cell, NS.s, 'v')?.textContent?.trim() ?? '';
    const type = attribute(cell, 't');
    if ((!type || type === 'n') && raw) {
      if (numericPrecisionRisk(raw)) mark('NUMERIC_PRECISION_RISK', ref);
      if (/^[+-]?0\d+/.test(raw)) mark('NUMERIC_LEADING_ZEROS', ref);
    }
    const text = textValue(cell, strings);
    if (text?.includes('\uFFFD')) mark('TEXT_REPLACEMENT_CHARACTER', ref);
    if (text !== null && !formula && /^\d+$/.test(text) && (/^0\d/.test(text) || text.length > 15)) mark('TEXT_IDENTIFIER_PRESERVED', ref);
    const { alignment, xf, fontSize } = effectiveAlignment(cell, styles);
    const format = styles.numberFormats.get(attribute(xf, 'numFmtId')) ?? '';
    if ((!type || type === 'n') && /^0{2,}$/.test(format.split(';')[0])) mark('ZERO_PADDED_DISPLAY', ref);

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
    NUMERIC_PRECISION_RISK: ['warning', '长数值需要核对精度', '这些单元格以数值保存，长度或有效数字超过 Excel 的 15 位精度范围。不能仅凭当前文件判断是否已经丢失数字，也不能恢复未知原值。', '如果它们是编号，请与原始数据核对，并在原应用中以文本重新导入；若用于计算，请核对所需精度。'],
    NUMERIC_LEADING_ZEROS: ['warning', '数值存储中出现前导零', 'XML 中带前导零的内容被标记为数值。目标软件可能按数值显示，不能假定这些零属于可丢弃格式或必须保留的编号。', '对照原始数据确认含义；本工具不会猜测并改写数据类型或补零。'],
    TEXT_IDENTIFIER_PRESERVED: ['info', '部分长编号或前导零已以文本保存', '这些内容当前以文本保存，前导零或长数字可保留。此项是保护提示，不是兼容性错误。', '保持文本类型，避免另存 CSV 后直接按数值重新导入。'],
    TEXT_REPLACEMENT_CHARACTER: ['warning', '文字中出现 Unicode 替换字符', '这些单元格保存了替换字符（�）。这可能来自较早的编码转换，也可能是有意输入；单凭当前文件不能确定或还原原字。', '与原始文稿核对；如属乱码，从原应用重新导出或从正确编码的数据源重新导入。'],
    ZERO_PADDED_DISPLAY: ['info', '部分前导零来自数字显示格式', '这些单元格用补零格式显示数值，前导零不是底层数值的一部分。XLSX 会保留格式，但复制到其他软件或导出纯文本时需核对。', '保留当前数字格式；如需作为外部编号使用，先核对导出的实际文本。'],
    TEXT_CLIPPING_HEURISTIC: ['info', '部分文字可能被相邻单元格挡住（估算）', '根据文本长度、列宽及相邻内容估计，部分未换行文字可能显示不全。此检查没有进行字体排版，也不能证明存在截断。', '在目标 Excel 中检查这些单元格，再决定是否调整列宽或换行；本工具不会自动改变布局。'],
    WRAPPED_TEXT_HEIGHT_HEURISTIC: ['info', '固定行高可能容纳不了换行文字（估算）', '这些单元格设置了换行且行高固定。估算的行数可能超出可见高度；实际显示取决于字体和渲染环境。', '在目标 Excel 中检查这些单元格，必要时手动调整行高。'],
  };
  for (const [code, group] of groups) {
    const [severity, title, description, recommendation] = definitions[code];
    add(code, part, { severity, title, description, recommendation, location: `${name} · ${sample(group.refs)}`, evidence: `${group.refs.length} 个单元格：${sample(group.refs)}。${group.extras.size ? `函数：${compact([...group.extras].join('；'), 500)}。` : ''}` });
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
  for (const [part, doc] of ctx.xml) {
    if (!part.startsWith('xl/')) continue;
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
  const issues = [];
  const counters = new Map();
  const add = (code, part, details) => {
    const key = `${code}:${part}`;
    const index = counters.get(key) ?? 0;
    counters.set(key, index + 1);
    issues.push({ id: `${code}:${part}:${index}`, code, severity: 'info', part, location: part, fixable: false, defaultSelected: false, ...details });
  };
  const inventory = { worksheets: [], formulaCount: 0, fonts: [], chartCount: 0, externalLinks: 0, connectionCount: 0, embeddedObjects: 0, dateSystem: '1900', hasPrintSettings: false };
  scanFonts(ctx, add, inventory);
  scanWorkbook(ctx, add, inventory);
  const styles = readStyles(ctx.xml);
  const strings = elements(ctx.xml.get('xl/sharedStrings.xml'), NS.s, 'si').map(richText);
  const names = worksheetNames(ctx.xml);
  for (const [part, doc] of ctx.xml) if (/^xl\/worksheets\/[^/]+\.xml$/i.test(part)) {
    scanSheet(ctx, part, doc, names.get(part) ?? part.split('/').at(-1), styles, strings, add, inventory);
  }
  scanPackageFeatures(ctx, add, inventory);
  scanOtherFormulas(ctx, add);
  return { issues, inventory };
}
