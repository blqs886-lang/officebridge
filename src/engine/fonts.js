import { listedFont, normalizeTargetFonts } from './target-fonts.js';

// Targets are common on Windows/Office, not evidence about a particular PC.
// Never embed proprietary font binaries or substitute math/symbol families.
const windowsMappings = new Map([
  ['pingfang sc', 'Microsoft YaHei'], ['苹方-简', 'Microsoft YaHei'],
  ['pingfang tc', 'Microsoft JhengHei'], ['苹方-繁', 'Microsoft JhengHei'],
  ['pingfang hk', 'Microsoft JhengHei'], ['heiti sc', 'Microsoft YaHei'],
  ['heiti tc', 'Microsoft JhengHei'], ['stheiti', 'Microsoft YaHei'],
  ['stheiti sc', 'Microsoft YaHei'], ['stheiti light', 'Microsoft YaHei'],
  ['stheiti medium', 'Microsoft YaHei'], ['hiragino sans gb', 'Microsoft YaHei'],
  ['songti sc', 'SimSun'], ['songti tc', 'PMingLiU'],
  ['kaiti sc', 'KaiTi'], ['kaiti tc', 'DFKai-SB'],
  ['helvetica', 'Arial'], ['helvetica neue', 'Arial'],
  ['times', 'Times New Roman'], ['courier', 'Courier New'],
  ['menlo', 'Consolas'], ['monaco', 'Consolas'],
  ['sf pro', 'Arial'], ['sf pro display', 'Arial'], ['sf pro text', 'Arial'],
  ['.applesystemuifont', 'Arial'], ['lucida grande', 'Arial'],
]);
const windowsKnown = new Set([
  'Arial', 'Arial Black', 'Arial Narrow', 'Calibri', 'Calibri Light', 'Cambria',
  'Cambria Math', 'Candara', 'Consolas', 'Constantia', 'Corbel', 'Courier New',
  'Georgia', 'Impact', 'Segoe UI', 'Segoe UI Symbol', 'Segoe UI Emoji', 'Symbol',
  'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Webdings', 'Wingdings',
  'Wingdings 2', 'Wingdings 3', 'Microsoft YaHei', 'Microsoft YaHei UI',
  '微软雅黑', 'SimSun', '宋体', 'NSimSun', '新宋体', 'SimHei', '黑体',
  'KaiTi', '楷体', 'FangSong', '仿宋', 'Microsoft JhengHei', 'Microsoft JhengHei UI',
  '微軟正黑體', 'MingLiU', 'PMingLiU', 'DFKai-SB', 'MS Gothic', 'MS PGothic',
  'MS Mincho', 'MS PMincho', 'Meiryo', 'Yu Gothic', 'Yu Mincho', 'Malgun Gothic',
].map(s => s.toLowerCase()));

// Apple documents these families as installed or downloadable. A profile gives
// optional candidates, while a target-machine list confirms available names.
// Common Office fonts can be supplied by Office itself, including cloud fonts.
const macMappings = new Map([
  ['microsoft yahei', 'PingFang SC'], ['microsoft yahei ui', 'PingFang SC'],
  ['微软雅黑', 'PingFang SC'], ['simhei', 'PingFang SC'], ['黑体', 'PingFang SC'],
  ['dengxian', 'PingFang SC'], ['dengxian light', 'PingFang SC'], ['等线', 'PingFang SC'],
  ['simsun', 'Songti SC'], ['宋体', 'Songti SC'], ['nsimsun', 'Songti SC'], ['新宋体', 'Songti SC'],
  ['kaiti', 'Kaiti SC'], ['楷体', 'Kaiti SC'], ['fangsong', 'STFangsong'], ['仿宋', 'STFangsong'],
  ['microsoft jhenghei', 'PingFang TC'], ['microsoft jhenghei ui', 'PingFang TC'],
  ['微軟正黑體', 'PingFang TC'], ['mingliu', 'Songti TC'], ['pmingliu', 'Songti TC'],
  ['dfkai-sb', 'Kaiti TC'], ['segoe ui', 'Helvetica Neue'], ['segoe ui variable', 'Helvetica Neue'],
]);
// Offer these only with evidence that the source family was not listed on the
// chosen Mac. Office's own or cached fonts may otherwise preserve the original.
const macMissingMappings = new Map([
  ['calibri', 'Arial'], ['calibri light', 'Arial'], ['cambria', 'Times New Roman'],
  ['consolas', 'Menlo'], ['aptos', 'Arial'], ['aptos display', 'Arial'],
  ['aptos narrow', 'Arial Narrow'], ['aptos mono', 'Menlo'], ['aptos serif', 'Times New Roman'],
]);
const macKnown = new Set([
  'Arial', 'Arial Black', 'Arial Narrow', 'Arial Unicode MS', 'Calibri', 'Calibri Light',
  'Cambria', 'Cambria Math', 'Candara', 'Consolas', 'Constantia', 'Corbel', 'Courier New',
  'Georgia', 'Impact', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Symbol',
  'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Webdings', 'Apple Color Emoji',
  'Helvetica', 'Helvetica Neue', 'Times', 'Courier', 'Menlo', 'Monaco', 'Lucida Grande',
  'PingFang SC', 'PingFang TC', 'PingFang HK', '苹方-简', '苹方-繁',
  'Heiti SC', 'Heiti TC', 'STHeiti', 'Songti SC', 'Songti TC', 'Kaiti SC', 'Kaiti TC',
  'STFangsong', 'Hiragino Sans GB', 'Hiragino Sans', 'Hiragino Mincho ProN',
  'Apple SD Gothic Neo',
].map(name => name.toLowerCase()));
for (const name of ['Aptos', 'Aptos Display', 'Aptos Narrow', 'Aptos Mono', 'Aptos Serif']) {
  windowsKnown.add(name.toLowerCase());
  macKnown.add(name.toLowerCase());
}

export function isKnownFont(name, options = {}) {
  if (!name?.trim() || /^\+(?:mj|mn)-/.test(name)) return true;
  if (options.targetFonts) return listedFont(options.targetFonts, name);
  const known = options.fontProfile === 'mac-office' ? macKnown : windowsKnown;
  return known.has(name.trim().toLowerCase());
}

export function isProtectedFont(name) {
  return /(?:symbol|wingdings|webdings|dingbat|math|stix|font\s*awesome|material\s*(?:icons|symbols)|emoji)/i.test(name);
}

function validName(value) {
  if (typeof value !== 'string' || !value.trim() || value.length > 128
    || /[\x00-\x1f\x7f<>\uFFFD\uFFFE\uFFFF]/.test(value)
    || /[\uD800-\uDFFF]/u.test(value)) {
    throw new Error('字体名称无效：请输入 1–128 个有效字符，不能只含空格或含控制字符。');
  }
  return value.trim();
}

export function normalizeFontOptions(options = {}) {
  if (!options || typeof options !== 'object' || Array.isArray(options)) throw new Error('检查选项无效。');
  const fontProfile = options.fontProfile ?? 'windows-office';
  if (!['windows-office', 'mac-office'].includes(fontProfile)) throw new Error('目标环境无效，请选择 Windows Office 或 Mac Office。');
  const source = options.fontMappings ?? {};
  if (typeof source !== 'object' || Array.isArray(source)) throw new Error('自定义字体规则必须是名称对应表。');
  const fontMappings = Object.create(null);
  if (Object.keys(source).length > 256) throw new Error('自定义字体规则最多 256 条。');
  for (const [from, target] of Object.entries(source)) {
    const clean = validName(from), to = validName(target);
    if (/^\+(?:mj|mn)-/.test(clean) || /^\+(?:mj|mn)-/.test(to)) throw new Error('主题引用不能作为字体名称。');
    if (isProtectedFont(clean) || isProtectedFont(to)) throw new Error('符号、数学或图标字体不能通过普通字体替换规则修改。');
    if (Object.hasOwn(fontMappings, clean)) throw new Error('自定义字体规则含有重复名称。');
    fontMappings[clean] = to;
  }
  const targetFonts = options.targetFonts == null ? undefined : normalizeTargetFonts(options.targetFonts);
  return { ...options, targetFonts, fontProfile, fontMappings };
}

export function fontAdvice(name, options = {}) {
  if (!name || /^\+(?:mj|mn)-/.test(name) || isProtectedFont(name)) return null;
  const clean = name.trim();
  const custom = Object.hasOwn(options.fontMappings || {}, clean) ? options.fontMappings[clean] : undefined;
  if (options.targetFonts && !custom && listedFont(options.targetFonts, clean)) return null;
  const mac = options.fontProfile === 'mac-office';
  const mappings = mac ? macMappings : windowsMappings;
  const missingFallback = mac && options.targetFonts ? macMissingMappings.get(clean.toLowerCase()) : undefined;
  const value = custom ?? mappings.get(clean.toLowerCase()) ?? missingFallback;
  if (value === undefined) return null;
  const to = validName(value);
  if (isProtectedFont(to)) throw new Error('符号、数学或图标字体不能作为普通字体替换目标。');
  if (to.toLowerCase() === clean.toLowerCase() || (options.targetFonts && !listedFont(options.targetFonts, to))) return null;
  return {
    to,
    reason: options.targetFonts
      ? '目标字体名称已在导入的名单中找到，仍需核对实际字形和换行。'
      : custom
      ? `使用你指定的目标字体；请确认 ${mac ? 'Mac' : 'Windows'} 电脑已安装该字体并检查版面。`
      : mac
      ? '此字体常见于 Windows，接收方 Mac 可能缺少，也可能由 Office 提供。仅在原件显示异常时考虑这个 Mac 字体候选，并核对换行。'
      : '此字体常见于 Mac；Windows 可能缺少。可替换为目标环境常见字体，但字宽、换行和外观可能变化。',
    confidence: 'medium',
    defaultSelected: false,
  };
}
