// Targets are common on Windows/Office, not evidence about a particular PC.
// Never embed proprietary font binaries or substitute math/symbol families.
const mappings = new Map([
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
const known = new Set([
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

export function isKnownFont(name) {
  return !name?.trim() || /^\+(?:mj|mn)-/.test(name) || known.has(name.trim().toLowerCase());
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
  return { ...options, fontProfile: 'windows-office', fontMappings };
}

export function fontAdvice(name, options = {}) {
  if (!name || /^\+(?:mj|mn)-/.test(name) || isProtectedFont(name)) return null;
  const clean = name.trim();
  const custom = Object.hasOwn(options.fontMappings || {}, clean) ? options.fontMappings[clean] : undefined;
  const value = custom ?? mappings.get(clean.toLowerCase());
  if (value === undefined) return null;
  const to = validName(value);
  if (isProtectedFont(to)) throw new Error('符号、数学或图标字体不能作为普通字体替换目标。');
  if (to === clean) return null;
  return {
    to,
    reason: custom
      ? '使用你指定的目标字体；请确认 Windows 电脑已安装该字体并检查版面。'
      : '此字体常见于 Mac；Windows 可能缺少。可替换为目标环境常见字体，但字宽、换行和外观可能变化。',
    confidence: 'medium',
    defaultSelected: false,
  };
}
