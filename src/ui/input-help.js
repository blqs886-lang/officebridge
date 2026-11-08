// File names can come from a drag or picker. Callers render the message as plain text.
export function unsupportedFileMessage(name) {
  const extension = /\.([^.]+)$/.exec(name)?.[1].toLowerCase();
  const hints = {
    key: '请用 Keynote 打开，导出为 PowerPoint（.pptx），再选择导出的文件。',
    numbers: '请用 Numbers 打开，导出为 Excel（.xlsx），再选择导出的文件。',
    ppt: '请用 PowerPoint 打开，把副本另存为 PowerPoint 演示文稿（.pptx），再来检查。',
    xls: '请用 Excel 打开，把副本另存为 Excel 工作簿（.xlsx），再来检查。',
    pptm: '这份演示含宏或使用了启用宏的格式。请保留原件，在目标电脑的 PowerPoint 中核对。',
    xlsm: '这份表格使用了启用宏的格式。请保留原件，在目标电脑的 Excel 中核对。',
    pdf: 'PDF 可以用来对照排版。要检查可编辑的内容，请选择导出这份 PDF 时使用的 PPTX 或 XLSX 原稿。',
    csv: '如果中文显示正确，请用 Excel 将表格另存为 .xlsx 后再来检查。若打开时已经乱码，请先在 Excel 导入时核对文本编码。',
  };
  const message = Object.hasOwn(hints, extension) ? hints[extension] : '请选择 PowerPoint 演示文稿（.pptx）或 Excel 工作簿（.xlsx）。';
  return name + '：' + message;
}

const displayNames = new Map([
  ['pingfang sc', '苹方（简体）'], ['pingfang tc', '苹方（繁体）'],
  ['heiti sc', '黑体（简体）'], ['songti sc', '宋体（简体）'], ['kaiti sc', '楷体（简体）'],
  ['microsoft yahei', '微软雅黑'], ['microsoft jhenghei', '微软正黑体'],
  ['simsun', '宋体'], ['kaiti', '楷体'],
]);

export function fontPairHint(from, to) {
  const source = from.toLowerCase();
  const target = to.toLowerCase();
  if (displayNames.has(source) || displayNames.has(target)) {
    return `将「${displayNames.get(source) || from}」换为「${displayNames.get(target) || to}」。`;
  }
  if (['menlo', 'monaco'].includes(source) && ['consolas', 'courier new'].includes(target)) {
    return '两种字体都让每个字母占相同的宽度。适合试着修复数字或代码错列。';
  }
  if (['helvetica', 'helvetica neue'].includes(source) && target === 'arial') {
    return '两种都是无衬线字体。请对照原稿看行末文字有没有换行。';
  }
  return '';
}
