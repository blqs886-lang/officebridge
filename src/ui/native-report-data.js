import { normalizeTargetFonts } from '../engine/target-fonts.js';

export const NATIVE_REPORT_LIMIT = 110 * 1024 * 1024;
const PDF_LIMIT = 25 * 1024 * 1024;
const fingerprint = value => typeof value === 'string' && /^[a-f0-9]{64}$/i.test(value);
const label = (value, fallback = '') => typeof value === 'string' ? value.slice(0, 500) : fallback;

export async function readNativeReport(text) {
  if (typeof text !== 'string' || text.length > NATIVE_REPORT_LIMIT) throw new Error('核对记录超过 110 MB，请重新导出较小的预览。');
  let data;
  try { data = JSON.parse(text.replace(/^\uFEFF/, '')); }
  catch { throw new Error('这份文件无法作为 Office 核对记录读取，请选择助手生成的 .officebridge-check.json。'); }
  if (data?.schema !== 'officebridge-native-check' || data.schemaVersion !== 1 ||
      !['windows', 'macos'].includes(data.environment?.platform) ||
      !Array.isArray(data.documents) || data.documents.length < 1 || data.documents.length > 2) {
    throw new Error('核对记录的格式或版本不支持，请使用当前版本的 OfficeBridge 助手重新核对。');
  }
  const seen = new Set();
  const documents = [];
  for (const item of data.documents) {
    const hasHash = fingerprint(item?.sha256);
    const unreadFailure = !hasHash && (item?.sha256 === null || item?.sha256 === undefined || item?.sha256 === '') &&
      item?.openedSuccessfully === false && item?.sourceUnchanged !== true && !item?.pdf &&
      typeof item.error === 'string' && item.error.trim().length > 0;
    if (!['original', 'copy'].includes(item?.role) || seen.has(item.role) ||
        !['pptx', 'xlsx'].includes(item.format) || (!hasHash && !unreadFailure)) {
      throw new Error('核对记录缺少有效的文件类型、指纹或原件标记。');
    }
    seen.add(item.role);
    const pdf = item.pdf ? await readPdf(item.pdf) : null;
    documents.push({
      role: item.role, format: item.format, fileName: label(item.fileName, '未命名文件'),
      sha256: hasHash ? item.sha256.toLowerCase() : '', openedSuccessfully: item.openedSuccessfully === true,
      sourceUnchanged: item.sourceUnchanged === true, officeApplication: label(item.officeApplication),
      officeVersion: label(item.officeVersion), officeBuild: label(item.officeBuild),
      error: label(item.error), pdfError: label(item.pdfError), cleanupNote: label(item.cleanupNote), pdf,
      slideCount: Number.isSafeInteger(item.slideCount) && item.slideCount >= 0 ? item.slideCount : undefined,
      sheetCount: Number.isSafeInteger(item.sheetCount) && item.sheetCount >= 0 ? item.sheetCount : undefined,
    });
  }
  if (!seen.has('original') || documents.some(item => item.format !== documents[0].format)) {
    throw new Error('请选择同一种格式的原件和副本进行核对。');
  }
  const comparisons = Array.isArray(data.comparisons) ? data.comparisons.slice(0, 50).flatMap(item => {
    if (typeof item === 'string') return [label(item)];
    return Array.isArray(item?.details) ? item.details.slice(0, 100).filter(value => typeof value === 'string').map(value => label(value)) : [];
  }) : [];
  let fontFamilies = null, fontListNote = label(data.environment.fontInventoryNote);
  if (data.environment.fontFamilies != null) {
    try { fontFamilies = normalizeTargetFonts(data.environment.fontFamilies); }
    catch { fontListNote = '记录里的字体名单无效，已跳过；Office 打开结果与 PDF 仍可查看。'; }
  }
  return {
    platform: data.environment.platform, createdAt: label(data.createdAt), documents, comparisons, fontFamilies, fontListNote,
    referencePdf: data.referencePdf ? await readPdf(data.referencePdf) : null,
  };
}

async function readPdf(pdf) {
  if (typeof pdf?.base64 !== 'string' || pdf.base64.length > Math.ceil(PDF_LIMIT / 3) * 4 ||
      !/^[a-z0-9+/]*={0,2}$/i.test(pdf.base64) || !fingerprint(pdf.sha256)) {
    throw new Error('PDF 预览缺少有效指纹或超过 25 MB。请用助手重新生成核对记录。');
  }
  let decoded;
  try { decoded = atob(pdf.base64); } catch { throw new Error('PDF 预览数据已损坏。'); }
  const bytes = Uint8Array.from(decoded, character => character.charCodeAt(0));
  if (bytes.length > PDF_LIMIT || bytes.length < 8 || String.fromCharCode(...bytes.subarray(0, 5)) !== '%PDF-') {
    throw new Error('记录中包含无效的 PDF 预览。');
  }
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  const hash = Array.from(digest, value => value.toString(16).padStart(2, '0')).join('');
  if (hash !== pdf.sha256.toLowerCase()) throw new Error('PDF 预览与记录中的指纹不一致，请用助手重新核对。');
  return { fileName: label(pdf.fileName, '预览.pdf'), bytes, sha256: hash };
}
