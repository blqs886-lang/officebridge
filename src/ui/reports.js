export function escapeHtml(value = '') {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[character]);
}

export function safeFileName(name = 'document', extension = '') {
  let clean = String(name).normalize('NFC').replace(/[\u0000-\u001f\u007f\u202a-\u202e\u2066-\u2069<>:"/\\|?*]/g, '_')
    .replace(/[. ]+$/g, '') || 'document';
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(clean)) clean = `_${clean}`;
  const suffix = /\.[a-z0-9]{1,10}$/i.exec(clean)?.[0] || '';
  if (clean.length > 160) clean = clean.slice(0, 160 - suffix.length) + suffix;
  return clean + extension;
}

export function formatBytes(value = 0) {
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  return `${(value / 1024 / 1024).toFixed(1)} MB`;
}

export function exportPayload(record) {
  const options = record.report?.analysisOptions || record.options || {};
  return {
    exportVersion: 1,
    generatedAt: new Date().toISOString(),
    analysisOptions: {
      fontProfile: options.fontProfile || 'windows-office',
      fontMappings: { ...(options.fontMappings || {}) },
    },
    analysis: record.report,
    repair: record.repaired ? {
      fileName: record.repaired.fileName,
      changes: record.repaired.changes,
      report: record.repaired.report,
    } : null,
    limitations: [
      '这是文件结构与兼容风险检查，不是 Windows Office 实际渲染验证。',
      '字体替换可能改变换行、行高和版式，仍需在目标电脑上复核。',
      '未安装字体、应用版本、外部链接和插件可能影响最终效果。',
    ],
  };
}

const severityLabels = { error: '需处理', warning: '需复核', info: '提示' };

export function buildHtmlReport(record) {
  const payload = exportPayload(record);
  const report = payload.analysis;
  const esc = escapeHtml;
  const rows = (report.issues || []).map((issue) => `<tr><td><span class="badge ${['error', 'warning', 'info'].includes(issue.severity) ? issue.severity : 'info'}">${esc(severityLabels[issue.severity] || '提示')}</span></td><td><strong>${esc(issue.title)}</strong><p>${esc(issue.description)}</p>${issue.evidence ? `<p>证据：${esc(issue.evidence)}</p>` : ''}${issue.recommendation ? `<p>建议：${esc(issue.recommendation)}</p>` : ''}${issue.fix ? `<p>字体建议：${esc(issue.fix.from)} → ${esc(issue.fix.to)}（需自行确认）</p>` : ''}</td><td>${esc(issue.location || issue.part || '文件整体')}${issue.part ? `<code>${esc(issue.part)}</code>` : ''}</td></tr>`).join('');
  const changes = (record.repaired?.changes || []).map((change) => `<tr><td>${esc(change.part)}</td><td>${esc(change.before)} → ${esc(change.after)}</td><td>${esc(change.count)}</td></tr>`).join('');
  const repairedIssues = record.repaired?.report?.issues || [];
  const mappingRows = Object.entries(payload.analysisOptions.fontMappings).map(([from, to]) => `<tr><td>${esc(from)}</td><td>${esc(to)}</td></tr>`).join('');
  const comparison = report.inventory?.targetFontComparison;
  const comparisonHtml = comparison ? `<h2>目标 Windows 字体名单核对</h2><p>已核对 ${comparison.nameCount} 个目标字体名称，含文档的主题备用声明。以下仅列出文档声明的字体。</p><p>名单有：${esc(comparison.listed.join('、') || '无')}</p><p>名单未列出：${esc(comparison.unlisted.join('、') || '无')}</p><p>名称匹配用于筛选替换建议。别名、云字体、嵌入字体、字形覆盖与实际换行需在目标电脑确认。</p>` : '';
  const repairedRows = repairedIssues.map((issue) => `<tr><td>${esc(severityLabels[issue.severity] || '提示')}</td><td><strong>${esc(issue.title)}</strong><p>${esc(issue.description)}</p>${issue.evidence ? `<p>证据：${esc(issue.evidence)}</p>` : ''}${issue.recommendation ? `<p>建议：${esc(issue.recommendation)}</p>` : ''}</td><td>${esc(issue.location || issue.part || '文件整体')}</td></tr>`).join('');
  return `<!doctype html>
<html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src 'none'; base-uri 'none'; form-action 'none'"><title>${esc(report.fileName)} — OfficeBridge 检查报告</title><style>
*{box-sizing:border-box}body{margin:0;overflow-wrap:anywhere;background:#f5f7fa;color:#1b2942;font:15px/1.7 system-ui,-apple-system,"Segoe UI",sans-serif}main{max-width:1050px;margin:40px auto;padding:36px;background:white;border:1px solid #dce3ec;border-radius:18px}header{border-bottom:1px solid #e5eaf1;padding-bottom:20px}.brand{color:#245fe5;font-weight:750;letter-spacing:.02em}h1{font-size:26px;overflow-wrap:anywhere}h2{font-size:20px;margin-top:32px}p{margin:8px 0}.muted{color:#65748c}.notice{padding:16px;border-left:4px solid #df9a23;background:#fff8e8;margin:24px 0}table{width:100%;border-collapse:collapse}td,th{vertical-align:top;text-align:left;border-bottom:1px solid #e5eaf1;padding:14px 10px;overflow-wrap:anywhere}th{background:#f5f7fb;font-size:13px}code{display:block;font-size:12px;white-space:normal;color:#65748c;margin-top:8px;overflow-wrap:anywhere}.badge{display:inline-block;white-space:nowrap;font-size:12px;border-radius:5px;padding:4px 7px}.error{color:#ad3542;background:#fff0f1}.warning{color:#8d5a0a;background:#fff7e7}.info{color:#315f98;background:#edf4ff}footer{border-top:1px solid #e5eaf1;margin-top:32px;padding-top:20px;font-size:12px;color:#65748c}@media(max-width:650px){main{margin:0;border:0;border-radius:0;padding:22px}table{font-size:12px}td,th{padding:10px 6px}}@media print{body{background:white}main{border:0;margin:0;padding:0}tr{break-inside:avoid}}
</style></head><body><main><header><div class="brand">OfficeBridge · 兼容检查报告</div><h1>${esc(report.fileName)}</h1><p class="muted">${esc(String(report.format || '').toUpperCase())} · ${esc(formatBytes(report.fileSize))} · ${esc(new Date(payload.generatedAt).toLocaleString('zh-CN'))}</p></header>
<div class="notice"><strong>尚未在 Windows Office 中实际验证</strong><p>本报告基于文档结构检查。没有发现风险，不等于已保证排版一致。请用目标电脑打开副本，检查版式与功能。</p></div>
<h2>本次检查参数</h2><p>目标环境：Windows Office。${mappingRows ? '以下为用户为此文件指定的字体规则，仅生成建议；只有修改记录中列出的替换已应用到副本。目标电脑仍须安装相应字体，且替换可能影响排版。' : '未设置自定义字体规则，使用内置字体建议。'}</p>${mappingRows ? `<table><thead><tr><th>原字体</th><th>用户指定的目标字体</th></tr></thead><tbody>${mappingRows}</tbody></table>` : ''}
${comparisonHtml}<h2>原文件检查结果</h2><p>${report.issues?.length || 0} 项发现。${report.repairBlocked ? '文件禁止自动修复；请在原始应用中处理后重新导出。' : '仅明确选择的字体替换会应用到副本，原件不变。'}</p>
${rows ? `<table><thead><tr><th>级别</th><th>发现与建议</th><th>位置</th></tr></thead><tbody>${rows}</tbody></table>` : '<p>没有发现当前规则覆盖的风险。</p>'}
${record.repaired ? `<h2>修复副本</h2><p>文件名：${esc(record.repaired.fileName)}</p><p>修复后复查仍有 ${repairedIssues.length} 项发现；这些发现不代表全部是可自动修复的问题。</p>${changes ? `<table><thead><tr><th>修改部件</th><th>修改内容</th><th>次数</th></tr></thead><tbody>${changes}</tbody></table>` : '<p>没有记录到文件修改。</p>'}<h2>修复后的复查发现</h2>${repairedRows ? `<table><thead><tr><th>级别</th><th>发现与建议</th><th>位置</th></tr></thead><tbody>${repairedRows}</tbody></table>` : '<p>没有发现当前规则覆盖的风险。仍需在目标 Windows 软件中复核。</p>'}` : ''}
<h2>交接前复核</h2><ol><li>在目标 Windows 电脑上用实际使用的 Office 或 WPS 版本打开。</li><li>核对字体、文字换行、图片位置、图表与打印版式。</li><li>演示文稿试播动画、音视频；表格检查公式结果、外部链接和日期。</li><li>保存、关闭并重新打开副本，确认仍然可编辑。</li></ol>
<footer><p>原件 SHA-256：<code>${esc(report.sha256 || '未提供')}</code></p><p>OfficeBridge ${esc(report.toolVersion || '')} · 本报告不包含 Office 渲染验证结论。文档在本机浏览器处理；报告可能包含文件名、字体和部件路径，分享前请检查。</p></footer></main></body></html>`;
}

export function downloadBlob(content, name, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = safeFileName(name);
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
