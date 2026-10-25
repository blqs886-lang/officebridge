import './style.css';
import EngineWorker from './ui/worker.js?worker&inline';
import { escapeHtml, formatBytes, safeFileName, exportPayload, buildHtmlReport, downloadBlob } from './ui/reports.js';
import { donationUrl, label as donationLabel, validateDonationUrl } from './support-config.js';

const MAX_FILE_SIZE = 100 * 1024 * 1024;
const MAX_FILES = 20;
const MAX_TOTAL_SIZE = 250 * 1024 * 1024;
const WORKER_TIMEOUT_MS = 60_000;
const state = { records: [], activeId: null, processing: false, filter: 'all', toastTimer: null, demoBusy: false };
const pending = new Map();
let nextRequest = 1;
let nextRecord = 1;
let worker = null;
let workerFailed = false;

function rejectPending(message) {
  for (const job of pending.values()) {
    clearTimeout(job.timer);
    job.reject(new Error(message));
  }
  pending.clear();
}

function initializeWorker() {
  try {
    const instance = new EngineWorker();
    worker = instance;
    workerFailed = false;
    instance.addEventListener('message', ({ data }) => {
      if (worker !== instance) return;
      const job = pending.get(data.id);
      if (!job) return;
      clearTimeout(job.timer);
      pending.delete(data.id);
      if (data.error) job.reject(Object.assign(new Error(data.error.message), { code: data.error.code }));
      else job.resolve(data.result);
    });
    instance.addEventListener('error', () => {
      if (worker !== instance) return;
      workerFailed = true;
      instance.terminate();
      rejectPending('本机检查程序遇到错误。请重新添加文件，或刷新页面后重试。');
      toast('本机检查程序遇到错误，原文件保持不变。可重新添加文件重试。', 'error');
    });
  } catch {
    workerFailed = true;
    worker = null;
  }
}

function restartAfterTimeout() {
  worker?.terminate();
  worker = null;
  rejectPending('文件处理超过 60 秒，已停止本次处理。原文件没有修改。请在原始应用中拆分或重新导出后再试。');
  initializeWorker();
}

function request(type, payload = {}, transfer = []) {
  if (!worker || workerFailed) initializeWorker();
  if (workerFailed) return Promise.reject(new Error('本机检查程序不可用，请刷新页面。'));
  const id = nextRequest++;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(restartAfterTimeout, WORKER_TIMEOUT_MS);
    pending.set(id, { resolve, reject, timer });
    try {
      worker.postMessage({ id, type, ...payload }, transfer);
    } catch (error) {
      clearTimeout(timer);
      pending.delete(id);
      reject(error);
    }
  });
}

const icons = {
  bridge: '<path d="M4 16V8h5v8m6 0V8h5v8M7 12h10"/>',
  upload: '<path d="M12 16V4m-4 4 4-4 4 4M4 15v4a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-4"/>',
  shield: '<path d="M12 3 4 6v6c0 4 8 9 8 9s8-5 8-9V6l-8-3Z"/><path d="m8 12 3 3 5-6"/>',
  file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9l-6-6Z"/><path d="M14 3v6h6M8 14h8M8 17h5"/>',
  arrow: '<path d="M4 12h16m-6-6 6 6-6 6"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  download: '<path d="M12 3v12m-4-4 4 4 4-4M4 16v4h16v-4"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6m0-10v.1"/>',
  warning: '<path d="m12 3 10 18H2L12 3Z"/><path d="M12 9v5m0 3v.1"/>',
  close: '<path d="m6 6 12 12M6 18 18 6"/>',
  sparkle: '<path d="m12 3 2.3 6.7L21 12l-6.7 2.3L12 21l-2.3-6.7L3 12l6.7-2.3L12 3Z"/>',
  pptx: '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M12 17v4m-4 0h8M7 8h5m-5 4h9"/>',
  xlsx: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M4 9h16M4 15h16M10 3v18"/>',
};
function icon(name, className = '') {
  return `<svg class="icon ${className}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] || icons.file}</svg>`;
}
const esc = escapeHtml;
const supportUrl = validateDonationUrl(donationUrl);
const supportLabel = typeof donationLabel === 'string' && donationLabel.trim() ? donationLabel.trim() : '请我喝杯咖啡';
const supportFooter = supportUrl ? `<div class="support-footer" id="support-footer"><a id="support-link" href="${esc(supportUrl)}" target="_blank" rel="noopener noreferrer">${esc(supportLabel)}</a><small>自愿支持，不影响免费功能</small></div>` : '';

document.querySelector('#app').innerHTML = `
  <a class="skip-link" href="#workspace">跳到工作台</a>
  <header class="site-header"><div class="header-inner">
    <a class="brand" href="#" aria-label="OfficeBridge 首页"><span class="brand-symbol">${icon('bridge')}</span><span>OfficeBridge<small>文档兼容工作台</small></span></a>
    <nav aria-label="主导航"><a class="nav-active" href="#workspace">工作台</a><a href="#verification">交接指南</a></nav>
    <span class="privacy-label">${icon('shield')} 文档仅在本机处理</span>
  </div></header>
  <main class="page-shell">
    <section class="hero" aria-labelledby="hero-title"><div><p class="eyebrow"><span class="eyebrow-line"></span> MAC <span class="direction">→</span> WINDOWS</p><h1 id="hero-title">好文档，不该困在一个系统。</h1><p class="hero-copy">检查演示文稿和表格的兼容风险，按你的选择生成可编辑副本。<br class="desktop-break">保留可编辑内容，让跨系统交接更有把握。</p></div><div class="hero-motif" aria-hidden="true"><span class="motif-doc">${icon('pptx')}</span><span class="motif-arrow">${icon('arrow')}</span><span class="motif-doc green">${icon('xlsx')}</span><span class="motif-caption">内容完整 · 继续编辑</span></div></section>
    <section class="upload-section" aria-labelledby="upload-title"><div id="drop-zone" class="drop-zone"><div class="upload-icon">${icon('upload')}</div><div class="upload-copy"><h2 id="upload-title">把文件拖到这里，先做一次检查</h2><p>PPTX / XLSX · 单个 100 MB · 每批最多 20 个 / 共 250 MB</p><span class="upload-detail">支持批量，依次在浏览器本机处理。不会上传文档。</span></div><div class="upload-action"><button class="button primary" id="choose-files">${icon('upload')} 选择文件</button><span>或直接拖放到这里</span></div><input type="file" id="file-input" accept=".pptx,.xlsx,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" multiple hidden></div><div class="demo-row"><span>还没有文件？用合成样例体验</span><button class="text-button" id="demo-pptx">演示文稿 ${icon('arrow')}</button><span class="demo-divider"></span><button class="text-button" id="demo-xlsx">电子表格 ${icon('arrow')}</button><span class="demo-note">样例仅用于展示检查流程</span></div></section>
    <section id="workspace" class="workspace" aria-label="文档工作台" tabindex="-1"><aside class="queue-panel"><div class="panel-label"><h2>文件列表 <span id="file-count">0</span></h2><button id="clear-files" class="subtle-button" hidden>清空</button></div><div id="file-list" class="file-list"></div><div class="queue-footer">${icon('shield')}<p>原文件始终保持不变。<br>关闭或刷新页面会清空本次工作台。</p></div></aside><div id="document-panel" class="document-panel" aria-live="off"></div></section>
    <section class="verification" id="verification" aria-labelledby="verification-title"><div class="verification-heading"><span class="section-kicker">最后一步，在目标电脑确认</span><h2 id="verification-title">自动检查之后，交接之前。</h2><p>结构检查能减少风险，但不能替代 Windows Office 实际打开验证。</p></div><div class="verification-steps"><div><span class="step-number">01</span><h3>打开修复副本</h3><p>用收件人实际使用的 Windows Office 或 WPS 版本打开文件。</p></div><div><span class="step-number">02</span><h3>检查外观与功能</h3><p>核对字体、换行和图表；试播动画，复核公式、日期及外部链接。</p></div><div><span class="step-number">03</span><h3>保存后再打开</h3><p>确认文件仍然可编辑。需要固定外观时，再附一份从原稿导出的 PDF。</p></div></div><details class="scope-details"><summary>这个工具可以处理什么？有哪些边界？</summary><div><p><strong>可以：</strong>检查当前规则覆盖的字体、对象、外链、公式及文档结构风险；经你确认后替换有明确映射建议的字体，并记录每一项修改。文档继续保留原有文本、公式和可编辑对象。</p><p><strong>仍需人工确认：</strong>字体替换可能改变换行和布局。字体可用性取决于目标电脑及 Office 安装；动画、宏、插件、链接、嵌入对象等不能靠此工具保证兼容。未知字体需由你指定目标字体，再确认是否采用替换。</p><p><strong>支持范围：</strong>未加密的标准 PPTX / XLSX。签名文件、宏文件、加密文件及存在危险结构的文件不会自动修复。Keynote、Numbers、Pages、旧版 PPT / XLS 需要先从原始软件导出。</p><p><strong>隐私：</strong>文件内容由浏览器本机处理，不传到服务器。报告可能包含文件名、字体和文档部件路径，分享前请检查。</p></div></details></section>
  </main>
  <footer class="site-footer"><span class="footer-brand">OfficeBridge</span><p>保留编辑能力，认真对待每一次交接。</p>${supportFooter}<span>结构检查 ≠ Windows 实测</span></footer>
  <div id="toast" class="toast" role="status" aria-live="polite" hidden></div>
  <div id="announcer" class="sr-only" role="status" aria-live="polite"></div>
`;

const fileInput = document.querySelector('#file-input');
document.querySelector('#choose-files').addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => { enqueue(fileInput.files); fileInput.value = ''; });
document.querySelector('#demo-pptx').addEventListener('click', () => loadDemo('pptx'));
document.querySelector('#demo-xlsx').addEventListener('click', () => loadDemo('xlsx'));
document.querySelector('#clear-files').addEventListener('click', () => {
  if (state.processing || state.records.some((record) => isDocumentBusy(record))) return;
  state.records = [];
  state.activeId = null;
  render();
  toast('本次文件列表已清空，原文件不受影响。');
});

const dropZone = document.querySelector('#drop-zone');
let dragDepth = 0;
window.addEventListener('dragover', (event) => {
  if (event.dataTransfer?.types.includes('Files')) event.preventDefault();
});
window.addEventListener('drop', (event) => {
  if (event.dataTransfer?.types.includes('Files')) event.preventDefault();
});
dropZone.addEventListener('dragenter', (event) => {
  if (!event.dataTransfer?.types.includes('Files')) return;
  event.preventDefault();
  dragDepth++;
  dropZone.classList.add('drag-active');
});
dropZone.addEventListener('dragover', (event) => { event.preventDefault(); event.dataTransfer.dropEffect = 'copy'; });
dropZone.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; dropZone.classList.remove('drag-active'); } });
dropZone.addEventListener('drop', (event) => {
  event.preventDefault(); dragDepth = 0; dropZone.classList.remove('drag-active'); enqueue(event.dataTransfer.files);
});

function toast(message, type = 'info') {
  const element = document.querySelector('#toast');
  if (!element) return;
  clearTimeout(state.toastTimer);
  element.textContent = message;
  element.className = `toast ${type}`;
  element.hidden = false;
  state.toastTimer = setTimeout(() => { element.hidden = true; }, 6500);
}
function announce(message) { document.querySelector('#announcer').textContent = message; }

function enqueue(files, demo = false) {
  const errors = [];
  const added = [];
  for (const file of Array.from(files)) {
    if (state.records.length >= MAX_FILES) { errors.push('每次工作台最多保留 20 个文件，请先清空已完成的列表。'); break; }
    if (!/\.(pptx|xlsx)$/i.test(file.name)) { errors.push(`${file.name}：目前只支持 PPTX 和 XLSX。`); continue; }
    if (file.size > MAX_FILE_SIZE) { errors.push(`${file.name}：超过 100 MB 上限。`); continue; }
    if (file.size === 0) { errors.push(`${file.name}：这是一个空文件。`); continue; }
    if (state.records.reduce((sum, item) => sum + item.file.size, 0) + file.size > MAX_TOTAL_SIZE) {
      errors.push(`${file.name}：当前批次总大小超过 250 MB，请先下载所需副本并清空列表。`); continue;
    }
    const record = { id: String(nextRecord++), file, name: file.name, format: file.name.split('.').pop().toLowerCase(), status: 'queued', report: null, selected: new Set(), repaired: null, error: null, demo, options: { fontProfile: 'windows-office', fontMappings: {} }, advancedOpen: false, reanalysisError: null };
    state.records.push(record);
    added.push(record);
  }
  if (errors.length) toast(errors.slice(0, 3).join(' '), 'error');
  if (added.length) {
    state.activeId = added[0].id;
    state.filter = 'all';
    announce(`已添加 ${added.length} 个文件，即将开始本机检查。`);
    render();
    processQueue();
  }
  return added;
}

async function processQueue() {
  if (state.processing) return;
  state.processing = true;
  let record;
  while ((record = state.records.find((item) => item.status === 'queued'))) {
    record.status = 'analyzing';
    renderRecord(record);
    try {
      const bytes = new Uint8Array(await record.file.arrayBuffer());
      record.report = await request('analyze', { bytes, fileName: record.name, options: record.options }, [bytes.buffer]);
      record.options = record.report.analysisOptions || record.options;
      record.status = 'ready';
      announce(`${record.name} 检查完成，发现 ${record.report.issues?.length || 0} 项待关注内容。`);
    } catch (error) {
      record.status = 'error'; record.error = error.message;
      announce(`${record.name} 无法检查：${error.message}`);
    }
    renderRecord(record);
  }
  state.processing = false;
  renderQueue();
}

async function loadDemo(format) {
  if (state.demoBusy) return;
  state.demoBusy = true;
  for (const button of document.querySelectorAll('.demo-row button')) button.disabled = true;
  try {
    const bytes = await request('demo', { format });
    const name = format === 'pptx' ? '跨系统交接示例.pptx' : '季度预算示例.xlsx';
    const added = enqueue([new File([bytes], name)], true);
    if (added.length) document.querySelector('#workspace').scrollIntoView({ behavior: reducedMotion() ? 'instant' : 'smooth', block: 'start' });
  } catch (error) { toast(error.message, 'error'); }
  finally {
    state.demoBusy = false;
    for (const button of document.querySelectorAll('.demo-row button')) button.disabled = false;
  }
}

function reducedMotion() { return window.matchMedia('(prefers-reduced-motion: reduce)').matches; }
function activeRecord() { return state.records.find((record) => record.id === state.activeId); }
function isDocumentBusy(record) { return ['repairing', 'reanalyzing'].includes(record.status); }
function fixableIssues(record) {
  if (!record.report || record.report.repairBlocked) return [];
  return (record.report.issues || []).filter((issue) => issue.fixable && issue.fix?.kind === 'font-replace');
}
function statusText(record) {
  return ({ queued: '等待检查', analyzing: '正在检查…', reanalyzing: '重新检查中…', repairing: '正在生成副本…', error: '无法检查' })[record.status]
    || (record.repaired ? '副本已生成' : record.report?.repairBlocked ? '仅可查看' : `${record.report?.issues?.length || 0} 项发现`);
}

function renderQueue() {
  document.querySelector('#file-count').textContent = state.records.length;
  const clear = document.querySelector('#clear-files');
  clear.hidden = !state.records.length;
  clear.disabled = state.processing || state.records.some((record) => isDocumentBusy(record));
  const list = document.querySelector('#file-list');
  list.innerHTML = state.records.length ? state.records.map((record) => `<button class="file-card ${record.id === state.activeId ? 'active' : ''}" data-file-id="${record.id}" aria-pressed="${record.id === state.activeId}"><span class="file-type ${record.format}">${icon(record.format)}</span><span class="file-card-content"><strong>${esc(record.name)}</strong><span>${formatBytes(record.file.size)} <span class="file-separator">·</span> ${esc(statusText(record))}</span>${record.demo ? '<small>合成样例</small>' : ''}</span>${['queued', 'analyzing', 'reanalyzing', 'repairing'].includes(record.status) ? '<span class="spinner small" aria-hidden="true"></span>' : record.repaired ? icon('check', 'queue-check') : ''}</button>`).join('') : `<div class="queue-empty">${icon('file')}<p>文件会显示在这里</p><span>添加文件，开始检查</span></div>`;
  list.querySelectorAll('[data-file-id]').forEach((button) => button.addEventListener('click', () => {
    state.activeId = button.dataset.fileId; state.filter = 'all'; render();
  }));
}

function render() { renderQueue(); renderDocument(); }
function renderRecord(record) {
  renderQueue();
  if (record.id === state.activeId) renderDocument();
}

function renderDocument() {
  const panel = document.querySelector('#document-panel');
  const record = activeRecord();
  if (!record) {
    panel.innerHTML = `<div class="welcome-panel"><span class="section-kicker">一个清楚、可追溯的交接流程</span><h2>先看风险，再决定怎么修。</h2><p class="welcome-copy">不把页面变成图片，不把公式变成数值。<br>你选择的修改，才会出现在新的文件里。</p><div class="process-list"><div><span class="process-icon">${icon('file')}</span><span><strong>01 / 检查文档</strong><p>识别字体、外链与功能兼容风险。</p></span></div><div><span class="process-icon">${icon('sparkle')}</span><span><strong>02 / 确认字体替换</strong><p>查看明确建议，自己选择是否应用。</p></span></div><div><span class="process-icon">${icon('download')}</span><span><strong>03 / 下载可编辑副本</strong><p>保留原件，附上修改记录与检查报告。</p></span></div></div><div class="welcome-note">${icon('info')}<p>不提供“兼容性分数”。每一项发现，都有具体原因与位置。</p></div></div>`;
    return;
  }
  const heading = `<div class="document-heading"><div class="document-title"><span class="file-type large ${record.format}">${icon(record.format)}</span><div><p class="section-kicker">${record.format === 'pptx' ? '演示文稿' : '电子表格'} ${record.demo ? '· 合成样例' : ''}</p><h2>${esc(record.name)}</h2><p class="document-meta">${formatBytes(record.file.size)} <span>·</span> 目标：Windows Office <span>·</span> 原件不修改</p></div></div></div>`;
  if (['queued', 'analyzing'].includes(record.status)) {
    panel.innerHTML = `${heading}<div class="processing-panel"><span class="spinner"></span><h3>${record.status === 'queued' ? '文件已加入检查队列' : '正在检查文档结构'}</h3><p>${record.status === 'queued' ? '正在依次处理文件，请稍候。' : '检查字体、引用关系和可能影响交接的功能。'}</p><span>文件留在本机，较大文件可能需要一点时间。</span></div>`;
    return;
  }
  if (record.status === 'error') {
    panel.innerHTML = `${heading}<div class="error-panel"><span class="error-icon">${icon('warning')}</span><h3>这个文件暂时无法检查</h3><p>${esc(record.error)}</p><div class="error-advice"><strong>建议从原始应用重新导出</strong><p>请确认文件未加密、没有损坏，并导出为标准 PPTX 或 XLSX。仅修改扩展名不能转换文件格式。</p></div></div>`;
    return;
  }
  const report = record.report;
  const issues = report.issues || [];
  const counts = { error: issues.filter((issue) => issue.severity === 'error').length, warning: issues.filter((issue) => issue.severity === 'warning').length, info: issues.filter((issue) => issue.severity === 'info').length };
  const fixable = fixableIssues(record);
  const inventory = report.inventory || {};
  const pages = record.format === 'pptx' ? inventory.slideCount : Array.isArray(inventory.worksheets) ? inventory.worksheets.length : inventory.worksheetCount;
  const fontsCount = Array.isArray(inventory.fonts) ? inventory.fonts.length : null;
  panel.innerHTML = `${heading}
    <div class="document-stats"><div><strong>${issues.length}</strong><span>检查发现</span></div><div><strong>${fixable.length}</strong><span>可选字体修复</span></div><div><strong>${pages ?? '—'}</strong><span>${record.format === 'pptx' ? '张幻灯片' : '个工作表'}</span></div><div><strong>${fontsCount ?? '—'}</strong><span title="包括主题中的备用声明，不代表都出现在可见文字中">声明字体</span></div></div>
    <div class="validation-note">${icon('info')}<p><strong>尚未在 Windows Office 实际验证</strong><span>这次检查基于文档结构，最终版式和功能仍需在目标电脑核对。</span></p></div>
    ${report.repairBlocked ? `<div class="blocked-notice">${icon('shield')}<div><strong>此文件不适合自动修复</strong><p>为保护文件完整性，已关闭修复功能。请查看下方发现，并在原始软件中处理。</p></div></div>` : ''}
    ${renderFontSettings(record)}
    <div class="findings-toolbar"><h3>检查发现</h3><div class="filter-tabs" role="group" aria-label="筛选检查发现">${[['all', '全部', issues.length], ['error', '需处理', counts.error], ['warning', '需复核', counts.warning], ['info', '提示', counts.info]].map(([value, label, count]) => `<button data-filter="${value}" class="filter-button ${state.filter === value ? 'selected' : ''}" aria-pressed="${state.filter === value}">${label}<span>${count}</span></button>`).join('')}</div></div>
    <div id="issue-list" class="issue-list"></div>
    <div class="repair-panel"><div class="repair-panel-heading"><span class="repair-symbol">${icon('sparkle')}</span><div><h3>生成可编辑副本</h3><p>仅应用你确认的字体替换，保留文档内容和其他对象。</p></div></div><p class="font-caution">字体替换可能改变换行和排版。建议的字体也需要目标电脑具备，请确认后再选择。</p><div class="repair-selection"><span id="selection-summary"></span><div><button id="select-fixes" class="text-button" ${!fixable.length || isDocumentBusy(record) ? 'disabled' : ''}>选择全部字体建议</button><button id="clear-fixes" class="subtle-button" ${!record.selected.size || isDocumentBusy(record) ? 'disabled' : ''}>取消选择</button></div></div><button id="repair-file" class="button primary repair-button" ${!record.selected.size || report.repairBlocked || isDocumentBusy(record) ? 'disabled' : ''}>${record.status === 'repairing' ? '<span class="spinner small light"></span> 正在生成副本…' : `${icon('sparkle')} 生成修复副本`}</button>${!fixable.length ? `<p class="no-repair-note">${report.repairBlocked ? '当前文件已禁止自动修复。' : '当前没有可自动修复的字体建议；仍可导出检查报告。'}</p>` : ''}<div id="repair-result"></div></div>
    <div class="report-export"><div><h3>带上这份检查记录</h3><p>报告包含发现、位置与修改记录，便于复核和交接。</p></div><div><button class="button secondary compact" id="download-html">${icon('download')} HTML 报告</button><button class="button secondary compact" id="download-json">JSON</button></div></div>
    <details class="technical-details"><summary>文件校验与检查说明</summary><dl><div><dt>原文件 SHA-256</dt><dd class="hash">${esc(report.sha256 || '未提供')}</dd></div><div><dt>工具版本</dt><dd>${esc(report.toolVersion || '开发版')}</dd></div><div><dt>结构检查</dt><dd>${report.validation?.packageChecked ? '已执行' : '未完整执行'}</dd></div></dl>${(report.validation?.notes || []).length ? `<ul>${report.validation.notes.map((note) => `<li>${esc(note)}</li>`).join('')}</ul>` : ''}</details>`;
  panel.querySelectorAll('[data-filter]').forEach((button) => button.addEventListener('click', () => {
    state.filter = button.dataset.filter;
    panel.querySelectorAll('[data-filter]').forEach((item) => { item.classList.toggle('selected', item === button); item.setAttribute('aria-pressed', String(item === button)); });
    renderIssues(record);
  }));
  panel.querySelector('#select-fixes').addEventListener('click', () => {
    record.selected = new Set(fixable.map((issue) => issue.id)); record.repaired = null;
    renderIssues(record); updateSelection(record); announce(`已选择 ${record.selected.size} 项字体建议，点击生成按钮才会修改副本。`);
  });
  panel.querySelector('#clear-fixes').addEventListener('click', () => { record.selected.clear(); record.repaired = null; renderIssues(record); updateSelection(record); });
  panel.querySelector('#repair-file').addEventListener('click', () => repair(record));
  panel.querySelector('#download-json').addEventListener('click', () => {
    downloadBlob(JSON.stringify(exportPayload(record), null, 2), `${record.name}.兼容报告.json`, 'application/json;charset=utf-8');
    toast('JSON 报告已准备下载。');
  });
  panel.querySelector('#download-html').addEventListener('click', () => {
    downloadBlob(buildHtmlReport(record), `${record.name}.兼容报告.html`, 'text/html;charset=utf-8');
    toast('HTML 报告已准备下载，可以独立打开或打印。');
  });
  bindFontSettings(record, panel);
  panel.querySelector('#download-html').disabled = isDocumentBusy(record);
  panel.querySelector('#download-json').disabled = isDocumentBusy(record);
  renderIssues(record);
  updateSelection(record);
}

function renderIssues(record) {
  const order = { error: 0, warning: 1, info: 2 };
  const issues = (record.report.issues || []).filter((issue) => state.filter === 'all' || issue.severity === state.filter)
    .sort((a, b) => (order[a.severity] ?? 2) - (order[b.severity] ?? 2));
  const list = document.querySelector('#issue-list');
  list.innerHTML = issues.length ? issues.map((issue) => {
    const fixable = !record.report.repairBlocked && issue.fixable && issue.fix?.kind === 'font-replace';
    const severity = ['error', 'warning', 'info'].includes(issue.severity) ? issue.severity : 'info';
    return `<article class="issue-card ${severity}"><div class="issue-main"><span class="issue-status">${icon(severity === 'info' ? 'info' : 'warning')}</span><div class="issue-content"><div class="issue-title"><h4>${esc(issue.title)}</h4><span class="severity-label">${({ error: '需处理', warning: '需复核', info: '提示' })[severity]}</span></div><p class="issue-description">${esc(issue.description)}</p><p class="issue-location"><span>位置</span>${esc(issue.location || issue.part || '文件整体')}</p>${fixable ? `<label class="fix-option"><input type="checkbox" data-issue-id="${esc(issue.id)}" ${record.selected.has(issue.id) ? 'checked' : ''} ${isDocumentBusy(record) ? 'disabled' : ''}><span><span class="font-map"><strong>${esc(issue.fix.from)}</strong>${icon('arrow')}<strong>${esc(issue.fix.to)}</strong></span><small>选择后应用到副本 · 字体变化可能影响排版</small></span></label>` : ''}${issue.recommendation && !fixable ? `<p class="issue-recommendation"><span>建议</span>${esc(issue.recommendation)}</p>` : ''}${issue.evidence || issue.part ? `<details class="issue-detail"><summary>查看检查依据</summary>${issue.evidence ? `<p>${esc(issue.evidence)}</p>` : ''}${issue.part ? `<code>${esc(issue.part)}</code>` : ''}<span class="issue-code">${esc(issue.code)}</span></details>` : ''}</div></div></article>`;
  }).join('') : `<div class="no-issues">${icon('check')}<h4>${state.filter === 'all' ? '没有发现当前规则覆盖的风险' : '这个级别没有检查发现'}</h4><p>${state.filter === 'all' ? '这不代表已保证兼容。交接前请在目标 Windows 电脑验证。' : '可以切换筛选，查看其他发现。'}</p></div>`;
  list.querySelectorAll('[data-issue-id]').forEach((input) => input.addEventListener('change', () => {
    if (input.checked) record.selected.add(input.dataset.issueId); else record.selected.delete(input.dataset.issueId);
    record.repaired = null; updateSelection(record);
  }));
}

function updateSelection(record) {
  const summary = document.querySelector('#selection-summary');
  if (!summary) return;
  summary.innerHTML = `已选择 <strong>${record.selected.size}</strong> 项 / 共 ${fixableIssues(record).length} 项建议`;
  document.querySelector('#repair-file').disabled = !record.selected.size || record.report.repairBlocked || isDocumentBusy(record);
  document.querySelector('#clear-fixes').disabled = !record.selected.size || isDocumentBusy(record);
  const result = document.querySelector('#repair-result');
  if (!record.repaired) { result.innerHTML = ''; return; }
  const total = (record.repaired.changes || []).reduce((sum, change) => sum + (change.count || 0), 0);
  const unresolved = record.repaired.report?.issues?.length || 0;
  result.innerHTML = `<div class="repair-success"><div class="repair-success-heading">${icon('check')}<strong>可编辑副本已生成</strong></div><p>记录了 ${record.repaired.changes?.length || 0} 项修改，共 ${total} 处替换。复查仍有 ${unresolved} 项发现，请查看报告并在 Windows 上验证。</p><p class="repaired-name">${esc(record.repaired.fileName)}</p><button class="button success" id="download-repaired">${icon('download')} 下载修复副本</button><details class="changes-details"><summary>查看本次修改记录</summary><ul>${(record.repaired.changes || []).map((change) => `<li><strong>${esc(change.before)} → ${esc(change.after)}</strong><span>${esc(change.part)} · ${change.count || 0} 处</span></li>`).join('')}</ul></details></div>`;
  document.querySelector('#download-repaired').addEventListener('click', () => {
    downloadBlob(record.repaired.bytes, safeFileName(record.repaired.fileName), record.format === 'pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    toast('修复副本已准备下载。请在目标 Windows 软件中复核。');
  });
}

async function repair(record) {
  if (!record.selected.size || record.report.repairBlocked || isDocumentBusy(record)) return;
  record.status = 'repairing';
  renderRecord(record);
  try {
    const bytes = new Uint8Array(await record.file.arrayBuffer());
    record.repaired = await request('repair', { bytes, fileName: record.name, options: { ...record.options, selectedIssueIds: [...record.selected] } }, [bytes.buffer]);
    record.status = 'ready';
    announce(`${record.name} 的修复副本已生成，可以下载。原文件保持不变。`);
    toast('副本已生成，原文件保持不变。');
  } catch (error) {
    record.status = 'ready'; record.repaired = null;
    toast(`没有生成副本：${error.message}`, 'error');
    announce(`修复未完成：${error.message}`);
  }
  renderRecord(record);
  if (record.id === state.activeId) document.querySelector(record.repaired ? '#download-repaired' : '#repair-file')?.focus({ preventScroll: true });
}

function availableFontNames(record) {
  return [...new Set((record.report?.inventory?.fonts || [])
    .map((font) => typeof font === 'string' ? font : font.name)
    .filter((name) => typeof name === 'string' && name.trim() && !/^\+(?:mj|mn)-/i.test(name.trim())))];
}

function renderFontSettings(record) {
  const fonts = availableFontNames(record);
  const mappings = Object.entries(record.options?.fontMappings || {});
  const disabled = isDocumentBusy(record) || record.report.repairBlocked;
  return `<details class="font-settings" id="font-settings" ${record.advancedOpen ? 'open' : ''}>
    <summary>高级：自定义字体替换 <span>${mappings.length ? `${mappings.length} 条当前文件规则` : '按需设置'}</span></summary>
    <div class="font-settings-body">
      <p>仅对当前文件生效。目标电脑必须已经安装所填字体；替换可能改变字宽、换行和排版。下方常见字体只是候选，不能证明目标电脑具备。</p>
      <p class="font-settings-caution">此处不适用于符号、数学或图标字体。普通字体替换可能改变字符含义，请在原始应用中处理。</p>
      ${record.report.repairBlocked ? '<p class="font-settings-error">当前文件禁止修复，不能添加字体规则。</p>' : ''}
      ${record.reanalysisError ? `<p class="font-settings-error" role="alert">规则未应用：${esc(record.reanalysisError)} 已保留上次有效的检查结果。</p>` : ''}
      <form id="font-rule-form"><fieldset ${disabled ? 'disabled' : ''}><div class="font-rule-fields">
        <label for="font-source">原字体<select id="font-source" name="source" required><option value="">选择文档中的字体</option>${fonts.map((font) => `<option value="${esc(font)}">${esc(font)}</option>`).join('')}</select></label>
        <span class="font-rule-arrow" aria-hidden="true">${icon('arrow')}</span>
        <label for="font-target">目标字体<input id="font-target" name="target" type="text" list="font-target-candidates" maxlength="128" required placeholder="例如 Microsoft YaHei" autocomplete="off"></label>
        <datalist id="font-target-candidates">${['Microsoft YaHei', 'Arial', 'SimSun', 'Microsoft JhengHei', 'Calibri', 'Times New Roman', 'Consolas', 'KaiTi'].map((font) => `<option value="${font}"></option>`).join('')}</datalist>
      </div><button class="button secondary compact" id="add-font-rule" type="submit" ${!fonts.length ? 'disabled' : ''}>${record.status === 'reanalyzing' ? '<span class="spinner small"></span> 正在重新检查…' : '添加规则并重新检查'}</button><p class="font-rule-help">同一原字体会更新现有规则。重新检查会清空已选项和旧副本；之后仍需勾选并生成，才会修改新副本。</p></fieldset></form>
      <div id="font-rules" class="font-rules">${mappings.length ? mappings.map(([from, to]) => `<div class="font-rule"><span><strong>${esc(from)}</strong>${icon('arrow')}<strong>${esc(to)}</strong></span><button type="button" class="subtle-button" data-remove-font="${esc(from)}" aria-label="移除 ${esc(from)} 替换规则" ${disabled ? 'disabled' : ''}>移除并重新检查</button></div>`).join('') : '<p class="font-rules-empty">尚未设置自定义规则，当前使用内置建议。</p>'}</div>
    </div></details>`;
}

function bindFontSettings(record, panel) {
  const details = panel.querySelector('#font-settings');
  details.addEventListener('toggle', () => { record.advancedOpen = details.open; });
  panel.querySelector('#font-rule-form').addEventListener('submit', (event) => {
    event.preventDefault();
    if (isDocumentBusy(record) || record.report.repairBlocked) return;
    const from = panel.querySelector('#font-source').value;
    const to = panel.querySelector('#font-target').value.trim();
    if (!availableFontNames(record).includes(from) || !to) return;
    if (from.trim().toLowerCase() === to.toLowerCase()) { toast('原字体与目标字体相同，无需添加替换规则。', 'error'); return; }
    const mappings = { ...(record.options?.fontMappings || {}), [from]: to };
    reanalyzeWithMappings(record, mappings);
  });
  panel.querySelectorAll('[data-remove-font]').forEach((button) => button.addEventListener('click', () => {
    if (isDocumentBusy(record) || record.report.repairBlocked) return;
    const mappings = Object.fromEntries(Object.entries(record.options?.fontMappings || {}).filter(([from]) => from !== button.dataset.removeFont));
    reanalyzeWithMappings(record, mappings);
  }));
}

async function reanalyzeWithMappings(record, fontMappings) {
  if (isDocumentBusy(record) || record.report.repairBlocked) return;
  record.status = 'reanalyzing';
  record.advancedOpen = true;
  record.reanalysisError = null;
  record.repaired = null;
  record.selected.clear();
  state.filter = 'all';
  renderRecord(record);
  const options = { ...record.options, fontMappings };
  try {
    const bytes = new Uint8Array(await record.file.arrayBuffer());
    const report = await request('analyze', { bytes, fileName: record.name, options }, [bytes.buffer]);
    record.report = report;
    record.options = report.analysisOptions || options;
    announce(`${record.name} 已按当前字体规则重新检查。尚未选择或应用任何替换。`);
    toast('字体规则已更新并重新检查，请确认后选择需要应用的替换。');
  } catch (error) {
    record.reanalysisError = error.message;
    announce(`字体规则没有应用：${error.message}`);
  }
  record.status = 'ready';
  renderRecord(record);
  if (record.id === state.activeId) document.querySelector('#font-settings summary')?.focus({ preventScroll: true });
}

render();
