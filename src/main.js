import './style.css';
import { fontPairHint, unsupportedFileMessage } from './ui/input-help.js';
import { bindFontList, renderFontList } from './ui/font-list.js';
import { listedFont } from './engine/target-fonts.js';
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
let fileAction = 0;
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
  <a class="skip-link" href="#choose-files">跳到选择文件</a>
  <header class="site-header"><div class="header-inner">
    <a class="brand" href="#start" aria-label="OfficeBridge 首页"><span class="brand-symbol">${icon('bridge')}</span><span>OfficeBridge</span></a>
    <nav aria-label="主导航"><a href="#verification" id="open-help">使用帮助</a></nav>
    <span class="privacy-label">${icon('shield')} 文件只在本机处理</span>
  </div></header>
  <main class="page-shell" id="start">
    <section class="hero" aria-labelledby="hero-title"><div><h1 id="hero-title">检查发到 Windows 的文件</h1><p class="hero-copy">看看字体、排版和播放有哪些地方要留意。需要换字体时，可另存一份可编辑的副本。</p></div></section>
    <section class="upload-section" aria-labelledby="upload-title">
      <div id="drop-zone" class="drop-zone"><div class="upload-icon">${icon('upload')}</div>
        <div class="upload-copy"><h2 id="upload-title">把演示文稿或表格拖到这里</h2><p>PowerPoint（.pptx）或 Excel（.xlsx），可一次选多个</p></div>
        <div class="upload-action"><button class="button primary" id="choose-files">${icon('file')} 选择文件</button></div>
        <input type="file" id="file-input" accept=".pptx,.xlsx,.key,.numbers,.ppt,.xls,.pptm,.xlsm,.pdf,.csv,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" multiple hidden>
      </div>
      <div class="demo-row"><span>想先试试看？</span><button class="text-button" id="demo-pptx">${icon('pptx')} 演示文稿样例</button><span class="demo-divider"></span><button class="text-button" id="demo-xlsx">${icon('xlsx')} 表格样例</button></div>
      <div class="input-errors" id="input-errors" role="group" aria-labelledby="input-errors-title" tabindex="-1" hidden><div><strong id="input-errors-title">有文件未能添加</strong><ul></ul></div><button id="dismiss-input-errors" class="subtle-button" aria-label="关闭添加文件提示">${icon('close')}</button></div>
    </section>
    <section id="workspace" class="workspace" aria-label="检查结果" tabindex="-1" hidden>
      <aside class="queue-panel"><div class="panel-label"><h2>本次文件 <span id="file-count">0</span></h2><button id="clear-files" class="subtle-button">清空列表</button></div>
        <div id="file-list" class="file-list"></div><div class="queue-footer">${icon('shield')}<p>关闭页面会清空列表。请先下载需要的副本。</p></div></aside>
      <div id="document-panel" class="document-panel" aria-live="off"></div>
    </section>
    <details class="verification" id="verification"><summary>第一次用？查看操作指南</summary>
      <div class="verification-steps">
        <div><h3>从 Mac 导出</h3><p>在 Keynote 或 Numbers 里打开原稿，选择“文件”→“导出为”，选 PowerPoint 或 Excel。已有 PPTX 或 XLSX 可以直接拖进来。</p></div>
        <div><h3>决定要不要换字体</h3><p>先在 Windows 打开原件。文字和分页都正常时，直接用原件。有错位时，再到检查结果里勾选相应字体，下载副本。</p></div>
        <div><h3>打开下载的副本</h3><p>在浏览器的下载记录里找到文件名带 .windows 的副本。在 Windows 试播幻灯片；表格看打印预览，再改一个单元格、保存并重开。</p></div>
      </div><details class="scope-details"><summary>支持的文件与隐私</summary><div>
        <p>每份文件最多 100 MB，本次最多 20 份、合计 250 MB。检查支持标准 PPTX 和 XLSX。受保护或含宏的文件无法生成字体替换副本。</p>
        <p>浏览器在本机检查文件结构，字体替换可能改变换行和分页。实际效果请在接收方电脑上核对。下载的报告可能包含文件名、文字片段或链接，分享前请先查看。</p>
      </div></details>
    </details>
  </main>
  ${supportFooter ? `<footer class="site-footer">${supportFooter}</footer>` : ''}
  <div id="toast" class="toast" role="status" aria-live="polite" hidden></div>
  <div id="announcer" class="sr-only" role="status" aria-live="polite"></div>
`;

const fileInput = document.querySelector('#file-input');
const workspace = document.querySelector('#workspace');
const inputErrors = document.querySelector('#input-errors');
document.querySelector('#open-help').addEventListener('click', () => {
  const help = document.querySelector('#verification');
  help.open = true;
  help.querySelector('summary').focus({ preventScroll: true });
});
document.querySelector('#dismiss-input-errors').addEventListener('click', () => {
  inputErrors.hidden = true;
  document.querySelector('#choose-files').focus();
});
document.querySelector('#choose-files').addEventListener('click', () => {
  fileAction++;
  fileInput.click();
});
fileInput.addEventListener('change', () => { enqueue(fileInput.files); fileInput.value = ''; });
document.querySelector('#demo-pptx').addEventListener('click', () => loadDemo('pptx'));
document.querySelector('#demo-xlsx').addEventListener('click', () => loadDemo('xlsx'));
document.querySelector('#clear-files').addEventListener('click', () => {
  if (state.processing || state.records.some((record) => isDocumentBusy(record))) return;
  fileAction++;
  state.records = [];
  state.activeId = null;
  inputErrors.hidden = true;
  render();
  document.querySelector('#choose-files').focus();
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
  if (!demo) fileAction++;
  const errors = [];
  const added = [];
  for (const file of Array.from(files)) {
    if (state.records.length >= MAX_FILES) { errors.push('每次工作台最多保留 20 个文件，请先清空已完成的列表。'); break; }
    if (!/\.(pptx|xlsx)$/i.test(file.name)) { errors.push(unsupportedFileMessage(file.name)); continue; }
    if (file.size > MAX_FILE_SIZE) { errors.push(`${file.name}：超过 100 MB 上限。`); continue; }
    if (file.size === 0) { errors.push(`${file.name}：这是一个空文件。`); continue; }
    if (state.records.reduce((sum, item) => sum + item.file.size, 0) + file.size > MAX_TOTAL_SIZE) {
      errors.push(`${file.name}：当前批次总大小超过 250 MB，请先下载所需副本并清空列表。`); continue;
    }
    const record = { id: String(nextRecord++), file, name: file.name, format: file.name.split('.').pop().toLowerCase(), status: 'queued', report: null, selected: new Set(), repaired: null, error: null, demo, options: { fontProfile: 'windows-office', fontMappings: {} }, advancedOpen: false, reanalysisError: null };
    state.records.push(record);
    added.push(record);
  }
  inputErrors.hidden = !errors.length;
  if (errors.length) {
    inputErrors.querySelector('ul').innerHTML = errors.map(message => `<li>${esc(message)}</li>`).join('');
    toast(errors[0], 'error');
  }
  if (added.length) {
    state.activeId = added[0].id;
    state.filter = 'all';
    announce(`已添加 ${added.length} 个文件，即将开始本机检查。`);
    render();
    processQueue();
  }
  if (errors.length || added.length) {
    const destination = errors.length ? inputErrors : workspace;
    destination.focus({ preventScroll: true });
    destination.scrollIntoView({ block: 'start' });
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
  const action = ++fileAction;
  for (const button of document.querySelectorAll('.demo-row button')) button.disabled = true;
  try {
    const bytes = await request('demo', { format });
    // A newer file choice or a cleared queue takes precedence over this sample.
    if (action !== fileAction) return;
    const name = format === 'pptx' ? '跨系统交接示例.pptx' : '季度预算示例.xlsx';
    enqueue([new File([bytes], name)], true);
  } catch (error) {
    if (action === fileAction) toast(error.message, 'error');
  } finally {
    state.demoBusy = false;
    for (const button of document.querySelectorAll('.demo-row button')) button.disabled = false;
  }
}

function activeRecord() { return state.records.find((record) => record.id === state.activeId); }
function isDocumentBusy(record) { return ['repairing', 'reanalyzing'].includes(record.status); }
function fixableIssues(record) {
  if (!record.report || record.report.repairBlocked) return [];
  return (record.report.issues || []).filter((issue) => issue.fixable && issue.fix?.kind === 'font-replace');
}
function statusText(record) {
  return ({ queued: '等待检查', analyzing: '正在检查…', reanalyzing: '重新检查中…', repairing: '正在生成副本…', error: '无法检查' })[record.status]
    || (record.repaired ? '副本已生成' : record.report?.repairBlocked ? '仅可检查' : '检查完成');
}

function renderQueue() {
  document.querySelector('#file-count').textContent = state.records.length;
  const clear = document.querySelector('#clear-files');
  clear.hidden = !state.records.length;
  clear.disabled = state.processing || state.records.some((record) => isDocumentBusy(record));
  const list = document.querySelector('#file-list');
  const focusedId = list.contains(document.activeElement) ? document.activeElement.dataset.fileId : null;
  list.innerHTML = state.records.length ? state.records.map((record) => `<button class="file-card ${record.id === state.activeId ? 'active' : ''}" data-file-id="${record.id}" aria-pressed="${record.id === state.activeId}"><span class="file-type ${record.format}">${icon(record.format)}</span><span class="file-card-content"><strong>${esc(record.name)}</strong><span>${formatBytes(record.file.size)} <span class="file-separator">·</span> ${esc(statusText(record))}</span>${record.demo ? '<small>合成样例</small>' : ''}</span>${['queued', 'analyzing', 'reanalyzing', 'repairing'].includes(record.status) ? '<span class="spinner small" aria-hidden="true"></span>' : record.repaired ? icon('check', 'queue-check') : ''}</button>`).join('') : '';
  if (focusedId) list.querySelector(`[data-file-id="${focusedId}"]`)?.focus({ preventScroll: true });
  list.querySelectorAll('[data-file-id]').forEach((button) => button.addEventListener('click', () => {
    fileAction++;
    state.activeId = button.dataset.fileId; state.filter = 'all'; render();
    document.querySelector('.document-title h2')?.focus();
  }));
}

function render() {
  workspace.hidden = !state.records.length;
  document.body.classList.toggle('has-files', !workspace.hidden);
  renderQueue();
  renderDocument();
}
function renderRecord(record) {
  renderQueue();
  if (record.id === state.activeId) renderDocument();
}

function renderSummary(record, fixable, manual) {
  const blocked = record.report.repairBlocked;
  let outcome = 'original', title = '先用原文件打开看看', text = '这次检查没有建议替换的字体。请在 Windows 打开原件核对显示和编辑，正常就可以直接用。';
  if (blocked) {
    outcome = 'blocked'; title = '这个文件只能检查';
    text = '请先在制作原稿的应用中处理下面的问题，再重新添加文件。';
  } else if (manual.length) {
    outcome = 'review'; title = '先在 Windows 核对原件';
    text = fixable.length ? '下面的内容需要实际打开查看。如果还遇到字体错位，可以选择字体做一份副本。' : '下面的内容需要实际打开查看。这次检查没有可供替换的字体。';
  } else if (fixable.length) {
    outcome = 'optional'; title = '原件显示正常，就可以直接用';
    text = '如果在 Windows 看到文字错位或换行变化，可以试试下面的字体替换。';
  }
  return `<div class="result-summary ${blocked ? 'blocked blocked-notice' : manual.length ? 'attention' : ''}" data-outcome="${outcome}">
    <span class="summary-icon">${icon(blocked ? 'shield' : manual.length ? 'warning' : 'check')}</span><div>
    <h3>${title}</h3><p>${text}</p>
    ${manual.length ? `<ul class="review-items">${manual.slice(0, 2).map(i => `<li>${esc(i.title)}</li>`).join('')}</ul><button class="text-button" id="open-findings">查看${manual.length > 2 ? '全部' : ''}核对说明 ${icon('arrow')}</button>` : ''}
    </div></div>`;
}

function choiceLocation(issue) {
  const name = issue.location || issue.part || '文件整体';
  if (name === 'xl/styles.xml') return '单元格样式';
  if (name === 'xl/sharedStrings.xml') return '单元格文字';
  if (/^(xl|ppt)\/theme\//.test(name)) return '文档主题';
  return name;
}

function fontChoices(issues) {
  const pairs = new Map();
  for (const issue of issues) {
    const key = JSON.stringify([issue.fix.from, issue.fix.to]);
    if (!pairs.has(key)) pairs.set(key, []);
    pairs.get(key).push(issue);
  }
  return [...pairs.values()];
}

function renderChoices(record, groups) {
  const disabled = isDocumentBusy(record) ? 'disabled' : '';
  function input(issue) {
    return `<input type="checkbox" data-issue-id="${esc(issue.id)}" aria-label="${esc(issue.title)}，${esc(choiceLocation(issue))}，替换为 ${esc(issue.fix.to)}" ${record.selected.has(issue.id) ? 'checked' : ''} ${disabled}>`;
  }
  return groups.map((group, index) => {
    const issue = group[0];
    const hint = fontPairHint(issue.fix.from, issue.fix.to);
    const map = `<span class="font-map"><strong>${esc(issue.fix.from)}</strong>${icon('arrow')}<strong>${esc(issue.fix.to)}</strong></span>`;
    const isGroup = group.length > 1;
    const location = isGroup ? `${group.length} 个位置` : esc(choiceLocation(issue));
    const descriptions = [...new Set(group.filter(i => i.fix.affectsDefaultFont || i.fix.affectsWrappedRows).map(i => i.description))];
    const examples = group.filter(i => i.fix.affectsWrappedRows && i.evidence).map(i => i.evidence);
    const note = descriptions.length ? `<p class="choice-note">${descriptions.map(esc).join(' ')}${examples.length ? ' ' + esc(examples[0]) : ''}</p>` : '';
    const control = isGroup ? `<input type="checkbox" data-font-group="${index}" aria-label="将 ${esc(issue.fix.from)} 替换为 ${esc(issue.fix.to)}，共 ${group.length} 个位置" ${disabled}>` : input(issue);
    return `<article class="font-choice issue-card${isGroup ? ' font-choice-group' : ''}"><label class="fix-option">${control}<span>${map}${hint ? `<small class="choice-hint">${esc(hint)}</small>` : ''}<small class="choice-location">${location}</small></span></label>${note}
      ${isGroup ? `<details class="font-places"><summary>逐处选择</summary>${group.map(i => `<label class="fix-option">${input(i)}<span>${esc(choiceLocation(i))}</span></label>`).join('')}</details>` : ''}</article>`;
  }).join('');
}

function renderDocument() {
  const panel = document.querySelector('#document-panel');
  const record = activeRecord();
  if (!record) {
    panel.replaceChildren();
    return;
  }
  const heading = `<div class="document-heading"><div class="document-title"><span class="file-type large ${record.format}">${icon(record.format)}</span><div><p class="section-kicker">${record.format === 'pptx' ? '演示文稿' : '电子表格'} ${record.demo ? '· 合成样例' : ''}</p><h2 tabindex="-1">${esc(record.name)}</h2><p class="document-meta">${formatBytes(record.file.size)} <span>·</span> 原文件</p></div></div></div>`;
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
  const groups = fontChoices(fixable);
  const manual = issues.filter(issue => !issue.fixable && (issue.severity === 'error' || issue.severity === 'warning'));
  manual.sort((a, b) => Number(b.severity === 'error') - Number(a.severity === 'error'));
  const busy = isDocumentBusy(record);
  const inventory = report.inventory || {};
  const pages = record.format === 'pptx' ? inventory.slideCount : inventory.worksheets?.length;
  const fontsCount = inventory.fonts?.length;
  panel.innerHTML = `${heading}
    ${renderSummary(record, fixable, manual)}
    <section class="repair-panel" ${groups.length ? '' : 'hidden'} aria-labelledby="repair-title">
      <div class="repair-panel-heading"><div><h3 id="repair-title">选择要替换的字体</h3><p>只勾选遇到问题的字体。修改会保存到新文件。</p></div></div>
      <div class="font-choice-list">${renderChoices(record, groups)}</div>
      <div class="repair-selection"><span id="selection-summary"></span><div><button id="select-fixes" class="text-button" ${!fixable.length || busy ? 'disabled' : ''}>全选</button><button id="clear-fixes" class="subtle-button" disabled>取消选择</button></div></div>
      <p id="font-caution" class="font-caution"></p>
      <button id="repair-file" class="button primary repair-button" disabled>${record.status === 'repairing' ? '<span class="spinner small light"></span> 正在生成副本…' : `${icon('download')} 生成并下载副本`}</button>
      <div id="repair-result"></div>
    </section>
    ${renderFontSettings(record)}
    <details class="findings-details" ${record.findingsOpen ? 'open' : ''}><summary>查看全部检查结果 <span>${issues.length} 项</span></summary>
      <div class="findings-toolbar"><div class="filter-tabs" role="group" aria-label="筛选检查发现">${[['all', '全部', issues.length], ['error', '需处理', counts.error], ['warning', '需复核', counts.warning], ['info', '提示', counts.info]].map(([value, label, count]) => `<button data-filter="${value}" class="filter-button ${state.filter === value ? 'selected' : ''}" aria-pressed="${state.filter === value}">${label}<span>${count}</span></button>`).join('')}</div></div>
      <div id="issue-list" class="issue-list"></div>
    </details>
    <details class="report-details" ${record.exportOpen ? 'open' : ''}><summary>下载检查报告</summary>
      <div class="report-export"><div><p>报告包含核对建议和已采用的修改，可另存后查看或打印。</p></div><div><button class="button secondary compact" id="download-html">${icon('download')} 下载可阅读的报告</button><button class="button secondary compact" id="download-json">下载 JSON</button></div></div>
      <details class="technical-details"><summary>文件信息与检查范围</summary><dl><div><dt>${record.format === 'pptx' ? '幻灯片' : '工作表'} / 声明字体</dt><dd>${pages ?? '—'} / ${fontsCount ?? '—'}</dd></div><div><dt>原文件 SHA-256</dt><dd class="hash">${esc(report.sha256 || '未提供')}</dd></div><div><dt>工具版本</dt><dd>${esc(report.toolVersion || '开发版')}</dd></div><div><dt>结构检查</dt><dd>${report.validation?.packageChecked ? '已执行' : '未完整执行'}</dd></div></dl>${(report.validation?.notes || []).length ? `<ul>${report.validation.notes.map(note => `<li>${esc(note)}</li>`).join('')}</ul>` : ''}</details>
    </details>`;
  const findings = panel.querySelector('.findings-details');
  const exports = panel.querySelector('.report-details');
  findings.addEventListener('toggle', () => { if (findings.isConnected) record.findingsOpen = findings.open; });
  exports.addEventListener('toggle', () => { if (exports.isConnected) record.exportOpen = exports.open; });
  panel.querySelector('#open-findings')?.addEventListener('click', () => {
    findings.open = true;
    panel.querySelector(`[data-filter="${report.repairBlocked ? 'error' : 'warning'}"]`).click();
    findings.querySelector('summary').focus();
  });
  panel.querySelectorAll('[data-filter]').forEach((button) => button.addEventListener('click', () => {
    state.filter = button.dataset.filter;
    panel.querySelectorAll('[data-filter]').forEach((item) => { item.classList.toggle('selected', item === button); item.setAttribute('aria-pressed', String(item === button)); });
    renderIssues(record);
  }));
  function select(ids, checked) {
    for (const id of ids) {
      if (checked) record.selected.add(id); else record.selected.delete(id);
    }
    record.repaired = null;
    record.repairError = null;
    updateSelection(record, groups);
    renderQueue();
  }
  panel.querySelectorAll('[data-issue-id]').forEach(check => check.addEventListener('change', () => select([check.dataset.issueId], check.checked)));
  panel.querySelectorAll('[data-font-group]').forEach(check => check.addEventListener('change', () => select(groups[Number(check.dataset.fontGroup)].map(i => i.id), check.checked)));
  panel.querySelector('#select-fixes').addEventListener('click', () => select(fixable.map(i => i.id), true));
  panel.querySelector('#clear-fixes').addEventListener('click', () => select([...record.selected], false));
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
  updateSelection(record, groups);
}

function renderIssues(record) {
  const order = { error: 0, warning: 1, info: 2 };
  const issues = (record.report.issues || []).filter((issue) => state.filter === 'all' || issue.severity === state.filter)
    .sort((a, b) => (order[a.severity] ?? 2) - (order[b.severity] ?? 2));
  const list = document.querySelector('#issue-list');
  list.innerHTML = issues.length ? issues.map((issue) => {
    const severity = ['error', 'warning', 'info'].includes(issue.severity) ? issue.severity : 'info';
    return `<article class="issue-card ${severity}"><div class="issue-main"><span class="issue-status">${icon(severity === 'info' ? 'info' : 'warning')}</span><div class="issue-content"><div class="issue-title"><h4>${esc(issue.title)}</h4><span class="severity-label">${({ error: '需处理', warning: '需复核', info: '提示' })[severity]}</span></div><p class="issue-description">${esc(issue.description)}</p><p class="issue-location"><span>位置</span>${esc(issue.location || issue.part || '文件整体')}</p>${issue.recommendation ? `<p class="issue-recommendation"><span>建议</span>${esc(issue.recommendation)}</p>` : ''}${issue.evidence || issue.part ? `<details class="issue-detail"><summary>查看检查依据</summary>${issue.evidence ? `<p>${esc(issue.evidence)}</p>` : ''}${issue.part ? `<code>${esc(issue.part)}</code>` : ''}<span class="issue-code">${esc(issue.code)}</span></details>` : ''}</div></div></article>`;
  }).join('') : `<div class="no-issues">${icon('check')}<h4>${state.filter === 'all' ? '没有发现当前规则覆盖的风险' : '这个级别没有检查发现'}</h4><p>${state.filter === 'all' ? '这不代表已保证兼容。交接前请在目标 Windows 电脑验证。' : '可以切换筛选，查看其他发现。'}</p></div>`;
}

function updateSelection(record, groups) {
  const panel = document.querySelector('#document-panel');
  for (const input of panel.querySelectorAll('[data-issue-id]')) input.checked = record.selected.has(input.dataset.issueId);
  for (const input of panel.querySelectorAll('[data-font-group]')) {
    const group = groups[Number(input.dataset.fontGroup)];
    const count = group.filter(issue => record.selected.has(issue.id)).length;
    input.checked = count === group.length;
    input.indeterminate = count > 0 && count < group.length;
  }
  const summary = document.querySelector('#selection-summary');
  if (!summary) return;
  const selected = fixableIssues(record).filter(issue => record.selected.has(issue.id));
  let caution = selected.length ? '下载后，在 Windows 对照原件检查换行和排版。' : '';
  if (selected.some(issue => issue.fix.affectsDefaultFont)) {
    caution = '已选择替换工作簿默认字体。实际列宽、自动行高和打印分页可能变化，请在下载后对比打印预览。';
  } else if (selected.some(issue => issue.fix.affectsWrappedRows)) {
    caution = '已选字体用于自动换行、自动行高的单元格。请在下载后对比打印预览的总页数和每页首末行。';
  }
  document.querySelector('#font-caution').textContent = caution;
  document.querySelector('#font-caution').hidden = !caution;
  summary.innerHTML = `已选 <strong>${record.selected.size}</strong> / ${fixableIssues(record).length} 处替换`;
  document.querySelector('#repair-file').disabled = !record.selected.size || record.report.repairBlocked || isDocumentBusy(record);
  document.querySelector('#clear-fixes').disabled = !record.selected.size || isDocumentBusy(record);
  document.querySelector('#select-fixes').disabled = record.selected.size === fixableIssues(record).length || isDocumentBusy(record);
  const result = document.querySelector('#repair-result');
  document.querySelector('#repair-file').hidden = !!record.repaired;
  if (!record.repaired) {
    result.innerHTML = record.repairError ? `<p class="font-settings-error" role="alert">${esc(record.repairError)}</p>` : '';
    return;
  }
  const total = (record.repaired.changes || []).reduce((sum, change) => sum + (change.count || 0), 0);
  const unresolved = record.repaired.report?.issues?.length || 0;
  result.innerHTML = `<div class="repair-success"><div class="repair-success-heading">${icon('check')}<strong>可编辑副本已生成</strong></div><p>替换了 ${total} 处字体。请在浏览器的下载记录中打开副本，用 Windows ${record.format === 'pptx' ? 'PowerPoint 试播' : 'Excel 检查打印预览'}。</p><p class="repaired-name">${esc(safeFileName(record.repaired.fileName))}</p><button class="button success" id="download-repaired">${icon('download')} 再次下载副本</button><p class="font-settings-error" id="copy-download-error" role="status" ${record.downloadError ? '' : 'hidden'}>${esc(record.downloadError || '')}</p><details class="changes-details"><summary>查看修改记录和复查结果</summary><p>复查还有 ${unresolved} 项提示，详情见检查报告。</p><ul>${(record.repaired.changes || []).map((change) => `<li><strong>${esc(change.before)} → ${esc(change.after)}</strong><span>${esc(change.part)} · ${change.count || 0} 处</span></li>`).join('')}</ul></details></div>`;
  document.querySelector('#download-repaired').addEventListener('click', () => {
    downloadCopy(record);
  });
}

function downloadCopy(record) {
  const mime = record.format === 'pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  record.downloadError = null;
  try {
    downloadBlob(record.repaired.bytes, safeFileName(record.repaired.fileName), mime);
    toast(`已请求下载「${safeFileName(record.repaired.fileName)}」。若没有看到文件，可再次下载。`);
  } catch {
    record.downloadError = '副本已生成，下载未能启动。请点击“再次下载副本”重试。';
    toast(record.downloadError, 'error');
  }
  const error = document.querySelector('#copy-download-error');
  if (error && activeRecord() === record) {
    error.textContent = record.downloadError || '';
    error.hidden = !record.downloadError;
  }
}

async function repair(record) {
  if (!record.selected.size || record.report.repairBlocked || isDocumentBusy(record)) return;
  record.status = 'repairing';
  record.repairError = null;
  renderRecord(record);
  try {
    const bytes = new Uint8Array(await record.file.arrayBuffer());
    record.repaired = await request('repair', { bytes, fileName: record.name, options: { ...record.options, selectedIssueIds: [...record.selected] } }, [bytes.buffer]);
    record.status = 'ready';
    announce(`${record.name} 的可编辑副本已生成。`);
    downloadCopy(record);
  } catch (error) {
    record.status = 'ready'; record.repaired = null;
    record.repairError = `没有生成副本：${error.message}`;
    toast(record.repairError, 'error');
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
    <summary>更多字体设置 <span>${mappings.length ? `${mappings.length} 条当前文件规则` : '目标电脑字体 · 自定义替换'}</span></summary>
    <div class="font-settings-body">
      ${renderFontList(record, isDocumentBusy(record))}
      <h4>自定义字体替换</h4><p>规则仅对当前文件生效。目标字体会影响字宽和换行。没有目标名单时，下方列出常见候选字体。</p>
      <p class="font-settings-caution">此处不适用于符号、数学或图标字体。普通字体替换可能改变字符含义，请在原始应用中处理。</p>
      ${record.report.repairBlocked ? '<p class="font-settings-error">当前文件禁止修复，不能添加字体规则。</p>' : ''}
      ${record.reanalysisError ? `<p class="font-settings-error" role="alert">设置未应用：${esc(record.reanalysisError)} 已保留上次有效的检查结果。</p>` : ''}
      <form id="font-rule-form"><fieldset ${disabled ? 'disabled' : ''}><div class="font-rule-fields">
        <label for="font-source">原字体<select id="font-source" name="source" required><option value="">选择文档中的字体</option>${fonts.map((font) => `<option value="${esc(font)}">${esc(font)}</option>`).join('')}</select></label>
        <span class="font-rule-arrow" aria-hidden="true">${icon('arrow')}</span>
        <label for="font-target">目标字体<input id="font-target" name="target" type="text" list="font-target-candidates" maxlength="128" required placeholder="例如 Microsoft YaHei" autocomplete="off"></label>
        <datalist id="font-target-candidates">${(record.options.targetFonts || ['Microsoft YaHei', 'Arial', 'SimSun', 'Microsoft JhengHei', 'Calibri', 'Times New Roman', 'Consolas', 'KaiTi']).map((font) => `<option value="${esc(font)}"></option>`).join('')}</datalist>
      </div><button class="button secondary compact" id="add-font-rule" type="submit" ${!fonts.length ? 'disabled' : ''}>${record.status === 'reanalyzing' ? '<span class="spinner small"></span> 正在重新检查…' : '添加规则并重新检查'}</button><p class="font-rule-help">同一原字体会更新现有规则。重新检查会清空已选项和旧副本；之后仍需勾选并生成，才会修改新副本。</p></fieldset></form>
      <div id="font-rules" class="font-rules">${mappings.length ? mappings.map(([from, to]) => `<div class="font-rule"><span><strong>${esc(from)}</strong>${icon('arrow')}<strong>${esc(to)}</strong>${record.options.targetFonts && !listedFont(record.options.targetFonts, to) ? '（目标名单未列出，此规则停用）' : ''}</span><button type="button" class="subtle-button" data-remove-font="${esc(from)}" aria-label="移除 ${esc(from)} 替换规则" ${disabled ? 'disabled' : ''}>移除并重新检查</button></div>`).join('') : '<p class="font-rules-empty">尚未设置自定义规则。</p>'}</div>
    </div></details>`;
}

function bindFontSettings(record, panel) {
  bindFontList(record, panel, changes => reanalyzeWithOptions(record, changes));
  const details = panel.querySelector('#font-settings');
  details.addEventListener('toggle', () => { if (details.isConnected) record.advancedOpen = details.open; });
  panel.querySelector('#font-rule-form').addEventListener('submit', (event) => {
    event.preventDefault();
    if (isDocumentBusy(record) || record.report.repairBlocked) return;
    const from = panel.querySelector('#font-source').value;
    const to = panel.querySelector('#font-target').value.trim();
    if (!availableFontNames(record).includes(from) || !to) return;
    if (from.trim().toLowerCase() === to.toLowerCase()) { toast('原字体与目标字体相同，无需添加替换规则。', 'error'); return; }
    if (record.options.targetFonts && !listedFont(record.options.targetFonts, to)) {
      toast('目标名单没有这个名称。请从名单中选一个已确认的字体。', 'error'); return;
    }
    const mappings = { ...(record.options?.fontMappings || {}), [from]: to };
    reanalyzeWithOptions(record, () => ({ fontMappings: mappings }));
  });
  panel.querySelectorAll('[data-remove-font]').forEach((button) => button.addEventListener('click', () => {
    if (isDocumentBusy(record) || record.report.repairBlocked) return;
    const mappings = Object.fromEntries(Object.entries(record.options?.fontMappings || {}).filter(([from]) => from !== button.dataset.removeFont));
    reanalyzeWithOptions(record, () => ({ fontMappings: mappings }));
  }));
}

async function reanalyzeWithOptions(record, readChanges) {
  if (isDocumentBusy(record)) return;
  record.status = 'reanalyzing';
  record.advancedOpen = true;
  record.reanalysisError = null;
  renderRecord(record);
  try {
    const options = { ...record.options, ...(await readChanges()) };
    const bytes = new Uint8Array(await record.file.arrayBuffer());
    const report = await request('analyze', { bytes, fileName: record.name, options }, [bytes.buffer]);
    record.report = report;
    record.options = { ...options, ...report.analysisOptions };
    record.repaired = null;
    record.repairError = null;
    record.selected.clear();
    if (record.id === state.activeId) state.filter = 'all';
    announce(`${record.name} 已按当前字体规则重新检查。尚未选择或应用任何替换。`);
    toast('字体设置已更新并重新检查，请确认后选择需要应用的替换。');
  } catch (error) {
    record.reanalysisError = error.message;
    announce(`字体设置没有应用：${error.message}`);
  }
  record.status = 'ready';
  renderRecord(record);
  if (record.id === state.activeId) document.querySelector('#font-settings summary')?.focus({ preventScroll: true });
}

render();
