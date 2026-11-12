// Load the optional PDF viewer only when someone chooses a proof. Both modules
// are shipped in the portable HTML. Neither the document nor a URL is fetched.
import displaySource from 'pdfjs-dist/build/pdf.min.mjs?raw';
import PdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?worker&inline';

const limit = 25 * 1024 * 1024;
const previewSize = 1280;
let library;
function moduleUrl(code) {
  return URL.createObjectURL(new Blob([code], { type: 'text/javascript' }));
}
function loadLibrary() {
  if (!library) {
    const url = moduleUrl(displaySource);
    library = import(/* @vite-ignore */ url).catch(error => { library = null; throw error; })
      .finally(() => URL.revokeObjectURL(url));
  }
  return library;
}

export const compareSection = `<details class="verification pdf-compare" id="pdf-compare">
  <summary>对照两份 PDF 的页数和排版</summary><div class="proof-body">
  <p>把制作端显示正确的原稿导出为 PDF，再用接收端的 Excel 或 PowerPoint 导出待核对稿。两边选相同的纸张、打印范围和缩放；也可以用上方的本机 Office 助手自动导出。</p>
  <div class="proof-inputs">${['Mac 原稿', 'Windows 原件或副本'].map((title, index) =>
    `<div><label for="proof-${index}">${title}的 PDF</label><input id="proof-${index}" type="file" accept=".pdf,application/pdf"><p class="proof-info" aria-live="polite"></p></div>`).join('')}</div>
  <p id="proof-status" role="status">选好两份文件后，会自动显示页数和同一页的预览。每份最多 25 MiB。</p>
  <div id="proof-view" hidden><div class="proof-scan">
    <button type="button" class="button secondary compact" id="proof-scan">连续核对所有页</button>
    <button type="button" class="subtle-button" id="proof-scan-stop" hidden>停止逐页核对</button>
    <p id="proof-scan-status" role="status"></p><div id="proof-scan-results" class="proof-page-results"></div>
  </div><div class="proof-nav">
    <button type="button" class="button secondary compact" id="proof-prev">上一页</button>
    <label>第 <input type="number" min="1" value="1" id="proof-page" aria-label="跳转到页码"> 页<span id="proof-total"></span></label>
    <button type="button" class="button secondary compact" id="proof-next">下一页</button>
    <label><input type="checkbox" id="proof-zoom"> 放大预览</label></div>
    <div class="proof-grid">${['Mac 原稿', 'Windows 待核对稿'].map(title => `<figure><figcaption>${title}<small></small></figcaption><div class="proof-paper"></div></figure>`).join('')}</div>
    <p id="proof-verdict" role="status"></p><label id="proof-overlay-control" hidden><input id="proof-overlay" type="checkbox"> 在 Windows 预览上标出差异</label>
    <p class="proof-tip">重点看行末、页尾和表格边界。颜色标记也会受字体平滑和导出方式影响。超大图片、特殊字体请以 Office 打印预览为准。动画请回到 PowerPoint 播放。</p>
  </div><button type="button" class="subtle-button" id="proof-reset" hidden>清空这两份 PDF</button></div></details>`;

export function openComparison() {
  const section = document.querySelector('#pdf-compare');
  section.open = true;
  section.querySelector('summary').focus();
}

function dispose(slot) {
  slot.renderTask?.cancel();
  if (!slot.port) return;
  function terminate() { slot.worker?.destroy(); slot.port.terminate(); }
  const deadline = setTimeout(terminate, 500);
  Promise.resolve(slot.loading?.destroy()).catch(() => {}).finally(() => {
    clearTimeout(deadline);
    terminate();
  });
}

// PDF.js can finish rendering after a failed print-operator stream, even when
// stopAtErrors is set (e.g. an oversized image). Retain failures for their
// own page. The protocol is tied to the locked PDF.js version and covered by
// the portable-browser regression with an actual oversized image declaration.
function trackPrintErrors(port) {
  const streams = new Map(), failedPages = new Set();
  const send = port.postMessage.bind(port);
  port.postMessage = (message, ...transfer) => {
    if (message?.action === 'GetOperatorList' && Number.isInteger(message.data?.pageIndex)) {
      streams.set(message.streamId, message.data.pageIndex + 1);
    }
    return send(message, ...transfer);
  };
  port.addEventListener('message', ({ data }) => {
    const number = streams.get(data?.streamId);
    if (!number) return;
    if (data.reason && !['AbortException', 'RenderingCancelledException'].includes(data.reason.name)) failedPages.add(number);
    if (data.reason || data.chunk?.lastChunk) streams.delete(data.streamId);
  });
  return failedPages;
}

// A failed XObject decode can resolve to null while PDF.js still considers
// the canvas complete. Inspect dependencies before its print renderer releases
// the page caches. This includes reusable images and masks inside form objects.
async function checkPageImages(page, ops) {
  const { fnArray, argsArray } = await page.getOperatorList({ intent: 'print' });
  const ids = new Set();
  function image(value) {
    if (typeof value === 'string') ids.add(value);
    else if (typeof value?.data === 'string') ids.add(value.data);
    else if (!value || !(value.width > 0 && value.height > 0)) throw new Error('Incomplete image');
  }
  for (let i = 0; i < fnArray.length; i++) {
    const fn = fnArray[i], args = argsArray[i];
    if (fn === ops.paintImageXObject || fn === ops.paintImageXObjectRepeat ||
        fn === ops.paintImageMaskXObject || fn === ops.paintImageMaskXObjectRepeat ||
        fn === ops.paintInlineImageXObject || fn === ops.paintInlineImageXObjectGroup) image(args[0]);
    else if (fn === ops.paintImageMaskXObjectGroup) for (const mask of args[0]) image(mask);
    else if (fn === ops.dependency) for (const id of args) {
      if (typeof id === 'string' && /(?:^|_)(?:img|mask)_/.test(id)) ids.add(id);
    }
  }
  await Promise.all(Array.from(ids, id => new Promise((resolve, reject) => {
    const objects = id.startsWith('g_') ? page.commonObjs : page.objs;
    objects.get(id, value => {
      if (value?.width > 0 && value?.height > 0) resolve();
      else reject(new Error('Incomplete image'));
    }, reject);
  })));
}

export function bindComparison() {
  const root = document.querySelector('#pdf-compare');
  const byId = id => root.querySelector('#' + id);
  const inputs = [...root.querySelectorAll('input[type=file]')];
  const info = [...root.querySelectorAll('.proof-info')];
  const papers = [...root.querySelectorAll('.proof-paper')];
  const dimensions = [...root.querySelectorAll('figcaption small')];
  const status = byId('proof-status'), view = byId('proof-view'), field = byId('proof-page');
  const overlay = byId('proof-overlay'), reset = byId('proof-reset'), verdict = byId('proof-verdict');
  const slots = [{}, {}];
  let labels = ['Mac 原稿', 'Windows 待核对稿'];
  let pendingLabels = labels;
  function setLabels(next) {
    labels = next;
    root.querySelectorAll('.proof-inputs label').forEach((node, index) => { node.textContent = labels[index] + '的 PDF'; });
    root.querySelectorAll('figcaption').forEach((node, index) => { node.firstChild.textContent = labels[index]; });
    byId('proof-overlay-control').lastChild.textContent = ' 在右侧预览上标出差异';
  }
  let current = 1, generation = 0, overlayCanvas;
  let scanGeneration = 0, scanning = false;
  const scanButton = byId('proof-scan'), scanStop = byId('proof-scan-stop');
  const scanStatus = byId('proof-scan-status'), scanResults = byId('proof-scan-results');
  function stopScan(clear = false) {
    scanGeneration++;
    if (scanning && !clear) scanStatus.textContent += ' 已停止，可点击页码查看已核对的页面。';
    scanning = false; scanButton.disabled = false; scanStop.hidden = true; field.disabled = false;
    byId('proof-prev').disabled = current <= 1;
    byId('proof-next').disabled = current >= total();
    if (clear) { scanStatus.textContent = ''; scanResults.replaceChildren(); }
  }

  function invalidatePreview() {
    generation++;
    for (const slot of slots) { slot.renderTask?.cancel(); slot.renderTask = null; }
    for (const target of [...papers, ...dimensions]) target.replaceChildren();
    overlayCanvas = null;
    byId('proof-overlay-control').hidden = true;
    verdict.textContent = '';
    view.hidden = true;
  }
  function total() { return Math.max(...slots.map(s => s.doc?.numPages || 0)); }
  async function update() {
    invalidatePreview();
    const serial = generation;
    if (slots.some(slot => !slot.doc)) {
      status.textContent = '选好两份文件后，会自动显示页数和同一页的预览。每份最多 25 MiB。';
      return;
    }
    const [left, right] = slots.map(slot => slot.doc.numPages);
    status.textContent = `${labels[0]} ${left} 页，${labels[1]} ${right} 页。` +
      (left === right ? '总页数相同，还可以逐页核对内容。' : `${labels[1]}${right > left ? '多' : '少'} ${Math.abs(left-right)} 页。`);
    view.hidden = false;
    field.value = current;
    field.max = total();
    byId('proof-total').textContent = ` / 共 ${total()} 页`;
    byId('proof-prev').disabled = scanning || current <= 1;
    byId('proof-next').disabled = scanning || current >= total();
    scanButton.textContent = total() > 200 ? '连续核对前 200 页' : '连续核对所有页';
    verdict.textContent = '正在准备本页预览…';
    const number = current;
    // Each operation owns its deadline even if the other PDF fails first.
    // Terminating an unresponsive worker also covers parsing before renderTask
    // exists. The other file remains available for a replacement comparison.
    async function track(index, action) {
      const slot = slots[index];
      let timer;
      const timeout = new Promise((_, reject) => {
        timer = setTimeout(() => {
          if (serial === generation && slots[index] === slot) {
            invalidatePreview(); dispose(slot); slots[index] = {}; inputs[index].value = '';
            info[index].textContent = '预览耗时过长，已停止读取。请重新选择 PDF，或用 PDF 阅读器核对。';
            status.textContent = '本页未完成对照。请替换耗时的 PDF 后再试。';
          }
          reject(new Error('PDF render timed out'));
        }, 30000);
      });
      try { return await Promise.race([action(), timeout]); }
      finally { clearTimeout(timer); }
    }
    try {
      const pages = await Promise.all(slots.map((slot, index) => number <= slot.doc.numPages ? track(index, () => slot.doc.getPage(number)) : null));
      if (serial !== generation) return;
      const sizes = pages.map(page => page?.getViewport({ scale: 1 }));
      const longest = Math.max(...sizes.filter(Boolean).flatMap(size => [size.width, size.height]));
      if (!Number.isFinite(longest) || longest <= 0) throw new Error('Invalid page size');
      // Office exporters round A4 dimensions differently (e.g. 595 vs 595.32 pt).
      // Keep one physical scale and pad sub-point differences to a common canvas.
      const paired = sizes.every(Boolean) && Math.abs(sizes[0].width - sizes[1].width) <= 1
        && Math.abs(sizes[0].height - sizes[1].height) <= 1;
      const scale = previewSize / longest;
      const widths = sizes.filter(Boolean).map(size => size.width);
      const widest = Math.max(...widths);
      const { OPS } = await loadLibrary();
      if (serial !== generation) return;
      const canvases = await Promise.all(pages.map((page, index) => track(index, async () => {
        if (!page) { papers[index].textContent = '这份 PDF 没有这一页'; return null; }
        const slot = slots[index];
        await checkPageImages(page, OPS);
        if (serial !== generation) return null;
        if (slot.failedPages.has(number)) throw new Error('Incomplete PDF page');
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.ceil((paired ? widest : sizes[index].width) * scale));
        canvas.height = Math.max(1, Math.ceil((paired ? Math.max(sizes[0].height, sizes[1].height) : sizes[index].height) * scale));
        canvas.style.width = (sizes[index].width / widest * 100) + '%';
        canvas.style.setProperty('--proof-width', String(sizes[index].width / widest));
        canvas.setAttribute('aria-label', labels[index] + ' 第 ' + number + ' 页预览');
        const mm = pt => Math.round(pt * 25.4 / 72);
        dimensions[index].textContent = mm(sizes[index].width) + ' × ' + mm(sizes[index].height) + ' 毫米';
        const task = page.render({ canvas, viewport, intent: 'print', background: '#fff' });
        slots[index].renderTask = task;
        await task.promise;
        if (slot.failedPages.has(number)) throw new Error('Incomplete PDF page');
        if (serial === generation) { papers[index].append(canvas); slots[index].renderTask = null; page.cleanup(); }
        return canvas;
      })));
      if (serial !== generation) return;
      if (canvases.some(canvas => !canvas)) { verdict.textContent = '只有一份文件包含这一页。可回到上一页核对分页位置。'; return 'missing'; }
      if (!paired) {
        verdict.textContent = '两份稿的纸张尺寸或方向不同。请核对导出时的纸张和方向。';
        return 'size';
      }
      const [a, b] = canvases.map(canvas => canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height));
      let changed = false, leftInk = false, rightInk = false;
      for (let i = 0; i < a.data.length; i += 4) {
        leftInk ||= Math.min(a.data[i], a.data[i+1], a.data[i+2]) < 250;
        rightInk ||= Math.min(b.data[i], b.data[i+1], b.data[i+2]) < 250;
        if (Math.max(Math.abs(a.data[i]-b.data[i]), Math.abs(a.data[i+1]-b.data[i+1]), Math.abs(a.data[i+2]-b.data[i+2])) < 18) continue;
        changed = true;
        b.data[i] = 213; b.data[i+1] = 18; b.data[i+2] = 99;
      }
      if (!leftInk || !rightInk) {
        verdict.textContent = '至少一份预览看起来是空白页。请用 PDF 阅读器确认这一页的内容，再判断排版。';
        return 'blank';
      }
      verdict.textContent = changed ? '本页预览有差异。勾选下方选项可查看具体位置。' : '在当前预览精度下，这一页未见明显图像差异。';
      if (changed) {
        overlayCanvas = document.createElement('canvas');
        overlayCanvas.width = b.width; overlayCanvas.height = b.height;
        overlayCanvas.setAttribute('aria-label', (labels[1] === 'Windows 待核对稿' ? 'Windows 预览，差异用玫红色标出' : labels[1] + '预览，差异用玫红色标出'));
        overlayCanvas.getContext('2d').putImageData(b, 0, 0);
        papers[1].append(overlayCanvas);
        byId('proof-overlay-control').hidden = false;
        toggleOverlay();
      }
      return changed ? 'different' : 'same';
    } catch {
      if (serial !== generation) return;
      invalidatePreview(); view.hidden = false;
      verdict.textContent = '本页预览未能完整显示，无法判断排版差异。请在 PDF 阅读器中核对这一页，也可以继续核对下一页。';
      return 'unavailable';
    }
  }
  async function scanPages() {
    if (slots.some(slot => !slot.doc)) return;
    stopScan(true);
    const serial = ++scanGeneration, count = Math.min(total(), 200);
    const outcomes = [], descriptions = {same:'预览未见明显差异',different:'预览有差异',missing:'一份缺页',size:'纸张不同',blank:'空白需确认',unavailable:'未能完整读取'};
    scanning = true; scanButton.disabled = true; scanStop.hidden = false; field.disabled = true;
    for (let number = 1; number <= count; number++) {
      if (serial !== scanGeneration) return;
      current = number;
      scanStatus.textContent = `正在核对第 ${number} / ${count} 页…`;
      const outcome = await update();
      if (serial !== scanGeneration) return;
      if (!outcome) { stopScan(); scanStatus.textContent = `逐页核对在第 ${number} 页中断，已完成 ${outcomes.length} 页。请查看 PDF 读取提示。`; return; }
      outcomes.push(outcome);
      const button = document.createElement('button');
      button.type = 'button'; button.className = 'proof-page-result ' + (outcome === 'same' ? 'unchanged' : 'review');
      button.textContent = `${number} · ${descriptions[outcome]}`;
      button.setAttribute('aria-label', `查看第 ${number} 页：${descriptions[outcome]}`);
      button.addEventListener('click', () => { stopScan(); navigate(number); });
      scanResults.append(button);
      // Let cancel, replacement and navigation events run between pages.
      await new Promise(resolve => setTimeout(resolve, 0));
    }
    if (serial !== scanGeneration) return;
    stopScan();
    const review = outcomes.filter(item => item !== 'same').length;
    scanStatus.textContent = `已核对 ${count} / ${total()} 页，${review} 页有差异或需要人工确认。` +
      (total() > count ? '其余页面尚未核对，请通过页码继续查看。' : '点击页码查看具体位置。') + ' 结果只代表当前预览精度。';
    current = Math.max(1, outcomes.findIndex(item => item !== 'same') + 1);
    void update();
  }
  scanButton.addEventListener('click', () => void scanPages());
  scanStop.addEventListener('click', () => { stopScan(); void update(); });
  function toggleOverlay() {
    if (!overlayCanvas) return;
    papers[1].firstElementChild.hidden = overlay.checked;
    overlayCanvas.hidden = !overlay.checked;
  }
  function navigate(number) {
    if (!Number.isInteger(number) || number < 1 || number > total()) { field.value = current; return; }
    current = number; void update();
  }
  async function choose(index, file) {
    stopScan(true);
    dispose(slots[index]);
    const slot = slots[index] = {};
    current = 1;
    void update();
    reset.hidden = !inputs.some(input => input.files.length);
    info[index].textContent = '';
    if (!file) return;
    if (!/\.pdf$/i.test(file.name) || !file.size || file.size > limit) {
      info[index].textContent = '请选择一份有效的 PDF，每份最多 25 MiB。';
      return;
    }
    info[index].textContent = '正在本机读取…';
    let timer;
    try {
      const [pdf, data] = await Promise.all([loadLibrary(), file.arrayBuffer()]);
      if (slots[index] !== slot) return;
      slot.port = new PdfWorker();
      slot.failedPages = trackPrintErrors(slot.port);
      slot.worker = new pdf.PDFWorker({ port: slot.port });
      slot.loading = pdf.getDocument({ data, worker: slot.worker, stopAtErrors: true, enableXfa: false,
        disableWorkerRendering: true, useWorkerFetch: false, useWasm: false, verbosity: 0,
        maxImageSize: 16 * 1024 * 1024, canvasMaxAreaInBytes: 32 * 1024 * 1024 });
      const failed = new Promise((_, reject) => {
        slot.port.addEventListener('error', () => reject(new Error('worker')), {once: true});
        timer = setTimeout(() => reject(new Error('timeout')), 30000);
      });
      const doc = await Promise.race([slot.loading.promise, failed]);
      if (slots[index] !== slot) return;
      slot.doc = doc;
      info[index].textContent = `已读取，共 ${doc.numPages} 页`;
      void update();
    } catch (error) {
      if (slots[index] !== slot) return;
      dispose(slot); slots[index] = {};
      info[index].textContent = error.name === 'PasswordException'
        ? '这份 PDF 有密码。请从你能打开的原稿重新导出一份用于对照的 PDF。'
        : '无法读取这份 PDF。请在最新版 Chrome 或 Edge 中重试，或从 Office 重新导出。';
    } finally { clearTimeout(timer); }
  }
  inputs.forEach((input, index) => input.addEventListener('change', () => void choose(index, input.files[0])));
  byId('proof-prev').addEventListener('click', () => navigate(current - 1));
  byId('proof-next').addEventListener('click', () => navigate(current + 1));
  field.addEventListener('change', () => navigate(field.valueAsNumber));
  overlay.addEventListener('change', toggleOverlay);
  byId('proof-zoom').addEventListener('change', event => {
    root.querySelector('.proof-grid').classList.toggle('expanded', event.target.checked);
  });
  reset.addEventListener('click', () => {
    stopScan(true);
    inputs.forEach(input => { input.value = ''; });
    for (const index of [0,1]) { dispose(slots[index]); slots[index] = {}; info[index].textContent = ''; }
    setLabels(pendingLabels);
    overlay.checked = false; current = 1; reset.hidden = true; void update(); inputs[0].focus();
  });
  return {
    setDirection(target) {
      pendingLabels = target === 'mac-office' ? ['Windows 原稿', 'Mac 待核对稿'] : ['Mac 原稿', 'Windows 待核对稿'];
      // A direction for a new document cannot relabel PDF evidence already
      // loaded from another document or a native Office record.
      if (inputs.every(input => !input.files.length)) setLabels(pendingLabels);
    },
    setFiles(files, names) {
      setLabels(names);
      for (const index of [0, 1]) {
        const transfer = new DataTransfer();
        if (files[index]) transfer.items.add(files[index]);
        inputs[index].files = transfer.files;
        void choose(index, files[index]);
      }
      root.open = true;
    },
  };
}
