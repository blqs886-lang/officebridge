import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { analyzeOffice } from '../src/engine/index.js';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { buildHtmlReport, safeFileName, escapeHtml } from '../src/ui/reports.js';
import { donationUrl, label as donationLabel, validateDonationUrl } from '../src/support-config.js';

// Windows Edge's own download flyout loads these built-in hosts. These exact
// scheme/host pairs were observed during download QA; other browser pages fail.
function isBrowserInternalRequest(request) {
  if (request.method !== 'GET') return false;
  try {
    const url = new URL(request.url);
    return url.protocol === 'edge:' && !url.username && !url.password && !url.port
      && ['downloads-hub', 'resources', 'fileicon'].includes(url.hostname);
  } catch { return false; }
}

function isDevelopmentServerWebSocket(value, base) {
  if (!['http:', 'https:'].includes(base.protocol)) return false;
  try {
    const url = new URL(value);
    return url.protocol === (base.protocol === 'https:' ? 'wss:' : 'ws:')
      && url.host === base.host && !url.username && !url.password;
  } catch { return false; }
}

function isVitePingFrame(frame) {
  return frame.type === 'text' && frame.payload === '{"type":"ping"}';
}

function isExpectedAppRequest(request, base, localReports) {
  try {
    const url = new URL(request.url);
    if (['blob:', 'data:'].includes(url.protocol) || localReports.has(url.href)) return true;
    if (base.protocol === 'file:') return url.href === base.href;
    if (isDevelopmentServerWebSocket(url.href, base)) return true;
    return ['http:', 'https:'].includes(base.protocol)
      && ['http:', 'https:'].includes(url.protocol) && url.origin === base.origin;
  } catch { return false; }
}

// Optional browser QA dependency. No production dependency or browser download is required.
// npm install --no-save playwright OR PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs
const playwright = await import('playwright').catch(async (error) => {
  if (!process.env.PLAYWRIGHT_MODULE) throw new Error('Set PLAYWRIGHT_MODULE to an existing Playwright index.mjs, or install Playwright for browser QA.', { cause: error });
  return import(pathToFileURL(path.resolve(process.env.PLAYWRIGHT_MODULE)).href);
});
const { chromium } = playwright;
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const baseUrl = process.env.APP_URL || 'http://127.0.0.1:5173/';
const outputDir = process.env.BROWSER_SMOKE_OUT ? path.resolve(process.env.BROWSER_SMOKE_OUT) : path.join(projectRoot, 'work', 'browser-smoke');
await fs.mkdir(outputDir, { recursive: true });
const browser = await chromium.launch({ headless: true, chromiumSandbox: true, ...(process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {}) });
const isLocalFile = new URL(baseUrl).protocol === 'file:';
const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true, offline: isLocalFile });
const page = await context.newPage();
const failures = [];
const requests = [];
const webSocketConnections = [];
let requestAudit;
function watchPage(observedPage) {
  observedPage.on('pageerror', (error) => failures.push(error.message));
  observedPage.on('websocket', (socket) => {
    const connection = { url: socket.url(), sentFrames: [] };
    webSocketConnections.push(connection);
    socket.on('framesent', ({ payload }) => connection.sentFrames.push(typeof payload === 'string'
      ? { type: 'text', payload }
      : { type: 'binary', payload: Buffer.from(payload).toString('base64') }));
  });
}
watchPage(page);
context.on('page', watchPage);
context.on('request', (request) => requests.push({ method: request.method(), url: request.url() }));
const checks = [];
const test = async (name, fn) => { await fn(); checks.push(name); console.log(`PASS ${name}`); };
async function expandSection(view, selector) {
  const section = view.locator(selector);
  if (!await section.evaluate(element => element.open)) await section.locator(':scope > summary').click();
}
const clickDownload = async (selector, name, view = page) => {
  if (['#download-json', '#download-html'].includes(selector)) await expandSection(view, '.report-details');
  const downloadPromise = view.waitForEvent('download');
  await view.locator(selector).click();
  const download = await downloadPromise;
  const target = path.join(outputDir, name || download.suggestedFilename());
  await download.saveAs(target);
  assert.equal(await download.failure(), null);
  return target;
};

try {
  await page.goto(baseUrl);
  await page.locator('#choose-files').waitFor({ state: 'visible' });
  await test('首页可用且没有横向溢出', async () => {
    assert.match(await page.title(), /OfficeBridge/);
    assert.equal(await page.locator('#workspace').isVisible(), false);
    assert.ok(await page.getByRole('button', { name: '选择文件', exact: true }).isVisible());
    const configuredSupportUrl = validateDonationUrl(donationUrl);
    if (configuredSupportUrl) {
      const supportLink = page.locator('#support-link');
      assert.equal(await supportLink.isVisible(), true);
      assert.equal(await supportLink.getAttribute('href'), configuredSupportUrl);
      assert.equal(await supportLink.innerText(), donationLabel.trim());
      assert.equal(await supportLink.getAttribute('target'), '_blank');
      const relations = (await supportLink.getAttribute('rel')).split(/\s+/);
      assert.ok(relations.includes('noopener') && relations.includes('noreferrer'));
      assert.equal(await page.locator('#support-footer small').innerText(), '自愿支持，不影响免费功能');
    } else {
      assert.equal(await page.locator('#support-link, #support-footer').count(), 0, 'Unconfigured donation link must be absent, including any placeholder');
    }
    assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await page.screenshot({ path: path.join(outputDir, 'desktop-home.png'), fullPage: true });
  });

  await test('新用户可用键盘打开文件选择器，检查后回到上传入口', async () => {
    const newcomer = await context.newPage();
    await newcomer.setViewportSize({ width: 320, height: 740 });
    await newcomer.goto(baseUrl);
    assert.ok(await newcomer.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await newcomer.screenshot({ path: path.join(outputDir, 'mobile-home.png'), fullPage: true });
    for (let step = 0; step < 12 && !await newcomer.locator('#choose-files').evaluate(node => node.matches(':focus')); step++) {
      await newcomer.keyboard.press('Tab');
    }
    assert.ok(await newcomer.locator('#choose-files').evaluate(node => node.matches(':focus')), 'Keyboard focus reaches the file chooser');
    const waiting = newcomer.waitForEvent('filechooser');
    await newcomer.keyboard.press('Enter');
    const chooser = await waiting;
    assert.equal(chooser.isMultiple(), true);
    await chooser.setFiles(path.join(projectRoot, 'public/examples/officebridge-demo.pptx'));
    await newcomer.locator('.result-summary').waitFor();
    assert.equal(await newcomer.locator('.document-title h2').innerText(), 'officebridge-demo.pptx');
    assert.ok(await newcomer.locator('#workspace').evaluate(node => node.contains(document.activeElement)), 'File choice moves focus into the workspace');
    assert.equal(await newcomer.locator('[data-issue-id]:checked').count(), 0);
    assert.equal(await newcomer.locator('.findings-details').evaluate(node => node.open), false);
    assert.equal(await newcomer.locator('.report-details').evaluate(node => node.open), false);
    await newcomer.locator('#clear-files').focus();
    await newcomer.keyboard.press('Enter');
    await newcomer.locator('#workspace').waitFor({ state: 'hidden' });
    assert.ok(await newcomer.locator('#choose-files').evaluate(node => node.matches(':focus')));
    await newcomer.close();
  });

  await test('拖入普通表格可直接沿用原件，没有替换操作', async () => {
    const original = await context.newPage();
    const parts = unzipSync(await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx')));
    parts['xl/styles.xml'] = strToU8(strFromU8(parts['xl/styles.xml']).replaceAll('PingFang SC', 'Arial').replaceAll('Helvetica Neue', 'Arial'));
    const simpleSheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>4</v></c></row></sheetData></worksheet>';
    for (const sheet of ['sheet1', 'sheet2']) parts[`xl/worksheets/${sheet}.xml`] = strToU8(simpleSheet);
    await original.goto(baseUrl);
    const data = await original.evaluateHandle(bytes => {
      const transfer = new DataTransfer();
      transfer.items.add(new File([Uint8Array.from(bytes)], '普通表格.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      return transfer;
    }, Array.from(zipSync(parts)));
    try {
      await original.locator('#drop-zone').dispatchEvent('drop', { dataTransfer: data });
    } finally { await data.dispose(); }
    await original.locator('.result-summary[data-outcome="original"]').waitFor();
    assert.match(await original.locator('.result-summary').innerText(), /原件/);
    assert.equal(await original.locator('[data-issue-id]').count(), 0);
    assert.equal(await original.locator('#repair-file').isEnabled(), false);
    assert.equal(await original.locator('#download-repaired').count(), 0);
    await original.close();
  });

  await test('批量上传 PPTX/XLSX，默认不启用字体替换', async () => {
    await page.locator('#file-input').setInputFiles([
      path.join(projectRoot, 'public/examples/officebridge-demo.xlsx'),
      path.join(projectRoot, 'public/examples/officebridge-demo.pptx'),
    ]);
    await page.locator('#repair-file').waitFor();
    await page.waitForFunction(() => document.querySelectorAll('.file-card .spinner').length === 0);
    assert.equal(await page.locator('.file-card').count(), 2);
    assert.equal(await page.locator('#repair-file').isEnabled(), false);
    assert.equal(await page.locator('[data-issue-id]:checked').count(), 0);
    assert.ok(await page.locator('[data-issue-id]').count() > 0);
  });

  let repairedPath;
  await test('一次点击生成并下载 XLSX，再次下载仍是同一份可编辑副本', async () => {
    await page.locator('#select-fixes').click();
    assert.equal(await page.locator('#repair-file').isEnabled(), true);
    repairedPath = await clickDownload('#repair-file', 'repaired.xlsx');
    await page.locator('#download-repaired').waitFor();
    const repairedBytes = await fs.readFile(repairedPath);
    const repeated = await clickDownload('#download-repaired', 'repaired-again.xlsx');
    assert.deepEqual(await fs.readFile(repeated), repairedBytes, 'Repeat download reuses the generated copy');
    const before = unzipSync(await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx')));
    const after = unzipSync(repairedBytes);
    for (const sheet of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) assert.deepEqual(after[sheet], before[sheet], 'Budget formula, input and text IDs must stay editable');
    const report = await analyzeOffice(repairedBytes, 'repaired.xlsx');
    assert.equal(report.repairBlocked, false);
    assert.equal(report.validation.windowsOfficeTested, false);
    assert.equal(report.issues.filter((issue) => issue.fixable).length, 0);
    assert.equal(report.inventory.worksheets.length, 2);
  });

  await test('HTML 和 JSON 报告包含复查结果与 Windows 未实测说明', async () => {
    const jsonPath = await clickDownload('#download-json', 'report.json');
    const exported = JSON.parse(await fs.readFile(jsonPath, 'utf8'));
    assert.equal(exported.analysis.validation.windowsOfficeTested, false);
    assert.ok(exported.repair.changes.length > 0);
    assert.equal(exported.repair.report.validation.windowsOfficeTested, false);
    const htmlPath = await clickDownload('#download-html', 'report.html');
    const html = await fs.readFile(htmlPath, 'utf8');
    assert.match(html, /Content-Security-Policy/);
    assert.match(html, /修复后的复查发现/);
    assert.match(html, /尚未在 Windows Office 中实际验证/);
    const reportPage = await context.newPage();
    await reportPage.setViewportSize({ width: 320, height: 740 });
    await reportPage.goto(pathToFileURL(htmlPath).href);
    assert.ok(await reportPage.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await reportPage.setViewportSize({ width: 794, height: 1123 });
    await reportPage.emulateMedia({ media: 'print' });
    assert.ok(await reportPage.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await reportPage.pdf({ path: path.join(outputDir, 'report-print.pdf'), format: 'A4', margin: { top: '15mm', bottom: '15mm', left: '15mm', right: '15mm' } });
    await reportPage.close();
    await page.locator('#repair-result').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(outputDir, 'repair-result.png') });
  });

  await test('批量结束后重新上传下载副本，能再次检查且无需重复字体修复', async () => {
    await page.locator('#file-input').setInputFiles(repairedPath);
    await page.waitForFunction(() => document.querySelectorAll('.file-card').length === 3 && document.querySelectorAll('.file-card .spinner').length === 0);
    await page.locator('.file-card').last().click();
    assert.equal(await page.locator('[data-issue-id]').count(), 0);
    assert.equal(await page.locator('#repair-file').isEnabled(), false);
  });

  await test('演示文稿也可生成可编辑副本', async () => {
    await page.locator('.file-card').filter({ hasText: 'officebridge-demo.pptx' }).click();
    await page.locator('#select-fixes').click();
    const output = await clickDownload('#repair-file', 'repaired.pptx');
    const bytes = await fs.readFile(output);
    const slides = unzipSync(bytes);
    assert.ok(Object.hasOwn(slides, 'ppt/charts/chart1.xml'), 'PowerPoint chart is still editable');
    assert.match(strFromU8(slides['ppt/slides/slide1.xml']), /<a:t>[^<]+<\/a:t>/, 'Slide still contains native text');
    const report = await analyzeOffice(bytes, 'repaired.pptx');
    assert.equal(report.repairBlocked, false);
    assert.equal(report.issues.filter((issue) => issue.fixable).length, 0);
    assert.equal(report.inventory.slideCount, 2);
  });

  await test('同种字体可整组选用或逐处取消，调整后只导出当前选择', async () => {
    const view = await context.newPage();
    const sample = path.join(projectRoot, 'public/examples/officebridge-demo.pptx');
    const input = await fs.readFile(sample);
    await view.goto(baseUrl);
    await view.locator('#file-input').setInputFiles(sample);
    const group = view.locator('[data-font-group]');
    await group.check();
    assert.equal(await view.locator('[data-issue-id]:checked').count(), 2);
    const both = await clickDownload('#repair-file', 'both-slides.pptx', view);
    await view.locator('#download-repaired').waitFor();
    const allChanges = JSON.parse(await fs.readFile(await clickDownload('#download-json', 'both-slides.json', view), 'utf8')).repair.changes;
    assert.deepEqual(allChanges.map(change => change.part).sort(), ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml']);
    await expandSection(view, '.font-places');
    await view.locator('.font-places [data-issue-id]').last().uncheck();
    assert.ok(await group.evaluate(node => node.indeterminate), 'Partial group communicates its mixed state');
    assert.equal(await view.locator('#download-repaired').count(), 0);
    const interim = JSON.parse(await fs.readFile(await clickDownload('#download-json', 'partial-pending.json', view), 'utf8'));
    assert.equal(interim.repair, null);
    const partial = await clickDownload('#repair-file', 'one-slide.pptx', view);
    const changes = JSON.parse(await fs.readFile(await clickDownload('#download-json', 'one-slide.json', view), 'utf8')).repair.changes;
    assert.equal(changes.length, 1);
    assert.equal(changes[0].before, 'PingFang SC');
    const source = unzipSync(input);
    const bothSlides = unzipSync(await fs.readFile(both));
    const oneSlide = unzipSync(await fs.readFile(partial));
    for (const { part } of allChanges) {
      assert.notDeepEqual(bothSlides[part], source[part]);
      assert.deepEqual(oneSlide[part], changes[0].part === part ? bothSlides[part] : source[part]);
    }
    await view.locator('#clear-fixes').click();
    assert.equal(await view.locator('[data-issue-id]:checked').count(), 0);
    assert.equal(await view.locator('#repair-file').isEnabled(), false);
    assert.equal(await view.locator('#download-repaired').count(), 0);
    await view.close();
  });

  await test('首次下载未启动时，保留生成的副本供手动重新下载', async () => {
    const retry = await context.newPage();
    await retry.addInitScript(() => {
      const normalClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function () {
        if (this.download.endsWith('.windows.xlsx')) {
          HTMLAnchorElement.prototype.click = normalClick;
          throw new DOMException('Test download refused', 'NotAllowedError');
        }
        normalClick.call(this);
      };
    });
    await retry.goto(baseUrl);
    await retry.locator('#demo-xlsx').click();
    await retry.locator('#select-fixes').click();
    await retry.locator('#repair-file').click();
    await retry.locator('#copy-download-error').waitFor();
    await retry.locator('#download-repaired').waitFor();
    assert.equal(await retry.locator('a[download]').count(), 0, 'A refused download must release its temporary link');
    const copy = await clickDownload('#download-repaired', 'manual-retry.xlsx', retry);
    assert.equal(await retry.locator('#copy-download-error').isVisible(), false);
    const report = await analyzeOffice(await fs.readFile(copy), 'manual-retry.xlsx');
    assert.equal(report.issues.filter(issue => issue.fixable).length, 0);
    assert.equal(report.inventory.worksheets.length, 2);
    await retry.close();
  });

  await test('Keynote、Numbers 与可用文件一起选入时，各自得到正确指引', async () => {
    const mixed = await context.newPage();
    await mixed.goto(baseUrl);
    const workbook = await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx'));
    await mixed.locator('#file-input').setInputFiles([
      { name: '演示.key', mimeType: 'application/octet-stream', buffer: Buffer.from('example') },
      { name: '预算.numbers', mimeType: 'application/octet-stream', buffer: Buffer.from('example') },
      { name: '可用表格.xlsx', mimeType: 'application/octet-stream', buffer: workbook },
    ]);
    await mixed.locator('.result-summary').waitFor();
    assert.equal(await mixed.locator('.file-card').count(), 1);
    assert.equal(await mixed.locator('.document-title h2').innerText(), '可用表格.xlsx');
    const errors = await mixed.locator('#input-errors li').allTextContents();
    assert.equal(errors.length, 2);
    assert.match(errors[0], /演示\.key.*Keynote.*PowerPoint.*\.pptx/);
    assert.match(errors[1], /预算\.numbers.*Numbers.*Excel.*\.xlsx/);
    await mixed.locator('#dismiss-input-errors').click();
    assert.equal(await mixed.locator('#input-errors').isVisible(), false);
    assert.equal(await mixed.locator('.file-card').count(), 1);
    await mixed.close();
  });

  await test('不支持、空文件和损坏文件给出可读错误', async () => {
    await page.locator('#file-input').setInputFiles({ name: 'unsupported.txt', mimeType: 'text/plain', buffer: Buffer.from('synthetic') });
    await page.locator('#input-errors').filter({ hasText: /unsupported\.txt.*\.pptx.*\.xlsx/i }).waitFor();
    assert.equal(await page.locator('.file-card').count(), 3);
    await page.locator('#file-input').setInputFiles({ name: 'empty.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.alloc(0) });
    await page.locator('#toast').filter({ hasText: '空文件' }).waitFor();
    await page.locator('#file-input').setInputFiles({ name: 'broken.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('not a ZIP') });
    await page.locator('.file-card').filter({ hasText: 'broken.xlsx' }).click();
    await page.getByRole('heading', { name: '这个文件暂时无法检查' }).waitFor();
  });

  await test('签名文件只读，禁止字体修复', async () => {
    const contents = unzipSync(new Uint8Array(await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx'))));
    contents['_xmlsignatures/sig1.xml'] = strToU8('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
    await page.locator('#file-input').setInputFiles({ name: 'signed.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from(zipSync(contents)) });
    await page.locator('.file-card').filter({ hasText: 'signed.xlsx' }).click();
    await page.locator('.result-summary[data-outcome="blocked"]').waitFor();
    assert.equal(await page.locator('[data-issue-id]').count(), 0);
    assert.equal(await page.locator('#repair-file').isEnabled(), false);
    assert.equal(await page.locator('#select-fixes').isEnabled(), false);
  });

  await test('文件名及报告文本转义，下载名符合 Windows 约束', async () => {
    const maliciousName = '<img src=x onerror=alert(1)>.xlsx';
    const payload = { fileName: maliciousName, format: 'xlsx', issues: [{ title: '<script>alert(1)</script>', description: '<img src=x onerror=alert(1)>', location: '" onclick="alert(1)', severity: 'warning' }] };
    const html = buildHtmlReport({ report: payload });
    assert.ok(!html.includes('<script>'));
    assert.ok(!html.includes('<img '));
    assert.match(html, /&lt;script&gt;/);
    assert.equal(escapeHtml('<&\"\''), '&lt;&amp;&quot;&#39;');
    assert.equal(safeFileName('CON.xlsx'), '_CON.xlsx');
    assert.ok(safeFileName('X'.repeat(200) + '.xlsx').endsWith('.xlsx'));
    assert.ok(!/[<>:"/\\|?*]/.test(safeFileName(maliciousName)));
    const longPart = 'xl/worksheets/' + 'synthetic-long-part-name'.repeat(20) + '.xml';
    const longReport = buildHtmlReport({
      report: { ...payload, fileName: 'synthetic'.repeat(35) + '.xlsx', analysisOptions: { fontMappings: { ['Synthetic Font '.repeat(20)]: 'Arial' } } },
      repaired: { fileName: 'synthetic'.repeat(35) + '-windows.xlsx', changes: [{ part: longPart, before: 'Synthetic Font', after: 'Arial', count: 1 }], report: { issues: [] } },
    });
    const longReportPath = path.join(outputDir, 'long-path-report.html');
    await fs.writeFile(longReportPath, longReport);
    const longPage = await context.newPage();
    await longPage.goto(pathToFileURL(longReportPath).href);
    for (const width of [320, 390, 794]) {
      await longPage.setViewportSize({ width, height: 1000 });
      await longPage.emulateMedia({ media: width === 794 ? 'print' : 'screen' });
      assert.ok(await longPage.evaluate(() => document.documentElement.scrollWidth <= innerWidth), `Long report overflows at ${width}px`);
    }
    await longPage.emulateMedia({ media: 'screen' });
    await longPage.setViewportSize({ width: 390, height: 844 });
    await longPage.screenshot({ path: path.join(outputDir, 'long-path-report-mobile.png'), fullPage: true });
    await longPage.emulateMedia({ media: 'print' });
    await longPage.setViewportSize({ width: 794, height: 1123 });
    await longPage.pdf({ path: path.join(outputDir, 'long-path-report-print.pdf'), format: 'A4', margin: { top: '15mm', bottom: '15mm', left: '15mm', right: '15mm' } });
    await longPage.close();
    await page.locator('#file-input').setInputFiles({ name: maliciousName, mimeType: 'application/octet-stream', buffer: await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx')) });
    await page.locator('.file-card').last().click();
    await page.locator('#repair-file').waitFor();
    assert.equal(await page.locator('.document-title h2').innerText(), maliciousName);
    assert.equal(await page.locator('#document-panel img, #document-panel script').count(), 0);
  });

  await test('窄屏工作台无横向溢出，可操作修复和筛选', async () => {
    await page.setViewportSize({ width: 390, height: 844 });
    await expandSection(page, '.findings-details');
    await page.locator('[data-filter="warning"]').click();
    assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await page.locator('#select-fixes').click();
    assert.equal(await page.locator('#repair-file').isEnabled(), true);
    await page.locator('.document-heading').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(outputDir, 'mobile-report.png') });
    await page.setViewportSize({ width: 320, height: 740 });
    assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
  });

  await test('超时会终止 Worker，并恢复后续文件的顺序检查', async () => {
    const recovery = await context.newPage();
    await recovery.addInitScript(() => {
      const OriginalWorker = window.Worker;
      const originalTimer = window.setTimeout.bind(window);
      let firstWorker = true;
      let firstTimeout = true;
      window.Worker = class {
        constructor(...args) {
          if (!firstWorker) return new OriginalWorker(...args);
          firstWorker = false;
          // Simulate a stuck worker without consuming CPU or waiting a real minute.
          return { addEventListener() {}, postMessage() {}, terminate() {} };
        }
      };
      window.setTimeout = (fn, delay, ...args) => {
        if (delay === 60_000 && firstTimeout) { firstTimeout = false; return originalTimer(fn, 20, ...args); }
        return originalTimer(fn, delay, ...args);
      };
    });
    await recovery.goto(baseUrl);
    await recovery.locator('#file-input').setInputFiles([
      path.join(projectRoot, 'public/examples/officebridge-demo.xlsx'),
      path.join(projectRoot, 'public/examples/officebridge-demo.pptx'),
    ]);
    await recovery.getByRole('heading', { name: '这个文件暂时无法检查' }).waitFor();
    assert.match(await recovery.locator('.error-panel').innerText(), /超过 60 秒/);
    await recovery.waitForFunction(() => document.querySelectorAll('.file-card .spinner').length === 0);
    await recovery.locator('.file-card').last().click();
    await recovery.locator('#repair-file').waitFor();
    assert.ok(await recovery.locator('[data-issue-id]').count() > 0);
    await recovery.close();
  });

  await test('单文件 100MB 和每批 250MB 限制在读取内容前执行', async () => {
    const limits = await context.newPage();
    await limits.addInitScript(() => {
      const originalSize = Object.getOwnPropertyDescriptor(Blob.prototype, 'size').get;
      Object.defineProperty(Blob.prototype, 'size', {
        get() {
          if (this.name?.startsWith('quota-')) return 100 * 1024 * 1024;
          if (this.name?.startsWith('oversize-')) return 100 * 1024 * 1024 + 1;
          return originalSize.call(this);
        },
      });
    });
    await limits.goto(baseUrl);
    const fixture = await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx'));
    await limits.locator('#file-input').setInputFiles([1, 2, 3].map((number) => ({ name: `quota-${number}.xlsx`, mimeType: 'application/octet-stream', buffer: fixture })));
    await limits.locator('#toast').filter({ hasText: '超过 250 MB' }).waitFor();
    assert.equal(await limits.locator('.file-card').count(), 2);
    await limits.locator('#file-input').setInputFiles({ name: 'oversize-test.xlsx', mimeType: 'application/octet-stream', buffer: fixture });
    await limits.locator('#toast').filter({ hasText: '超过 100 MB' }).waitFor();
    assert.equal(await limits.locator('.file-card').count(), 2);
    await limits.close();
  });

  await test('自定义字体规则仅影响当前文件，重新检查不自动修改，报告披露参数', async () => {
    const custom = await context.newPage();
    const fixture = new Uint8Array(await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx')));
    const parts = unzipSync(fixture);
    parts['xl/styles.xml'] = strToU8(new TextDecoder().decode(parts['xl/styles.xml']).replaceAll('PingFang SC', 'OfficeBridge Test Font'));
    const customBytes = zipSync(parts);
    await custom.goto(baseUrl);
    await custom.locator('#file-input').setInputFiles([
      { name: 'custom-font.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from(customBytes) },
      { name: 'another-file.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from(fixture) },
    ]);
    await custom.locator('#repair-file').waitFor();
    await custom.waitForFunction(() => document.querySelectorAll('.file-card .spinner').length === 0);
    await custom.locator('#select-fixes').click();
    await custom.locator('#repair-file').click();
    await custom.locator('#download-repaired').waitFor();
    await custom.locator('#font-settings > summary').click();
    await custom.locator('#font-source').selectOption('OfficeBridge Test Font');
    await custom.locator('#font-target').fill('Arial');
    await custom.locator('#add-font-rule').click();
    await custom.locator('[data-remove-font="OfficeBridge Test Font"]').waitFor();
    assert.equal(await custom.locator('#download-repaired').count(), 0);
    assert.equal(await custom.locator('[data-issue-id]:checked').count(), 0);
    assert.equal(await custom.locator('#repair-file').isEnabled(), false);
    const downloadCustom = (selector, name) => clickDownload(selector, name, custom);
    const pendingReportPath = await downloadCustom('#download-json', 'custom-before-repair.json');
    const before = JSON.parse(await fs.readFile(pendingReportPath, 'utf8'));
    assert.deepEqual(before.analysisOptions.fontMappings, { 'OfficeBridge Test Font': 'Arial' });
    assert.equal(before.repair, null);
    const originalReport = await analyzeOffice(customBytes, 'custom-font.xlsx');
    assert.equal(before.analysis.sha256, originalReport.sha256);
    await custom.locator('.file-card').filter({ hasText: 'another-file.xlsx' }).click();
    assert.equal(await custom.locator('.font-rule').count(), 0);
    assert.ok((await custom.locator('.font-map').allTextContents()).some((text) => text.includes('Microsoft YaHei')));
    const otherPath = await downloadCustom('#download-json', 'custom-other-file.json');
    assert.deepEqual(JSON.parse(await fs.readFile(otherPath, 'utf8')).analysisOptions.fontMappings, {});
    await custom.locator('.file-card').filter({ hasText: 'custom-font.xlsx' }).click();
    await custom.locator('.font-choice').filter({ hasText: 'OfficeBridge Test Font' }).locator('[data-font-group], [data-issue-id]').first().check();
    await custom.locator('#repair-file').click();
    await custom.locator('#download-repaired').waitFor();
    const afterPath = await downloadCustom('#download-json', 'custom-after-repair.json');
    const after = JSON.parse(await fs.readFile(afterPath, 'utf8'));
    assert.ok(after.repair.changes.length > 0);
    assert.ok(after.repair.changes.every((change) => change.before === 'OfficeBridge Test Font' && change.after === 'Arial'));
    const customOutput = await downloadCustom('#download-repaired', 'custom-repaired.xlsx');
    const customOutputReport = await analyzeOffice(new Uint8Array(await fs.readFile(customOutput)), 'custom-repaired.xlsx');
    assert.ok(!customOutputReport.inventory.fonts.some((font) => font.name === 'OfficeBridge Test Font'));
    assert.ok(customOutputReport.inventory.fonts.some((font) => font.name === 'Helvetica Neue'));
    const html = await fs.readFile(await downloadCustom('#download-html', 'custom-report.html'), 'utf8');
    assert.match(html, /用户为此文件指定的字体规则/);
    assert.match(html, /OfficeBridge Test Font/);
    assert.match(html, /只有修改记录中列出的替换已应用/);
    await custom.locator('#font-settings').scrollIntoViewIfNeeded();
    await custom.screenshot({ path: path.join(outputDir, 'advanced-font-rules.png') });
    await custom.locator('[data-remove-font="OfficeBridge Test Font"]').click();
    await custom.waitForFunction(() => document.querySelectorAll('.font-rule').length === 0 && !document.querySelector('#add-font-rule').disabled);
    assert.equal(await custom.locator('#download-repaired').count(), 0);
    assert.equal(await custom.locator('[data-issue-id]:checked').count(), 0);
    await custom.setViewportSize({ width: 320, height: 740 });
    assert.ok(await custom.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await custom.locator('#font-source').selectOption('OfficeBridge Test Font');
    await custom.locator('#font-target').fill('Wingdings');
    await custom.locator('#add-font-rule').click();
    await custom.locator('#font-settings .font-settings-error').waitFor();
    assert.equal(await custom.locator('.font-rule').count(), 0);
    assert.equal(await custom.locator('#repair-file').isEnabled(), false);
    await custom.close();
  });

  await test('样例后添加文件自动切换，原选择与副本保留，无效新增不切换', async () => {
    const firstUse = await context.newPage();
    await firstUse.goto(baseUrl);
    await firstUse.locator('#demo-xlsx').click();
    await firstUse.locator('#repair-file').waitFor();
    await firstUse.locator('#select-fixes').click();
    const selectedCount = await firstUse.locator('[data-issue-id]:checked').count();
    assert.ok(selectedCount > 0);
    await firstUse.locator('#repair-file').click();
    await firstUse.locator('#download-repaired').waitFor();
    const originalCopyName = await firstUse.locator('.repaired-name').innerText();
    await expandSection(firstUse, '.findings-details');
    await firstUse.locator('[data-filter="info"]').click();
    await firstUse.locator('#file-input').setInputFiles({ name: 'my-first-workbook.xlsx', mimeType: 'application/octet-stream', buffer: await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.xlsx')) });
    await firstUse.locator('.document-title h2').filter({ hasText: 'my-first-workbook.xlsx' }).waitFor();
    await firstUse.locator('#repair-file').waitFor();
    assert.equal(await firstUse.locator('.file-card[aria-pressed="true"] strong').innerText(), 'my-first-workbook.xlsx');
    assert.equal(await firstUse.locator('[data-filter="all"]').getAttribute('aria-pressed'), 'true');
    assert.equal(await firstUse.locator('[data-issue-id]:checked').count(), 0);
    assert.equal(await firstUse.locator('#download-repaired').count(), 0);
    await firstUse.locator('.file-card').filter({ hasText: '季度预算示例.xlsx' }).click();
    assert.equal(await firstUse.locator('[data-issue-id]:checked').count(), selectedCount);
    assert.equal(await firstUse.locator('.repaired-name').innerText(), originalCopyName);
    assert.equal(await firstUse.locator('#download-repaired').isVisible(), true);
    await firstUse.locator('#file-input').setInputFiles({ name: 'unsupported.txt', mimeType: 'text/plain', buffer: Buffer.from('synthetic') });
    await firstUse.locator('#input-errors').filter({ hasText: /unsupported\.txt.*\.pptx.*\.xlsx/i }).waitFor();
    assert.equal(await firstUse.locator('.document-title h2').innerText(), '季度预算示例.xlsx');
    assert.equal(await firstUse.locator('.file-card').count(), 2);
    assert.equal(await firstUse.locator('[data-issue-id]:checked').count(), selectedCount);
    assert.equal(await firstUse.locator('.repaired-name').innerText(), originalCopyName);
    await firstUse.close();
  });

  await test('晚到的样例不会抢走后来选的文件，也不会重新填入已清空的列表', async () => {
    const view = await context.newPage();
    await view.addInitScript(() => {
      const send = Worker.prototype.postMessage;
      Worker.prototype.postMessage = function (job, transfer) {
        if (job.type !== 'demo') return send.call(this, job, transfer);
        window.finishTestDemo = () => new Promise(resolve => {
          delete window.finishTestDemo;
          const onReply = ({ data }) => {
            if (data.id !== job.id) return;
            this.removeEventListener('message', onReply);
            setTimeout(resolve, 0);
          };
          this.addEventListener('message', onReply);
          send.call(this, job, transfer);
        });
      };
    });
    await view.goto(baseUrl);
    async function startSlowDemo() {
      await view.locator('#demo-xlsx').click();
      await view.waitForFunction(() => typeof window.finishTestDemo === 'function');
    }
    async function finishDemo() {
      await view.evaluate(() => window.finishTestDemo());
      await view.waitForFunction(() => !document.querySelector('#demo-xlsx').disabled);
    }
    await startSlowDemo();
    await view.locator('#file-input').setInputFiles(path.join(projectRoot, 'public/examples/officebridge-demo.pptx'));
    await view.locator('.result-summary').waitFor();
    await view.locator('#select-fixes').click();
    const selected = await view.locator('[data-issue-id]:checked').count();
    assert.ok(selected > 0);
    await finishDemo();
    assert.equal(await view.locator('.document-title h2').innerText(), 'officebridge-demo.pptx');
    assert.equal(await view.locator('[data-issue-id]:checked').count(), selected);
    await startSlowDemo();
    await view.locator('#clear-files').click();
    await finishDemo();
    assert.equal(await view.locator('.file-card').count(), 0);
    assert.equal(await view.locator('#workspace').isVisible(), false);
    assert.ok(await view.locator('#choose-files').evaluate(node => node.matches(':focus')));
    await view.close();
  });

  await test('生成期间切换文件，下载与修改记录仍属于最初选择的文档', async () => {
    const view = await context.newPage();
    await view.addInitScript(() => {
      // Hold only the repair request so a real tab switch can happen before its worker responds.
      const send = Worker.prototype.postMessage;
      Worker.prototype.postMessage = function (job, transfer) {
        if (job.type !== 'repair') return send.call(this, job, transfer);
        window.continueTestRepair = () => send.call(this, job, transfer);
      };
    });
    await view.goto(baseUrl);
    await view.locator('#file-input').setInputFiles(['xlsx', 'pptx'].map(ext => path.join(projectRoot, `public/examples/officebridge-demo.${ext}`)));
    await view.locator('#select-fixes').click();
    await view.locator('#repair-file').click();
    await view.waitForFunction(() => typeof window.continueTestRepair === 'function');
    await view.locator('.file-card').filter({ hasText: 'officebridge-demo.pptx' }).click();
    const download = view.waitForEvent('download');
    await view.evaluate(() => window.continueTestRepair());
    assert.equal((await download).suggestedFilename(), 'officebridge-demo.windows.xlsx');
    assert.equal(await view.locator('.document-title h2').innerText(), 'officebridge-demo.pptx');
    assert.equal(await view.locator('#download-repaired').count(), 0);
    const untouched = JSON.parse(await fs.readFile(await clickDownload('#download-json', 'untouched.json', view), 'utf8'));
    assert.equal(untouched.repair, null);
    await view.locator('.file-card').filter({ hasText: 'officebridge-demo.xlsx' }).click();
    assert.equal(await view.locator('.repaired-name').innerText(), 'officebridge-demo.windows.xlsx');
    await view.close();
  });

  await test('目标字体名单能筛选修复、保存再导入，每个文稿独立使用', async () => {
    const sheet = await context.newPage();
    await sheet.goto(baseUrl);
    await sheet.locator('#demo-xlsx').click();
    await sheet.locator('#font-settings > summary').click();
    const importList = buffer => sheet.locator('#font-list-input').setInputFiles({
      name: 'windows-fonts.txt', mimeType: 'text/plain', buffer,
    });
    await importList(Buffer.from('Arial\r\nPingFang SC\r\nPrivate QA unreferenced family'));
    await sheet.locator('#font-comparison').waitFor();
    const choices = sheet.locator('[data-issue-id]');
    assert.ok(await choices.count() > 0);
    assert.ok((await sheet.locator('.font-map').allTextContents()).every(s => s.includes('Arial') && !s.includes('Microsoft YaHei')));
    assert.equal(await sheet.locator('[data-issue-id]:checked').count(), 0);
    await sheet.setViewportSize({ width: 320, height: 740 });
    assert.ok(await sheet.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    const save = async (selector, file) => fs.readFile(await clickDownload(selector, file, sheet));
    const list = await save('#save-font-list', 'target-list.txt');
    assert.match(list.toString(), /Private QA unreferenced family/);
    await sheet.locator('#select-fixes').click();
    await sheet.locator('#repair-file').click();
    await sheet.locator('#download-repaired').waitFor();
    const json = (await save('#download-json', 'target-report.json')).toString();
    assert.doesNotMatch(json, /Private QA/);
    const payload = JSON.parse(json);
    assert.ok(payload.repair.changes.every(change => change.after === 'Arial'));
    assert.ok(payload.repair.report.inventory.targetFontComparison.listed.includes('Arial'));
    await importList(Buffer.from([0xff, 0x51]));
    await sheet.locator('#font-settings .font-settings-error').waitFor();
    assert.ok(await sheet.locator('#download-repaired').isVisible(), 'Invalid import must retain the last good output');
    await sheet.locator('#remove-font-list').click();
    await sheet.waitForFunction(() => !document.querySelector('#font-comparison') && !document.querySelector('#choose-font-list').disabled);
    assert.equal(await sheet.locator('#download-repaired').count(), 0);
    assert.ok((await sheet.locator('.font-map').allTextContents()).some(s => s.includes('Microsoft YaHei')));
    await importList(list);
    await sheet.locator('#font-comparison').waitFor();
    await sheet.locator('#demo-pptx').click();
    await sheet.locator('#font-settings > summary').click();
    assert.equal(await sheet.locator('#font-comparison').count(), 0);
    await sheet.close();
  });

  await test('浏览器拒绝读取或提供空名单时，保留当前核对并提示导入 TXT', async () => {
    // Edge can resolve to an empty array on denial. Other browsers reject.
    // Simulate both responses here; real OS enumeration is a separate VM test.
    const denied = await context.newPage();
    await denied.addInitScript(() => {
      Object.defineProperty(navigator, 'userAgent', { get: () => 'Windows test browser' });
      let calls = 0;
      Object.defineProperty(window, 'queryLocalFonts', { value: async () => {
        if (++calls === 1) return [];
        throw new DOMException('Permission denied', 'NotAllowedError');
      } });
    });
    await denied.goto(baseUrl);
    await denied.locator('#demo-xlsx').click();
    await denied.locator('#font-settings > summary').click();
    await denied.locator('#font-list-input').setInputFiles({
      name: 'known.txt', mimeType: 'text/plain', buffer: Buffer.from('Arial'),
    });
    await denied.locator('#font-comparison').waitFor();
    for (const expected of ['没有提供可读取的字体', '未提供本机字体']) {
      await denied.locator('#read-local-fonts').click();
      await denied.waitForFunction(msg => document.querySelector('#font-settings .font-settings-error')?.textContent.includes(msg), expected);
      assert.match(await denied.locator('#font-settings .font-settings-error').innerText(), /导入字体名单/);
      assert.match(await denied.locator('#font-comparison').innerText(), /已核对 1 个/);
    }
    await denied.close();
  });

  await test('快捷选择能撤销高影响勾选，下载后原文本框保留原字体', async () => {
    const sample = await context.newPage();
    const parts = unzipSync(await fs.readFile(path.join(projectRoot, 'public/examples/officebridge-demo.pptx')));
    const key = 'ppt/slides/slide1.xml';
    const before = strFromU8(parts[key]);
    const box = '<p:sp><p:nvSpPr><p:cNvPr id="90" name="窄框样例"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="650000" cy="300000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="2200"><a:latin typeface="PingFang SC"/></a:rPr><a:t>QA lines</a:t></a:r></a:p></p:txBody></p:sp>';
    parts[key] = strToU8(before.replace('</p:spTree>', box + '</p:spTree>'));
    await sample.goto(baseUrl);
    await sample.locator('#file-input').setInputFiles({ name: '对照稿.pptx', mimeType: 'application/octet-stream', buffer: Buffer.from(zipSync(parts)) });
    await sample.locator('#select-fixes').click();
    const all = await sample.locator('[data-issue-id]:checked').count();
    await sample.locator('#select-lower-risk').click();
    const kept = await sample.locator('[data-issue-id]:checked').count();
    assert.ok(kept > 0 && kept < all);
    assert.match(await sample.locator('.font-choice-list').innerText(), /窄框样例/);
    const downloaded = await clickDownload('#repair-file', 'layout-choice.pptx', sample);
    const actual = strFromU8(unzipSync(await fs.readFile(downloaded))[key]);
    assert.ok(actual.includes(box));
    assert.ok(actual.includes('typeface="Microsoft YaHei"'));
    await sample.locator('#compare-result').click();
    assert.equal(await sample.locator('#pdf-compare').evaluate(el => el.open), true);
    await sample.close();
  });

  await test('两份 PDF 本地逐页对照，能区分同页、换行位置变化和缺页', async () => {
    const sample = await context.newPage();
    // Generate deterministic print proofs in memory, with no Office dependency.
    async function proof(pages, offset) {
      await sample.setContent(`<style>@page{size:210mm 297mm;margin:15mm}body{margin:0}section+section{break-before:page}h1{font:24pt Arial;margin-left:${offset}mm}</style>`
        + Array.from({ length: pages }, (_, i) => `<section><h1>Row ${i + 1}</h1></section>`).join(''));
      return sample.pdf({ preferCSSPageSize: true });
    }
    const original = await proof(2, 0);
    const different = await proof(1, 18);
    await sample.close();
    const compare = await context.newPage();
    await compare.goto(baseUrl);
    await compare.locator('#pdf-compare > summary').click();
    const attach = (id, name, buffer) => compare.locator('#proof-' + id).setInputFiles({ name, mimeType: 'application/pdf', buffer });
    const verdict = text => compare.waitForFunction(value => document.querySelector('#proof-verdict')?.innerText.includes(value), text);
    await attach(0, '原稿.pdf', original);
    await attach(1, '原稿复核.pdf', original);
    await verdict('未见明显图像差异');
    assert.match(await compare.locator('#proof-status').innerText(), /2 页.*2 页/);
    await compare.locator('#proof-scan').click();
    await compare.waitForFunction(() => document.querySelector('#proof-scan-status').textContent.includes('已核对 2 / 2 页'));
    assert.match(await compare.locator('#proof-scan-status').innerText(), /0 页有差异/);
    assert.equal(await compare.locator('.proof-page-result.unchanged').count(), 2);
    await attach(1, '位移.pdf', different);
    await verdict('本页预览有差异');
    assert.equal(await compare.locator('.proof-page-result').count(), 0);
    await compare.locator('#proof-scan').click();
    await compare.waitForFunction(() => document.querySelector('#proof-scan-status').textContent.includes('已核对 2 / 2 页'));
    assert.match(await compare.locator('#proof-scan-status').innerText(), /2 页有差异/);
    await compare.getByRole('button', {name:'查看第 2 页：一份缺页', exact:true}).click();
    await verdict('只有一份文件');
    await compare.getByRole('button', {name:'查看第 1 页：预览有差异', exact:true}).click();
    await verdict('本页预览有差异');
    assert.match(await compare.locator('#proof-status').innerText(), /少 1 页/);
    const normalWidth = (await compare.locator('.proof-paper').first().boundingBox()).width;
    await compare.locator('#proof-zoom').check();
    assert.ok((await compare.locator('.proof-paper').first().boundingBox()).width > normalWidth);
    await compare.locator('#proof-overlay').check();
    assert.equal(await compare.locator('.proof-paper').nth(1).locator('canvas:visible').getAttribute('aria-label'), 'Windows 预览，差异用玫红色标出');
    await compare.locator('#proof-next').click();
    await verdict('只有一份文件');
    assert.match(await compare.locator('.proof-paper').nth(1).innerText(), /没有这一页/);
    await compare.locator('#proof-prev').click();
    await verdict('本页预览有差异');
    await attach(1, '损坏.pdf', Buffer.from('damaged'));
    await compare.waitForFunction(() => document.querySelectorAll('.proof-info')[1].textContent.includes('无法读取'));
    assert.equal(await compare.locator('#proof-view').isVisible(), false);
    await attach(1, '再选.pdf', original);
    await verdict('未见明显图像差异');
    await compare.setViewportSize({ width: 320, height: 740 });
    assert.ok(await compare.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await compare.locator('#proof-reset').click();
    assert.equal(await compare.locator('#proof-view').isVisible(), false);
    assert.equal(await compare.locator('#proof-0').evaluate(input => input.files.length), 0);
    await compare.close();
  });

  await test('PDF 缺少图片像素时保留页数，停止给出排版结论，换回正常图片仍可比较', async () => {
    function imageProof(payload, filter = '', size = 64) {
      const paint = 'q 150 0 0 150 0 0 cm /im Do Q';
      const objects = [
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 150 150] /Contents 4 0 R /Resources << /XObject << /im 5 0 R >> >> >>',
        `<< /Length ${paint.length} >>\nstream\n${paint}\nendstream`,
        Buffer.concat([Buffer.from(`<< /Type /XObject /Subtype /Image /Width ${size} /Height ${size} /ColorSpace /DeviceRGB /BitsPerComponent 8 ${filter} /Length ${payload.length} >>\nstream\n`), payload, Buffer.from('\nendstream')]),
      ];
      const chunks = [Buffer.from('%PDF-1.7\n')];
      let offset = chunks[0].length;
      const xref = ['0000000000 65535 f '];
      for (const [index, object] of objects.entries()) {
        xref.push(String(offset).padStart(10, '0') + ' 00000 n ');
        const bytes = Buffer.concat([Buffer.from(`${index + 1} 0 obj\n`), Buffer.from(object), Buffer.from('\nendobj\n')]);
        chunks.push(bytes); offset += bytes.length;
      }
      chunks.push(Buffer.from(`xref\n0 ${xref.length}\n${xref.join('\n')}\ntrailer\n<< /Root 1 0 R /Size ${xref.length} >>\nstartxref\n${offset}\n%%EOF\n`));
      return Buffer.concat(chunks);
    }
    // Two valid OpenJPEG JPEG 2000 codestreams for solid 64 × 64 RGB images.
    // The offline PDF reader omits this decoder. The decoded pages must never
    // be classified as matching empty canvases.
    const red = '/0//UQAvAAAAAABAAAAAQAAAAAAAAAAAAAAAQAAAAEAAAAAAAAAAAAADBwEBBwEBBwEB/1IADAAAAAEABQQEAAH/XAATQEBISFBISFBISFBISFBISFD/ZAAlAAFDcmVhdGVkIGJ5IE9wZW5KUEVHIHZlcnNpb24gMi41LjT/kAAKAAAAAAAtAAH/k8+0DAHKv9+AEAiQ34AQCJCAgICAgICAgICAgICAgID/2Q==';
    const blue = '/0//UQAvAAAAAABAAAAAQAAAAAAAAAAAAAAAQAAAAEAAAAAAAAAAAAADBwEBBwEBBwEB/1IADAAAAAEABQQEAAH/XAATQEBISFBISFBISFBISFBISFD/ZAAlAAFDcmVhdGVkIGJ5IE9wZW5KUEVHIHZlcnNpb24gMi41LjT/kAAKAAAAAAAtAAH/k9+AEAiQ34AQCJDPtAwByr+AgICAgICAgICAgICAgID/2Q==';
    const compare = await context.newPage();
    await compare.addInitScript(() => {
      const Original = window.Worker;
      window.Worker = class extends Original {
        postMessage(message, ...rest) {
          if (window.__stallNextPdf && message?.action === 'GetOperatorList') {
            window.__stallNextPdf = false;
            this.stalled = true;
            return;
          }
          return super.postMessage(message, ...rest);
        }
        terminate() {
          if (this.stalled) window.__stalledPdfTerminated = true;
          return super.terminate();
        }
      };
      const timer = window.setTimeout.bind(window);
      window.setTimeout = (callback, delay, ...rest) => timer(callback,
        window.__testPdfDeadline && delay === 30000 ? 1500 : delay, ...rest);
    });
    await compare.goto(baseUrl);
    await compare.locator('#pdf-compare > summary').click();
    const attach = (id, name, buffer) => compare.locator('#proof-' + id).setInputFiles({ name, mimeType: 'application/pdf', buffer });
    const verdict = text => compare.waitForFunction(value => document.querySelector('#proof-verdict').textContent.includes(value), text);
    const rawRed = imageProof(Buffer.from([255, 0, 0]), '', 1);
    await attach(0, '红.pdf', imageProof(Buffer.from(red, 'base64'), '/Filter /JPXDecode'));
    await attach(1, '蓝.pdf', imageProof(Buffer.from(blue, 'base64'), '/Filter /JPXDecode'));
    await verdict('无法判断排版差异');
    await compare.locator('#proof-scan').click();
    await compare.waitForFunction(() => document.querySelector('#proof-scan-status').textContent.includes('已核对 1 / 1 页'));
    assert.match(await compare.locator('#proof-scan-status').innerText(), /1 页有差异或需要人工确认/);
    assert.equal(await compare.locator('.proof-page-result.unchanged').count(), 0);
    assert.match(await compare.locator('.proof-page-result').innerText(), /未能完整读取/);
    assert.match(await compare.locator('#proof-status').innerText(), /1 页.*1 页/);
    assert.ok(await compare.locator('#proof-overlay-control').isHidden());
    await attach(0, '普通红.pdf', rawRed);
    await attach(1, '损坏图片.pdf', imageProof(Buffer.from('broken JPEG'), '/Filter /DCTDecode'));
    await verdict('无法判断排版差异');
    await attach(1, '超大图.pdf', imageProof(Buffer.from([0]), '', 8000));
    await verdict('无法判断排版差异');
    await attach(1, '普通蓝.pdf', imageProof(Buffer.from([0, 0, 255]), '', 1));
    await verdict('本页预览有差异');
    await attach(1, '同一图.pdf', rawRed);
    await verdict('未见明显图像差异');
    await compare.evaluate(() => { window.__stallNextPdf = window.__testPdfDeadline = true; });
    await attach(1, '短暂卡住.pdf', rawRed);
    await compare.waitForFunction(() => window.__stalledPdfTerminated);
    const stalled = await compare.locator('.proof-info').allTextContents();
    assert.equal(stalled.filter(message => message.includes('预览耗时过长')).length, 1);
    const retry = stalled.findIndex(message => message.includes('预览耗时过长'));
    assert.ok((await compare.locator('#proof-' + (1 - retry)).evaluate(input => input.files.length)) === 1);
    await compare.evaluate(() => { window.__testPdfDeadline = false; });
    await attach(retry, '重试成功.pdf', rawRed);
    await verdict('未见明显图像差异');
    await compare.close();
  });

  await test('连续逐页核对可以停止，更换 PDF 会清除旧结果', async () => {
    const source = await context.newPage();
    await source.setContent('<style>@page{size:A4}section+section{break-before:page}</style>' + Array.from({length:30},(_,i)=>`<section><h1>Page ${i+1}</h1></section>`).join(''));
    const pdf = await source.pdf({preferCSSPageSize:true});
    await source.close();
    const view = await context.newPage();
    await view.addInitScript(() => {
      const original = window.setTimeout.bind(window);
      window.setTimeout = (callback, delay, ...rest) => original(callback,
        delay === 30000 && window.__testStaleDeadline ? 1500 : delay === 0 && window.__slowPageScan ? 200 : delay, ...rest);
      const WorkerClass = window.Worker;
      window.Worker = class extends WorkerClass {
        postMessage(message, ...rest) {
          if (window.__stallSecondPage && message?.action === 'GetOperatorList' && message.data?.pageIndex === 1) {
            window.__stallSecondPage = false; window.__pageStalled = true; return;
          }
          return super.postMessage(message, ...rest);
        }
      };
    });
    await view.goto(baseUrl);
    await view.locator('#pdf-compare > summary').click();
    for (const side of [0,1]) await view.locator('#proof-'+side).setInputFiles({name:'many-pages.pdf',mimeType:'application/pdf',buffer:pdf});
    await view.waitForFunction(()=>document.querySelector('#proof-verdict').textContent.includes('未见明显'));
    await view.evaluate(()=>{window.__slowPageScan=true;});
    await view.locator('#proof-scan').click();
    await view.locator('.proof-page-result').first().waitFor();
    await view.locator('#proof-scan-stop').click();
    assert.match(await view.locator('#proof-scan-status').innerText(),/已停止/);
    assert.ok(await view.locator('.proof-page-result').count() < 30);
    await view.locator('#proof-1').setInputFiles({name:'replacement.pdf',mimeType:'application/pdf',buffer:pdf});
    await view.waitForFunction(()=>document.querySelector('#proof-verdict').textContent.includes('未见明显'));
    assert.equal(await view.locator('.proof-page-result').count(),0);
    assert.equal(await view.locator('#proof-scan-status').innerText(),'');
    assert.equal(await view.locator('#proof-scan').isEnabled(),true);
    await view.evaluate(()=>{window.__testStaleDeadline=window.__stallSecondPage=true;});
    await view.locator('#proof-page').fill('2');
    await view.locator('#proof-page').dispatchEvent('change');
    await view.waitForFunction(()=>window.__pageStalled);
    await view.locator('#proof-page').fill('1');
    await view.locator('#proof-page').dispatchEvent('change');
    await view.waitForFunction(()=>document.querySelector('#proof-verdict').textContent.includes('未见明显'));
    await view.evaluate(()=>new Promise(resolve=>setTimeout(resolve,1800)));
    assert.doesNotMatch((await view.locator('.proof-info').allTextContents()).join(''),/耗时过长/);
    assert.match(await view.locator('#proof-verdict').innerText(),/未见明显/);
    await view.close();
  });

  await test('两个接收系统独立检查，改变方向后清空旧选择并生成 Mac 副本', async () => {
    const reverse = await context.newPage();
    await reverse.goto(baseUrl);
    await reverse.locator('#target-platform').selectOption('mac-office');
    await reverse.locator('#demo-xlsx').click();
    await reverse.locator('#repair-file').waitFor();
    assert.equal(await reverse.locator('#record-target').inputValue(), 'mac-office');
    assert.match(await reverse.locator('.font-choice-list').innerText(), /Microsoft YaHei|Segoe UI/);
    await reverse.locator('#select-lower-risk').click();
    const output = await clickDownload('#repair-file', 'mac-demo.xlsx', reverse);
    assert.match(await reverse.locator('.repaired-name').innerText(), /\.mac\.xlsx$/);
    const report = await analyzeOffice(new Uint8Array(await fs.readFile(output)), 'mac-demo.xlsx', {fontProfile:'mac-office'});
    assert.equal(report.analysisOptions.fontProfile, 'mac-office');
    assert.equal(report.repairBlocked, false);
    await reverse.locator('#record-target').selectOption('windows-office');
    await reverse.waitForFunction(() => document.querySelector('#record-target')?.value === 'windows-office' && !document.querySelector('#record-target').disabled);
    assert.equal(await reverse.locator('[data-issue-id]:checked').count(), 0);
    assert.equal(await reverse.locator('#download-repaired').count(), 0);
    await reverse.close();
  });

  await test('本机 Office 助手结果自动显示真实 PDF，受损记录可恢复且文件名作为文字显示', async () => {
    const sample = await context.newPage();
    await sample.setContent('<p style="font:24px sans-serif">OfficeBridge native proof 2026</p>');
    const pdf = await sample.pdf();
    await sample.close();
    const report = {schema:'officebridge-native-check',schemaVersion:1,createdAt:'2026-11-12T00:00:00Z',environment:{platform:'macos'},documents:[
      {role:'original',format:'xlsx',fileName:'Budget.xlsx',sha256:'a'.repeat(64),openedSuccessfully:true,sourceUnchanged:true,officeApplication:'Excel',officeVersion:'QA',
        pdf:{fileName:'before.pdf',base64:pdf.toString('base64'),sha256:createHash('sha256').update(pdf).digest('hex')}},
      {role:'copy',format:'xlsx',fileName:'Budget.mac.xlsx',sha256:'b'.repeat(64),openedSuccessfully:true,sourceUnchanged:true,officeApplication:'Excel',officeVersion:'QA',
        pdf:{fileName:'after.pdf',base64:pdf.toString('base64'),sha256:createHash('sha256').update(pdf).digest('hex')}}],
      comparisons:[{details:['The tested worksheet has the same page count.']}]};
    const view = await context.newPage();
    await view.goto(baseUrl);
    await view.locator('#native-check > summary').click();
    await view.locator('#native-check details > summary').click();
    const attach = data => view.locator('#native-report-input').setInputFiles({name:'check.officebridge-check.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(data))});
    report.environment.fontFamilies=['Microsoft YaHei','Segoe UI','PingFang SC','Helvetica Neue','Arial','Calibri'];
    await attach(report);
    await view.waitForFunction(() => document.querySelector('#proof-verdict')?.innerText.includes('未见明显'));
    assert.match(await view.locator('#native-results').innerText(), /Mac Office.*已成功打开/s);
    assert.match(await view.locator('#proof-status').innerText(), /Mac 原件.*Mac 副本/);
    await view.locator('.native-fonts > summary').click();
    const fontFile = await clickDownload('#native-font-download','native-office-fonts.txt',view);
    assert.match(await fs.readFile(fontFile,'utf8'),/Microsoft YaHei/);
    await view.locator('#target-platform').selectOption('mac-office');
    assert.match(await view.locator('label[for="proof-1"]').innerText(),/Mac 副本/);
    await view.locator('#demo-xlsx').click();
    await view.locator('#font-settings > summary').click();
    await view.locator('#font-list-input').setInputFiles(fontFile);
    await view.waitForFunction(()=>document.querySelector('#font-comparison')?.textContent.includes('已核对 6 个'));
    assert.equal(await view.locator('[data-issue-id]').count(),0,'Available Office fonts must not be offered for unnecessary replacement');
    const bad = structuredClone(report);
    bad.documents[0].pdf.sha256 = '0'.repeat(64);
    await attach(bad);
    await view.waitForFunction(() => document.querySelector('#native-import-status').innerText.includes('指纹不一致'));
    report.documents[0].fileName = '<img src=x onerror=alert(1)>.xlsx';
    await attach(report);
    await view.waitForFunction(() => document.querySelector('#native-import-status').innerText.includes('已打开'));
    assert.equal(await view.locator('#native-results img').count(),0);
    assert.match(await view.locator('#native-results').innerText(), /<img/);
    const unread = structuredClone(report);
    Object.assign(unread.documents[0],{openedSuccessfully:false,sourceUnchanged:false,sha256:null,pdf:null,error:'文件被占用，请关闭后重新核对。'});
    await attach(unread);
    await view.waitForFunction(()=>document.querySelector('#native-results').textContent.includes('文件被占用'));
    assert.match(await view.locator('#native-results').innerText(),/未完成项目/);
    assert.match(await view.locator('label[for="proof-1"]').innerText(),/Mac 副本/);
    assert.equal(await view.locator('#proof-0').evaluate(e=>e.files.length),0);
    await attach(report);
    await view.waitForFunction(()=>document.querySelector('#proof-verdict').textContent.includes('未见明显'));
    await view.setViewportSize({width:320,height:740});
    assert.ok(await view.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await view.close();
    if (isLocalFile) {
      const html = await fs.readFile(new URL(baseUrl), 'utf8');
      const end = html.lastIndexOf('</body>');
      const safe = JSON.stringify(report).replaceAll('<', '\\u003c');
      const resultPath = path.join(outputDir, 'native-result.html');
      await fs.writeFile(resultPath, html.slice(0,end) + '<script type="application/json" id="officebridge-native-data">' + safe + '</script>' + html.slice(end));
      const automatic = await context.newPage();
      await automatic.goto(pathToFileURL(resultPath).href);
      await automatic.waitForFunction(() => document.querySelector('#proof-verdict')?.innerText.includes('未见明显'));
      assert.equal(await automatic.locator('#officebridge-native-data').count(),0);
      assert.match(await automatic.locator('#native-results').innerText(), /<img/);
      await automatic.close();
    }
  });

  await test('检查及修复过程没有文档上传或第三方请求', async () => {
    const base = new URL(baseUrl);
    const localReports = new Set(['report.html', 'long-path-report.html', 'native-result.html'].map((name) => pathToFileURL(path.join(outputDir, name)).href));
    const excludedBrowserInternalRequests = requests.filter(isBrowserInternalRequest);
    const checkedRequests = requests.filter((request) => !isBrowserInternalRequest(request));
    const unexpectedRequests = checkedRequests.filter((request) => !isExpectedAppRequest(request, base, localReports));
    const externalNetworkRequests = unexpectedRequests.filter((request) => /^(?:https?|wss?):/i.test(request.url));
    const developmentServerWebSockets = webSocketConnections.filter((connection) => isDevelopmentServerWebSocket(connection.url, base));
    const unexpectedWebSockets = webSocketConnections.filter((connection) => !isDevelopmentServerWebSocket(connection.url, base));
    const unexpectedWebSocketFrames = developmentServerWebSockets.flatMap((connection) => connection.sentFrames
      .filter((frame) => !isVitePingFrame(frame)).map((frame) => ({ url: connection.url, ...frame })));
    requestAudit = {
      observedRequestCount: requests.length,
      checkedRequestCount: checkedRequests.length,
      excludedBrowserInternalRequestCount: excludedBrowserInternalRequests.length,
      excludedBrowserInternalRequests,
      externalNetworkRequestCount: externalNetworkRequests.length,
      externalNetworkRequests,
      webSocketConnectionCount: webSocketConnections.length,
      webSocketConnections,
      developmentServerWebSocketCount: developmentServerWebSockets.length,
      developmentServerWebSockets,
      unexpectedWebSockets,
      unexpectedWebSocketFrames,
      unexpectedRequests,
    };
    await fs.writeFile(path.join(outputDir, 'request-audit.json'), JSON.stringify(requestAudit, null, 2));
    console.log(`Request audit: checked ${checkedRequests.length}; excluded ${excludedBrowserInternalRequests.length} Edge internal UI requests; ${developmentServerWebSockets.length} development-server WebSockets.`);
    assert.deepEqual(requests.filter((request) => request.method !== 'GET'), []);
    assert.deepEqual(unexpectedRequests, []);
    assert.deepEqual(unexpectedWebSockets, []);
    assert.deepEqual(unexpectedWebSocketFrames, []);
    assert.deepEqual(failures, []);
  });
  const appSha256 = isLocalFile ? createHash('sha256').update(await fs.readFile(fileURLToPath(baseUrl))).digest('hex') : null;
  await fs.writeFile(path.join(outputDir, 'result.json'), JSON.stringify({ passed: checks.length, checks, browserVersion: browser.version(), offline: isLocalFile, appSha256, pageErrors: failures, checkedRequestCount: requestAudit.checkedRequestCount, observedRequestCount: requestAudit.observedRequestCount, externalNetworkRequestCount: requestAudit.externalNetworkRequestCount, excludedBrowserInternalRequestCount: requestAudit.excludedBrowserInternalRequestCount, webSocketConnectionCount: requestAudit.webSocketConnectionCount, developmentServerWebSocketCount: requestAudit.developmentServerWebSocketCount, requestAuditFile: 'request-audit.json' }, null, 2));
  console.log(`Browser smoke checks passed: ${checks.length}. Artifacts: ${outputDir}`);
} finally {
  await browser.close();
}
