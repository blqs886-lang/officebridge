import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { analyzeOffice } from '../src/engine/index.js';
import { unzipSync, zipSync, strToU8 } from 'fflate';
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
const outputDir = path.join(projectRoot, 'work', 'browser-smoke');
await fs.mkdir(outputDir, { recursive: true });
const browser = await chromium.launch({ headless: true, chromiumSandbox: true, ...(process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {}) });
const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
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
const clickDownload = async (selector, name) => {
  const downloadPromise = page.waitForEvent('download');
  await page.locator(selector).click();
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
    assert.equal(await page.title(), 'OfficeBridge · 文档兼容工作台');
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
  await test('选择建议后生成并下载可编辑 XLSX 副本', async () => {
    await page.locator('#select-fixes').click();
    assert.equal(await page.locator('#repair-file').isEnabled(), true);
    await page.locator('#repair-file').click();
    await page.locator('#download-repaired').waitFor();
    repairedPath = await clickDownload('#download-repaired', 'repaired.xlsx');
    const report = await analyzeOffice(new Uint8Array(await fs.readFile(repairedPath)), 'repaired.xlsx');
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
    await page.locator('#repair-file').click();
    await page.locator('#download-repaired').waitFor();
    const output = await clickDownload('#download-repaired', 'repaired.pptx');
    const report = await analyzeOffice(new Uint8Array(await fs.readFile(output)), 'repaired.pptx');
    assert.equal(report.repairBlocked, false);
    assert.equal(report.issues.filter((issue) => issue.fixable).length, 0);
    assert.equal(report.inventory.slideCount, 2);
  });

  await test('不支持、空文件和损坏文件给出可读错误', async () => {
    await page.locator('#file-input').setInputFiles({ name: 'unsupported.txt', mimeType: 'text/plain', buffer: Buffer.from('synthetic') });
    await page.locator('#toast').filter({ hasText: '目前只支持 PPTX 和 XLSX' }).waitFor();
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
    await page.getByText('此文件不适合自动修复').waitFor();
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
    const downloadCustom = async (selector, name) => {
      const waiting = custom.waitForEvent('download');
      await custom.locator(selector).click();
      const download = await waiting;
      const file = path.join(outputDir, name);
      await download.saveAs(file);
      return file;
    };
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
    await custom.locator('.issue-card').filter({ hasText: '字体「OfficeBridge Test Font」' }).locator('[data-issue-id]').check();
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
    await custom.locator('.font-settings-error').waitFor();
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
    await firstUse.locator('#toast').filter({ hasText: '目前只支持 PPTX 和 XLSX' }).waitFor();
    assert.equal(await firstUse.locator('.document-title h2').innerText(), '季度预算示例.xlsx');
    assert.equal(await firstUse.locator('.file-card').count(), 2);
    assert.equal(await firstUse.locator('[data-issue-id]:checked').count(), selectedCount);
    assert.equal(await firstUse.locator('.repaired-name').innerText(), originalCopyName);
    await firstUse.close();
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
    const save = async (selector, file) => {
      const waiting = sheet.waitForEvent('download');
      await sheet.locator(selector).click();
      const download = await waiting;
      const target = path.join(outputDir, file);
      await download.saveAs(target);
      return fs.readFile(target);
    };
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
    await sheet.locator('.font-settings-error').waitFor();
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
      await denied.waitForFunction(msg => document.querySelector('.font-settings-error')?.textContent.includes(msg), expected);
      assert.match(await denied.locator('.font-settings-error').innerText(), /导入字体名单/);
      assert.match(await denied.locator('#font-comparison').innerText(), /已核对 1 个/);
    }
    await denied.close();
  });

  await test('检查及修复过程没有文档上传或第三方请求', async () => {
    const base = new URL(baseUrl);
    const localReports = new Set(['report.html', 'long-path-report.html'].map((name) => pathToFileURL(path.join(outputDir, name)).href));
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
    console.log(`Request audit: checked ${checkedRequests.length}; excluded ${excludedBrowserInternalRequests.length} Edge internal UI requests; ${developmentServerWebSockets.length} development-server WebSockets. Full list: work/browser-smoke/request-audit.json`);
    assert.deepEqual(requests.filter((request) => request.method !== 'GET'), []);
    assert.deepEqual(unexpectedRequests, []);
    assert.deepEqual(unexpectedWebSockets, []);
    assert.deepEqual(unexpectedWebSocketFrames, []);
    assert.deepEqual(failures, []);
  });
  await fs.writeFile(path.join(outputDir, 'result.json'), JSON.stringify({ passed: checks.length, checks, pageErrors: failures, checkedRequestCount: requestAudit.checkedRequestCount, observedRequestCount: requestAudit.observedRequestCount, externalNetworkRequestCount: requestAudit.externalNetworkRequestCount, excludedBrowserInternalRequestCount: requestAudit.excludedBrowserInternalRequestCount, webSocketConnectionCount: requestAudit.webSocketConnectionCount, developmentServerWebSocketCount: requestAudit.developmentServerWebSocketCount, requestAuditFile: 'request-audit.json' }, null, 2));
  console.log(`Browser smoke checks passed: ${checks.length}. Artifacts: work/browser-smoke/`);
} finally {
  await browser.close();
}
