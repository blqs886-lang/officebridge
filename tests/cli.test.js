import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { mkdir, mkdtemp, writeFile, readFile, readdir, rm, access } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { createDemo } from '../src/engine/demo.js';
import { readArchive, writeArchive } from '../src/engine/archive.js';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';

const execute = promisify(execFile);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const script = path.join(root, 'scripts/cli.mjs');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');

async function temporary(t) {
  await mkdir(path.join(root, 'work'), { recursive: true });
  const directory = await mkdtemp(path.join(root, 'work', 'cli-tests-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  return directory;
}

async function source(directory, name, bytes = createDemo('pptx')) {
  const file = path.join(directory, name);
  await mkdir(path.dirname(file), { recursive: true });
  await writeFile(file, bytes, { flag: 'wx' });
  return file;
}

async function cli(directory, args) {
  try {
    const result = await execute(process.execPath, [script, ...args], { cwd: directory, encoding: 'utf8', timeout: 20_000, maxBuffer: 256 * 1024 });
    return { code: 0, ...result };
  } catch (error) {
    if (typeof error.code !== 'number') throw error;
    return { code: error.code, stdout: error.stdout, stderr: error.stderr };
  }
}

async function exists(file) {
  try { await access(file); return true; } catch { return false; }
}

async function json(file) { return JSON.parse(await readFile(file, 'utf8')); }

test('CLI help describes explicit repair consent and does not create output files', async t => {
  const directory = await temporary(t);
  const result = await cli(directory, ['--help']);
  assert.equal(result.code, 0);
  assert.match(result.stdout, /--accept-font-substitutions/);
  assert.match(result.stdout, /未.*Windows Office|不等同于 Windows Office/);
  assert.deepEqual(await readdir(directory), []);
});

test('default check handles Chinese names and spaces, records the source hash and leaves the original intact', async t => {
  const directory = await temporary(t);
  const bytes = createDemo('pptx');
  const input = await source(directory, '资料 目录/中文 演示.PPTX', bytes);
  const result = await cli(directory, [input]);
  assert.equal(result.code, 0, result.stderr);
  const report = await json(path.join(directory, 'officebridge-reports/中文 演示.PPTX.report.json'));
  assert.equal(report.sha256, hash(bytes));
  assert.equal(report.validation.windowsOfficeTested, false);
  assert.equal(report.format, 'pptx');
  assert.equal(hash(await readFile(input)), hash(bytes));
  assert.deepEqual(await readdir(path.join(directory, 'officebridge-reports')), ['中文 演示.PPTX.report.json']);
});

test('batch check distinguishes identical stems with different Office extensions', async t => {
  const directory = await temporary(t);
  const deck = await source(directory, 'summary.pptx');
  const workbook = await source(directory, 'summary.xlsx', createDemo('xlsx'));
  const result = await cli(directory, ['check', deck, workbook, '--out', '报告 输出']);
  assert.equal(result.code, 0, result.stderr);
  assert.equal((await json(path.join(directory, '报告 输出/summary.pptx.report.json'))).format, 'pptx');
  assert.equal((await json(path.join(directory, '报告 输出/summary.xlsx.report.json'))).format, 'xlsx');
});

test('repair without explicit acceptance fails before creating a directory or changing the source', async t => {
  const directory = await temporary(t);
  const bytes = createDemo('pptx');
  const input = await source(directory, 'deck.pptx', bytes);
  const result = await cli(directory, ['repair', input, '--out', 'would-be-created']);
  assert.equal(result.code, 2);
  assert.match(result.stderr, /--accept-font-substitutions/);
  assert.equal(await exists(path.join(directory, 'would-be-created')), false);
  assert.equal(hash(await readFile(input)), hash(bytes));
});

test('accepted repair creates an editable uppercase-extension copy and a provenance report', async t => {
  const directory = await temporary(t);
  const bytes = createDemo('pptx');
  const input = await source(directory, '演示 File.PPTX', bytes);
  const result = await cli(directory, ['repair', input, '--accept-font-substitutions', '--out=交付 文件']);
  assert.equal(result.code, 0, result.stderr);
  const repaired = await readFile(path.join(directory, '交付 文件/演示 File.windows.PPTX'));
  const report = await json(path.join(directory, '交付 文件/演示 File.windows.PPTX.report.json'));
  assert.equal(report.sha256, hash(repaired));
  assert.equal(report.repair.originalSha256, hash(bytes));
  assert.ok(report.repair.changes.length > 0);
  assert.equal(report.validation.windowsOfficeTested, false);
  assert.equal(report.repair.unchangedPartsVerified, true);
  assert.equal(hash(await readFile(input)), hash(bytes));
  assert.ok(readArchive(repaired).entries.has('ppt/slides/slide1.xml'));
});

test('signed documents receive a blocked report and never produce a repaired copy', async t => {
  const directory = await temporary(t);
  const { entries } = readArchive(createDemo('pptx'));
  entries.set('_xmlsignatures/sig1.xml', new TextEncoder().encode('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>'));
  const input = await source(directory, 'signed.pptx', writeArchive(entries));
  const result = await cli(directory, ['repair', input, '--accept-font-substitutions', '--out', 'reports']);
  assert.equal(result.code, 1);
  const report = await json(path.join(directory, 'reports/signed.pptx.report.json'));
  assert.equal(report.repairBlocked, true);
  assert.equal(await exists(path.join(directory, 'reports/signed.windows.pptx')), false);
});

test('repair without remaining font suggestions saves only an honest check report', async t => {
  const directory = await temporary(t);
  const bytes = createDemo('pptx');
  const report = await analyzeOffice(bytes, 'clean.pptx');
  const repaired = await repairOffice(bytes, 'clean.pptx', { selectedIssueIds: report.issues.filter(issue => issue.fixable).map(issue => issue.id) });
  const input = await source(directory, 'clean.pptx', repaired.bytes);
  const result = await cli(directory, ['repair', input, '--accept-font-substitutions']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /没有可应用的字体建议/);
  assert.deepEqual(await readdir(path.join(directory, 'officebridge-reports')), ['clean.pptx.report.json']);
  assert.equal((await json(path.join(directory, 'officebridge-reports/clean.pptx.report.json'))).repair, undefined);
});

test('existing reports are preserved and the batch continues to a different valid file', async t => {
  const directory = await temporary(t);
  const first = await source(directory, 'first.pptx');
  const second = await source(directory, 'second.xlsx', createDemo('xlsx'));
  await mkdir(path.join(directory, 'reports'));
  await writeFile(path.join(directory, 'reports/first.pptx.report.json'), 'EXISTING USER RESULT', { flag: 'wx' });
  const result = await cli(directory, ['check', first, second, '--out', 'reports']);
  assert.equal(result.code, 1);
  assert.equal(await readFile(path.join(directory, 'reports/first.pptx.report.json'), 'utf8'), 'EXISTING USER RESULT');
  assert.equal((await json(path.join(directory, 'reports/second.xlsx.report.json'))).format, 'xlsx');
});

test('existing repair report prevents creation of a partial document-report pair', async t => {
  const directory = await temporary(t);
  const input = await source(directory, 'deck.pptx');
  await mkdir(path.join(directory, 'reports'));
  await writeFile(path.join(directory, 'reports/deck.windows.pptx.report.json'), 'KEEP ME', { flag: 'wx' });
  const result = await cli(directory, ['repair', input, '--accept-font-substitutions', '--out', 'reports']);
  assert.equal(result.code, 1);
  assert.equal(await exists(path.join(directory, 'reports/deck.windows.pptx')), false);
  assert.equal(await readFile(path.join(directory, 'reports/deck.windows.pptx.report.json'), 'utf8'), 'KEEP ME');
});

test('identical and case-folded output names from different input directories never overwrite each other', async t => {
  const directory = await temporary(t);
  const first = await source(directory, 'source-a/Deck.PPTX');
  const second = await source(directory, 'source-b/deck.pptx');
  const result = await cli(directory, ['check', first, second, '--out', 'reports']);
  assert.equal(result.code, 1);
  assert.deepEqual(await readdir(path.join(directory, 'reports')), ['Deck.PPTX.report.json']);
  assert.equal((await json(path.join(directory, 'reports/Deck.PPTX.report.json'))).fileName, 'Deck.PPTX');
});

test('one corrupt input does not stop a later valid input and does produce an aggregate failure status', async t => {
  const directory = await temporary(t);
  const broken = await source(directory, 'broken.pptx', new TextEncoder().encode('PRIVATE_CONTENT_NOT_FOR_TERMINAL'));
  const valid = await source(directory, 'valid.xlsx', createDemo('xlsx'));
  const result = await cli(directory, ['check', broken, valid, '--out', 'reports']);
  assert.equal(result.code, 1);
  assert.ok(await exists(path.join(directory, 'reports/valid.xlsx.report.json')));
  assert.equal(await exists(path.join(directory, 'reports/broken.pptx.report.json')), false);
  assert.doesNotMatch(result.stdout + result.stderr, /PRIVATE_CONTENT_NOT_FOR_TERMINAL/);
});

test('unsupported extensions and missing files are reported without obstructing subsequent inputs', async t => {
  const directory = await temporary(t);
  const unsupported = await source(directory, 'raw.key');
  const valid = await source(directory, 'valid.pptx');
  const result = await cli(directory, ['check', unsupported, 'missing.xlsx', valid]);
  assert.equal(result.code, 1);
  assert.match(result.stdout, /2 个输入失败或受阻/);
  assert.ok(await exists(path.join(directory, 'officebridge-reports/valid.pptx.report.json')));
});

test('dash-prefixed filenames work after an explicit argument separator', async t => {
  const directory = await temporary(t);
  await source(directory, '-deck.pptx');
  const result = await cli(directory, ['check', '--out', 'reports', '--', '-deck.pptx']);
  assert.equal(result.code, 0, result.stderr);
  assert.ok(await exists(path.join(directory, 'reports/-deck.pptx.report.json')));
});

test('output paths cannot escape the chosen directory through relative input paths', async t => {
  const directory = await temporary(t);
  await source(directory, 'inputs/deck.pptx');
  const result = await cli(directory, ['check', 'inputs/../inputs/deck.pptx', '--out', 'reports']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(await readdir(path.join(directory, 'reports')), ['deck.pptx.report.json']);
  assert.deepEqual(await readdir(path.join(directory, 'inputs')), ['deck.pptx']);
});

test('Windows-reserved output basenames are made portable while preserving source identity in the report', { skip: process.platform === 'win32' }, async t => {
  const directory = await temporary(t);
  const input = await source(directory, 'CON.pptx');
  const result = await cli(directory, ['check', input]);
  assert.equal(result.code, 0, result.stderr);
  const report = await json(path.join(directory, 'officebridge-reports/_CON.pptx.report.json'));
  assert.equal(report.fileName, 'CON.pptx');
});

test('directories with an Office extension are rejected instead of being treated as documents', async t => {
  const directory = await temporary(t);
  await mkdir(path.join(directory, 'folder.pptx'));
  const result = await cli(directory, ['check', 'folder.pptx']);
  assert.equal(result.code, 1);
  assert.match(result.stderr, /普通文件/);
  assert.deepEqual(await readdir(path.join(directory, 'officebridge-reports')), []);
});

test('named-pipe inputs are rejected promptly without waiting for a writer', { skip: process.platform === 'win32' }, async t => {
  const directory = await temporary(t);
  const pipe = path.join(directory, 'pipe.pptx');
  await execute('mkfifo', [pipe]);
  const result = await cli(directory, ['check', pipe]);
  assert.equal(result.code, 1);
  assert.match(result.stderr, /普通文件/);
});

test('invalid option combinations are usage failures with no filesystem writes', async t => {
  const directory = await temporary(t);
  for (const args of [[], ['--out'], ['check', 'file.pptx', '--accept-font-substitutions'], ['check', '--bad-option', 'file.pptx']]) {
    assert.equal((await cli(directory, args)).code, 2);
  }
  assert.deepEqual(await readdir(directory), []);
});

test('Mac CLI target checks and repairs to a distinct copy while preserving the original', async t => {
  const directory = await temporary(t);
  const bytes = createDemo('xlsx', { fontProfile: 'mac-office' });
  const input = await source(directory, '跨系统 表格.XLSX', bytes);
  const checked = await cli(directory, ['check', input, '--target', 'mac', '--out', 'checks']);
  assert.equal(checked.code, 0, checked.stderr);
  const original = await json(path.join(directory, 'checks/跨系统 表格.XLSX.report.json'));
  assert.equal(original.analysisOptions.fontProfile, 'mac-office');
  assert.equal(original.validation.targetPlatform, 'mac');
  const repaired = await cli(directory, ['repair', input, '--target=mac', '--accept-font-substitutions', '--out', 'repairs']);
  assert.equal(repaired.code, 0, repaired.stderr);
  assert.match(repaired.stdout, /未进行 Mac Office 实测/);
  const report = await json(path.join(directory, 'repairs/跨系统 表格.mac.XLSX.report.json'));
  assert.equal(report.analysisOptions.fontProfile, 'mac-office');
  assert.equal(report.validation.targetOfficeTested, false);
  assert.equal(report.repair.originalSha256, hash(bytes));
  assert.ok(report.repair.changes.some(change => change.after === 'PingFang SC'));
  assert.ok(report.repair.changes.some(change => change.after === 'Helvetica Neue'));
  const output = await readFile(path.join(directory, 'repairs/跨系统 表格.mac.XLSX'));
  assert.equal(hash(output), report.sha256);
  assert.equal(hash(await readFile(input)), hash(bytes));
});

test('CLI rejects unsupported and missing targets before creating any output', async t => {
  const directory = await temporary(t);
  for (const args of [['--target', 'linux'], ['--target'], ['--target='], ['--target=Mac']]) {
    const result = await cli(directory, ['input.pptx', ...args]);
    assert.equal(result.code, 2);
    assert.match(result.stderr, /--target.*windows.*mac/);
  }
  assert.deepEqual(await readdir(directory), []);
});
