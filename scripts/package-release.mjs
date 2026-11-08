import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { zipSync } from 'fflate';
import { build } from 'vite';
import { buildPortable } from './build-portable.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
const release = path.join(root, 'release');
const name = `OfficeBridge-v${pkg.version}`;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const encoder = new TextEncoder();

function runNode(args) {
  const result = spawnSync(process.execPath, args, { cwd: root, stdio: 'inherit' });
  if (result.error || result.status !== 0) throw new Error(`Release check failed: ${args[0]}`);
}

// Tests and build run against this checkout before packaging. No publishing,
// authentication, dependency installation or network requests happen here.
const tests = (await fs.readdir(path.join(root, 'tests'))).filter(n => n.endsWith('.test.js')).sort();
runNode(['--test', ...tests.map(n => `tests/${n}`)]);
runNode(['scripts/update-notices.mjs']);
await build({ root });
await buildPortable(root);

async function walk(relative) {
  const entry = await fs.lstat(path.join(root, relative));
  if (entry.isSymbolicLink()) throw new Error(`Refusing symlink in source package: ${relative}`);
  if (entry.isFile()) return [relative];
  if (!entry.isDirectory()) throw new Error(`Unexpected source entry: ${relative}`);
  const names = (await fs.readdir(path.join(root, relative))).sort();
  const files = [];
  for (const child of names) files.push(...await walk(`${relative}/${child}`));
  return files;
}

// Explicit allowlist: user documents, local reports, screenshots, environments,
// dependency folders, Git credentials and unrelated workspace files stay out.
const sourceRoots = ['.github', '.gitattributes', '.gitignore', 'AGENTS.md', 'CHANGELOG.md', 'CONTRIBUTING.md',
  'LICENSE', 'README.md', 'SECURITY.md', 'THIRD_PARTY_NOTICES.md', 'docs', 'index.html',
  'package.json', 'package-lock.json', 'public', 'scripts', 'src', 'tests', 'vite.config.js'];
const sourceFiles = [];
for (const relative of sourceRoots) sourceFiles.push(...await walk(relative));
sourceFiles.sort();
const source = Object.create(null), sourceManifest = [];
const zipEntry = bytes => [bytes, { level: 6, mtime: new Date('2026-01-01T00:00:00Z') }];
for (const relative of sourceFiles) {
  const bytes = await fs.readFile(path.join(root, relative));
  source[`${name}/${relative}`] = zipEntry(bytes);
  sourceManifest.push({ path: relative, size: bytes.length, sha256: hash(bytes) });
}
await fs.writeFile(path.join(release, `${name}-source.zip`), zipSync(source));

const quickStart = `OfficeBridge ${pkg.version}

1. 用 Chrome 或 Windows Edge 打开 OfficeBridge.html。如果你下载了 ZIP，请先完整解压。
2. 点“选择文件”，或拖入 .pptx 演示文稿、.xlsx 表格。也可以先点“表格样例”体验。
3. 先看检查结果。原件在 Windows 上已经显示、打印正常时，可以直接使用原件。
4. 需要换字体时，勾选对应的字体，再点“生成并下载副本”。浏览器会下载文件，可以从“下载”中找到名称含 .windows 的副本。
5. 用 Windows PowerPoint 或 Excel 打开副本。演示文稿核对换行、试播动画；表格在打印预览中核对每页首末行，改个测试值并重开。

文件只在本机处理，支持断网。原文件会保留。刷新或关闭页面会清空本次列表，下载好的文件仍在磁盘上。
Keynote 或 Numbers 请先从原应用导出成 PowerPoint 或 Excel。换字体可能改变字宽和分页；它不能恢复已经丢失的文字。

需要更多说明时，请点页面上的“使用帮助”。完整便携 ZIP 中另附 docs/QUICKSTART.md。
`;
await fs.writeFile(path.join(release, 'START-HERE.txt'), quickStart);
const portable = Object.create(null);
portable['OfficeBridge.html'] = zipEntry(await fs.readFile(path.join(release, 'OfficeBridge.html')));
portable['START-HERE.txt'] = zipEntry(encoder.encode(quickStart));
for (const file of ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'docs/QUICKSTART.md', 'docs/WINDOWS-VERIFICATION.md', 'docs/TESTING.md', 'docs/images/start.png',
  'scripts/verify-windows.ps1', 'public/examples/officebridge-demo.pptx', 'public/examples/officebridge-demo.xlsx']) {
  portable[file] = zipEntry(await fs.readFile(path.join(root, file)));
}
await fs.writeFile(path.join(release, `${name}-portable.zip`), zipSync(portable));

const artifacts = [];
for (const file of ['OfficeBridge.html', 'START-HERE.txt', `${name}-portable.zip`, `${name}-source.zip`]) {
  const bytes = await fs.readFile(path.join(release, file));
  artifacts.push({ file, size: bytes.length, sha256: hash(bytes) });
}
const manifest = { version: pkg.version, builtAt: new Date().toISOString(), node: process.version,
  validation: { automatedTestsPassed: true, browserTestsRunByPackager: false, windowsOfficeTested: false },
  sourceFiles: sourceManifest, artifacts };
const manifestBytes = encoder.encode(JSON.stringify(manifest, null, 2) + '\n');
await fs.writeFile(path.join(release, 'release-manifest.json'), manifestBytes);
await fs.writeFile(path.join(release, 'SHA256SUMS.txt'), [...artifacts, { file: 'release-manifest.json', sha256: hash(manifestBytes) }].map(x => `${x.sha256}  ${x.file}`).join('\n') + '\n');
console.log(`Release ready in ${release}`);
for (const artifact of artifacts) console.log(`${artifact.file}: ${artifact.size} bytes`);
