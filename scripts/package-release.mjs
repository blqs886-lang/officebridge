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

const quickStart = `OfficeBridge ${pkg.version}\n\n1. 解压本文件夹，用 Chrome 或 Edge 打开 OfficeBridge.html。无需安装、登录或联网。\n2. 点击“演示文稿”或“电子表格”样例，先熟悉检查和下载流程。\n3. 放入 .pptx 或 .xlsx 文件，阅读报告，勾选要采用的字体替换。\n4. 生成并下载副本。原文件不会被覆盖；浏览器页面不会上传文稿。\n5. 在目标 Windows Office 中核对文字、公式、图表、动画和打印。\n\n目前自动修复范围是字体名称，不能保证所有版式完全相同，不能还原已丢失文字。\n字体替换需要目标电脑安装相应字体；保留文稿的原生编辑结构。\nKeynote / Numbers 请先导出 Office 格式；宏、签名和 Strict 文件不自动修复。\n\nWindows 验证脚本：参见 docs/WINDOWS-VERIFICATION.md。脚本需要本机已装 Office，不安装软件。\n源码另见同版本 -source.zip；MIT 开源，第三方许可见 THIRD_PARTY_NOTICES.md。\n\n当前验收状态详见 docs/TESTING.md；不能把 Mac 验证当成 Windows Office 验证。\n`;
await fs.writeFile(path.join(release, 'START-HERE.txt'), quickStart);
const portable = Object.create(null);
portable['OfficeBridge.html'] = zipEntry(await fs.readFile(path.join(release, 'OfficeBridge.html')));
portable['START-HERE.txt'] = zipEntry(encoder.encode(quickStart));
for (const file of ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'docs/WINDOWS-VERIFICATION.md', 'docs/TESTING.md',
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
