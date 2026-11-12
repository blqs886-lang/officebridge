import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { zipSync } from 'fflate';
import { build } from 'vite';
import { buildPortable } from './build-portable.mjs';
import { VERSION } from '../src/engine/index.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
const lock = JSON.parse(await fs.readFile(path.join(root, 'package-lock.json'), 'utf8'));
if (VERSION !== pkg.version || lock.version !== pkg.version || lock.packages[''].version !== pkg.version) {
  throw new Error('Package, lockfile and engine versions must match before packaging.');
}
const release = path.join(root, 'release');
const name = `OfficeBridge-v${pkg.version}`;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const encoder = new TextEncoder();

function runNode(args) {
  const result = spawnSync(process.execPath, args, { cwd: root, stdio: 'inherit' });
  if (result.error || result.status !== 0) throw new Error(`Release check failed: ${args[0]}`);
}

// Use only tracked files inside this explicit allowlist. Staged new source files
// are included; untracked local reports, bytecode and screenshots are never swept
// up by recursive directory traversal. Content is read from the working tree.
const sourceRoots = ['.github', '.gitattributes', '.gitignore', 'AGENTS.md', 'CHANGELOG.md', 'CONTRIBUTING.md',
  'LICENSE', 'README.md', 'README.en.md', 'SECURITY.md', 'THIRD_PARTY_NOTICES.md', 'docs', 'index.html',
  'OfficeBridge-Windows.cmd', 'OfficeBridge-Mac.command',
  'package.json', 'package-lock.json', 'public', 'scripts', 'src', 'tests', 'vite.config.js'];
const git = spawnSync('git', ['ls-files', '-z', '--cached'], { cwd: root, encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 });
if (git.error || git.status !== 0) throw new Error('Cannot obtain tracked source inventory. Package from a Git checkout.');
const tracked = new Set(git.stdout.split('\0').filter(Boolean));
const sourceFiles = [...tracked].filter(file => sourceRoots.some(allowed => file === allowed || file.startsWith(allowed + '/'))).sort();
const windowsFiles = ['OfficeBridge-Windows.cmd', 'scripts/office-check-windows.ps1', 'scripts/verify-windows.ps1'];
const macFiles = ['OfficeBridge-Mac.command', 'scripts/office-check-mac.js', 'scripts/office-check-mac-office.js', 'scripts/office-check-mac-ppt.applescript'];
const commonFiles = ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'docs/QUICKSTART.md', 'docs/WINDOWS-VERIFICATION.md', 'docs/TESTING.md',
  'docs/images/start.png', 'docs/images/proof.png', 'public/examples/officebridge-demo.pptx', 'public/examples/officebridge-demo.xlsx'];
// Optional during development, included in every app ZIP once the guide exists.
const nativeGuide = 'docs/NATIVE-OFFICE-CHECK.md';
try { await fs.lstat(path.join(root, nativeGuide)); commonFiles.push(nativeGuide); }
catch (error) { if (error.code !== 'ENOENT') throw error; }
for (const file of ['README.en.md', ...windowsFiles, ...macFiles, ...commonFiles]) {
  if (!tracked.has(file)) throw new Error(`Required release source is not tracked; stage it before packaging: ${file}`);
}
async function readSource(relative) {
  if (relative.startsWith('/') || relative.split('/').some(segment => !segment || segment === '.' || segment === '..')
    || /(?:^|\/)(?:work|node_modules|release|\.git|dist|coverage|__pycache__)(?:\/|$)/.test(relative)) {
    throw new Error(`Refusing private or generated source path: ${relative}`);
  }
  const segments = relative.split('/');
  for (let index = 1; index <= segments.length; index += 1) {
    const entry = await fs.lstat(path.join(root, ...segments.slice(0, index)));
    if (entry.isSymbolicLink() || (index === segments.length ? !entry.isFile() : !entry.isDirectory())) {
      throw new Error(`Refusing symlink or non-regular source path: ${relative}`);
    }
  }
  return fs.readFile(path.join(root, relative));
}

// Tests and build run against this checkout before packaging. No publishing,
// authentication, dependency installation or network requests happen here.
const tests = (await fs.readdir(path.join(root, 'tests'))).filter(n => n.endsWith('.test.js')).sort();
runNode(['--test', ...tests.map(n => `tests/${n}`)]);
runNode(['scripts/update-notices.mjs']);
await build({ root });
await buildPortable(root);

// Explicit Unix modes make the Mac launcher double-clickable after extraction.
// Archive metadata is deterministic and independent of the builder's umask/OS.
const zipEntry = (bytes, file = '') => [bytes, { level: 6, mtime: new Date('2026-01-01T00:00:00Z'),
  os: 3, attrs: ((file.endsWith('.command') ? 0o100755 : 0o100644) << 16) >>> 0 }];
const source = Object.create(null), sourceManifest = [], content = new Map();
for (const relative of sourceFiles) {
  const bytes = await readSource(relative);
  content.set(relative, bytes);
  source[`${name}/${relative}`] = zipEntry(bytes, relative);
  sourceManifest.push({ path: relative, size: bytes.length, sha256: hash(bytes) });
}
await fs.writeFile(path.join(release, `${name}-source.zip`), zipSync(source));

const quickStart = `OfficeBridge ${pkg.version}

1. 先完整解压 ZIP，再用 Chrome 或 Windows Edge 打开 OfficeBridge.html。
2. 选择接收方电脑：Windows 或 Mac。点“选择文件”或拖入 .pptx 演示文稿、.xlsx 表格，也可以先试样例。
3. 先看检查结果。原件在接收方显示、打印正常时，可以直接保留原件。
4. 需要换字体时，选择要采用的建议，再生成副本。下载的文件名含 .windows 或 .mac；原文件不会被覆盖。
5. 在接收方电脑完整解压对应的 Windows / Mac ZIP，双击 OfficeBridge-Windows.cmd 或 OfficeBridge-Mac.command，选择原件以及可选的修复副本。本机需要装好桌面版 PowerPoint / Excel。
6. 本机 Office 核对完成后，会打开结果页面，显示实际导出的 PDF 和检查记录。Mac 首次使用时可能请求自动化或文件访问权限，请阅读系统提示。播放动画和复杂排版仍需在 Office 里检查。

有源电脑导出的参照 PDF 时，可以一起选择；它有助于判断跨系统后是否真的变化。也可以打开页面中的 PDF 对照工具，手动选两份 PDF 比较页数与排版。

文件只在本机处理，支持断网。刷新或关闭页面会清空本次列表，下载好的文件仍在磁盘上。Keynote / Numbers 请先导出成 PowerPoint / Excel。
换字体可能改变字宽和分页；工具不能恢复已经丢失的文字，也不保证所有 Office 版本显示完全一致。

只想检查和修复字体时，直接打开 OfficeBridge.html 即可，不需要安装 Node.js、Python 或其他开发工具。更多步骤见 docs/QUICKSTART.md。
`;
await fs.writeFile(path.join(release, 'START-HERE.txt'), quickStart);
const html = await fs.readFile(path.join(release, 'OfficeBridge.html'));
for (const [suffix, helpers] of [['Windows', windowsFiles], ['Mac', macFiles], ['portable', [...windowsFiles, ...macFiles]]]) {
  const archive = Object.create(null);
  archive['OfficeBridge.html'] = zipEntry(html);
  archive['START-HERE.txt'] = zipEntry(encoder.encode(quickStart));
  for (const file of [...commonFiles, ...helpers]) archive[file] = zipEntry(content.get(file), file);
  await fs.writeFile(path.join(release, `${name}-${suffix}.zip`), zipSync(archive));
}

const artifacts = [];
for (const file of ['OfficeBridge.html', 'START-HERE.txt', `${name}-Windows.zip`, `${name}-Mac.zip`, `${name}-portable.zip`, `${name}-source.zip`]) {
  const bytes = await fs.readFile(path.join(release, file));
  artifacts.push({ file, size: bytes.length, sha256: hash(bytes) });
}
const manifest = { version: pkg.version, builtAt: new Date().toISOString(), node: process.version,
  validation: { automatedTestsPassed: true, browserTestsRunByPackager: false, windowsOfficeTested: false, macOfficeTested: false },
  sourceFiles: sourceManifest, artifacts };
const manifestBytes = encoder.encode(JSON.stringify(manifest, null, 2) + '\n');
await fs.writeFile(path.join(release, 'release-manifest.json'), manifestBytes);
await fs.writeFile(path.join(release, 'SHA256SUMS.txt'), [...artifacts, { file: 'release-manifest.json', sha256: hash(manifestBytes) }].map(x => `${x.sha256}  ${x.file}`).join('\n') + '\n');
console.log(`Release ready in ${release}`);
for (const artifact of artifacts) console.log(`${artifact.file}: ${artifact.size} bytes`);
