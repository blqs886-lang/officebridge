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

const quickStart = `OfficeBridge ${pkg.version}\n\n不需要安装、账号或联网；正常使用也不需要源码、Node.js 或 PowerShell。\n\n1. 如果拿到 ZIP，先完整解压。右键 OfficeBridge.html，选择“打开方式 → Google Chrome”或 Windows 的 Microsoft Edge。\n2. 第一次可点“电子表格”样例。检查完成后会看到文件名、检查发现和可选字体修复。\n3. 在样例中点“选择全部字体建议”，再点“生成修复副本”。这里只生成副本，不会改原件。\n4. 看到“可编辑副本已生成”后，还要点“下载修复副本”才会保存文件。随后点“HTML 报告”保存检查记录。\n5. 下载文件通常在浏览器的“下载内容”或系统“下载”文件夹里；新文稿名含 .windows。\n6. 再点“选择文件”或拖入自己的 .pptx / .xlsx。先确认页面标题是自己的文件名，再阅读每项字体建议、选择需要的替换。\n7. 用目标电脑上的 Office 打开下载的副本，核对文字、公式、图表、动画和打印。\n\n怎样判断成功：能下载 Office 格式的新副本与报告；报告列出实际修改；副本在 Office 中还能选中文字或编辑公式。\n“仍有检查发现”不一定是修复失败；有些项目只是外部依赖或需要人工核对的提示。不需要追求所有提醒归零。\n\n目前自动修复范围是字体名称，不能保证所有版式完全相同，不能还原已丢失文字。\n字体替换需要目标电脑安装相应字体；工具不安装或嵌入字体，保留文稿的原生编辑结构。\nKeynote / Numbers 请先从原软件导出 PPTX / XLSX；仅改扩展名不能转换格式。宏、签名和 Strict 文件不自动修复。\n关闭或刷新页面会清空本次工作台；先下载要保留的副本和报告。原始文稿始终保留。\n\n详细试用步骤及常见问题：docs/QUICKSTART.md，可用文本编辑或记事本打开，也可到 GitHub 仓库在线阅读。\nWindows 可选验证脚本：docs/WINDOWS-VERIFICATION.md。它需要本机已装桌面 Office，不是正常使用必需步骤。\n源码另见 OfficeBridge-v${pkg.version}-source.zip；MIT 开源，第三方许可见 THIRD_PARTY_NOTICES.md。\n当前验证范围详见 docs/TESTING.md；浏览器检查通过不代表 Windows Office 排版验收通过。\n`;
await fs.writeFile(path.join(release, 'START-HERE.txt'), quickStart);
const portable = Object.create(null);
portable['OfficeBridge.html'] = zipEntry(await fs.readFile(path.join(release, 'OfficeBridge.html')));
portable['START-HERE.txt'] = zipEntry(encoder.encode(quickStart));
for (const file of ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'docs/QUICKSTART.md', 'docs/WINDOWS-VERIFICATION.md', 'docs/TESTING.md',
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
