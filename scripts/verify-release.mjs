// Verify our release archive, source inventory and checksums without executing
// anything from the archives or extracting user-supplied paths.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { readArchive } from '../src/engine/archive.js';
import { VERSION } from '../src/engine/index.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const release = path.join(root, 'release');
const pkg = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
const prefix = `OfficeBridge-v${pkg.version}`;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const manifest = JSON.parse(await fs.readFile(path.join(release, 'release-manifest.json'), 'utf8'));
assert.equal(manifest.version, pkg.version);
assert.equal(VERSION, pkg.version, 'Engine and package versions differ');
const artifactNames = ['OfficeBridge.html', 'START-HERE.txt', `${prefix}-Windows.zip`, `${prefix}-Mac.zip`, `${prefix}-portable.zip`, `${prefix}-source.zip`];
const expected = new Set([...artifactNames, 'release-manifest.json']);
const lines = (await fs.readFile(path.join(release, 'SHA256SUMS.txt'), 'utf8')).trim().split(/\r?\n/);
assert.equal(lines.length, expected.size);
for (const line of lines) {
  const match = /^([a-f0-9]{64})  (.+)$/.exec(line);
  assert.ok(match && expected.delete(match[2]), 'Unexpected or duplicate checksum entry');
  assert.equal(hash(await fs.readFile(path.join(release, match[2]))), match[1], match[2]);
}
assert.equal(expected.size, 0);
assert.deepEqual(manifest.artifacts.map(item => item.file).sort(), [...artifactNames].sort(), 'Artifact manifest file set is incomplete');
for (const artifact of manifest.artifacts) {
  const bytes = await fs.readFile(path.join(release, artifact.file));
  assert.equal(artifact.size, bytes.length, artifact.file);
  assert.equal(artifact.sha256, hash(bytes), artifact.file);
}

const sourceRoots = ['.github', '.gitattributes', '.gitignore', 'AGENTS.md', 'CHANGELOG.md', 'CONTRIBUTING.md',
  'LICENSE', 'README.md', 'README.en.md', 'SECURITY.md', 'THIRD_PARTY_NOTICES.md', 'docs', 'index.html',
  'OfficeBridge-Windows.cmd', 'OfficeBridge-Mac.command',
  'package.json', 'package-lock.json', 'public', 'scripts', 'src', 'tests', 'vite.config.js'];
const git = spawnSync('git', ['ls-files', '-z', '--cached'], { cwd: root, encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 });
assert.ok(!git.error && git.status === 0, 'Cannot obtain tracked source inventory');
const trackedAllowed = [...new Set(git.stdout.split('\0').filter(file => file && sourceRoots.some(allowed => file === allowed || file.startsWith(allowed + '/'))))].sort();
assert.deepEqual(manifest.sourceFiles.map(file => file.path).sort(), trackedAllowed, 'Source manifest must exactly match tracked allowlisted files');
const sourceZip = new Uint8Array(await fs.readFile(path.join(release, `${prefix}-source.zip`)));
const source = readArchive(sourceZip).entries;
assert.equal(source.size, manifest.sourceFiles.length);
for (const file of manifest.sourceFiles) {
  const key = `${prefix}/${file.path}`;
  assert.ok(!/(?:^|\/)(?:work|node_modules|release|\.git|dist|coverage|__pycache__)(?:\/|$)/.test(file.path), 'Private or generated directory in source package');
  const bytes = source.get(key);
  assert.ok(bytes, `Missing source entry: ${file.path}`);
  assert.equal(bytes.length, file.size, file.path);
  assert.equal(hash(bytes), file.sha256, file.path);
  assert.equal(hash(await fs.readFile(path.join(root, file.path))), file.sha256, `Source changed after packaging: ${file.path}`);
}
const windowsFiles = ['OfficeBridge-Windows.cmd', 'scripts/office-check-windows.ps1', 'scripts/verify-windows.ps1'];
const macFiles = ['OfficeBridge-Mac.command', 'scripts/office-check-mac.js', 'scripts/office-check-mac-office.js', 'scripts/office-check-mac-ppt.applescript'];
for (const required of ['package.json', 'package-lock.json', 'LICENSE', 'THIRD_PARTY_NOTICES.md',
  'README.md', 'README.en.md', 'CONTRIBUTING.md', 'SECURITY.md', 'CHANGELOG.md', 'docs/QUICKSTART.md',
  'docs/DONATIONS.md', '.github/FUNDING.yml', 'src/support-config.js', 'tests/support.test.js',
  'docs/ARCHITECTURE.md', 'docs/ENGINE-CONTRACT.md', 'docs/PUBLISHING.md', 'docs/TESTING.md', 'docs/WINDOWS-VERIFICATION.md',
  'scripts/package-release.mjs', 'scripts/build-portable.mjs', 'scripts/verify-release.mjs',
  'src/engine/index.js', 'src/main.js', ...windowsFiles, ...macFiles]) {
  assert.ok(source.has(`${prefix}/${required}`), `Missing required source entry: ${required}`);
}
const decodeJson = relative => JSON.parse(new TextDecoder().decode(source.get(`${prefix}/${relative}`)));
assert.equal(decodeJson('package.json').version, pkg.version);
const lock = decodeJson('package-lock.json');
assert.equal(lock.version, pkg.version);
assert.equal(lock.packages[''].version, pkg.version);

// readArchive already checked central-directory boundaries and file names. Inspect
// only the extra metadata needed to ensure .command can run after ZIP extraction.
function verifyLauncherMode(zip, launcher) {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  let end = -1;
  for (let i = zip.length - 22; i >= Math.max(0, zip.length - 65557); i -= 1) {
    if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === zip.length) { end = i; break; }
  }
  assert.ok(end >= 0, 'ZIP directory not found');
  let cursor = view.getUint32(end + 16, true), found = false;
  for (let i = 0; i < view.getUint16(end + 10, true); i += 1) {
    const length = view.getUint16(cursor + 28, true);
    const file = new TextDecoder().decode(zip.subarray(cursor + 46, cursor + 46 + length));
    if (file === launcher) {
      found = true;
      assert.equal(view.getUint16(cursor + 4, true) >>> 8, 3, 'Mac launcher must carry Unix file attributes');
      assert.equal(view.getUint32(cursor + 38, true) >>> 16, 0o100755, 'Mac launcher must be a regular executable file (0755)');
    }
    cursor += 46 + length + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
  }
  assert.ok(found, `Missing executable launcher: ${launcher}`);
}
verifyLauncherMode(sourceZip, `${prefix}/OfficeBridge-Mac.command`);

const commonFiles = ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'docs/QUICKSTART.md', 'docs/WINDOWS-VERIFICATION.md', 'docs/TESTING.md',
  'docs/images/start.png', 'docs/images/proof.png', 'public/examples/officebridge-demo.pptx', 'public/examples/officebridge-demo.xlsx'];
if (source.has(`${prefix}/docs/NATIVE-OFFICE-CHECK.md`)) commonFiles.push('docs/NATIVE-OFFICE-CHECK.md');
let appEntries = 0;
for (const [suffix, helpers] of [['Windows', windowsFiles], ['Mac', macFiles], ['portable', [...windowsFiles, ...macFiles]]]) {
  const zip = new Uint8Array(await fs.readFile(path.join(release, `${prefix}-${suffix}.zip`)));
  const entries = readArchive(zip).entries;
  const originals = new Map([
    ['OfficeBridge.html', path.join(release, 'OfficeBridge.html')],
    ['START-HERE.txt', path.join(release, 'START-HERE.txt')],
    ...[...commonFiles, ...helpers].map(file => [file, path.join(root, file)]),
  ]);
  assert.deepEqual([...entries.keys()].sort(), [...originals.keys()].sort(), `${suffix} archive entries do not match the required file set`);
  for (const [file, original] of originals) {
    assert.deepEqual(entries.get(file), new Uint8Array(await fs.readFile(original)), `${suffix} content differs from its source: ${file}`);
  }
  if (suffix !== 'Windows') verifyLauncherMode(zip, 'OfficeBridge-Mac.command');
  appEntries += entries.size;
}
console.log(`Verified ${manifest.sourceFiles.length} tracked source files, ${appEntries} app ZIP entries and ${lines.length} artifact checksums.`);
