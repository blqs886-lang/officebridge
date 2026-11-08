// Verify our release archive, source inventory and checksums without executing
// anything from the archives or extracting user-supplied paths.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { readArchive } from '../src/engine/archive.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const release = path.join(root, 'release');
const pkg = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
const prefix = `OfficeBridge-v${pkg.version}`;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const manifest = JSON.parse(await fs.readFile(path.join(release, 'release-manifest.json'), 'utf8'));
assert.equal(manifest.version, pkg.version);
const expected = new Set(['OfficeBridge.html', 'START-HERE.txt', `${prefix}-portable.zip`, `${prefix}-source.zip`, 'release-manifest.json']);
const lines = (await fs.readFile(path.join(release, 'SHA256SUMS.txt'), 'utf8')).trim().split(/\r?\n/);
assert.equal(lines.length, expected.size);
for (const line of lines) {
  const match = /^([a-f0-9]{64})  (.+)$/.exec(line);
  assert.ok(match && expected.delete(match[2]), 'Unexpected or duplicate checksum entry');
  assert.equal(hash(await fs.readFile(path.join(release, match[2]))), match[1], match[2]);
}
assert.equal(expected.size, 0);
const source = readArchive(new Uint8Array(await fs.readFile(path.join(release, `${prefix}-source.zip`)))).entries;
assert.equal(source.size, manifest.sourceFiles.length);
for (const file of manifest.sourceFiles) {
  const key = `${prefix}/${file.path}`;
  assert.ok(!/(?:^|\/)(?:work|node_modules|release|\.git)(?:\/|$)/.test(file.path), 'Private or generated directory in source package');
  const bytes = source.get(key);
  assert.ok(bytes, `Missing source entry: ${file.path}`);
  assert.equal(bytes.length, file.size, file.path);
  assert.equal(hash(bytes), file.sha256, file.path);
  assert.equal(hash(await fs.readFile(path.join(root, file.path))), file.sha256, `Source changed after packaging: ${file.path}`);
}
for (const required of ['package.json', 'package-lock.json', 'LICENSE', 'THIRD_PARTY_NOTICES.md',
  'README.md', 'CONTRIBUTING.md', 'SECURITY.md', 'CHANGELOG.md', 'docs/QUICKSTART.md',
  'docs/DONATIONS.md', '.github/FUNDING.yml', 'src/support-config.js', 'tests/support.test.js',
  'docs/ARCHITECTURE.md', 'docs/ENGINE-CONTRACT.md', 'docs/PUBLISHING.md', 'docs/TESTING.md', 'docs/WINDOWS-VERIFICATION.md',
  'scripts/package-release.mjs', 'scripts/build-portable.mjs', 'scripts/verify-release.mjs',
  'src/engine/index.js', 'src/main.js']) assert.ok(source.has(`${prefix}/${required}`), `Missing required source entry: ${required}`);
const portable = readArchive(new Uint8Array(await fs.readFile(path.join(release, `${prefix}-portable.zip`)))).entries;
const portableSources = new Map([
  ['OfficeBridge.html', path.join(release, 'OfficeBridge.html')],
  ['START-HERE.txt', path.join(release, 'START-HERE.txt')],
  ...['LICENSE', 'THIRD_PARTY_NOTICES.md', 'docs/QUICKSTART.md', 'docs/WINDOWS-VERIFICATION.md', 'docs/TESTING.md', 'docs/images/start.png',
    'scripts/verify-windows.ps1', 'public/examples/officebridge-demo.pptx', 'public/examples/officebridge-demo.xlsx']
    .map(file => [file, path.join(root, file)]),
]);
assert.deepEqual([...portable.keys()].sort(), [...portableSources.keys()].sort(), 'Portable archive entries do not match the required file set');
for (const [file, original] of portableSources) {
  assert.deepEqual(portable.get(file), new Uint8Array(await fs.readFile(original)), `Portable content differs from its source: ${file}`);
}
console.log(`Verified ${manifest.sourceFiles.length} source files, ${portable.size} portable entries and ${lines.length} artifact checksums.`);
