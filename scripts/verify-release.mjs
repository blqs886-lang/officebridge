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
  'scripts/package-release.mjs', 'scripts/build-portable.mjs', 'src/engine/index.js', 'src/main.js']) assert.ok(source.has(`${prefix}/${required}`));
const portable = readArchive(new Uint8Array(await fs.readFile(path.join(release, `${prefix}-portable.zip`)))).entries;
assert.deepEqual(portable.get('OfficeBridge.html'), new Uint8Array(await fs.readFile(path.join(release, 'OfficeBridge.html'))));
assert.ok(portable.has('LICENSE') && portable.has('THIRD_PARTY_NOTICES.md'));
assert.ok(portable.has('scripts/verify-windows.ps1') && portable.has('docs/WINDOWS-VERIFICATION.md'));
console.log(`Verified ${manifest.sourceFiles.length} source files, ${portable.size} portable entries and ${lines.length} artifact checksums.`);
