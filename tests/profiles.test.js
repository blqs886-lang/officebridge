import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { createDemo } from '../src/engine/demo.js';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';
import { readArchive } from '../src/engine/archive.js';
import { fontAdvice, isKnownFont, normalizeFontOptions } from '../src/engine/fonts.js';

const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const decoder = new TextDecoder();
const mac = { fontProfile: 'mac-office' };

test('profiles are explicit and unknown targets fail before inspecting a document', async () => {
  assert.equal(normalizeFontOptions().fontProfile, 'windows-office');
  assert.equal(normalizeFontOptions(mac).fontProfile, 'mac-office');
  for (const fontProfile of ['linux-office', 'Mac', '', false, {}]) {
    assert.throws(() => normalizeFontOptions({ fontProfile }), /目标环境/);
    await assert.rejects(analyzeOffice(new Uint8Array(), 'input.pptx', { fontProfile }), /目标环境/);
    assert.throws(() => createDemo('pptx', { fontProfile }), /目标环境/);
  }
});

test('Mac candidates are conservative and preserve common Office and Apple families', () => {
  for (const [from, to] of [
    ['Microsoft YaHei', 'PingFang SC'], ['微软雅黑', 'PingFang SC'],
    ['Microsoft JhengHei', 'PingFang TC'], ['SimSun', 'Songti SC'],
    ['KaiTi', 'Kaiti SC'], ['FangSong', 'STFangsong'], ['Segoe UI', 'Helvetica Neue'],
  ]) {
    const choice = fontAdvice(from, mac);
    assert.equal(choice.to, to);
    assert.equal(choice.defaultSelected, false);
    assert.match(choice.reason, /原件显示异常/);
  }
  for (const face of ['Arial', 'Arial Narrow', 'Times New Roman', 'Calibri', 'Consolas', 'Cambria', 'Aptos', 'PingFang SC', 'Helvetica Neue', 'Menlo']) {
    assert.equal(fontAdvice(face, mac), null, face);
    assert.equal(isKnownFont(face, mac), true, face);
  }
  assert.equal(fontAdvice('Custom Scientific Sans', mac), null);
  assert.equal(isKnownFont('Custom Scientific Sans', mac), false);
  assert.equal(fontAdvice('PingFang SC').to, 'Microsoft YaHei');
  assert.equal(fontAdvice('Helvetica Neue').to, 'Arial');
});

test('a concrete Mac list suppresses present sources and requires available targets, including Office fallback', () => {
  const options = targetFonts => normalizeFontOptions({ ...mac, targetFonts });
  assert.equal(fontAdvice('Microsoft YaHei', options(['Microsoft YaHei', 'PingFang SC'])), null);
  assert.equal(fontAdvice('Microsoft YaHei', options(['Arial'])), null);
  assert.equal(fontAdvice('Calibri', options(['Calibri', 'Arial'])), null);
  assert.equal(fontAdvice('Calibri', options(['PingFang SC'])), null);
  assert.equal(fontAdvice('Calibri', options(['Arial'])).to, 'Arial');
  assert.equal(fontAdvice('Consolas', options(['Menlo'])).to, 'Menlo');
  assert.equal(fontAdvice('Segoe UI', options(['helvetica neue'])).to, 'Helvetica Neue');
  const custom = normalizeFontOptions({ ...mac, targetFonts: ['Arial'], fontMappings: { 'Project Sans': 'Arial' } });
  assert.equal(fontAdvice('Project Sans', custom).to, 'Arial');
  for (const family of ['Cambria Math', 'STIX Two Math', 'Wingdings', 'Segoe UI Symbol', 'Apple Color Emoji']) {
    assert.equal(fontAdvice(family, options(['Arial'])), null, family);
    assert.throws(() => normalizeFontOptions({ ...mac, fontMappings: { [family]: 'Arial' } }), /符号、数学或图标/);
  }
});

for (const format of ['pptx', 'xlsx']) {
  test(`${format}: default demo remains identical to published bytes; Mac demo is synthetic and directional`, async () => {
    const original = createDemo(format);
    assert.deepEqual(original, createDemo(format, { fontProfile: 'windows-office' }));
    const published = await readFile(new URL(`../public/examples/officebridge-demo.${format}`, import.meta.url));
    assert.equal(hash(original), hash(published));
    const directional = createDemo(format, mac);
    assert.notEqual(hash(original), hash(directional));
    const report = await analyzeOffice(directional, `sample.${format}`, mac);
    assert.equal(report.repairBlocked, false);
    assert.equal(report.analysisOptions.fontProfile, 'mac-office');
    assert.equal(report.validation.targetPlatform, 'mac');
    assert.equal(report.validation.targetOfficeTested, false);
    assert.equal(report.validation.windowsOfficeTested, false);
    assert.match(report.validation.notes.join(' '), /目标 Mac Office/);
    const names = new Set(report.issues.filter(issue => issue.fixable).map(issue => issue.fix.from));
    assert.ok(names.has('Microsoft YaHei'));
    assert.ok(names.has('Segoe UI'));
    assert.ok(report.issues.every(issue => !issue.defaultSelected));
    // Inspection and generating another direction must never mutate the original.
    assert.equal(hash(original), hash(createDemo(format)));
    await assert.rejects(repairOffice(directional, `sample.${format}`, mac), /先选择/);
  });

  test(`${format}: Mac repair touches only the selected family in its exact part and retains editable content`, async () => {
    const original = createDemo(format, mac), beforeHash = hash(original);
    const report = await analyzeOffice(original, `sample.${format}`, mac);
    const selected = report.issues.find(issue => issue.fix?.from === 'Segoe UI');
    assert.ok(selected);
    const result = await repairOffice(original, `sample.${format}`, { ...mac, selectedIssueIds: [selected.id] });
    assert.equal(result.fileName, `sample.mac.${format}`);
    assert.equal(result.report.analysisOptions.fontProfile, 'mac-office');
    assert.equal(result.report.validation.targetOfficeTested, false);
    assert.equal(result.report.repair.originalSha256, beforeHash);
    assert.equal(result.report.repair.unchangedPartsVerified, true);
    assert.equal(hash(original), beforeHash);
    assert.equal(result.changes.length, 1);
    assert.deepEqual([result.changes[0].before, result.changes[0].after], ['Segoe UI', 'Helvetica Neue']);
    const before = readArchive(original).entries, after = readArchive(result.bytes).entries;
    assert.equal(before.size, after.size);
    for (const [part, bytes] of before) {
      if (part !== selected.part) assert.deepEqual(after.get(part), bytes, part);
      else {
        // Match only font attributes: a visible font-name explanation remains intact.
        const changed = decoder.decode(bytes).replace(/(typeface|val)="Segoe UI"/g, '$1="Helvetica Neue"');
        assert.equal(decoder.decode(after.get(part)), changed);
      }
    }
    assert.ok(result.report.issues.some(issue => issue.fix?.from === 'Microsoft YaHei'));
    if (format === 'xlsx') assert.equal(result.report.inventory.formulaCount, report.inventory.formulaCount);
    else assert.ok([...after.keys()].some(part => /^ppt\/charts\/chart/.test(part)));
  });
}
