import test from 'node:test';
import assert from 'node:assert/strict';
import { createDemo } from '../src/engine/demo.js';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';
import { encodeXml, parseXml, replaceFontAttributes, NS } from '../src/engine/xml.js';
import { normalizeTargetFonts, parseFontList, listedFont } from '../src/engine/target-fonts.js';
import { buildHtmlReport, exportPayload } from '../src/ui/reports.js';

test('font list accepts Windows text encodings, aliases and canonical Unicode', () => {
  for (const encoding of ['utf-8', 'utf-16le', 'utf-16be']) {
    const names = parseFontList(encodeXml(' Arial \r\nARIAL\r\n微软雅黑\rCafé\nCafe\u0301\n', { encoding, bom: true }));
    assert.deepEqual(names, ['Arial', '微软雅黑', 'Café']);
    assert.equal(listedFont(names, 'Cafe\u0301'), true);
    assert.equal(listedFont(names, 'Ari'), false);
  }
});

test('font list rejects empty, corrupt, excessive and markup inputs', () => {
  for (const bytes of [new Uint8Array(0), Uint8Array.of(0xff, 0x55), new TextEncoder().encode('Arial\n<x/>'),
    new Uint8Array(512 * 1024 + 1), Uint8Array.of(65, 0, 98, 0)]) assert.throws(() => parseFontList(bytes));
  for (const names of [[], 'Arial', Array(4097).fill('Arial'), ['A'.repeat(129)], [null]]) assert.throws(() => normalizeTargetFonts(names));
});

for (const format of ['pptx', 'xlsx']) {
  test(`${format}: a target list preserves listed originals and only offers listed replacements`, async () => {
    const bytes = createDemo(format), name = 'snapshot.' + format;
    const options = { targetFonts: ['PINGFANG SC', 'Arial', 'Private unrelated family'] };
    const report = await analyzeOffice(bytes, name, options);
    const fixes = report.issues.filter(i => i.fixable);
    assert.ok(fixes.length);
    assert.ok(fixes.every(i => i.fix.from === 'Helvetica Neue' && i.fix.to === 'Arial' && !i.defaultSelected));
    assert.ok(report.inventory.targetFontComparison.listed.includes('PingFang SC'));
    assert.ok(report.inventory.targetFontComparison.unlisted.includes('Helvetica Neue'));
    const repaired = await repairOffice(bytes, name, { ...options, selectedIssueIds: fixes.map(i => i.id) });
    assert.equal(repaired.report.issues.filter(i => i.fixable).length, 0);
    assert.ok(repaired.report.inventory.targetFontComparison.listed.includes('Arial'));
    const record = { report, options, repaired };
    assert.doesNotMatch(JSON.stringify(exportPayload(record)), /Private unrelated family/);
    const html = buildHtmlReport(record);
    assert.match(html, /目标 Windows 字体名单核对/);
    assert.doesNotMatch(html, /Private unrelated family/);
    const restricted = await analyzeOffice(bytes, name, { targetFonts: ['Unrelated'], fontMappings: { 'Helvetica Neue': 'Arial' } });
    assert.equal(restricted.issues.some(i => i.fixable), false);
    assert.ok((await analyzeOffice(bytes, name)).issues.some(i => i.fix?.from === 'PingFang SC'));
  });
}

test('surgical edits handle Unicode extension names and leave unrelated attributes byte-for-byte intact', () => {
  const xml = `<r xmlns:a="${NS.a}" xmlns:s="${NS.s}" xmlns:x="urn:x"><x:备注 提示=" typeface='Old' "/><a:latin x:备注=" typeface='Old' " typeface='Old'/><a:unknown typeface="Old"/><x:font><s:name val="Old"/><s:rFont val="Old"/></x:font><s:rPr><s:rFont val="Old"/></s:rPr></r>`;
  const fixed = replaceFontAttributes(xml, parseXml(xml), [
    { issueId: 'd', from: 'Old', to: 'New', scope: NS.a }, { issueId: 's', from: 'Old', to: 'New', scope: NS.s },
  ]);
  assert.equal(fixed.text, xml.replace("typeface='Old'/>", "typeface='New'/>").replace('<s:rPr><s:rFont val="Old"', '<s:rPr><s:rFont val="New"'));
  assert.equal(fixed.counts.get('d'), 1);
  assert.equal(fixed.counts.get('s'), 1);
});

test('dense font edits are simultaneous and each occurrence is counted once', () => {
  const xml = '<r xmlns:a="' + NS.a + '">' + '<a:latin typeface="A"/><a:ea typeface="B"/>'.repeat(12000) + '</r>';
  const fixed = replaceFontAttributes(xml, parseXml(xml), [
    { issueId: 'b', from: 'B', to: 'C', scope: NS.a }, { issueId: 'a', from: 'A', to: 'B', scope: NS.a },
  ]);
  assert.equal(fixed.text, xml.replaceAll('"B"', '"C"').replaceAll('"A"', '"B"'));
  assert.deepEqual([...fixed.counts], [['a', 12000], ['b', 12000]]);
});
