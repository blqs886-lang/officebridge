import test from 'node:test';
import assert from 'node:assert/strict';
import { buildHtmlReport } from '../src/ui/reports.js';
import { encodingExample } from '../src/engine/text-encoding.js';
import { readArchive, writeArchive } from '../src/engine/archive.js';
import { createDemo } from '../src/engine/demo.js';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';

test('reversible Chinese text previews require adjacent valid byte triples and stay bounded', () => {
  assert.deepEqual(encodingExample('ABC ä¸­æ–‡ PDF'), { stored: 'ä¸­æ–‡', preview: '中文' });
  const latin = String.fromCharCode(...new TextEncoder().encode('保留编辑'));
  assert.equal(encodingExample(latin).preview, '保留编辑');
  assert.equal(encodingExample(latin.repeat(10)).preview.length, 12);
  for (const text of ['中文', 'Résumé déjà vu', 'ä¸­', 'ä¸­/æ–‡', '�', '汉字 ñ 1200.00', '田中さん', '', null]) {
    assert.equal(encodingExample(text), null);
  }
});

for (const format of ['pptx', 'xlsx']) {
  test(`${format}: likely mojibake is located and previewed without altering stored text during repair`, async () => {
    const entries = readArchive(createDemo(format)).entries;
    const where = format === 'pptx' ? 'ppt/slides/slide1.xml' : 'xl/worksheets/sheet1.xml';
    const xml = new TextDecoder().decode(entries.get(where));
    const marker = format === 'pptx' ? 'OfficeBridge · editable-first' : 'SUM(B2:B4)';
    assert.ok(xml.includes(marker));
    const altered = format === 'pptx' ? xml.replace(marker, 'ä¸­æ–‡') : xml.replace(/<t>[^<]*<\/t>/, '<t>ä¸­æ–‡</t>');
    assert.notEqual(altered, xml);
    entries.set(where, new TextEncoder().encode(altered));
    const bytes = writeArchive(entries);
    const report = await analyzeOffice(bytes, 'example.' + format);
    const issue = report.issues.find(item => item.code.endsWith('TEXT_ENCODING_REVIEW'));
    assert.match(issue.evidence, /ä¸­æ–‡ → 中文/u);
    assert.equal(issue.fixable, false);
    assert.match(buildHtmlReport({ report }), /ä¸­æ–‡ → 中文/u);
    const fixed = await repairOffice(bytes, 'example.' + format, { selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id) });
    assert.ok(new TextDecoder().decode(readArchive(fixed.bytes).entries.get(where)).includes('ä¸­æ–‡'));
  });
}
