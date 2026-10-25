import test from 'node:test';
import assert from 'node:assert/strict';
import { createDemo } from '../src/engine/demo.js';
import { analyzeOffice, repairOffice, resolveTarget } from '../src/engine/index.js';
import { readArchive, writeArchive, crc32, LIMITS } from '../src/engine/archive.js';
import { NS, parseXml, replaceFontAttributes, encodeXml, decodeXml } from '../src/engine/xml.js';
import { fontAdvice, normalizeFontOptions } from '../src/engine/fonts.js';

const encoder = new TextEncoder();
function altered(format, change) {
  const entries = readArchive(createDemo(format)).entries;
  change(entries);
  return writeArchive(entries);
}
function xmlChange(entries, part, change) {
  entries.set(part, encoder.encode(change(new TextDecoder().decode(entries.get(part)))));
}
async function fixed(format) {
  const input = createDemo(format), fileName = `sample.${format}`;
  const report = await analyzeOffice(input, fileName);
  return { input, report, result: await repairOffice(input, fileName, { selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id) }) };
}

for (const format of ['pptx', 'xlsx']) {
  test(`${format}: real generated Office package analyzes deterministically without mutation`, async () => {
    const input = createDemo(format), copy = input.slice();
    const report = await analyzeOffice(input, `test.${format}`);
    assert.equal(report.repairBlocked, false);
    assert.equal(report.validation.windowsOfficeTested, false);
    assert.match(report.sha256, /^[a-f\d]{64}$/);
    assert.ok(report.issues.some(i => i.fixable));
    assert.deepEqual(input, copy);
    assert.deepEqual(await analyzeOffice(input, `test.${format}`), report);
  });
  test(`${format}: repairing preserves every unrelated package part byte for byte`, async () => {
    const { input, result } = await fixed(format);
    const before = readArchive(input).entries, after = readArchive(result.bytes).entries;
    const changed = new Set(result.changes.map(c => c.part));
    assert.deepEqual([...after.keys()], [...before.keys()]);
    assert.ok(changed.size > 0);
    for (const [part, bytes] of before) {
      if (!changed.has(part)) assert.deepEqual(after.get(part), bytes, part);
      else {
        let expected = new TextDecoder().decode(bytes);
        for (const change of result.changes.filter(c => c.part === part)) {
          const key = format === 'pptx' ? 'typeface' : 'val';
          expected = expected.replaceAll(`${key}="${change.before}"`, `${key}="${change.after}"`);
        }
        assert.equal(new TextDecoder().decode(after.get(part)), expected, `Only selected attributes may change: ${part}`);
      }
    }
    assert.equal(result.report.repair.unchangedPartsVerified, true);
    assert.equal(result.report.repairBlocked, false);
    assert.equal(result.report.validation.windowsOfficeTested, false);
  });
  test(`${format}: repair requires an explicit current choice`, async () => {
    const input = createDemo(format);
    await assert.rejects(repairOffice(input, `a.${format}`), /选择/);
    await assert.rejects(repairOffice(input, `a.${format}`, { selectedIssueIds: ['invented-id'] }), /失效/);
  });
}

test('XLSX formulas, numeric values, long identifiers and formatting are untouched', async () => {
  const { input, result } = await fixed('xlsx');
  const a = readArchive(input).entries, b = readArchive(result.bytes).entries;
  for (const name of a.keys()) {
    if (/worksheets\/|workbook.xml|theme\//.test(name)) assert.deepEqual(b.get(name), a.get(name));
  }
  const xml = new TextDecoder().decode(b.get('xl/worksheets/sheet1.xml'));
  assert.match(xml, /<f>SUM\(B2:B4\)<\/f>/);
  assert.match(new TextDecoder().decode(b.get('xl/worksheets/sheet2.xml')), /00123/);
  assert.match(new TextDecoder().decode(b.get('xl/worksheets/sheet2.xml')), /12345678901234567890/);
});

test('unknown opaque extensions and embedded editable chart data are retained', async () => {
  const input = altered('pptx', entries => entries.set('customData/unknown.dat', Uint8Array.of(0, 1, 255, 7, 128)));
  const report = await analyzeOffice(input, 'extra.pptx');
  const result = await repairOffice(input, 'extra.pptx', { selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id) });
  const a = readArchive(input).entries, b = readArchive(result.bytes).entries;
  assert.deepEqual(b.get('customData/unknown.dat'), a.get('customData/unknown.dat'));
  assert.deepEqual(b.get('ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'), a.get('ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'));
});

test('partial selection changes only one suggested font in one part', async () => {
  const input = createDemo('pptx'), report = await analyzeOffice(input, 'a.pptx');
  const choice = report.issues.find(i => i.fixable && i.fix.from === 'Helvetica Neue');
  const result = await repairOffice(input, 'a.pptx', { selectedIssueIds: [choice.id] });
  assert.equal(result.changes.length, 1);
  assert.ok(result.report.issues.some(i => i.fix?.from === 'PingFang SC'));
  assert.ok(!result.report.issues.some(i => i.fix?.from === 'Helvetica Neue'));
});

test('signed package is analyzable but repair is blocked', async () => {
  const input = altered('pptx', entries => entries.set('_xmlsignatures/sig1.xml', encoder.encode('<Signature/>')));
  const report = await analyzeOffice(input, 'signed.pptx');
  assert.equal(report.repairBlocked, true);
  await assert.rejects(repairOffice(input, 'signed.pptx', { selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id) }), /不能自动修复/);
});

test('missing relationships prevent repair instead of silently dropping content', async () => {
  const input = altered('pptx', entries => entries.delete('ppt/slides/slide1.xml'));
  const report = await analyzeOffice(input, 'broken.pptx');
  assert.equal(report.repairBlocked, true);
  assert.ok(report.issues.some(i => i.title === '内部关联文件缺失'));
});

test('invalid root relationship or duplicate relationship IDs block repair', async () => {
  const input = altered('xlsx', entries => xmlChange(entries, '_rels/.rels', text => text.replace('Id="rId2"', 'Id="rId1"')));
  const report = await analyzeOffice(input, 'broken.xlsx');
  assert.equal(report.repairBlocked, true);
  assert.ok(report.issues.some(i => i.code === 'RELATIONSHIP_INVALID'));
});

test('unsupported or misleading extension/content is rejected', async () => {
  await assert.rejects(analyzeOffice(createDemo('pptx'), 'a.key'), /目前支持/);
  await assert.rejects(analyzeOffice(createDemo('pptx'), 'a.xlsx'), /不匹配/);
  await assert.rejects(analyzeOffice(Uint8Array.of(0xd0, 0xcf, 0, 0), 'encrypted.xlsx'), /加密/);
  await assert.rejects(analyzeOffice(Uint8Array.of(1, 2, 3), 'bad.pptx'), /不是完整/);
});

test('malformed, entity-bearing and invalid UTF-8 XML are rejected', async () => {
  assert.throws(() => parseXml('<a><b></a>'), /损坏|不完整/);
  assert.throws(() => parseXml('<!DOCTYPE a [<!ENTITY x SYSTEM "file:///x">]><a>&x;</a>'), /DTD/);
  assert.throws(() => decodeXml(Uint8Array.of(0xff, 0, 0x3c), 'broken.xml'), /编码/);
  const input = altered('xlsx', entries => entries.set('customXml/item1.xml', encoder.encode('<!DOCTYPE x><x/>')));
  await assert.rejects(analyzeOffice(input, 'doctype.xlsx'), /DTD/);
});

test('legal replacement characters remain available for diagnostics', () => {
  assert.equal(parseXml('<a>\uFFFD</a>').documentElement.textContent, '\uFFFD');
});

test('surgical font edits respect namespaces, quoting, entity values, comments and CDATA', () => {
  const xml = `<?xml version="1.0"?><r xmlns:a="${NS.a}" xmlns:x="urn:test"><!-- <a:latin typeface="A &amp; B"/> --><a:latin typeface='A &amp; B'/><x:latin typeface="A &amp; B"/><![CDATA[<a:latin typeface="A &amp; B"/>]]><a:t>A &amp; B</a:t></r>`;
  const fixed = replaceFontAttributes(xml, parseXml(xml), [{ issueId: 'one', from: 'A & B', to: "O'Brien & Co", scope: NS.a }]);
  assert.equal(fixed.counts.get('one'), 1);
  assert.equal(fixed.text, xml.replace("typeface='A &amp; B'", "typeface='O&apos;Brien &amp; Co'"));
  assert.doesNotThrow(() => parseXml(fixed.text));
});

test('UTF-8 BOM and UTF-16 byte order survive edits', () => {
  for (const encoding of ['utf-8', 'utf-16le', 'utf-16be']) for (const bom of [false, true]) {
    const text = `<?xml version="1.0" encoding="${encoding === 'utf-8' ? 'UTF-8' : 'UTF-16'}"?><a:latin xmlns:a="${NS.a}" typeface="苹方-简"/>`;
    const original = encodeXml(text, { encoding, bom });
    const decoded = decodeXml(original, 'a.xml');
    assert.equal(decoded.text, text);
    assert.equal(decoded.encoding, encoding);
    const updated = replaceFontAttributes(text, parseXml(text), [{ issueId: 'x', from: '苹方-简', to: '微软雅黑', scope: NS.a }]);
    const reread = decodeXml(encodeXml(updated.text, decoded), 'a.xml');
    assert.equal(reread.text, text.replace('苹方-简', '微软雅黑'));
    assert.equal(reread.bom, bom);
  }
});

test('ZIP paths, duplicate entries, CRC, local/central mismatches and expansion limits are checked', () => {
  assert.throws(() => readArchive(writeArchive(new Map([['../evil.xml', encoder.encode('<x/>')]]))), /不安全/);
  assert.throws(() => readArchive(writeArchive(new Map([['%2e%2e/evil.xml', encoder.encode('<x/>')]]))), /不安全/);
  const duplicate = writeArchive(new Map([['a.xml', encoder.encode('<x/>')], ['b.xml', encoder.encode('<y/>')]]));
  for (let i = 0; i < duplicate.length - 5; i++) if (String.fromCharCode(...duplicate.slice(i, i + 5)) === 'b.xml') duplicate[i] = 97;
  assert.throws(() => readArchive(duplicate), /重复/);
  const mismatch = writeArchive(new Map([['a.xml', encoder.encode('<x/>')]]));
  mismatch[30] = 98;
  assert.throws(() => readArchive(mismatch), /名称不一致/);
  const brokenCrc = writeArchive(new Map([['a.bin', Uint8Array.of(1, 2, 3)]]));
  const data = new DataView(brokenCrc.buffer);
  for (let i = 0; i < brokenCrc.length - 4; i++) {
    if (data.getUint32(i, true) === 0x04034b50) data.setUint32(i + 14, 0, true);
    if (data.getUint32(i, true) === 0x02014b50) data.setUint32(i + 16, 0, true);
  }
  assert.throws(() => readArchive(brokenCrc), /校验失败/);
  const bomb = writeArchive(new Map([['large.bin', new Uint8Array(2 * 1024 * 1024)]]));
  assert.throws(() => readArchive(bomb), /压缩比例异常/);
  assert.equal(crc32(encoder.encode('123456789')), 0xcbf43926);
  assert.equal(LIMITS.fileBytes, 100 * 1024 * 1024);
});

test('relationship path normalization accepts package-root and parent paths but blocks escape', () => {
  assert.equal(resolveTarget('ppt/slides/slide1.xml', '../media/a.png'), 'ppt/media/a.png');
  assert.equal(resolveTarget('xl/workbook.xml', '/xl/worksheets/sheet1.xml'), 'xl/worksheets/sheet1.xml');
  assert.equal(resolveTarget('', '../../etc/passwd'), null);
  assert.equal(resolveTarget('xl/workbook.xml', 'file:///a'), null);
});

test('wrong or nested relationship roots cannot masquerade as valid package relations', async () => {
  for (const change of [
    s => s.replaceAll('Relationships', 'OtherRoot'),
    s => s.replace('<Relationship ', '<Container><Relationship ').replace('</Relationships>', '</Container></Relationships>'),
  ]) {
    const bytes = altered('xlsx', entries => xmlChange(entries, '_rels/.rels', change));
    assert.equal((await analyzeOffice(bytes, 'bad.xlsx')).repairBlocked, true);
  }
});

test('ambiguous or incomplete content-type declarations stay read-only in both formats', async () => {
  for (const format of ['pptx', 'xlsx']) {
    const mainPart = format === 'pptx' ? 'ppt/presentation.xml' : 'xl/workbook.xml';
    const original = new TextDecoder().decode(readArchive(createDemo(format)).entries.get('[Content_Types].xml'));
    const main = original.match(/<Override\b[^>]*\/>/g).find(s => s.includes(`PartName="/${mainPart}"`));
    const defaultXml = original.match(/<Default\b[^>]*\/>/g).find(s => s.includes('Extension="xml"'));
    const conflict = main.replace(/ContentType="[^"]+"/, 'ContentType="application/octet-stream"');
    const variants = [
      ['conflicting Override last', s => s.replace('</Types>', `${conflict}</Types>`), 'DUPLICATE'],
      ['conflicting Override first', s => s.replace(main, `${conflict}${main}`), 'DUPLICATE'],
      ['identical Override', s => s.replace('</Types>', `${main}</Types>`), 'DUPLICATE'],
      ['case alias Override', s => s.replace('</Types>', `${main.replace(mainPart, mainPart.toUpperCase())}</Types>`), 'DUPLICATE'],
      ['URI alias Override', s => s.replace('</Types>', `${main.replace(mainPart, `%${mainPart.charCodeAt(0).toString(16)}${mainPart.slice(1)}`)}</Types>`), 'DUPLICATE'],
      ['conflicting Default', s => s.replace('</Types>', '<Default Extension="xml" ContentType="application/octet-stream"/></Types>'), 'DUPLICATE'],
      ['case alias Default', s => s.replace('</Types>', '<Default Extension="XML" ContentType="application/xml"/></Types>'), 'DUPLICATE'],
      ['URI alias Default', s => s.replace('</Types>', '<Default Extension="%78ml" ContentType="application/xml"/></Types>'), 'DUPLICATE'],
      ['nested main Override', s => s.replace(main, `<Wrapper>${main}</Wrapper>`), 'STRUCTURE'],
      ['nested Default', s => s.replace(defaultXml, `<Wrapper>${defaultXml}</Wrapper>`), 'STRUCTURE'],
      ['Default missing Extension', s => s.replace('</Types>', '<Default ContentType="image/png"/></Types>'), 'INVALID'],
      ['Default missing ContentType', s => s.replace('</Types>', '<Default Extension="png"/></Types>'), 'INVALID'],
      ['Override missing PartName', s => s.replace('</Types>', '<Override ContentType="application/xml"/></Types>'), 'INVALID'],
      ['Override missing ContentType', s => s.replace('</Types>', '<Override PartName="/customData/item.xml"/></Types>'), 'INVALID'],
      ['empty required attribute', s => s.replace('</Types>', '<Default Extension="   " ContentType="application/xml"/></Types>'), 'INVALID'],
    ];
    for (const [label, change, code] of variants) {
      const input = altered(format, entries => xmlChange(entries, '[Content_Types].xml', change));
      const saved = input.slice();
      const report = await analyzeOffice(input, `ambiguous.${format}`);
      assert.equal(report.repairBlocked, true, `${format}: ${label}`);
      assert.ok(report.issues.some(i => i.code === `CONTENT_TYPE_${code}`), `${format}: ${label}`);
      await assert.rejects(repairOffice(input, `ambiguous.${format}`, {
        selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id),
      }), /不能自动修复/, `${format}: ${label}`);
      assert.deepEqual(input, saved, `${format}: ${label}`);
    }
    const mismatch = altered(format, entries => xmlChange(entries, '[Content_Types].xml', s => s.replace(main, conflict)));
    await assert.rejects(analyzeOffice(mismatch, `mismatch.${format}`), /扩展名与内部内容不匹配/);
  }
});

test('encoded main relationship names use the same existence check without enabling encoded traversal', async () => {
  for (const format of ['pptx', 'xlsx']) {
    const mainPart = format === 'pptx' ? 'ppt/presentation.xml' : 'xl/workbook.xml';
    const encodedPart = `%${mainPart.charCodeAt(0).toString(16)}${mainPart.slice(1)}`;
    const originalTarget = new RegExp(`Target="/?${mainPart.replaceAll('.', '\\.')}"`);
    const input = altered(format, entries => xmlChange(entries, '_rels/.rels', s => s.replace(originalTarget, `Target="${encodedPart}"`)));
    const report = await analyzeOffice(input, `encoded.${format}`);
    assert.equal(report.repairBlocked, false);
    const result = await repairOffice(input, `encoded.${format}`, { selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id) });
    assert.equal(result.report.repairBlocked, false);
    assert.deepEqual(readArchive(result.bytes).entries.get('_rels/.rels'), readArchive(input).entries.get('_rels/.rels'));
    for (const target of [`%2e%2e/${mainPart}`, `${mainPart.split('/')[0]}/%2e%2e/${mainPart}`, `%2e%2e%2f${mainPart}`]) {
      const unsafe = altered(format, entries => xmlChange(entries, '_rels/.rels', s => s.replace(originalTarget, `Target="${target}"`)));
      const blocked = await analyzeOffice(unsafe, `traversal.${format}`);
      assert.equal(blocked.repairBlocked, true, target);
      assert.ok(blocked.issues.some(i => i.code.startsWith('MISSING_PART_')), target);
      assert.ok(blocked.issues.some(i => i.code === 'MAIN_RELATIONSHIP'), target);
    }
  }
});

test('XLM macros are blocked even when only content types or relationships reveal them', async () => {
  const variants = [
    entries => entries.set('xl/macrosheets/sheet1.xml', encoder.encode(`<worksheet xmlns="${NS.s}"/>`)),
    entries => xmlChange(entries, '[Content_Types].xml', s => s.replace('</Types>', '<Default Extension="secret" ContentType="application/vnd.ms-excel.macrosheet+xml"/></Types>')),
    entries => xmlChange(entries, '_rels/.rels', s => s.replace('</Relationships>', '<Relationship Id="macro" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="xl/workbook.xml"/></Relationships>')),
  ];
  for (const variant of variants) assert.equal((await analyzeOffice(altered('xlsx', variant), 'macro.xlsx')).repairBlocked, true);
});

test('contradictory UTF-16 endianness and forbidden XML characters are rejected', () => {
  assert.throws(() => decodeXml(encodeXml('<?xml version="1.0" encoding="UTF-16BE"?><x/>', { encoding: 'utf-16le', bom: true }), 'a.xml'), /字节顺序/);
  for (const c of ['\u0000', '\u000b', '\ud800', '\udc00', '\ufffe', '\uffff']) assert.throws(() => parseXml(`<x>${c}</x>`), /不允许/);
  assert.equal(parseXml('<x>中文 😀\n\t</x>').documentElement.textContent, '中文 😀\n\t');
});

test('custom font rules reject empty, unsafe, theme and symbol names before touching a file', async () => {
  for (const value of ['', '   ', 'A\u0000', 'A\ud800', 'A\uffff', '<x>', 42, '+mn-lt', 'Wingdings', 'Cambria Math', 'Font Awesome 6 Free', 'Material Icons', 'Segoe UI Emoji']) {
    await assert.rejects(analyzeOffice(createDemo('xlsx'), 'a.xlsx', { fontMappings: { 'PingFang SC': value } }), /字体|主题/);
  }
  assert.throws(() => normalizeFontOptions({ fontMappings: ['Arial'] }), /对应表/);
  assert.throws(() => normalizeFontOptions({ fontMappings: { Symbol: 'Arial' } }), /符号/);
  for (const name of ['Font Awesome 6 Free', 'Material Icons', 'Segoe UI Emoji']) assert.throws(() => normalizeFontOptions({ fontMappings: { [name]: 'Arial' } }), /符号/);
  assert.equal(fontAdvice('Unknown family', { fontMappings: { 'Unknown family': '  Arial  ' } }).to, 'Arial');
});

test('custom mappings are disclosed and can modify unknown fonts without touching content', async () => {
  const input = altered('xlsx', entries => xmlChange(entries, 'xl/styles.xml', s => s.replaceAll('PingFang SC', 'An Unknown Family')));
  const options = { fontMappings: { 'An Unknown Family': '  Arial  ' } };
  const report = await analyzeOffice(input, 'custom.xlsx', options);
  const chosen = report.issues.filter(i => i.fixable && i.fix.from === 'An Unknown Family');
  assert.ok(chosen.length);
  assert.deepEqual(report.analysisOptions, { fontProfile: 'windows-office', fontMappings: { 'An Unknown Family': 'Arial' } });
  const result = await repairOffice(input, 'custom.xlsx', { ...options, selectedIssueIds: chosen.map(i => i.id) });
  assert.ok(result.changes.every(c => c.after === 'Arial'));
  assert.deepEqual(result.report.analysisOptions, report.analysisOptions);
  assert.deepEqual(readArchive(result.bytes).entries.get('xl/worksheets/sheet1.xml'), readArchive(input).entries.get('xl/worksheets/sheet1.xml'));
  assert.deepEqual(options.fontMappings, { 'An Unknown Family': '  Arial  ' });
});
