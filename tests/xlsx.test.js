import test from 'node:test';
import assert from 'node:assert/strict';
import { XMLSerializer } from '@xmldom/xmldom';
import { scanXlsx } from '../src/engine/xlsx.js';
import { NS, parseXml, escapeAttribute } from '../src/engine/xml.js';
import { createDemo } from '../src/engine/demo.js';
import { readArchive, writeArchive } from '../src/engine/archive.js';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';

const S = NS.s;
const A = NS.a;
const sheet = body => `<worksheet xmlns="${S}">${body}</worksheet>`;
const workbook = body => `<workbook xmlns="${S}" xmlns:r="${NS.r}">${body}</workbook>`;
const styles = body => `<styleSheet xmlns="${S}">${body}</styleSheet>`;
const inline = (ref, text, extra = '') => `<c r="${ref}" t="inlineStr" ${extra}><is><t>${text}</t></is></c>`;
const formula = (ref, value, extra = '') => `<c r="${ref}" ${extra}><f>${value}</f><v>1</v></c>`;
function context(parts, options = {}) {
  const texts = new Map(Object.entries(parts));
  return { texts, xml: new Map([...texts].map(([part, value]) => [part, parseXml(value, part)])), entries: new Map([...texts].map(([part, value]) => [part, new TextEncoder().encode(value)])), options };
}
function inspect(parts, options) { return scanXlsx(context(parts, options)); }
function codes(report) { return report.issues.map(issue => issue.code); }
function issue(report, code) { return report.issues.find(item => item.code === code); }

test('finds style, rich-text, theme and chart fonts with explicit opt-in repair', () => {
  const report = inspect({
    'xl/styles.xml': styles('<fonts><font><name val="PingFang SC"/><sz val="11"/></font></fonts><dxfs><dxf><font><name val="Helvetica"/></font></dxf></dxfs>'),
    'xl/sharedStrings.xml': `<sst xmlns="${S}"><si><r><rPr><rFont val="PingFang SC"/></rPr><t>文字</t></r></si></sst>`,
    'xl/theme/theme1.xml': `<a:theme xmlns:a="${A}"><a:themeElements><a:fontScheme name="example"><a:majorFont><a:latin typeface="Helvetica Neue"/><a:font script="Hans" typeface="PingFang SC"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`,
    'xl/charts/chart1.xml': `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="${A}"><a:rPr><a:latin typeface="Menlo"/><a:ea typeface="PingFang SC"/></a:rPr></c:chart>`,
  });
  const replacements = report.issues.filter(item => item.fixable);
  assert.equal(replacements.length, 7);
  assert.ok(replacements.every(item => item.defaultSelected === false && item.fix.kind === 'font-replace'));
  assert.equal(replacements.find(item => item.part === 'xl/sharedStrings.xml').fix.scope, S);
  assert.equal(replacements.find(item => item.part === 'xl/theme/theme1.xml').fix.scope, A);
  assert.equal(report.inventory.fonts.find(item => item.name === 'PingFang SC').occurrences, 4);
  assert.equal(report.inventory.chartCount, 1);
});

test('known and theme fonts are not presented as missing; unknown remains unverified', () => {
  const report = inspect({
    'xl/styles.xml': styles('<fonts><font><name val="Arial"/></font><font><name val="Imaginary Bespoke Font"/></font></fonts>'),
    'xl/drawings/drawing1.xml': `<a:txBody xmlns:a="${A}"><a:p><a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr></a:r></a:p></a:txBody>`,
  });
  assert.deepEqual(codes(report), ['FONT_UNVERIFIED']);
  assert.equal(report.issues[0].fixable, false);
  assert.ok(report.issues[0].title.includes('Imaginary Bespoke Font'));
});

test('custom mapping is opt-in and does not treat unrelated names as fonts', () => {
  const report = inspect({
    'xl/styles.xml': styles('<fonts><font><name val="Example Serif"/></font></fonts><extLst><name val="Example Serif"/></extLst>'),
  }, { fontMappings: { 'Example Serif': 'Cambria' } });
  assert.equal(report.issues.length, 1);
  assert.equal(report.issues[0].fix.to, 'Cambria');
  assert.equal(report.issues[0].defaultSelected, false);
  assert.equal(report.inventory.fonts[0].occurrences, 1);
});

test('future functions are version requirements, not Windows errors; ordinary formulas and strings are ignored', () => {
  const report = inspect({
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1">${formula('A1', '_xlfn._xlws.FILTER(A2:A8,A2:A8&gt;0)')}${formula('B1', 'XLOOKUP(1,A2:A8,B2:B8)')}${formula('C1', 'SUM(A2:A8)')}${formula('D1', 'IF(A1=1,&quot;XLOOKUP(A1)&quot;,&quot;_xlfn.FILTER(A1)&quot;)')}${formula('E1', "'XLOOKUP(Weird)'!A1+1")}</row></sheetData>`),
  });
  const found = issue(report, 'FORMULA_VERSION_REQUIREMENT');
  assert.ok(found);
  assert.ok(found.evidence.includes('A1、B1'));
  assert.ok(!found.evidence.includes('C1'));
  assert.ok(!found.evidence.includes('D1'));
  assert.ok(!found.evidence.includes('E1'));
  assert.match(found.description, /不代表公式损坏或 Windows 不支持/);
  assert.equal(found.fixable, false);
  assert.equal(report.inventory.formulaCount, 5);
});

test('distinguishes external workbook formulas from table structured references', () => {
  const report = inspect({
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1">${formula('A1', "'[1]Sheet 1'!A1")}${formula('B1', '[2]Sheet1!B5')}${formula('C1', 'SUM(Table1[Amount])+Sheet2!A1')}${formula('D1', 'IF(A1=1,&quot;[2]Sheet!A1&quot;,0)')}</row></sheetData>`),
  });
  const found = issue(report, 'EXTERNAL_WORKBOOK_FORMULA');
  assert.match(found.evidence, /2 个单元格：A1、B1/);
  assert.ok(!found.evidence.includes('C1'));
  assert.ok(!found.evidence.includes('D1'));
});

test('reports environment-dependent formulas without rewriting them', () => {
  const report = inspect({
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1">${formula('A1', 'WEBSERVICE(&quot;https://example.invalid&quot;)')}${formula('B1', '_xll.CustomFunction(A1)')}${formula('C1', '_xlfn.IMAGE(A1)')}</row></sheetData>`),
  });
  const found = issue(report, 'FORMULA_ENVIRONMENT_DEPENDENCY');
  assert.match(found.evidence, /3 个单元格/);
  assert.equal(found.fixable, false);
});

test('separates long numeric values, preserved text, lexical zeros and number display formats', () => {
  const report = inspect({
    'xl/styles.xml': styles('<numFmts><numFmt numFmtId="164" formatCode="000000"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs>'),
    'xl/sharedStrings.xml': `<sst xmlns="${S}"><si><t>000123</t><rPh sb="0" eb="1"><t>phonetic</t></rPh></si></sst>`,
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1"><c r="A1"><v>12345678901234567</v></c><c r="B1" t="n"><v>1.2345678901234567E+16</v></c>${inline('C1', '12345678901234567')}<c r="D1" t="s"><v>0</v></c><c r="E1"><v>00123</v></c><c r="F1" s="1"><v>123</v></c><c r="G1"><v>12345</v></c><c r="H1" t="b"><v>1</v></c></row></sheetData>`),
  });
  assert.match(issue(report, 'NUMERIC_PRECISION_RISK').evidence, /2 个单元格：A1、B1/);
  assert.match(issue(report, 'TEXT_IDENTIFIER_PRESERVED').evidence, /2 个单元格：C1、D1/);
  assert.match(issue(report, 'NUMERIC_LEADING_ZEROS').evidence, /1 个单元格：E1/);
  assert.match(issue(report, 'ZERO_PADDED_DISPLAY').evidence, /1 个单元格：F1/);
  assert.ok(report.issues.every(item => !item.fixable));
  assert.match(issue(report, 'NUMERIC_PRECISION_RISK').description, /不能仅凭当前文件判断/);
});

test('1904 date system is informational and is never automatically converted', () => {
  for (const flag of ['1', 'true']) {
    const report = inspect({ 'xl/workbook.xml': workbook(`<workbookPr date1904="${flag}"/>`) });
    const found = issue(report, 'DATE_SYSTEM_1904');
    assert.equal(found.severity, 'info');
    assert.equal(found.fixable, false);
    assert.match(found.description, /Windows 和 Mac Excel 都支持/);
    assert.equal(report.inventory.dateSystem, '1904');
  }
  assert.ok(!codes(inspect({ 'xl/workbook.xml': workbook('<workbookPr date1904="false"/>') })).includes('DATE_SYSTEM_1904'));
});

test('preserves sheet names and reports manual calculation and named future formulas', () => {
  const report = inspect({
    'xl/workbook.xml': workbook('<sheets><sheet name="收入汇总" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Latest">_xlfn.XLOOKUP(1,Sheet1!A:A,Sheet1!B:B)</definedName><definedName name="_xlnm.Print_Area">Sheet1!A1:B9</definedName></definedNames><calcPr calcMode="manual"/>'),
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="${NS.r}/worksheet"/></Relationships>`,
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1">${formula('A1', '_xlfn.XLOOKUP(1,B:B,C:C)')}</row></sheetData>`),
  });
  assert.equal(report.inventory.worksheets[0].name, '收入汇总');
  assert.match(issue(report, 'FORMULA_VERSION_REQUIREMENT').location, /^收入汇总/);
  assert.ok(codes(report).includes('MANUAL_CALCULATION'));
  assert.ok(codes(report).includes('NAMED_FORMULA_VERSION'));
  assert.equal(report.inventory.hasPrintSettings, true);
});

test('clipping heuristic requires occupied adjacent cell and skips wrapped, merged and hidden content', () => {
  const report = inspect({
    'xl/styles.xml': styles('<fonts><font><sz val="11"/></font></fonts><cellXfs><xf fontId="0"/><xf fontId="0"><alignment wrapText="1"/></xf><xf fontId="0"><alignment shrinkToFit="1"/></xf></cellXfs>'),
    'xl/worksheets/sheet1.xml': sheet(`<cols><col min="1" max="10" width="5"/><col min="11" max="12" width="2" hidden="1"/></cols><sheetData><row r="1">${inline('A1', '一段很长的中文文字')}${inline('B1', '占用')}${inline('D1', '一段很长的中文文字')}${inline('F1', '一段很长的中文文字', 's="1"')}${inline('G1', '占用')}${inline('I1', '一段很长的中文文字')}${inline('J1', '占用')}${inline('K1', '一段很长的中文文字')}${inline('L1', '占用')}</row><row r="2">${inline('A2', '一段很长的中文文字', 's="2"')}${inline('B2', '占用')}</row></sheetData><mergeCells><mergeCell ref="I1:J1"/></mergeCells>`),
  });
  const found = issue(report, 'TEXT_CLIPPING_HEURISTIC');
  assert.match(found.evidence, /^1 个单元格：A1/);
  assert.equal(found.severity, 'info');
  assert.match(found.description, /不能证明存在截断/);
  assert.equal(found.fixable, false);
});

test('wrapped text in short fixed-height rows and print settings generate review hints only', () => {
  const report = inspect({
    'xl/styles.xml': styles('<fonts><font><sz val="11"/></font></fonts><cellXfs><xf fontId="0"><alignment wrapText="1"/></xf></cellXfs>'),
    'xl/worksheets/sheet1.xml': sheet(`<cols><col min="1" max="1" width="4"/></cols><sheetData><row r="1" customHeight="1" ht="15">${inline('A1', '很长的换行文字需要多行才能放得下')}</row></sheetData><pageSetup paperSize="9" orientation="landscape"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`),
  });
  assert.ok(codes(report).includes('WRAPPED_TEXT_HEIGHT_HEURISTIC'));
  assert.ok(codes(report).includes('PRINT_LAYOUT_REVIEW'));
  assert.ok(report.issues.every(item => !item.fixable));
});

test('detects external connections, OLE, controls, macros and genuine dynamic-array metadata', () => {
  const ctx = context({
    'xl/externalLinks/externalLink1.xml': `<externalLink xmlns="${S}"><externalBook/></externalLink>`,
    'xl/connections.xml': `<connections xmlns="${S}"><connection id="1" name="Example source"/></connections>`,
    'xl/worksheets/sheet1.xml': sheet('<oleObjects><oleObject progId="Excel.Sheet.12"/></oleObjects><controls><control shapeId="1"/></controls>'),
    'xl/metadata.xml': `<metadata xmlns="${S}" xmlns:da="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><da:dynamicArrayProperties fDynamic="1"/><da:dynamicArrayProperties fDynamic="0"/></metadata>`,
    'xl/workbook.xml': workbook('<definedNames><definedName name="RunMacro" xlm="1">Sheet1!A1</definedName></definedNames>'),
  });
  ctx.entries.set('xl/vbaProject.bin', new Uint8Array([1, 2, 3]));
  ctx.entries.set('xl/embeddings/oleObject1.bin', new Uint8Array([4, 5]));
  const report = scanXlsx(ctx);
  for (const code of ['EXTERNAL_LINK_PART', 'DATA_CONNECTION', 'EMBEDDED_OBJECT', 'WORKBOOK_CONTROL', 'DYNAMIC_ARRAY_METADATA', 'MACRO_CONTENT', 'MACRO_DEFINED_NAME']) assert.ok(codes(report).includes(code), code);
  assert.equal(report.inventory.externalLinks, 1);
  assert.equal(report.inventory.connectionCount, 1);
  assert.equal(report.inventory.embeddedObjects, 1);
  assert.match(issue(report, 'DYNAMIC_ARRAY_METADATA').evidence, /1 处/);
});

test('scan is deterministic and leaves every byte and DOM value untouched', () => {
  const ctx = context({
    'xl/styles.xml': styles('<fonts><font><name val="PingFang SC"/></font></fonts>'),
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1">${formula('A1', '_xlfn.XLOOKUP(1,B:B,C:C)')}<c r="B1"><v>12345678901234567</v></c></row></sheetData>`),
  });
  const serializer = new XMLSerializer();
  const domBefore = [...ctx.xml].map(([part, doc]) => [part, serializer.serializeToString(doc)]);
  const bytesBefore = [...ctx.entries].map(([part, bytes]) => [part, Array.from(bytes)]);
  const first = scanXlsx(ctx);
  const second = scanXlsx(ctx);
  assert.deepEqual(first, second);
  assert.deepEqual([...ctx.xml].map(([part, doc]) => [part, serializer.serializeToString(doc)]), domBefore);
  assert.deepEqual([...ctx.entries].map(([part, bytes]) => [part, Array.from(bytes)]), bytesBefore);
});

test('empty and minimal packages can be inspected without inventing findings', () => {
  assert.deepEqual(scanXlsx(context({})).issues, []);
  const report = inspect({ 'xl/worksheets/sheet1.xml': sheet('<sheetData/>') });
  assert.deepEqual(report.issues, []);
  assert.equal(report.inventory.worksheets.length, 1);
  assert.equal(report.inventory.formulaCount, 0);
});

test('checks conditional formatting, data validation, table and named dependency formulas', () => {
  const report = inspect({
    'xl/workbook.xml': workbook('<definedNames><definedName name="Remote">\'[1]Source\'!A1</definedName><definedName name="Online">WEBSERVICE("https://example.invalid")</definedName></definedNames>'),
    'xl/worksheets/sheet1.xml': sheet('<conditionalFormatting sqref="A1:A9"><cfRule type="expression"><formula>_xlfn.LET(x,A1,x&gt;0)</formula></cfRule></conditionalFormatting><dataValidations><dataValidation sqref="B1"><formula1>_xlfn._xlws.FILTER(A1:A9,A1:A9&gt;0)</formula1></dataValidation></dataValidations>'),
    'xl/tables/table1.xml': `<table xmlns="${S}"><tableColumns><tableColumn id="1" name="Result"><calculatedColumnFormula>_xlfn.XLOOKUP([@ID],B:B,C:C)</calculatedColumnFormula></tableColumn></tableColumns></table>`,
  });
  assert.equal(report.issues.filter(item => item.code === 'OTHER_FORMULA_VERSION').length, 2);
  assert.match(report.issues.find(item => item.code === 'OTHER_FORMULA_VERSION' && item.part.includes('worksheets')).evidence, /formula A1:A9/);
  assert.ok(codes(report).includes('OTHER_FORMULA_EXTERNAL_LINK'));
  assert.ok(codes(report).includes('OTHER_FORMULA_DEPENDENCY'));
  assert.ok(!codes(report).includes('EXTERNAL_WORKBOOK_FORMULA'));
});

test('large merge-rule cross products skip only layout heuristics and disclose the limit', () => {
  const cells = Array.from({ length: 1500 }, (_, index) => `<row r="${index + 1}">${inline(`A${index + 1}`, '一段特别长的文字')}${inline(`B${index + 1}`, '占用')}</row>`).join('');
  const merges = Array.from({ length: 1000 }, (_, index) => `<mergeCell ref="C${index + 1}:D${index + 1}"/>`).join('');
  const report = inspect({ 'xl/worksheets/sheet1.xml': sheet(`<sheetData>${cells}</sheetData><mergeCells>${merges}</mergeCells>`) });
  assert.ok(codes(report).includes('LAYOUT_HEURISTIC_LIMIT'));
  assert.ok(!codes(report).includes('TEXT_CLIPPING_HEURISTIC'));
  assert.equal(report.inventory.worksheets[0].cellCount, 3000);
});

test('unused supplemental theme fonts are aggregated while explicit fonts stay individually visible', () => {
  const report = inspect({
    'xl/theme/theme1.xml': `<a:theme xmlns:a="${A}"><a:themeElements><a:fontScheme name="Example"><a:majorFont><a:latin typeface="Explicit Unknown Serif"/><a:font script="Thai" typeface="Unknown Thai"/><a:font script="Tibt" typeface="Unknown Tibetan"/></a:majorFont><a:minorFont><a:font script="Thai" typeface="Unknown Thai"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
  });
  assert.equal(report.issues.length, 2);
  assert.equal(report.issues.filter(item => item.code === 'FONT_UNVERIFIED').length, 1);
  const grouped = issue(report, 'THEME_FONTS_UNVERIFIED');
  assert.match(grouped.title, /^2 种/);
  assert.match(grouped.description, /不代表可见文字实际用到了/);
  assert.equal(grouped.fixable, false);
  assert.equal(report.inventory.fonts.find(font => font.name === 'Unknown Thai').occurrences, 2);
});

test('default printer margins and empty print containers do not create noise', () => {
  const report = inspect({
    'xl/worksheets/sheet1.xml': sheet('<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="portrait" scale="100"/><headerFooter/><rowBreaks count="0"/><colBreaks count="0"/>'),
  });
  assert.ok(!codes(report).includes('PRINT_LAYOUT_REVIEW'));
  const otherDefault = inspect({ 'xl/worksheets/sheet1.xml': sheet('<sheetData/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>') });
  assert.ok(!codes(otherDefault).includes('PRINT_LAYOUT_REVIEW'));
  const custom = inspect({ 'xl/worksheets/sheet1.xml': sheet('<sheetData/><pageMargins left="0.1" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>') });
  assert.ok(codes(custom).includes('PRINT_LAYOUT_REVIEW'));
});

test('reports replacement characters as evidence without guessing original text', () => {
  const report = inspect({
    'xl/sharedStrings.xml': `<sst xmlns="${S}"><si><t>有�字符</t></si></sst>`,
    'xl/worksheets/sheet1.xml': sheet(`<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c>${inline('B1', '_xFFFD_')}${inline('C1', '_x005F_xFFFD_')}${inline('D1', '_x0030_123')}</row></sheetData>`),
  });
  const found = issue(report, 'TEXT_REPLACEMENT_CHARACTER');
  assert.match(found.evidence, /^2 个单元格：A1、B1/);
  assert.equal(found.fixable, false);
  assert.match(found.description, /不能确定或还原原字/);
  assert.match(issue(report, 'TEXT_IDENTIFIER_PRESERVED').evidence, /D1/);
});

test('symbol and mathematical declarations block substitutions even when custom mappings exist', () => {
  const report = inspect({
    'xl/styles.xml': styles('<fonts><font><name val="Wingdings"/></font><font><name val="Cambria Math"/></font><font><name val="SpecialGlyphs"/><charset val="2"/></font></fonts>'),
    'xl/charts/chart1.xml': `<a:txBody xmlns:a="${A}"><a:p><a:r><a:rPr><a:sym typeface="Helvetica"/><a:latin typeface="Helvetica"/></a:rPr></a:r></a:p></a:txBody>`,
  }, { fontMappings: { Wingdings: 'Arial', 'Cambria Math': 'Arial', SpecialGlyphs: 'Arial', Helvetica: 'Arial' } });
  assert.ok(report.issues.every(item => !item.fixable));
  const warnings = report.issues.filter(item => item.code === 'SYMBOL_FONT_REVIEW');
  assert.ok(warnings.some(item => item.title.includes('SpecialGlyphs')));
  assert.ok(warnings.some(item => item.title.includes('Helvetica')));
  // Known mathematical/symbol families may stay quiet when the shared font
  // policy already suppresses their mappings; they must never become fixes.
});

test('local, relative, UNC and localhost spreadsheet hyperlinks are actionable dependencies', () => {
  const targets = [
    'file:///Users/example/Documents/attachment.pdf', '/Users/example/attachment.pdf',
    '../attachments/说明.pdf', '\\\\server\\share\\attachment.pdf',
    'http://localhost:8080/report', 'https://reports.localhost/report',
    'http://127.0.0.1:8080/report', 'http://[::1]:8080/report',
  ];
  const rels = targets.map((target, i) => `<Relationship Id="r${i}" Type="${NS.r}/hyperlink" Target="${escapeAttribute(target)}" TargetMode="External"/>`).join('');
  const report = inspect({
    'xl/workbook.xml': workbook('<sheets><sheet name="附件列表" sheetId="1" r:id="sheet"/></sheets>'),
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="sheet" Type="${NS.r}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': sheet(`<sheetData/><hyperlinks xmlns:r="${NS.r}">${targets.map((_, i) => `<hyperlink ref="A${i + 1}" r:id="r${i}"/>`).join('')}</hyperlinks>`),
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${NS.rel}">${rels}</Relationships>`,
  });
  const found = report.issues.filter(item => item.code === 'EXTERNAL_RESOURCE_RELATIONSHIP');
  assert.equal(found.length, targets.length);
  targets.forEach((target, i) => {
    assert.ok(found.some(item => item.location === `附件列表 · A${i + 1}` && item.evidence.endsWith(target)));
  });
  assert.ok(found.every(item => item.severity === 'warning' && !item.fixable));
  assert.ok(!codes(report).includes('EXTERNAL_WEB_HYPERLINK'));
});

test('ordinary web links are informational, linked images remain dependencies, and internal links stay quiet', () => {
  const report = inspect({
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${NS.rel}">
      <Relationship Id="web" Type="${NS.r}/hyperlink" Target="https://example.invalid/reference" TargetMode="External"/>
      <Relationship Id="mail" Type="${NS.r}/hyperlink" Target="mailto:help@example.invalid" TargetMode="External"/>
      <Relationship Id="phone" Type="${NS.r}/hyperlink" Target="tel:+1234567890" TargetMode="External"/>
      <Relationship Id="internal" Type="${NS.r}/hyperlink" Target="#A1"/>
    </Relationships>`,
    'xl/drawings/_rels/drawing1.xml.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="image" Type="${NS.r}/image" Target="https://example.invalid/logo.png" TargetMode="External"/></Relationships>`,
  });
  const web = report.issues.filter(item => item.code === 'EXTERNAL_WEB_HYPERLINK');
  assert.equal(web.length, 3);
  assert.ok(web.every(item => item.severity === 'info' && !item.fixable));
  assert.ok(web.every(item => item.description.includes('本身不是兼容性错误')));
  const resource = issue(report, 'EXTERNAL_RESOURCE_RELATIONSHIP');
  assert.equal(resource.severity, 'warning');
  assert.equal(resource.location, 'xl/drawings/drawing1.xml');
  assert.match(resource.evidence, /logo\.png/);
  assert.equal(report.issues.length, 4);
});

test('real-package hyperlink diagnostics survive font repair without changing links or worksheet bytes', async () => {
  const entries = readArchive(createDemo('xlsx')).entries;
  const encoder = new TextEncoder(), decoder = new TextDecoder();
  const sheetPart = 'xl/worksheets/sheet1.xml', relPart = 'xl/worksheets/_rels/sheet1.xml.rels';
  const links = `<hyperlinks xmlns:r="${NS.r}"><hyperlink ref="A2" r:id="local"/><hyperlink ref="A3" r:id="web"/></hyperlinks>`;
  entries.set(sheetPart, encoder.encode(decoder.decode(entries.get(sheetPart)).replace('<printOptions', `${links}<printOptions`)));
  entries.set(relPart, encoder.encode(`<Relationships xmlns="${NS.rel}"><Relationship Id="local" Type="${NS.r}/hyperlink" Target="file:///Users/example/attachment.pdf" TargetMode="External"/><Relationship Id="web" Type="${NS.r}/hyperlink" Target="https://example.invalid/reference" TargetMode="External"/></Relationships>`));
  const bytes = writeArchive(entries), original = bytes.slice();
  const report = await analyzeOffice(bytes, 'attachments.xlsx');
  assert.equal(report.repairBlocked, false);
  assert.equal(report.inventory.externalRelationshipCount, 2);
  assert.equal(issue(report, 'EXTERNAL_RESOURCE_RELATIONSHIP').location, '预算示例 · A2');
  assert.equal(issue(report, 'EXTERNAL_WEB_HYPERLINK').location, '预算示例 · A3');
  const repaired = await repairOffice(bytes, 'attachments.xlsx', { selectedIssueIds: report.issues.filter(item => item.fixable).map(item => item.id) });
  const after = readArchive(repaired.bytes).entries;
  assert.deepEqual(after.get(relPart), entries.get(relPart));
  assert.deepEqual(after.get(sheetPart), entries.get(sheetPart));
  assert.deepEqual(bytes, original);
  assert.ok(issue(repaired.report, 'EXTERNAL_RESOURCE_RELATIONSHIP'));
  assert.ok(issue(repaired.report, 'EXTERNAL_WEB_HYPERLINK'));
});
