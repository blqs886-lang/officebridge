import test from 'node:test';
import assert from 'node:assert/strict';
import { XMLSerializer } from '@xmldom/xmldom';
import { scanPptx } from '../src/engine/pptx.js';
import { NS, parseXml } from '../src/engine/xml.js';
import { readArchive } from '../src/engine/archive.js';
import { createDemo } from '../src/engine/demo.js';

const namespaces = `xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"`;
const presentation = `<p:presentation ${namespaces}><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`;
const slide = content => `<p:sld ${namespaces}><p:cSld><p:spTree>${content}</p:spTree></p:cSld></p:sld>`;
const drawing = content => `<a:theme xmlns:a="${NS.a}">${content}</a:theme>`;
const font = name => `<a:latin typeface="${name}"/>`;

function context(parts = {}, options = {}, binaries = {}) {
  const entries = new Map();
  const xml = new Map();
  const texts = new Map();
  for (const [part, text] of Object.entries(parts)) {
    entries.set(part, new TextEncoder().encode(text));
    texts.set(part, text);
    xml.set(part, parseXml(text, part));
  }
  for (const [part, bytes] of Object.entries(binaries)) entries.set(part, bytes);
  return { entries, xml, texts, options: { fontProfile: 'windows-office', ...options } };
}

function textShape({ text = '示例文字', x = 0, y = 0, width = 4000000, height = 2000000, size = 2400,
  lines = 1, autoFit = 'noAutofit', rotation = '', bodyAttributes = '', paragraphProperties = '', geometry = 'rect' } = {}) {
  const paragraphs = Array.from({ length: lines }, () => `<a:p>${paragraphProperties}<a:r><a:rPr${size == null ? '' : ` sz="${size}"`}/><a:t>${text}</a:t></a:r></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="示例文本框"/></p:nvSpPr><p:spPr><a:xfrm ${rotation}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="${geometry}"/></p:spPr><p:txBody><a:bodyPr ${bodyAttributes}>${autoFit ? `<a:${autoFit}/>` : ''}</a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

test('scans DrawingML fonts in slides, themes, masters, layouts and charts, including bullets and supplemental fonts', () => {
  const parts = {
    'ppt/presentation.xml': presentation,
    'ppt/slides/slide1.xml': slide(`<a:rPr>${font('PingFang SC')}</a:rPr><a:pPr><a:buFont typeface="Helvetica"/></a:pPr>`),
    'ppt/theme/theme1.xml': drawing('<a:majorFont><a:font script="Hans" typeface="PingFang SC"/></a:majorFont>'),
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster ${namespaces}>${font('PingFang SC')}</p:sldMaster>`,
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${namespaces}>${font('PingFang SC')}</p:sldLayout>`,
    'ppt/charts/chart1.xml': `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="${NS.a}">${font('PingFang SC')}</c:chartSpace>`,
  };
  const result = scanPptx(context(parts));
  const replacements = result.issues.filter(issue => issue.code === 'PPTX_FONT_REPLACEMENT');
  assert.equal(replacements.length, 6);
  assert.ok(replacements.every(issue => issue.fix.scope === NS.a && issue.fix.kind === 'font-replace' && !issue.defaultSelected));
  assert.equal(result.inventory.fonts.find(item => item.name === 'PingFang SC').parts.length, 5);
  assert.equal(result.inventory.fonts.find(item => item.name === 'Helvetica').declarations, 1);
  assert.equal(result.inventory.slideCount, 1);
  assert.equal(result.inventory.chartCount, 1);
});

test('theme references are retained without substituting symbolic font names', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(`${font('+mj-lt')}${font('+mn-ea')}`) }));
  assert.equal(result.issues.length, 0);
  assert.ok(result.inventory.fonts.every(item => item.status === 'theme-reference'));
});

test('common target fonts are not flagged, while unknown fonts are described as unverified', () => {
  const result = scanPptx(context({ 'ppt/theme/theme1.xml': drawing(`${font('Arial')}${font('Calibri')}${font('Custom Studio Typeface')}`) }));
  assert.equal(result.issues.length, 1);
  assert.equal(result.issues[0].code, 'PPTX_FONT_UNVERIFIED');
  assert.equal(result.issues[0].severity, 'info');
  assert.equal(result.issues[0].fixable, false);
  assert.match(result.issues[0].description, /不将未知字体判定为不兼容/);
});

test('deduplicates each font per namespace and part, while retaining declaration counts', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(`${font('Helvetica')}${font('Helvetica')}<a:ea typeface="Helvetica"/>`) }));
  assert.equal(result.issues.length, 1);
  assert.equal(result.inventory.fonts[0].declarations, 3);
  assert.equal(result.issues[0].fix.from, 'Helvetica');
});

test('repair candidates retain exact source attribute values even with whitespace', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(font(' Helvetica ')) }));
  assert.equal(result.issues[0].fix.from, ' Helvetica ');
  assert.equal(result.issues[0].fix.to, 'Arial');
});

test('unknown supplemental theme fonts are grouped without hiding directly declared unknown fonts', () => {
  const result = scanPptx(context({ 'ppt/theme/theme1.xml': drawing('<a:majorFont><a:font script="Hans" typeface="Studio Han"/><a:font script="Arab" typeface="Studio Arabic"/><a:font script="Hebr" typeface="Studio Hebrew"/></a:majorFont><a:latin typeface="Studio Han"/>') }));
  const aggregate = result.issues.find(issue => issue.code === 'PPTX_THEME_FONTS_UNVERIFIED');
  assert.ok(aggregate);
  assert.match(aggregate.evidence, /Studio Arabic/);
  assert.match(aggregate.evidence, /Studio Hebrew/);
  assert.doesNotMatch(aggregate.evidence, /Studio Han/);
  assert.equal(result.issues.filter(issue => issue.code === 'PPTX_FONT_UNVERIFIED').length, 1);
  assert.equal(result.inventory.fonts.length, 3);
});

test('font names in unrelated namespaces are not treated as DrawingML fonts', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide('<x:latin xmlns:x="https://example.invalid/custom" typeface="Helvetica"/>') }));
  assert.equal(result.inventory.fonts.length, 0);
  assert.equal(result.issues.length, 0);
});

test('symbol and mathematical font declarations never offer ordinary font replacement', () => {
  const result = scanPptx(context({
    'ppt/slides/slide1.xml': slide(`${font('Wingdings')}${font('Cambria Math')}<a:sym typeface="Helvetica"/>${font('Helvetica')}`),
  }, { fontMappings: { Wingdings: 'Arial', 'Cambria Math': 'Arial', Helvetica: 'Arial' } }));
  assert.equal(result.issues.filter(issue => issue.code === 'PPTX_SYMBOL_FONT').length, 3);
  assert.equal(result.issues.filter(issue => issue.fixable).length, 0);
});

test('symbol charset declarations and private-use bullet characters retain a glyph warning', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide('<a:pPr><a:buFont typeface="Helvetica" charset="2"/><a:buChar char="&#xE000;"/></a:pPr>') }));
  assert.ok(result.issues.some(issue => issue.code === 'PPTX_SYMBOL_FONT'));
  assert.ok(result.issues.some(issue => issue.code === 'PPTX_PRIVATE_USE_GLYPHS'));
  assert.equal(result.issues.filter(issue => issue.fixable).length, 0);
});

test('flags private-use characters, replacement characters and complex emoji without promising recovery', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(textShape({ text: '\uE123 \u{F0123} � 👩‍💻' })) }));
  for (const code of ['PPTX_PRIVATE_USE_GLYPHS', 'PPTX_REPLACEMENT_CHARACTER', 'PPTX_COMPLEX_GLYPHS']) {
    assert.ok(result.issues.some(issue => issue.code === code && !issue.fixable));
  }
  assert.match(result.issues.find(issue => issue.code === 'PPTX_REPLACEMENT_CHARACTER').description, /不能恢复原字符/);
});

test('does not claim ordinary Chinese characters are missing glyphs', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(textShape({ text: '中文简体 繁體 日本語 한글' })) }));
  assert.equal(result.issues.length, 0);
});

test('ordinary web hyperlinks are informational, while linked files and media are dependencies', () => {
  const result = scanPptx(context({ 'ppt/slides/_rels/slide1.xml.rels': `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${NS.r}/hyperlink" Target="https://example.invalid/slide" TargetMode="External"/>
    <Relationship Id="rId2" Type="${NS.r}/hyperlink" Target="file:///Users/example/notes.pdf" TargetMode="External"/>
    <Relationship Id="rId3" Type="${NS.r}/video" Target="../movie.mp4" TargetMode="External"/>
    <Relationship Id="rId4" Type="${NS.r}/image" Target="../media/image.png"/>
  </Relationships>` }));
  assert.equal(result.inventory.externalLinks.length, 3);
  assert.equal(result.issues.filter(issue => issue.code === 'PPTX_EXTERNAL_DEPENDENCY').length, 2);
  assert.equal(result.issues.find(issue => issue.code === 'PPTX_WEB_HYPERLINK').severity, 'info');
  assert.equal(result.inventory.externalLinks[0].sourcePart, 'ppt/slides/slide1.xml');
});

test('embedded media uses MIME declarations where available and never infers codec support from extension', () => {
  const result = scanPptx(context({
    '[Content_Types].xml': `<Types xmlns="${NS.ct}"><Default Extension="bin" ContentType="video/mp4"/><Override PartName="/ppt/media/sound.mp4" ContentType="audio/mp4"/></Types>`,
  }, {}, {
    'ppt/media/sound.mp4': new Uint8Array([1, 2]),
    'ppt/media/movie.bin': new Uint8Array([3, 4]),
    'ppt/media/image.png': new Uint8Array([5, 6]),
  }));
  assert.equal(result.inventory.media.length, 3);
  assert.equal(result.inventory.media.find(item => item.part.endsWith('sound.mp4')).kind, 'audio');
  assert.equal(result.inventory.media.find(item => item.part.endsWith('movie.bin')).kind, 'video');
  const mediaIssues = result.issues.filter(issue => issue.code === 'PPTX_EMBEDDED_MEDIA');
  assert.equal(mediaIssues.length, 2);
  assert.ok(mediaIssues.every(issue => /扩展名不能证明内部编码兼容/.test(issue.description)));
});

test('embedded chart workbooks alone do not produce OLE or external-application warnings', () => {
  const result = scanPptx(context({}, {}, { 'ppt/embeddings/Microsoft_Excel_Worksheet.xlsx': new Uint8Array([1, 2]) }));
  assert.equal(result.inventory.embeddedFiles.length, 1);
  assert.equal(result.inventory.embeddedObjects.length, 0);
  assert.equal(result.issues.length, 0);
});

test('actual OLE objects and controls are reported without executing or flattening them', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide('<p:oleObj name="Workbook" progId="Excel.Sheet.12"/><p:control name="Button"/>') }));
  assert.equal(result.inventory.embeddedObjects.length, 2);
  assert.equal(result.issues[0].code, 'PPTX_EMBEDDED_OBJECT');
  assert.equal(result.issues[0].fixable, false);
});

test('empty timing and ordinary preset or custom shapes are not compatibility warnings', () => {
  const content = textShape({ geometry: 'roundRect' }) + '<p:sp><p:spPr><a:custGeom/></p:spPr></p:sp><p:timing><p:tnLst/></p:timing>';
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(content) }));
  assert.equal(result.issues.length, 0);
});

test('animation and advanced-graphics findings describe review needs, not confirmed breakage', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide('<p:timing><p:animMotion/></p:timing><p:transition><p:fade/></p:transition><a:scene3d/><a:prstTxWarp prst="textArchUp"/><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>') }));
  assert.equal(result.inventory.animations[0].animationNodes, 1);
  assert.equal(result.inventory.animations[0].transitions, 1);
  assert.equal(result.issues.find(issue => issue.code === 'PPTX_ADVANCED_GRAPHICS').severity, 'info');
  assert.match(result.issues.find(issue => issue.code === 'PPTX_ANIMATION_REVIEW').description, /不意味着不兼容/);
});

test('3D definitions in an unused theme style library do not assert visible 3D effects', () => {
  const result = scanPptx(context({ 'ppt/theme/theme1.xml': drawing('<a:fmtScheme><a:effectStyleLst><a:effectStyle><a:scene3d/><a:sp3d/></a:effectStyle></a:effectStyleLst></a:fmtScheme>') }));
  assert.equal(result.issues.filter(issue => issue.code === 'PPTX_ADVANCED_GRAPHICS').length, 0);
});

test('text bounds use slide dimensions but label intentional off-canvas placement as heuristic', () => {
  const result = scanPptx(context({ 'ppt/presentation.xml': presentation, 'ppt/slides/slide9.xml': slide(textShape({ x: 8000000, width: 3000000 })) }));
  const issue = result.issues.find(item => item.code === 'PPTX_TEXT_OUTSIDE_SLIDE');
  assert.ok(issue);
  assert.match(issue.title, /启发式/);
  assert.match(issue.location, /不一定是播放顺序/);
  assert.equal(issue.severity, 'info');
  assert.equal(issue.fixable, false);
});

test('overflow findings require explicit font sizes, fixed fit and large height discrepancy', () => {
  const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(textShape({ height: 200000, lines: 3, size: 2400 })) }));
  const issue = result.issues.find(item => item.code === 'PPTX_TEXT_OVERFLOW_HEURISTIC');
  assert.ok(issue);
  assert.match(issue.description, /不能确认实际溢出/);
  assert.equal(result.inventory.heuristicChecks.explicitLineHeightChecks, 1);
});

test('overflow estimation skips inherited fonts, autofit, line spacing, rotation, vertical text and groups', () => {
  const cases = [
    textShape({ height: 200000, lines: 4, size: null }),
    textShape({ height: 200000, lines: 4, autoFit: 'normAutofit' }),
    textShape({ height: 200000, lines: 4, paragraphProperties: '<a:pPr><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>' }),
    textShape({ height: 200000, lines: 4, rotation: 'rot="5400000"' }),
    textShape({ height: 200000, lines: 4, bodyAttributes: 'vert="vert"' }),
    `<p:grpSp>${textShape({ height: 200000, lines: 4 })}</p:grpSp>`,
  ];
  for (const content of cases) {
    const result = scanPptx(context({ 'ppt/slides/slide1.xml': slide(content) }));
    assert.equal(result.issues.filter(issue => issue.code.includes('TEXT_')).length, 0);
  }
});

test('embedded font declarations are not taken as proof of usable or complete glyphs', () => {
  const result = scanPptx(context({ 'ppt/presentation.xml': `<p:presentation ${namespaces}><p:embeddedFontLst><p:embeddedFont><p:font typeface="Sample Font"/><p:regular r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>` }));
  assert.deepEqual(result.inventory.embeddedFontFamilies, ['Sample Font']);
  assert.match(result.issues[0].description, /不能据此保证字体一致/);
  assert.equal(result.issues[0].fixable, false);
});

test('strict DrawingML fonts retain their real namespace in repair choices', () => {
  const strict = 'http://purl.oclc.org/ooxml/drawingml/main';
  const result = scanPptx(context({ 'ppt/theme/theme1.xml': `<a:theme xmlns:a="${strict}">${font('Helvetica')}</a:theme>` }));
  assert.equal(result.issues[0].fix.scope, strict);
});

test('scan is deterministic and leaves XML documents, source text and entry bytes unchanged', () => {
  const input = context({ 'ppt/slides/slide1.xml': slide(`${font('Helvetica')}${textShape({ text: '你好 🌏' })}`) });
  const serialize = new XMLSerializer();
  const beforeXml = [...input.xml].map(([name, doc]) => [name, serialize.serializeToString(doc)]);
  const beforeBytes = [...input.entries].map(([name, bytes]) => [name, [...bytes]]);
  const beforeTexts = [...input.texts];
  const first = scanPptx(input);
  assert.deepEqual(scanPptx(input), first);
  assert.deepEqual([...input.xml].map(([name, doc]) => [name, serialize.serializeToString(doc)]), beforeXml);
  assert.deepEqual([...input.entries].map(([name, bytes]) => [name, [...bytes]]), beforeBytes);
  assert.deepEqual([...input.texts], beforeTexts);
});

test('complete synthetic presentation remains concise despite a large default theme font/style library', () => {
  const archive = readArchive(createDemo('pptx'));
  const parts = {};
  for (const [name, bytes] of archive.entries) {
    if (/\.(?:xml|rels)$/i.test(name)) parts[name] = new TextDecoder().decode(bytes);
  }
  const input = context(parts);
  input.entries = archive.entries;
  const result = scanPptx(input);
  assert.ok(result.inventory.slideCount > 0);
  assert.ok(result.inventory.fonts.length > 10);
  assert.ok(result.issues.some(issue => issue.code === 'PPTX_FONT_REPLACEMENT'));
  const themeFindings = result.issues.filter(issue => issue.code === 'PPTX_THEME_FONTS_UNVERIFIED');
  assert.ok(themeFindings.length > 0);
  assert.equal(new Set(themeFindings.map(issue => issue.part)).size, themeFindings.length);
  assert.equal(result.issues.filter(issue => issue.code === 'PPTX_ADVANCED_GRAPHICS').length, 0);
  assert.ok(result.issues.length < 20, `Unexpected report noise: ${result.issues.length} findings`);
});
