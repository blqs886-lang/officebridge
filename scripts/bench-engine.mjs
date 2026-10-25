// Optional synthetic load check. Uses no user documents and writes no files.
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
import { createDemo } from '../src/engine/demo.js';
import { readArchive, writeArchive } from '../src/engine/archive.js';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';

const rows = Number(process.argv[2] || 20000);
if (!Number.isSafeInteger(rows) || rows < 1 || rows > 50000) throw new Error('Choose 1–50000 synthetic rows.');
const entries = readArchive(createDemo('xlsx')).entries;
let sheet = new TextDecoder().decode(entries.get('xl/worksheets/sheet1.xml'));
const data = [];
for (let i = 1; i <= rows; i++) data.push(`<row r="${i}"><c r="A${i}" s="1" t="inlineStr"><is><t>Synthetic ${i}</t></is></c><c r="B${i}"><v>${i * 17}</v></c><c r="C${i}"><f>B${i}*2</f><v>${i * 34}</v></c></row>`);
sheet = sheet.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${data.join('')}</sheetData>`)
  .replace(/<dimension\b[^>]*\/>/, `<dimension ref="A1:C${rows}"/>`);
entries.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(sheet));
const bytes = writeArchive(entries), copy = bytes.slice();
const start = performance.now();
const report = await analyzeOffice(bytes, 'synthetic-load.xlsx');
const checked = performance.now();
const repaired = await repairOffice(bytes, 'synthetic-load.xlsx', {
  selectedIssueIds: report.issues.filter(i => i.fixable).map(i => i.id),
});
const finished = performance.now();
assert.equal(report.repairBlocked, false);
assert.equal(report.inventory.formulaCount, rows);
assert.deepEqual(bytes, copy);
assert.deepEqual(readArchive(repaired.bytes).entries.get('xl/worksheets/sheet1.xml'), entries.get('xl/worksheets/sheet1.xml'));
console.log(JSON.stringify({ rows, cells: rows * 3, inputBytes: bytes.length,
  xmlBytes: entries.get('xl/worksheets/sheet1.xml').length, formulas: report.inventory.formulaCount,
  checkMs: Math.round(checked - start), repairMs: Math.round(finished - checked),
  sourceUnchanged: true, worksheetBytesUnchanged: true, windowsOfficeTested: false }, null, 2));
