import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readNativeReport } from '../src/ui/native-report-data.js';

const pdf = Buffer.from('%PDF-1.7\nexample bytes for integrity verification');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
function record() {
  return {schema:'officebridge-native-check',schemaVersion:1,environment:{platform:'macos'},createdAt:'2026-11-12T00:00:00Z',
    documents:[{role:'original',format:'xlsx',sha256:'a'.repeat(64),fileName:'test.xlsx',openedSuccessfully:true,sourceUnchanged:true,
      pdf:{fileName:'test.pdf',sha256:hash(pdf).toUpperCase(),base64:pdf.toString('base64')}}],
    comparisons:[{details:['one page']}]};
}
test('native reports retain exact PDF bytes and distinguish target platform and native Office checks', async () => {
  const actual=await readNativeReport('\ufeff'+JSON.stringify(record()));
  assert.equal(actual.platform,'macos');
  assert.deepEqual(actual.documents[0].pdf.bytes,new Uint8Array(pdf));
  assert.equal(actual.documents[0].pdf.sha256,hash(pdf));
  assert.equal(actual.documents[0].openedSuccessfully,true);
  assert.deepEqual(actual.comparisons,['one page']);
});
test('native import rejects a damaged preview instead of attaching it to the wrong Office result', async () => {
  const wrong=record();wrong.documents[0].pdf.base64=Buffer.from('%PDF-1.7\ndifferent file').toString('base64');
  await assert.rejects(readNativeReport(JSON.stringify(wrong)),/指纹不一致/);
  const url=record();url.documents[0].pdf={url:'https://example.invalid/user-document.pdf'};
  await assert.rejects(readNativeReport(JSON.stringify(url)),/预览/);
});
test('native import rejects unsupported schema, duplicate roles, mixed formats and missing original', async () => {
  const variants=[];
  const unknown=record();unknown.schemaVersion=90;variants.push(unknown);
  const double=record();double.documents.push({...double.documents[0]});variants.push(double);
  const mixed=record();mixed.documents.push({...mixed.documents[0],role:'copy',format:'pptx'});variants.push(mixed);
  const missing=record();missing.documents[0].role='copy';variants.push(missing);
  for(const input of variants)await assert.rejects(readNativeReport(JSON.stringify(input)));
});
test('failed Office reads remain failures, and untrusted source labels are never treated as verification', async () => {
  const failed=record();failed.documents[0].openedSuccessfully=false;failed.documents[0].sourceUnchanged='true';
  delete failed.documents[0].pdf;failed.documents[0].fileName='<img src=x onerror=alert(1)>';
  failed.documents[0].error='Office needs sign-in';
  const actual=await readNativeReport(JSON.stringify(failed));
  assert.equal(actual.documents[0].openedSuccessfully,false);
  assert.equal(actual.documents[0].sourceUnchanged,false);
  assert.equal(actual.documents[0].pdf,null);
  assert.equal(actual.documents[0].fileName,failed.documents[0].fileName);
});

test('early failures without a readable fingerprint remain actionable, never verified', async () => {
  const failed=record();
  Object.assign(failed.documents[0],{sha256:null,openedSuccessfully:false,sourceUnchanged:false,pdf:null,error:'文件已被锁定，请关闭后重试。'});
  const actual=await readNativeReport(JSON.stringify(failed));
  assert.equal(actual.documents[0].sha256,'');
  assert.match(actual.documents[0].error,/锁定/);
  for (const misleading of [{openedSuccessfully:true},{sourceUnchanged:true},{pdf:record().documents[0].pdf},{sha256:'not-a-hash'}]) {
    const input=structuredClone(failed);Object.assign(input.documents[0],misleading);
    await assert.rejects(readNativeReport(JSON.stringify(input)),/指纹/);
  }
});

test('Office font inventory is normalized independently from document and PDF evidence', async () => {
  const input=record();input.environment.fontFamilies=['Arial','Microsoft YaHei','arial','PingFang SC'];
  const result=await readNativeReport(JSON.stringify(input));
  assert.deepEqual(result.fontFamilies,['Arial','Microsoft YaHei','PingFang SC']);
  input.environment.fontFamilies=['Arial','<script>'];
  const bad=await readNativeReport(JSON.stringify(input));
  assert.equal(bad.fontFamilies,null);
  assert.match(bad.fontListNote,/名单无效/);
  assert.equal(bad.documents[0].openedSuccessfully,true);
  assert.ok(bad.documents[0].pdf);
});
