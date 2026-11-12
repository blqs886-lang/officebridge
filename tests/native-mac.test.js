import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {unzipSync,zipSync,strToU8,strFromU8} from 'fflate';
import {createDemo} from '../src/engine/demo.js';

const supported = process.platform === 'darwin';
const script = path.resolve('scripts/office-check-mac.js');
const base = () => unzipSync(createDemo('xlsx'));
async function check(entries, success) {
  const dir=await mkdtemp(path.join(tmpdir(),'officebridge-mac-guard-'));
  try {
    const file=path.join(dir,'input.xlsx');
    await writeFile(file,zipSync(entries));
    const run=spawnSync('/usr/bin/osascript',['-l','JavaScript',script,'--original',file,'--validate-only'],{encoding:'utf8',timeout:15000,maxBuffer:1024*1024});
    assert.equal(run.status===0,success,run.stdout+'\n'+run.stderr);
    if(success)assert.equal(JSON.parse(run.stdout).format,'xlsx');
  } finally {await rm(dir,{recursive:true,force:true});}
}
test('mac native preflight accepts normal local formulas and an ordinary hyperlink',{skip:!supported},async()=>{
  const e=base();e['xl/worksheets/_rels/sheet1.xml.rels']=strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External" Target="https://example.com/"/></Relationships>');await check(e,true);
});
test('mac native preflight rejects malformed XML inside an otherwise valid ZIP',{skip:!supported},async()=>{
 const e=base();e['xl/workbook.xml']=strToU8('<workbook><broken></workbook>');await check(e,false);
});
test('mac native preflight blocks online formulas in cells, names and validation',{skip:!supported},async()=>{
 for(const [part,change] of [
  ['xl/worksheets/sheet1.xml',s=>s.replace('</worksheet>','<dataValidations><dataValidation><formula1>STOCKHISTORY(A1)</formula1></dataValidation></dataValidations></worksheet>')],
  ['xl/workbook.xml',s=>s.replace('</workbook>','<definedNames><definedName name="fetch">WEBSERVICE("https://example.com")</definedName></definedNames></workbook>')],
  ['xl/worksheets/sheet1.xml',s=>s.replace(/<f>[^<]+<\/f>/,'<f>_xll.Run(A1)</f>')],
 ]){const e=base();e[part]=strToU8(change(strFromU8(e[part])));await check(e,false);}
});
test('mac native preflight rejects VBA, active names, external media and unsafe paths',{skip:!supported},async()=>{
 for(const [part,value] of [
  ['xl/vbaProject.bin','dummy'],
  ['../outside.xml','<x/>'],
  ['xl/worksheets/_rels/sheet1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" TargetMode="External" Target="https://example.com/a.png"/></Relationships>'],
 ]){const e=base();e[part]=strToU8(value);await check(e,false);}
 const e=base();e['xl/workbook.xml']=strToU8(strFromU8(e['xl/workbook.xml']).replace('</workbook>','<definedNames><definedName name="Auto_Open" xlm="1">1</definedName></definedNames></workbook>'));await check(e,false);
});
