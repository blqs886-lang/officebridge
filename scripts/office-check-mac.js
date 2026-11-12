/* OfficeBridge local desktop Office checker. Uses only macOS built-in frameworks/tools.
 * Originals are never opened by Office: each run owns a uniquely named input copy.
 * No network, application-wide preferences, printer settings or existing documents are changed.
 */
ObjC.import('Foundation');
ObjC.import('AppKit');
ObjC.import('PDFKit');
ObjC.import('CoreText');
const fm = $.NSFileManager.defaultManager;
const sa = Application.currentApplication();
sa.includeStandardAdditions = true;
const MiB = 1024 * 1024;
function unwrap(s) { return ObjC.unwrap(s); }
function basename(p) { return unwrap($(p).lastPathComponent); }
function exists(p) { return !!fm.fileExistsAtPath($(p)); }
function absolute(p) { const s=unwrap($(p).stringByStandardizingPath);return s[0]==='/'?s:unwrap($(unwrap(fm.currentDirectoryPath)+'/'+s).stringByStandardizingPath); }
function bytes(p) {
  const attr = fm.attributesOfItemAtPathError($(p), null);
  if (!attr || unwrap(attr.objectForKey('NSFileType')) !== 'NSFileTypeRegular') throw Error('请选择本机的普通文件。');
  const n = Number(unwrap(attr.objectForKey('NSFileSize')));
  if (!Number.isSafeInteger(n) || n > 100 * MiB) throw Error('文件超过 100 MiB，请分成较小的文件后核对。');
  const data = $.NSData.dataWithContentsOfFile($(p));
  if (!data) throw Error('无法读取文件，请先下载到本机再试。');
  return data;
}
function textOf(data) { const s = $.NSString.alloc.initWithDataEncoding(data, $.NSUTF8StringEncoding); if (!s) throw Error('文件不是有效 UTF-8 文本。'); return unwrap(s); }
function taskData(command, args, limit) {
  const t = $.NSTask.alloc.init;
  if(command==='/usr/bin/unzip') { args=['-c','ulimit -t 15; exec /usr/bin/unzip "$@"','officebridge-unzip'].concat(args);command='/bin/sh'; }
  t.launchPath = command; t.arguments = $(args);
  const pipe = $.NSPipe.pipe;t.standardOutput=pipe;t.standardError=pipe;
  t.launch;pipe.fileHandleForWriting.closeFile;
  const result=$.NSMutableData.data;
  while(true){const chunk=pipe.fileHandleForReading.readDataOfLength(65536);if(Number(chunk.length)===0)break;result.appendData(chunk);if(result.length>(limit||14*MiB)){t.terminate;t.waitUntilExit;throw Error('处理内容超过本机核对限制。');}}
  t.waitUntilExit;
  if(t.terminationStatus!==0)throw Error(textOf(result).trim().slice(0,1200)||'系统操作未完成。');
  return result;
}
function task(command,args) {return textOf(taskData(command,args));}
// Native Office can display a permission or licensing dialog. Bound only our worker,
// never terminate Excel/PowerPoint (which may contain the user's other documents).
function officeTask(args) {
  const log=unwrap($.NSTemporaryDirectory())+'officebridge-worker-'+unwrap($.NSUUID.UUID.UUIDString)+'.txt';
  if(!fm.createFileAtPathContentsAttributes($(log),$.NSData.data,$()))throw Error('不能创建本次核对日志。');
  const h=$.NSFileHandle.fileHandleForWritingAtPath($(log));
  const t=$.NSTask.alloc.init;t.launchPath='/usr/bin/osascript';t.arguments=$(args);t.standardOutput=h;t.standardError=h;
  let timedOut=false;
  try{
    t.launch;const start=Date.now();
    while(t.running){
      if(Date.now()-start>180000){timedOut=true;t.terminate;break;}
      $.NSThread.sleepForTimeInterval(0.2);
    }
    if(timedOut){for(let i=0;i<20&&t.running;i++)$.NSThread.sleepForTimeInterval(0.1);if(t.running)task('/bin/kill',['-KILL',String(t.processIdentifier)]);}
    t.waitUntilExit;h.closeFile;
    if(timedOut)throw Error('Office 等待超过 3 分钟。请完成 Office 的登录或本次文件夹访问提示，再重新核对；本次副本保留在结果文件夹。如仍在 Office 打开，请仅关闭名称以 OfficeBridge- 开头的本次副本，不要保存。');
    const data=bytes(log);if(Number(data.length)>MiB)throw Error('Office 返回的记录过大。');
    const output=textOf(data);if(t.terminationStatus!==0)throw Error(output.trim().slice(0,1000)||'Office 核对未完成。');
    return output;
  }finally{try{h.closeFile;}catch(_){}fm.removeItemAtPathError($(log),null);}
}
function sha(p) { return task('/usr/bin/shasum', ['-a', '256', p]).trim().split(/\s+/)[0].toLowerCase(); }
function newDir(p) {
  if (exists(p)) throw Error('结果文件夹已存在。请另选一个新名称，防止覆盖已有结果。');
  const err = Ref();
  if (!fm.createDirectoryAtPathWithIntermediateDirectoriesAttributesError($(p), false, $(), err)) throw Error('不能创建结果文件夹：' + unwrap(err[0].localizedDescription));
}
function writeNew(p, s) {
  if (exists(p)) throw Error('拒绝覆盖已有文件：' + basename(p));
  if (!$(s).writeToFileAtomicallyEncodingError($(p), true, $.NSUTF8StringEncoding, null)) throw Error('结果未能保存。');
}
function copyNew(source, dest) {
  if (!fm.copyItemAtPathToPathError($(source), $(dest), null)) throw Error('不能创建本次核对副本。请检查磁盘剩余空间和文件权限。');
}
function decodeBase64(s) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  const out = new Uint8Array(Math.floor(s.length * 3 / 4) - (s.endsWith('==') ? 2 : s.endsWith('=') ? 1 : 0));
  let buffer = 0, bits = 0, n = 0;
  for (let i=0; i<s.length; i++) { const v=chars.indexOf(s[i]); if(v<0)continue; buffer=(buffer<<6)|v; bits+=6; if(bits>=8){bits-=8;out[n++]=(buffer>>bits)&255;} }
  return out;
}
// Validate the central directory before asking unzip to decompress anything.
function packageEntries(path) {
  const b=decodeBase64(unwrap(bytes(path).base64EncodedStringWithOptions(0)));
  function u16(p){if(p<0||p+2>b.length)throw Error('Office 文件结构不完整。');return b[p]|b[p+1]<<8;}
  function u32(p){return (u16(p)+(u16(p+2)*65536))>>>0;}
  let end=-1;
  for(let p=b.length-22;p>=Math.max(0,b.length-65557);p--) if(u32(p)===0x06054b50 && p+22+u16(p+20)===b.length){end=p;break;}
  if(end<0||u16(end+4)||u16(end+6)||u16(end+8)!==u16(end+10))throw Error('只支持完整、未加密的 PPTX 和 XLSX 文件。');
  const count=u16(end+10), size=u32(end+12), start=u32(end+16);
  if(!count||count>5000||start+size!==end)throw Error('文件结构过大或不受支持。');
  let pos=start,total=0,xmlTotal=0;
  const entries=[],seen={};
  for(let i=0;i<count;i++){
    if(u32(pos)!==0x02014b50)throw Error('Office 文件目录损坏。');
    if(((u32(pos+38)>>>16)&0xF000)===0xA000)throw Error('Office 包中含符号链接，未自动打开。');
    const flags=u16(pos+8), method=u16(pos+10), packed=u32(pos+20),unpacked=u32(pos+24),len=u16(pos+28),extra=u16(pos+30),comment=u16(pos+32),local=u32(pos+42);
    if(flags&1||![0,8].includes(method)||u16(pos+34)||packed===0xffffffff||unpacked===0xffffffff)throw Error('不支持加密、分卷或特殊压缩的 Office 文件。');
    if(pos+46+len+extra+comment>end||local+30>=start||u32(local)!==0x04034b50)throw Error('Office 文件目录不完整。');
    const raw=[];for(let j=0;j<len;j++)raw.push(b[pos+46+j]);
    let name;try{name=decodeURIComponent(raw.map(x=>'%'+x.toString(16).padStart(2,'0')).join(''));}catch(_){throw Error('文件内部名称编码无效。');}
    if(!name||/[\\\x00-\x1f\x7f]/.test(name)||name[0]==='/'||/^[a-z]:/i.test(name)||name.split('/').some(x=>x==='..'||x==='.')||seen[name.normalize('NFC').toLowerCase()])throw Error('文件内部路径不安全或重复。');
    seen[name.normalize('NFC').toLowerCase()]=true;
    if(u16(local+6)!==flags||u16(local+8)!==method||u16(local+26)!==len)throw Error('Office 文件目录不一致。');
    for(let j=0;j<len;j++)if(b[local+30+j]!==raw[j])throw Error('Office 文件内部名称不一致。');
    const dataStart=local+30+len+u16(local+28);
    if(dataStart+packed>start||unpacked>200*MiB||(unpacked>MiB&&unpacked/Math.max(1,packed)>300))throw Error('文件展开后过大，请分成较小的文件再核对。');
    total+=unpacked;
    if(/\.(xml|rels)$/i.test(name)){xmlTotal+=unpacked;if(unpacked>12*MiB)throw Error('文件内部 XML 过大。');}
    if(total>200*MiB||xmlTotal>60*MiB)throw Error('文件展开后过大。');
    if(/(^|\/)(_xmlsignatures|activeX|macrosheets|dialogsheets|externalLinks|queryTables|webextensions|richData|model)(\/|$)|vbaProject|connections\.xml/i.test(name))throw Error('这份文件包含宏、签名、嵌入对象或数据连接，未自动打开。请在 Office 里手动核对。');
    entries.push(name);pos+=46+len+extra+comment;
  }
  if(pos!==end)throw Error('Office 文件目录长度不一致。');
  task('/usr/bin/unzip',['-t',path]);
  return entries;
}
function xmlPart(path,name) {
  const glob=name.replace(/[\[\]*?\\]/g,'\\$&');
  const s=task('/usr/bin/unzip',['-p',path,glob]);
  if(/<!DOCTYPE|<!ENTITY/i.test(s))throw Error('文件含不支持的 XML 实体。');
  const err=Ref(); const doc=$.NSXMLDocument.alloc.initWithXMLStringOptionsError($(s), $.NSXMLNodeLoadExternalEntitiesNever, err);
  if(!Number.isFinite(Number(doc.childCount))||Number(doc.childCount)<1||!unwrap(doc.rootElement.name))throw Error('文件内部 XML 损坏：'+name);
  return {text:s,doc:doc};
}
function nodes(doc,xpath) { const err=Ref();const a=doc.nodesForXPathError($(xpath),err);if(!Number.isFinite(Number(a.count)))throw Error('文件 XML 无法解析。');return a; }
function attr(node,key) {const a=node.attributeForName($(key));return a?(unwrap(a.stringValue)||''):'';}
function validate(path,depth) {
  depth=depth||0;
  const format=basename(path).split('.').pop().toLowerCase();
  if(!['pptx','xlsx'].includes(format))throw Error('请选择 .pptx 演示文稿或 .xlsx 表格。');
  const entries=packageEntries(path);
  if(!entries.includes('[Content_Types].xml')||!entries.includes(format==='pptx'?'ppt/presentation.xml':'xl/workbook.xml'))throw Error('扩展名与文件内容不一致。');
  for(const member of entries.filter(x=>/(^|\/)embeddings\//i.test(x))){
    if(depth || format!=='pptx' || !/^ppt\/embeddings\/[^/]+\.xlsx$/i.test(member))throw Error('文件含嵌入对象，请在 Office 里手动核对。');
    const temp=unwrap($.NSTemporaryDirectory())+'officebridge-chart-'+unwrap($.NSUUID.UUID.UUIDString);
    newDir(temp);const nested=temp+'/chart.xlsx';
    try{const data=taskData('/usr/bin/unzip',['-p',path,member.replace(/[\[\]*?\\]/g,'\\$&')],12*MiB);if(!data.writeToFileOptionsError($(nested),$.NSDataWritingWithoutOverwriting,null))throw Error('无法检查内嵌图表。');validate(nested,depth+1);}
    finally{if(exists(nested))fm.removeItemAtPathError($(nested),null);fm.removeItemAtPathError($(temp),null);}
  }
  const contentTypes=xmlPart(path,'[Content_Types].xml');const types=contentTypes.text;
  const overrides=nodes(contentTypes.doc,'//*[local-name()="Override"]');let main=false;
  const part=format==='xlsx'?'/xl/workbook.xml':'/ppt/presentation.xml',type=format==='xlsx'?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml':'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
  for(let i=0;i<overrides.count;i++){const n=overrides.objectAtIndex(i);if(attr(n,'PartName')===part&&attr(n,'ContentType')===type)main=true;}if(!main)throw Error('文件不是受支持的无宏 PPTX / XLSX 格式。');
  if(/macroEnabled|vbaProject|digital-signature|oleObject|activeX|webextension|richvalue|richdata|model\+data/i.test(types))throw Error('这份文件含宏、签名或嵌入对象，请在 Office 里手动核对。');
  let formulas=0,animations=0;const requestedFonts=new Set();
  for(const name of entries.filter(x=>/\.(xml|rels)$/i.test(x))){
    const parsed=xmlPart(path,name);
    const fontNodes=nodes(parsed.doc,'//*[local-name()="latin" or local-name()="ea" or local-name()="cs" or local-name()="rFont"] | //*[local-name()="font"]/*[local-name()="name"]');
    for(let i=0;i<fontNodes.count;i++){const n=fontNodes.objectAtIndex(i),f=attr(n,'typeface')||attr(n,'val');if(f&&f[0]!=="+"&&f.length<=128)requestedFonts.add(f);}
    if(Number(nodes(parsed.doc,'//*[local-name()="connection" or local-name()="queryTable" or local-name()="oleObject" or local-name()="control" or local-name()="externalBook" or local-name()="ddeLink" or local-name()="oleLink"]').count)>0)throw Error('文件含活动内容或外部连接，请在 Office 里手动核对。');
    if(name.endsWith('.rels')){
      const rels=nodes(parsed.doc,'//*[local-name()="Relationship"]');
      for(let i=0;i<rels.count;i++){const r=rels.objectAtIndex(i);if(/vbaProject|xl(intl)?Macrosheet|oleObject|activeX|connections|externalLink|webextension|richdata|richvalue/i.test(attr(r,'Type')))throw Error('文件含活动内容或外部关系，请手动核对。');if(attr(r,'TargetMode').toLowerCase()==='external'&&attr(r,'Type')!=='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink')throw Error('文件引用了外部数据或媒体，未自动打开。普通网页超链接仍可核对。');}
    }
    const names=nodes(parsed.doc,'//*[local-name()="definedName"]');
    for(let i=0;i<names.count;i++)for(const flag of ['xlm','function','vbProcedure'])if(/^(1|true)$/i.test(attr(names.objectAtIndex(i),flag)))throw Error('名称含宏或自定义函数标记，请在 Office 里手动核对。');
    const allFormulas=nodes(parsed.doc,'//*[local-name()="f" or local-name()="formula" or local-name()="formula1" or local-name()="formula2" or local-name()="definedName" or local-name()="calculatedColumnFormula" or local-name()="totalsRowFormula"]');
    if(/^xl\/worksheets\/.*\.xml$/.test(name))formulas+=Number(nodes(parsed.doc,'//*[local-name()="f"]').count);
    for(let i=0;i<allFormulas.count;i++){
      const value=unwrap(allFormulas.objectAtIndex(i).stringValue)||'';
      const f=value.replace(/"(?:[^"]|"")*"/g,'""');
      if(/(?:^|[^A-Z0-9_.])(?:_xlfn\.)?(WEBSERVICE|FILTERXML|IMAGE|STOCKHISTORY|RTD|DDE|PY|CUBEVALUE|CUBEMEMBER|CUBESET|CUBESETCOUNT|CUBERANKEDMEMBER|CUBEMEMBERPROPERTY|CUBEKPIMEMBER|CALL|REGISTER|EXEC|EVALUATE)\s*\(|_xll\./i.test(f)||/\|[^!]*!|\[[^\]]+\][^!]*!/i.test(f))throw Error('公式可能调用在线服务、外部工作簿或活动功能，请在 Office 里手动核对。');
    }
    if(/^ppt\/slides\/slide\d+\.xml$/.test(name))animations+=Number(nodes(parsed.doc,'//*[local-name()="timing"]').count);
  }
  return {format:format,formulaCount:formulas,slidesWithAnimation:animations,requestedFonts:Array.from(requestedFonts)};
}
function fontInventory(format) {
  const fonts=new Set(ObjC.deepUnwrap($.NSFontManager.sharedFontManager.availableFontFamilies));
  const id=format==='xlsx'?'com.microsoft.Excel':'com.microsoft.Powerpoint';
  const appURL=$.NSWorkspace.sharedWorkspace.URLForApplicationWithBundleIdentifier($(id));
  const appPath=unwrap(appURL.path);
  if(appPath){
    const folder=appPath+'/Contents/Resources/DFonts';
    if(exists(folder)){
      const names=ObjC.deepUnwrap(fm.contentsOfDirectoryAtPathError($(folder),null))||[];
      for(const file of names.filter(n=>/\.(ttf|ttc|otf)$/i.test(n)).slice(0,4096))try{
        const descriptors=ObjC.castRefToObject($.CTFontManagerCreateFontDescriptorsFromURL($.NSURL.fileURLWithPath($(folder+'/'+file))));
        for(let i=0;i<Number(descriptors.count);i++){
          const family=unwrap(descriptors.objectAtIndex(i).objectForKey('NSFontFamilyAttribute'));
          if(typeof family==='string'&&family.length>0&&family.length<=128)fonts.add(family);
        }
      }catch(_){}
    }
  }
  return Array.from(fonts).sort().slice(0,4096);
}
function pdfPayload(path,name) {
  const data=bytes(path);if(data.length>25*MiB)throw Error('预览 PDF 超过 25 MiB，请分段核对。');
  if(Number(data.length)<8||textOf(data.subdataWithRange($.NSMakeRange(0,5)))!=='%PDF-')throw Error('Office 未生成真正的 PDF 预览，请在 Office 里导出 PDF 后手动导入。');
  const doc=$.PDFDocument.alloc.initWithData(data);
  if(!Number.isFinite(Number(doc.pageCount))||doc.isLocked||Number(doc.pageCount)<1)throw Error('未得到有效的 PDF 预览。');
  return {fileName:name,sha256:sha(path),base64:unwrap(data.base64EncodedStringWithOptions(0)),pageCount:Number(doc.pageCount)};
}
function observedFontFamilies(path,requested) {
  const raw=unwrap($.NSString.alloc.initWithDataEncoding(bytes(path),$.NSISOLatin1StringEncoding));
  const normalize=x=>x.toLowerCase().replace(/[\s_-]/g,'');
  const names=new Set();
  const regex=/\/BaseFont\s*\/([^\s<>()\[\]{}\/]+)/g;let m;
  while((m=regex.exec(raw))!==null){let n=m[1].replace(/#([0-9a-f]{2})/gi,(_,v)=>String.fromCharCode(parseInt(v,16))).replace(/^[A-Z]{6}\+/,'');names.add(normalize(n));}
  return requested.filter(f=>names.has(normalize(f)));
}
function pick(title,type) {return String(sa.chooseFile({withPrompt:title,ofType:type}));}
function choice(message,yes,no){return sa.displayDialog(message,{withTitle:'OfficeBridge · 本机 Office 核对',buttons:[no,yes],defaultButton:yes}).buttonReturned===yes;}
function parseArgs(args){const o={};for(let i=0;i<args.length;i++){const a=args[i];if(['--no-open','--validate-only','--fonts-only'].includes(a))o[a.slice(2)]=true;else if(['--original','--copy','--reference','--out','--bundle'].includes(a)&&args[i+1])o[a.slice(2)]=args[++i];else throw Error('未知参数：'+a);}return o;}
function run(args) {
  const options=parseArgs(args),interactive=!options.original;
  try{
    if(interactive){
      options.original=pick('第 1 步：选择要在这台 Mac 核对的 PPTX 或 XLSX 原件',['pptx','xlsx']);
      if(choice('有 OfficeBridge 生成的修复副本吗？添加后，可以并排核对修复前后。','添加副本','只有原件')) options.copy=pick('选择修复后的副本',['pptx','xlsx']);
      if(choice('如果要核对 Windows → Mac，请添加在 Windows 正确显示时导出的 PDF。没有也可以继续。','添加原稿 PDF','暂时没有'))options.reference=pick('选择原来那台电脑导出的参照 PDF',['pdf']);
    }
    for(const k of ['original','copy','reference'])if(options[k])options[k]=absolute(options[k]);
    if(options['validate-only'])return JSON.stringify(validate(options.original));
    if(options['fonts-only'])return JSON.stringify(fontInventory(basename(options.original).split('.').pop().toLowerCase()));
    const orig=validate(options.original),cop=options.copy?validate(options.copy):null;
    if(!unwrap($.NSWorkspace.sharedWorkspace.URLForApplicationWithBundleIdentifier($(orig.format==='xlsx'?'com.microsoft.Excel':'com.microsoft.Powerpoint')).path))throw Error('请先安装并激活桌面版 Microsoft '+(orig.format==='xlsx'?'Excel':'PowerPoint')+'，然后重新核对。');
    if(cop&&cop.format!==orig.format)throw Error('原件和修复副本必须是同一种文件格式。');
    const reference=options.reference?pdfPayload(options.reference,basename(options.reference)):null;
    const uuid=unwrap($.NSUUID.UUID.UUIDString).toLowerCase();
    const folder=options.out?absolute(options.out):unwrap($.NSHomeDirectory())+'/Downloads/OfficeBridge-核对-'+uuid.slice(0,8);
    newDir(folder);
    const bundle=options.bundle||unwrap($.NSProcessInfo.processInfo.arguments.lastObject);
    const report={schema:'officebridge-native-check',schemaVersion:1,toolVersion:'1.1.0',createdAt:new Date().toISOString(),environment:{platform:'macos',os:unwrap($.NSProcessInfo.processInfo.operatingSystemVersionString)},documents:[],comparisons:[]};
    try{report.environment.fontFamilies=fontInventory(orig.format);report.environment.fontInventoryNote='包含 macOS 系统字体及本次 Office 应用随附字体的名称；不包含尚未下载的云字体，仍以实际预览为准。';}catch(_){report.environment.fontInventoryNote='字体名单读取未完成，请以实际预览为准。';}
    if(reference)report.referencePdf=reference;
    if(interactive)sa.displayDialog('即将用本机 Office 打开本次核对副本并导出预览。\n\n如果 Office 提示“授予文件访问权限”，请选择本次结果文件夹：\n'+basename(folder)+'\n\n结果保存在“下载”文件夹；原件不会被修改。',{withTitle:'开始核对',buttons:['开始核对'],defaultButton:1});
    const choices=[{role:'original',path:options.original,stats:orig}];if(options.copy)choices.push({role:'copy',path:options.copy,stats:cop});
    let officeBlocked=false;
    for(const c of choices){
      const hash=sha(c.path),name='OfficeBridge-'+uuid.slice(0,8)+'-'+c.role+'.'+c.stats.format;
      const isolated=folder+'/'+name,pdf=folder+'/'+name.replace(/\.[^.]+$/,'.pdf');
      const item={role:c.role,fileName:basename(c.path),sha256:hash,format:c.stats.format,openedSuccessfully:false,sourceUnchanged:false,pdfExported:false,officeApplication:c.stats.format==='xlsx'?'Microsoft Excel':'Microsoft PowerPoint',formulaCount:c.stats.formulaCount,slidesWithAnimation:c.stats.slidesWithAnimation};
      try{
        if(officeBlocked)throw Error('前一份文件未完成，请先处理 Office 窗口中的提示，再重新核对。');
        copyNew(c.path,isolated);
        if(sha(isolated)!==hash)throw Error('读取期间文件发生变化，请保存文件后重新核对。');
        validate(isolated);
        const workerArgs=c.stats.format==='pptx'?[bundle+'/scripts/office-check-mac-ppt.applescript',isolated,pdf,name]:['-l','JavaScript',bundle+'/scripts/office-check-mac-office.js',isolated,pdf,c.stats.format];
        const result=JSON.parse(officeTask(workerArgs));
        Object.keys(result).forEach(k=>item[k]=result[k]);
        if(result.error||result.cleanupNote)officeBlocked=true;
        if(result.pdfExported)try{item.pdf=pdfPayload(pdf,basename(c.path).replace(/\.[^.]+$/,'.pdf'));item.renderedFontFamilies=observedFontFamilies(pdf,c.stats.requestedFonts);}catch(e){item.pdfExported=false;item.pdfError=String(e.message||e);}
      }catch(e){officeBlocked=true;item.error=String(e.message||e).replace(folder,'本次结果文件夹').slice(0,1200);}
      finally{try{item.sourceUnchanged=sha(c.path)===hash;}catch(_){item.sourceUnchanged=false;}}
      report.documents.push(item);
    }
    if(report.documents.length===2&&report.documents.every(d=>d.pdf))report.comparisons.push({kind:'page-count',original:report.documents[0].pdf.pageCount,copy:report.documents[1].pdf.pageCount,equal:report.documents[0].pdf.pageCount===report.documents[1].pdf.pageCount});
    report.environment.fontFamilies=Array.from(new Set((report.environment.fontFamilies||[]).concat.apply(report.environment.fontFamilies||[],report.documents.map(d=>d.renderedFontFamilies||[])))).sort().slice(0,4096);
    report.environment.fontInventoryNote='包含 macOS 系统字体、本次 Office 应用随附字体，以及实际 PDF 字体名称与文稿请求名称完全匹配的字体；名单可能不含未下载的云字体，仍以实际预览为准。';
    writeNew(folder+'/OfficeBridge-Mac-fonts.txt',report.environment.fontFamilies.join('\n')+'\n');
    const json=JSON.stringify(report,null,2);if($(json).lengthOfBytesUsingEncoding($.NSUTF8StringEncoding)>110*MiB)throw Error('核对记录过大，请分次核对。');
    writeNew(folder+'/OfficeBridge.officebridge-check.json',json);
    const template=exists(bundle+'/OfficeBridge.html')?bundle+'/OfficeBridge.html':bundle+'/release/OfficeBridge.html';
    const safe=json.replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
    let html=exists(template)?textOf(bytes(template)):'';
    if(!/<\/body>/i.test(html))throw Error('缺少完整的 OfficeBridge.html。核对记录与 PDF 已保存在结果文件夹，请用完整离线包导入核对记录。');
    const at=html.toLowerCase().lastIndexOf('</body>');html=html.slice(0,at)+'<script type="application/json" id="officebridge-native-data">'+safe+'</script>'+html.slice(at);
    const result=folder+'/OfficeBridge-result.html';writeNew(result,html);
    if(!options['no-open'])$.NSWorkspace.sharedWorkspace.openURL($.NSURL.fileURLWithPath($(result)));
    return JSON.stringify({report:folder+'/OfficeBridge.officebridge-check.json',result:result,documents:report.documents.map(d=>({fileName:d.fileName,opened:d.openedSuccessfully,pdf:d.pdfExported,error:d.error}))});
  }catch(e){if(e.errorNumber===-128)return '已取消。';if(interactive)sa.displayDialog(String(e.message||e),{withTitle:'OfficeBridge · 未完成核对',buttons:['知道了'],defaultButton:1});throw e;}
}
