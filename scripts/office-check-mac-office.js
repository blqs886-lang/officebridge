/* Excel worker receives only a validated task-owned copy. Do not change app-wide state.
 * PowerPoint uses office-check-mac-ppt.applescript because its JXA PDF enum is unreliable. */
function run(args) {
  const path=args[0],pdf=args[1],kind=args[2],name=path.split('/').pop();
  if(kind!=='xlsx')throw Error('此工作进程只处理 Excel；PowerPoint 使用随包的专用导出脚本。');
  const app=Application('Microsoft Excel');
  const result={openedSuccessfully:false,pdfExported:false,officeVersion:app.version()};
  let doc;
  try{
    const collection=kind==='xlsx'?app.workbooks:app.presentations;
    if(collection.byName(name).exists())throw Error('本次副本已经在 Office 打开，请关闭该副本后重新核对。');
    app.open(Path(path));
    doc=collection.byName(name);
    result.openedSuccessfully=doc.exists();
    if(kind==='xlsx'){
      result.sheetCount=doc.worksheets.length;
      result.sheets=[];
      for(let i=0;i<doc.worksheets.length;i++){
        const s=doc.worksheets[i],item={name:s.name()};
        try{const p=s.pageSetupObject;item.printArea=p.printArea();item.zoom=p.zoom();item.fitToPagesWide=p.fitToPagesWide();item.fitToPagesTall=p.fitToPagesTall();}catch(_){}
        result.sheets.push(item);
      }
      app.saveWorkbookAs(doc,{filename:pdf,fileFormat:'PDF file format',addToMostRecentlyUsedList:false});
    }
    result.pdfExported=true;
  }catch(e){result.error=String(e.message||e).slice(0,1000);}
  finally{if(doc)try{
    const owned=(kind==='xlsx'?app.workbooks():app.presentations()).filter(p=>{try{return p.fullName()===path||p.fullName()===pdf;}catch(_){return false;}});
    if(owned.length===1)app.close(owned[0],{saving:'no'});
    else if(owned.length>1)result.cleanupNote='本次核对副本仍在 Office 打开，可手动关闭而不保存。';
  }catch(e){result.cleanupNote='本次核对副本可能仍在 Office 打开，可手动关闭而不保存。';}}
  return JSON.stringify(result);
}
