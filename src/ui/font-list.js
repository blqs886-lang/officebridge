import { FONT_LIST_MAX_BYTES, normalizeTargetFonts, parseFontList } from '../engine/target-fonts.js';
import { downloadBlob, escapeHtml } from './reports.js';

export function renderFontList(record, busy) {
  const comparison = record.report.inventory.targetFontComparison;
  const canReadLocal = /Windows/i.test(navigator.userAgent) && typeof window.queryLocalFonts === 'function';
  const disabled = busy ? 'disabled' : '';
  return `<section class="target-font-list" aria-label="目标 Windows 字体名单">
    <h4>目标 Windows 的字体名单</h4>
    <p>手边有接收文件的电脑时，可以用它的字体名单核对这个文件。未导入名单时使用常见字体建议。</p>
    <div class="font-list-actions">
      <button type="button" class="button secondary compact" id="choose-font-list" ${disabled}>导入字体名单 (.txt)</button>
      <input type="file" id="font-list-input" accept=".txt,text/plain" aria-label="目标电脑字体名单" hidden>
      ${canReadLocal ? `<button type="button" class="button secondary compact" id="read-local-fonts" ${disabled}>读取这台 Windows 的字体</button>` : ''}
    </div>
    ${comparison ? `<div id="font-comparison"><p><strong>已核对 ${comparison.nameCount} 个目标字体名称。</strong> 文档声明中有 ${comparison.listed.length} 个列入名单，${comparison.unlisted.length} 个未列入，含主题备用字体。</p>
      <details><summary>查看文档字体对照</summary><p>名单有：${escapeHtml(comparison.listed.join('、') || '无')}</p><p>名单未列出：${escapeHtml(comparison.unlisted.join('、') || '无')}</p></details>
      <p>只建议替换为名单中已有的名称。字体别名、嵌入字体、Office 云字体和实际字形仍需在目标电脑核对。</p>
      <div class="font-list-actions"><button type="button" class="text-button" id="save-font-list" ${disabled}>保存名单到 TXT</button><button type="button" class="subtle-button" id="remove-font-list" ${disabled}>移除名单并重新检查</button></div></div>` : ''}
    <details class="font-list-help"><summary>怎样从 Windows 取得名单？</summary>
      <p>在目标 Windows 的 Edge 或 Chrome 打开本工具，点“电子表格”样例，展开字体设置，点“读取这台 Windows 的字体”，允许浏览器访问后点“保存名单到 TXT”。把这个文本文件带回 Mac，点“导入字体名单”。</p>
      <p>也可以在记事本中每行写一个已确认可用的字体家族名称，保存为 UTF-8 文本。手工名单只覆盖你填写的名称。完整系统导出方法见便携包内的使用指南。</p></details>
    </section>`;
}

export function bindFontList(record, panel, recheck) {
  const input = panel.querySelector('#font-list-input');
  panel.querySelector('#choose-font-list').onclick = () => input.click();
  input.onchange = () => {
    const file = input.files[0];
    if (!file) return;
    recheck(async () => {
      if (file.size > FONT_LIST_MAX_BYTES) throw new Error('字体名单上限为 512 KiB。');
      return { targetFonts: parseFontList(new Uint8Array(await file.arrayBuffer())) };
    });
  };
  const read = panel.querySelector('#read-local-fonts');
  if (read) read.onclick = () => recheck(async () => {
    let fonts;
    try { fonts = await window.queryLocalFonts(); }
    catch { throw new Error('浏览器未提供本机字体，请用“导入字体名单”选择文本文件。'); }
    const families = [...new Set((fonts || []).map(font => font.family).filter(Boolean))];
    if (!families.length) throw new Error('浏览器没有提供可读取的字体，请用“导入字体名单”选择文本文件。');
    return { targetFonts: normalizeTargetFonts(families) };
  });
  const remove = panel.querySelector('#remove-font-list');
  if (remove) remove.onclick = () => recheck(() => ({ targetFonts: undefined }));
  const save = panel.querySelector('#save-font-list');
  if (save) save.onclick = () => downloadBlob('\uFEFF' + record.options.targetFonts.join('\r\n') + '\r\n', 'OfficeBridge-Windows-fonts.txt', 'text/plain;charset=utf-8');
}
