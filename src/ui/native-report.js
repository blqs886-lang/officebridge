import { escapeHtml as esc, downloadBlob } from './reports.js';
import { NATIVE_REPORT_LIMIT, readNativeReport } from './native-report-data.js';

const releaseBase = 'https://github.com/blqs886-lang/officebridge/releases/download/v1.1.0/';
export const nativeSection = `<details class="verification native-check" id="native-check">
 <summary>用本机 Office 一键核对</summary><div class="proof-body">
 <p>助手会打开所选文件，自动导出 PDF 并显示实际分页和排版。请在接收文件的电脑运行，需要已安装并激活的桌面 Excel 或 PowerPoint。</p>
 <div class="native-downloads">
  <a class="button secondary compact" href="${releaseBase}OfficeBridge-v1.1.0-Windows.zip" target="_blank" rel="noopener noreferrer">下载 Windows 助手</a>
  <a class="button secondary compact" href="${releaseBase}OfficeBridge-v1.1.0-Mac.zip" target="_blank" rel="noopener noreferrer">下载 Mac 助手</a>
 </div><p>下载后完整解压：Windows 双击 OfficeBridge-Windows.cmd，Mac 双击 OfficeBridge-Mac.command。选择原件，按需加入副本或制作端的 PDF 参照。核对结束后会自动打开结果。</p>
 <details><summary>已有核对记录？重新打开</summary>
 <label for="native-report-input">选择助手生成的 .officebridge-check.json</label>
 <input id="native-report-input" type="file" accept=".json,application/json">
 </details><p id="native-import-status" role="status"></p><div id="native-results" tabindex="-1"></div></div></details>`;

export function bindNativeReport(comparison) {
  const root = document.querySelector('#native-check');
  const input = document.querySelector('#native-report-input');
  const status = document.querySelector('#native-import-status');
  const output = document.querySelector('#native-results');
  let requestId = 0;
  function showReport(report) {
    const target = report.platform === 'macos' ? 'Mac' : 'Windows';
    const allOpened = report.documents.every(item => item.openedSuccessfully && item.sourceUnchanged);
    const sources = report.documents.filter(item => item.pdf);
    const byRole = role => report.documents.find(item => item.role === role);
    const pairOptions = [];
    const original = byRole('original'), copy = byRole('copy');
    if (report.referencePdf) {
      for (const item of sources) pairOptions.push({ label: `参照 PDF 与${item.role === 'copy' ? '副本' : '原件'}`,
        left: report.referencePdf, right: item.pdf, leftName: '制作端参照', rightName: `${target} ${item.role === 'copy' ? '副本' : '原件'}` });
    }
    if (original?.pdf && copy?.pdf) pairOptions.push({ label: '本机原件与副本',
      left: original.pdf, right: copy.pdf, leftName: `${target} 原件`, rightName: `${target} 副本` });
    const overview = allOpened ? `${target} Office 已打开所选文件` : `${target} Office 核对有未完成项目`;
    output.innerHTML = `<div class="native-result-heading"><h3>${esc(overview)}</h3><p>记录时间 ${esc(report.createdAt || '未记录')}。以下结果对应助手当时读取的文件版本。</p></div>
      <div class="native-document-list">${report.documents.map(item => `<article><strong>${item.role === 'copy' ? '副本' : '原件'} · ${esc(item.fileName)}</strong>
       <p>${item.openedSuccessfully ? '已成功打开' : '未能打开'}${item.pdf ? ' · 已生成 PDF 预览' : ' · 未生成 PDF 预览'}${item.sourceUnchanged ? ' · 源文件指纹一致' : ' · 源文件状态需核对'}</p>
       ${item.error ? `<p class="font-settings-error">${esc(item.error)}</p>` : ''}
       ${item.pdfError ? `<p class="font-settings-error">PDF 预览未完成：${esc(item.pdfError)}</p>` : ''}
       ${item.cleanupNote ? `<p>${esc(item.cleanupNote)}</p>` : ''}
       <details><summary>查看本次软件与文件信息</summary><p>${esc(item.officeApplication)} ${esc(item.officeVersion)}${item.officeBuild ? `（${esc(item.officeBuild)}）` : ''}</p><p class="hash">${esc(item.sha256 || '未能读取文件指纹')}</p></details></article>`).join('')}</div>
      ${report.comparisons.length ? `<ul>${report.comparisons.map(text => `<li>${esc(text)}</li>`).join('')}</ul>` : ''}
      ${report.fontFamilies ? `<details class="native-fonts"><summary>使用这台 Office 的字体名单</summary><p>记录包含 ${report.fontFamilies.length} 个字体名称。可以下载后，在文档的“更多字体设置”中导入，减少对已可用字体的替换。名称名单不保证字形和排版一致。</p><button type="button" class="button secondary compact" id="native-font-download">下载这台 Office 的字体名单</button></details>` : ''}
      ${report.fontListNote ? `<p>${esc(report.fontListNote)}</p>` : ''}
      ${pairOptions.length ? `<label class="native-pair-label" for="native-pair">查看哪两份排版</label><select id="native-pair">${pairOptions.map((item, index) => `<option value="${index}">${esc(item.label)}</option>`).join('')}</select>` : `<p>${sources.length ? '已有一份本机 PDF。可以在下方添加制作端的参照继续比较。' : '请按上面的原因处理后，再用助手核对。'}</p>`}
      <p class="proof-tip">PDF 用于核对静态排版。需要使用的动画、音视频与公式仍请在 Office 中实际操作。</p>`;
    output.querySelector('#native-font-download')?.addEventListener('click', () => {
      downloadBlob('\uFEFF' + report.fontFamilies.join('\r\n') + '\r\n', `OfficeBridge-${target}-Office-fonts.txt`, 'text/plain;charset=utf-8');
    });
    function displayPair(index) {
      const pair = pairOptions[index];
      comparison.setFiles([new File([pair.left.bytes], pair.left.fileName, {type:'application/pdf'}),
        new File([pair.right.bytes], pair.right.fileName, {type:'application/pdf'})], [pair.leftName, pair.rightName]);
    }
    if (pairOptions.length) {
      output.querySelector('#native-pair').addEventListener('change', event => displayPair(Number(event.target.value)));
      displayPair(0);
    } else if (sources.length) {
      comparison.setFiles([null, new File([sources[0].pdf.bytes], sources[0].pdf.fileName, {type:'application/pdf'})], ['制作端参照', target + (sources[0].role === 'copy' ? ' 副本' : ' 原件')]);
    }
    root.open = true;
    output.focus({preventScroll:true});
    output.scrollIntoView({block:'start'});
  }
  async function read(textOrPromise) {
    const id = ++requestId;
    status.textContent = '正在本机读取核对记录…';
    try {
      const text = await textOrPromise;
      if (id !== requestId) return;
      const report = await readNativeReport(text);
      if (id !== requestId) return;
      showReport(report);
      status.textContent = '核对记录已打开。预览也在本机处理。';
    } catch (error) {
      if (id === requestId) status.textContent = error.message + (output.childElementCount ? ' 下方保留的是上一份已打开的记录。' : '');
    }
  }
  input.addEventListener('change', async () => {
    const file = input.files[0];
    if (!file) return;
    if (file.size > NATIVE_REPORT_LIMIT) { requestId++; status.textContent = '核对记录超过 110 MB，请用助手重新导出。' + (output.childElementCount ? ' 下方保留的是上一份已打开的记录。' : ''); return; }
    await read(file.text());
    input.value = '';
  });
  const embedded = document.querySelector('#officebridge-native-data');
  if (embedded) {
    const text = embedded.textContent;
    embedded.remove();
    void read(text);
  }
}
