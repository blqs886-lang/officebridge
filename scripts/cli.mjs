#!/usr/bin/env node
import { mkdir, open, readdir, lstat, unlink } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { analyzeOffice, repairOffice } from '../src/engine/index.js';
import { LIMITS } from '../src/engine/archive.js';

const HELP = `OfficeBridge：本地批量检查与主动选择的字体替换

用法：
  node scripts/cli.mjs [check] <文件.pptx|文件.xlsx> [...] [--target windows|mac] [--out 目录]
  node scripts/cli.mjs repair <文件.pptx|文件.xlsx> [...] --accept-font-substitutions [--target windows|mac] [--out 目录]

省略命令时只检查。默认目标为 Windows；--target mac 检查发往 Mac 的文稿。
默认输出目录为 ./officebridge-reports。
含空格的路径请加引号；以 - 开头的文件路径请放在 -- 之后。

check 保存 JSON 报告；警告不改变退出状态，结构错误会返回非零状态。
repair 会采用报告中所有可用的字体替换建议，必须明确传入确认参数。
替换可能改变换行、外观及版面。请先 check 并阅读报告，再决定是否采用。
修复只生成 .windows.pptx / .windows.xlsx 或 .mac.pptx / .mac.xlsx 副本及对应 JSON，不覆盖原文件。
签名、宏或结构异常文件仅保存检查报告；没有建议时也只保存报告。
输出文件名保留中文和空格，但会替换 Windows 不允许的文件名字符。
已有同名、大小写或 Unicode 规范化冲突的输出会被保留，该项任务会失败。
单个失败不会阻断后续文件；退出码 0 表示完成，1 表示有文件失败，2 表示用法错误。
全部处理在本机进行，不访问网络；检查结果不等同于 Windows Office 或 Mac Office 实测。
`;

class CliError extends Error {
  constructor(message) { super(message); this.name = 'CliError'; }
}

function parseArguments(args) {
  if (args.includes('--help') || args.includes('-h')) return { help: true };
  const rest = [...args];
  const command = ['check', 'repair'].includes(rest[0]) ? rest.shift() : 'check';
  const files = [];
  let output = 'officebridge-reports';
  let target = 'windows';
  let accepted = false;
  let positionalOnly = false;
  for (let index = 0; index < rest.length; index += 1) {
    const value = rest[index];
    if (!positionalOnly && value === '--') { positionalOnly = true; continue; }
    if (!positionalOnly && value === '--accept-font-substitutions') { accepted = true; continue; }
    if (!positionalOnly && (value === '--target' || value.startsWith('--target='))) {
      target = value === '--target' ? rest[++index] : value.slice('--target='.length);
      if (!['windows', 'mac'].includes(target)) throw new CliError('--target 必须是 windows 或 mac。');
      continue;
    }
    if (!positionalOnly && value === '--out') {
      output = rest[++index];
      if (!output || output.startsWith('--')) throw new CliError('--out 后需要提供输出目录。');
      continue;
    }
    if (!positionalOnly && value.startsWith('--out=')) {
      output = value.slice('--out='.length);
      if (!output) throw new CliError('--out 后需要提供输出目录。');
      continue;
    }
    if (!positionalOnly && value.startsWith('-')) throw new CliError('无法识别参数；请用 --help 查看用法，或用 -- 分隔文件路径。');
    files.push(value);
  }
  if (!files.length) throw new CliError('需要至少提供一个 PPTX 或 XLSX 文件。');
  if (command === 'repair' && !accepted) throw new CliError('repair 必须明确传入 --accept-font-substitutions；请先检查报告并了解字体替换对版面的影响。');
  if (command === 'check' && accepted) throw new CliError('--accept-font-substitutions 只能与 repair 一起使用；check 不修改文稿。');
  return { command, files, output: path.resolve(output), accepted, target };
}

function quoted(value) {
  // File names and file-system messages can contain terminal control characters.
  return JSON.stringify(String(value));
}

function canonicalName(name) {
  return name.normalize('NFC').toLowerCase();
}

function portableName(name) {
  let result = path.basename(name).normalize('NFC')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_').replace(/[ .]+$/g, '');
  if (/^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(result)) result = `_${result}`;
  if (!result || result === '.' || result === '..') throw new CliError('无法生成安全的输出文件名。');
  return result;
}

function errorMessage(error) {
  if (error.code === 'ENOENT') return '找不到输入文件或输出目录。';
  if (error.code === 'EACCES' || error.code === 'EPERM') return '没有读取文件或创建输出的权限。';
  if (error.code === 'EEXIST') return '输出已存在；为保护已有成果，未覆盖。请指定新的 --out 目录。';
  if (error.code === 'ENOTDIR' || error.code === 'EISDIR') return '输入必须是普通文件，--out 必须是目录。';
  if (error.code === 'ENOSPC') return '输出位置空间不足，未完成的本次输出已尽量移除。';
  if (error.code === 'ENAMETOOLONG') return '输出路径或文件名过长，请使用较短的文件名或目录。';
  return error instanceof Error ? error.message : '处理失败。';
}

async function readInput(input) {
  // Nonblocking open lets us reject named pipes/devices after fstat without
  // waiting indefinitely for a pipe writer. It does not affect regular files.
  const handle = await open(input, constants.O_RDONLY | (constants.O_NONBLOCK || 0));
  try {
    const before = await handle.stat();
    if (!before.isFile()) throw new CliError('输入必须是普通文件，不能是目录或设备。');
    if (before.size > LIMITS.fileBytes) throw new CliError('输入超过 100 MB，已在读取完整内容之前停止。');
    // Read at most the initial size plus one byte, even if another process grows
    // the file. This is a bounded snapshot, never an unbounded readFile call.
    const buffer = new Uint8Array(before.size + 1);
    let length = 0;
    while (length < buffer.length) {
      const { bytesRead } = await handle.read(buffer, length, buffer.length - length, length);
      if (!bytesRead) break;
      length += bytesRead;
    }
    const after = await handle.stat();
    if (length !== before.size || after.size !== before.size
      || after.mtimeMs !== before.mtimeMs || after.ctimeMs !== before.ctimeMs) {
      throw new CliError('读取期间输入文件发生变化；请保存并停止写入后重试。');
    }
    return buffer.subarray(0, length);
  } finally { await handle.close(); }
}

async function writeExclusiveGroup(directory, outputs) {
  const existing = new Set((await readdir(directory)).map(canonicalName));
  const planned = new Set();
  for (const { name } of outputs) {
    const key = canonicalName(name);
    if (existing.has(key) || planned.has(key)) throw new CliError('输出名称与已有文件或另一项输出冲突；原成果已保留。请使用新的 --out 目录。');
    planned.add(key);
  }
  const created = [];
  try {
    // Reserve every member before writing either document or report. wx also
    // rejects an existing symlink, and closes the race after the directory scan.
    for (const output of outputs) {
      const target = path.join(directory, output.name);
      const handle = await open(target, 'wx', 0o600);
      const item = { target, handle, identity: null, output };
      created.push(item);
      item.identity = await handle.stat();
    }
    for (const item of created) {
      await item.handle.writeFile(item.output.bytes);
      await item.handle.sync();
    }
    await Promise.all(created.map(item => item.handle.close()));
  } catch (error) {
    await Promise.allSettled(created.map(item => item.handle.close()));
    // Remove only this invocation's incomplete new outputs. If another process
    // replaced a path, its file must be left alone.
    for (const { target, identity } of created) {
      try {
        const current = await lstat(target);
        if (identity && current.isFile() && current.dev === identity.dev && current.ino === identity.ino) await unlink(target);
      } catch { /* Preserve the original failure; never delete uncertain files. */ }
    }
    throw error;
  }
}

function reportOutput(name, report) {
  return { name, bytes: `${JSON.stringify(report, null, 2)}\n` };
}

async function processFile(input, settings) {
  const absolute = path.resolve(input);
  const name = path.basename(absolute);
  if (!/\.(?:pptx|xlsx)$/i.test(name)) throw new CliError('目前仅支持 .pptx 和 .xlsx；文件扩展名大小写均可。');
  const bytes = await readInput(absolute);
  const options = { fontProfile: `${settings.target}-office` };
  const report = await analyzeOffice(bytes, name, options);
  const originalReportName = `${portableName(name)}.report.json`;
  if (settings.command === 'check' || report.repairBlocked) {
    await writeExclusiveGroup(settings.output, [reportOutput(originalReportName, report)]);
    return { status: report.repairBlocked ? 'blocked' : 'checked', report, outputs: [originalReportName] };
  }
  const selectedIssueIds = report.issues.filter(issue => issue.fixable && issue.fix?.kind === 'font-replace').map(issue => issue.id);
  if (!selectedIssueIds.length) {
    await writeExclusiveGroup(settings.output, [reportOutput(originalReportName, report)]);
    return { status: 'unchanged', report, outputs: [originalReportName] };
  }
  const repaired = await repairOffice(bytes, name, { ...options, selectedIssueIds });
  const documentName = portableName(repaired.fileName);
  const reportName = `${documentName}.report.json`;
  const resultReport = { ...repaired.report, fileName: documentName };
  await writeExclusiveGroup(settings.output, [
    { name: documentName, bytes: repaired.bytes },
    reportOutput(reportName, resultReport),
  ]);
  return { status: 'repaired', report: resultReport, outputs: [documentName, reportName] };
}

async function main(args) {
  let settings;
  try {
    const [major, minor] = process.versions.node.split('.').map(Number);
    if (major < 22 || (major === 22 && minor < 13)) throw new CliError('请使用 Node.js 22.13 或更新版本。');
    settings = parseArguments(args);
  } catch (error) {
    process.stderr.write(`参数错误：${quoted(errorMessage(error))}\n使用 --help 查看用法。\n`);
    return 2;
  }
  if (settings.help) { process.stdout.write(HELP); return 0; }
  try { await mkdir(settings.output, { recursive: true }); }
  catch (error) {
    process.stderr.write(`无法使用输出目录：${quoted(errorMessage(error))}\n`);
    return 1;
  }
  let failures = 0;
  let completed = 0;
  for (const file of settings.files) {
    try {
      const result = await processFile(file, settings);
      completed += 1;
      const warningCount = result.report.issues.filter(issue => issue.severity === 'warning').length;
      const status = { checked: '检查完成', repaired: '已生成字体替换副本', unchanged: '没有可应用的字体建议，仅保存报告', blocked: '检查完成，但文件结构或保护状态禁止修复' }[result.status];
      process.stdout.write(`${status}：${quoted(path.basename(file))}；${warningCount} 项警告。\n`);
      for (const output of result.outputs) process.stdout.write(`  输出：${quoted(path.join(settings.output, output))}\n`);
      if (result.status === 'blocked') failures += 1;
    } catch (error) {
      failures += 1;
      process.stderr.write(`处理失败：${quoted(path.basename(file))}；${quoted(errorMessage(error))}\n`);
    }
  }
  process.stdout.write(`已保存 ${completed} 个输入的结果；${failures} 个输入失败或受阻。未进行 ${settings.target === 'mac' ? 'Mac' : 'Windows'} Office 实测。\n`);
  return failures ? 1 : 0;
}

process.exitCode = await main(process.argv.slice(2));
