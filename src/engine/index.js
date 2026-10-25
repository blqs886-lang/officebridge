import { readArchive, writeArchive } from './archive.js';
import { NS, parseXml, decodeXml, encodeXml, elements, replaceFontAttributes } from './xml.js';
import { scanPptx } from './pptx.js';
import { scanXlsx } from './xlsx.js';
import { normalizeFontOptions } from './fonts.js';

export const VERSION = '0.1.0';
const mainTypes = {
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
};

export async function digest(bytes) {
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(hash), n => n.toString(16).padStart(2, '0')).join('');
}

function packageIssue(code, title, description, part = '', severity = 'error') {
  return { id: `${code}:${part}`, code, severity, title, description, part,
    location: part || '文档包', fixable: false, defaultSelected: false };
}

function contentTypeIdentity(value) {
  // Compare aliases without rewriting either the declaration or the ZIP entry.
  try { value = decodeURIComponent(value); } catch { /* retain malformed escapes for inspection */ }
  return value.replace(/[A-Z]/g, c => c.toLowerCase());
}

function inspectContentTypes(doc) {
  const issues = [], seen = { Default: new Set(), Override: new Set() };
  const nodes = [...elements(doc, NS.ct, 'Default'), ...elements(doc, NS.ct, 'Override')];
  for (const node of nodes) {
    if (node.parentNode !== doc.documentElement) {
      issues.push(packageIssue('CONTENT_TYPE_STRUCTURE', '内容类型声明层级异常',
        'Default 和 Override 必须直接位于 Types 根节点下。此文件仅可检查，请在原软件中修复。', '[Content_Types].xml'));
    }
    const key = node.getAttribute(node.localName === 'Default' ? 'Extension' : 'PartName');
    const type = node.getAttribute('ContentType');
    if (!key?.trim() || !type?.trim()) {
      issues.push(packageIssue('CONTENT_TYPE_INVALID', '内容类型声明缺少必要属性',
        'Default 需要 Extension 和 ContentType，Override 需要 PartName 和 ContentType。无法可靠判断包内声明，禁止修复。', '[Content_Types].xml'));
    }
    if (!key?.trim()) continue;
    const identity = contentTypeIdentity(key);
    if (seen[node.localName].has(identity)) {
      issues.push(packageIssue('CONTENT_TYPE_DUPLICATE', '内容类型声明重复',
        '同一个扩展名或部件名称存在多条声明，包括大小写或 URI 编码别名。即使声明内容相同，也禁止自动修复。', '[Content_Types].xml'));
    }
    seen[node.localName].add(identity);
  }
  return issues;
}

function relationshipSource(part) {
  if (part === '_rels/.rels') return '';
  const m = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(part);
  return m ? `${m[1] || ''}${m[2]}` : null;
}

export function resolveTarget(source, target) {
  const withoutFragment = target.split('#')[0];
  if (!withoutFragment) return source;
  if (/^[a-z][a-z\d+.-]*:/i.test(withoutFragment) || /[\\\x00-\x1f]/.test(withoutFragment)) return null;
  const path = withoutFragment.startsWith('/') ? withoutFragment.slice(1)
    : source.slice(0, source.lastIndexOf('/') + 1) + withoutFragment;
  const stack = [];
  for (const segment of path.split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') { if (!stack.length) return null; stack.pop(); }
    else stack.push(segment);
  }
  return stack.join('/');
}

function resolveExistingTarget(source, target, entries) {
  const resolved = resolveTarget(source, target);
  if (!resolved) return null;
  if (entries.has(resolved)) return resolved;
  let decoded;
  try { decoded = decodeURIComponent(resolved); } catch { return null; }
  // Do not normalize again after decoding: encoded dot segments must not turn
  // into parent traversal. readArchive already rejects such entry names.
  return entries.has(decoded) ? decoded : null;
}

function prepare(bytes, fileName, options) {
  options = normalizeFontOptions(options);
  const format = /\.(pptx|xlsx)$/i.exec(fileName)?.[1]?.toLowerCase();
  if (!format) throw new Error('目前支持 .pptx 和 .xlsx。Keynote / Numbers 请先导出为对应 Office 格式；旧版、宏文档和 CSV 暂不处理。');
  const archive = readArchive(bytes), { entries } = archive;
  const xml = new Map(), texts = new Map(), encodings = new Map();
  for (const [part, data] of entries) {
    if (!/\.(xml|rels)$/i.test(part)) continue;
    const decoded = decodeXml(data, part);
    texts.set(part, decoded.text); encodings.set(part, decoded);
    xml.set(part, parseXml(decoded.text, part));
  }
  if (!xml.has('[Content_Types].xml') || !xml.has('_rels/.rels')) throw new Error('缺少 Office 文档包的必要目录信息。');
  const mainPart = format === 'pptx' ? 'ppt/presentation.xml' : 'xl/workbook.xml';
  const types = xml.get('[Content_Types].xml');
  if (types.documentElement.namespaceURI !== NS.ct || types.documentElement.localName !== 'Types') throw new Error('内容类型目录不是有效的 Office Open XML。');
  const issues = inspectContentTypes(types);
  const mainDeclarations = elements(types, NS.ct, 'Override')
    .filter(n => contentTypeIdentity(n.getAttribute('PartName') || '') === `/${mainPart}`);
  // Ambiguous declarations remain readable when the expected main type is
  // present, but inspectContentTypes blocks every repair regardless of order.
  if (!xml.has(mainPart) || !mainDeclarations.some(n => n.getAttribute('ContentType') === mainTypes[format])) throw new Error('文件扩展名与内部内容不匹配，或这是宏/模板等暂不支持的文档类型。');
  const mainNS = format === 'pptx' ? NS.p : NS.s;
  const root = xml.get(mainPart).documentElement;
  const expectedRoot = format === 'pptx' ? 'presentation' : 'workbook';
  if (root.localName !== expectedRoot) throw new Error('主文档结构无效。');
  if (root.namespaceURI !== mainNS) {
    if (!root.namespaceURI?.startsWith('http://purl.oclc.org/ooxml/')) throw new Error('不支持此文档命名空间。');
    issues.push(packageIssue('STRICT_READONLY', 'Strict 格式仅检查', '此文件使用 Strict OOXML。当前版本仅作只读检查，不生成修复副本。', mainPart));
  }
  for (const name of entries.keys()) {
    if (/^_xmlsignatures\//i.test(name) || /vbaproject|vbasignature|^xl\/(?:intl)?macrosheets\//i.test(name)) {
      issues.push(packageIssue('PROTECTED_PACKAGE', '签名或宏内容需要原软件处理', '修改可能使数字签名失效或影响宏。当前工具保留只读检查，并禁止修复。', name));
    }
  }
  for (const node of [...elements(types, NS.ct, 'Override'), ...elements(types, NS.ct, 'Default')]) {
    if (/digital-signature|vbaProject|macroEnabled|macrosheet/i.test(node.getAttribute('ContentType') || '')) {
      issues.push(packageIssue('PROTECTED_CONTENT_TYPE', '包含受保护的扩展内容', '检测到签名或宏内容类型。为保留完整性，此文件仅可检查。', node.getAttribute('PartName') || ''));
    }
  }
  const relsBySource = new Map();
  let externalRelationshipCount = 0;
  for (const [part, doc] of xml) {
    if (!part.endsWith('.rels')) continue;
    const source = relationshipSource(part);
    if (source === null || (source && !entries.has(source)) || doc.documentElement.namespaceURI !== NS.rel
      || doc.documentElement.localName !== 'Relationships') {
      issues.push(packageIssue('RELATIONSHIP_STRUCTURE', '关系文件结构异常', '关系文件位置或对应的源部件不正确；请在原软件中修复。', part));
      continue;
    }
    const ids = new Map();
    for (const rel of elements(doc, NS.rel, 'Relationship')) {
      if (rel.parentNode !== doc.documentElement) {
        issues.push(packageIssue('RELATIONSHIP_NESTED', '关系文件层级异常', '内部关系必须直接位于关系根节点下，禁止自动修复。', part)); continue;
      }
      const id = rel.getAttribute('Id'), target = rel.getAttribute('Target'), type = rel.getAttribute('Type');
      if (!id || ids.has(id) || !target || !type) {
        issues.push(packageIssue('RELATIONSHIP_INVALID', '关系定义缺失或重复', '文件内部关联无法可靠解析；为避免进一步损坏，禁止修复。', part));
        continue;
      }
      ids.set(id, rel);
      if (/digital-signature|vbaProject|macrosheet/i.test(type)) issues.push(packageIssue('PROTECTED_RELATIONSHIP', '包含签名或宏关系', '签名/宏文件只能检查，不能自动改写。', part));
      if (rel.getAttribute('TargetMode') === 'External') { externalRelationshipCount++; continue; }
      if (rel.hasAttribute('TargetMode') && rel.getAttribute('TargetMode') !== 'Internal') {
        issues.push(packageIssue('RELATIONSHIP_MODE', '关系模式无效', '内部关系使用了无法识别的目标模式。', part)); continue;
      }
      if (!resolveExistingTarget(source, target, entries)) issues.push(packageIssue(`MISSING_PART_${id}`, '内部关联文件缺失', `关系 ${id} 指向不存在或无效的部件：${target}。无法自动重建原内容。`, part));
    }
    relsBySource.set(source, ids);
  }
  const rootMain = [...(relsBySource.get('')?.values() || [])].filter(r => /\/officeDocument$/.test(r.getAttribute('Type') || ''));
  if (rootMain.length !== 1 || rootMain[0].getAttribute('TargetMode') === 'External' || resolveExistingTarget('', rootMain[0].getAttribute('Target'), entries) !== mainPart) {
    issues.push(packageIssue('MAIN_RELATIONSHIP', '主文档关系异常', '文档入口与内部主文件不一致，禁止修复。', '_rels/.rels'));
  }
  for (const [part, doc] of xml) {
    if (part.endsWith('.rels')) continue;
    for (const node of Array.from(doc.getElementsByTagName('*'))) {
      for (const attr of Array.from(node.attributes || [])) {
        if (attr.namespaceURI !== NS.r || !['id', 'embed', 'link'].includes(attr.localName) || !attr.value) continue;
        if (!relsBySource.get(part)?.has(attr.value)) issues.push(packageIssue(`MISSING_REL_${attr.value}`, '引用的内部关系不存在', `部件引用关系 ${attr.value}，但没有对应关系定义。`, part));
      }
    }
  }
  return { ...archive, format, fileName, options, xml, texts, encodings, issues, externalRelationshipCount };
}

async function buildReport(ctx) {
  const scanned = ctx.format === 'pptx' ? scanPptx(ctx) : scanXlsx(ctx);
  const issues = [...ctx.issues, ...scanned.issues];
  // Deduplicate package findings produced by multiple references.
  const unique = [...new Map(issues.map(i => [i.id, i])).values()];
  return {
    schemaVersion: 1, toolVersion: VERSION, fileName: ctx.fileName, format: ctx.format,
    sha256: await digest(ctx.bytes), fileSize: ctx.bytes.length,
    analysisOptions: { fontProfile: ctx.options.fontProfile, fontMappings: { ...ctx.options.fontMappings } },
    issues: unique,
    inventory: { ...scanned.inventory, packageParts: ctx.entries.size, expandedBytes: ctx.expandedBytes,
      externalRelationshipCount: ctx.externalRelationshipCount },
    repairBlocked: ctx.issues.some(i => i.severity === 'error'),
    validation: { packageChecked: true, windowsOfficeTested: false,
      notes: ['已检查压缩包校验、XML 与内部关系；这不是完整 OOXML 架构验证。',
        '未在目标 Windows Office 中打开，也未执行宏、公式、动画或媒体。',
        '字体替换仅修改名称，不安装或嵌入字体；字宽变化可能影响换行，仍需核对。',
        '未发现风险不代表全部兼容；当前规则不覆盖所有 Office 功能。'] },
  };
}

export async function analyzeOffice(bytes, fileName, options = {}) {
  return buildReport(prepare(bytes, fileName, options));
}

export async function repairOffice(bytes, fileName, options = {}) {
  const ctx = prepare(bytes, fileName, options);
  const beforeReport = await buildReport(ctx);
  if (beforeReport.repairBlocked) throw new Error('此文件存在签名、宏、结构异常或不支持的格式，不能自动修复。请查看报告。');
  const selected = new Set(options.selectedIssueIds || []);
  if (!selected.size) throw new Error('请先选择需要采用的字体替换建议。');
  const fixes = [];
  for (const id of selected) {
    const issue = beforeReport.issues.find(i => i.id === id);
    if (!issue?.fixable || issue.fix?.kind !== 'font-replace') throw new Error('选择的修复项目已失效或不支持，请重新检查文件。');
    fixes.push({ ...issue.fix, issueId: issue.id, part: issue.part });
  }
  const entries = new Map(ctx.entries), changes = [];
  for (const part of new Set(fixes.map(f => f.part))) {
    const partFixes = fixes.filter(f => f.part === part);
    const { text, counts } = replaceFontAttributes(ctx.texts.get(part), ctx.xml.get(part), partFixes);
    parseXml(text, part);
    for (const fix of partFixes) {
      const count = counts.get(fix.issueId) || 0;
      if (!count) throw new Error(`未定位到 ${part} 的目标字体；为保护内容，未生成修复副本。`);
      changes.push({ issueId: fix.issueId, part, kind: fix.kind, before: fix.from, after: fix.to, count });
    }
    entries.set(part, encodeXml(text, ctx.encodings.get(part)));
  }
  const output = writeArchive(entries);
  const outputName = fileName.replace(/\.(pptx|xlsx)$/i, `.windows.$1`);
  const report = await analyzeOffice(output, outputName, options);
  if (report.repairBlocked) throw new Error('修复副本未通过内部结构检查，已停止输出。原文件不受影响。');
  const changedParts = new Set(changes.map(c => c.part));
  const verified = readArchive(output);
  if (verified.entries.size !== ctx.entries.size) throw new Error('修复后文件数量不一致，已停止输出。');
  for (const [name, original] of ctx.entries) {
    if (changedParts.has(name)) continue;
    const next = verified.entries.get(name);
    if (!next || next.length !== original.length || next.some((b, i) => b !== original[i])) throw new Error('未修改部件的完整性核对失败，已停止输出。');
  }
  report.repair = { originalFileName: fileName, originalSha256: beforeReport.sha256,
    changes, changedPartCount: changedParts.size, unchangedPartsVerified: true,
    beforeIssueCount: beforeReport.issues.length, afterIssueCount: report.issues.length };
  return { bytes: output, fileName: outputName, report, changes };
}
