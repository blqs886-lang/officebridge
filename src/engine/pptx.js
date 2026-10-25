import { NS, attribute } from './xml.js';
import { fontAdvice, isKnownFont, isProtectedFont } from './fonts.js';

const DRAWING_NS = new Set([NS.a, 'http://purl.oclc.org/ooxml/drawingml/main']);
const PRESENTATION_NS = new Set([NS.p, 'http://purl.oclc.org/ooxml/presentationml/main']);
const RELATIONSHIP_NS = new Set([NS.rel, 'http://purl.oclc.org/ooxml/package/relationships']);
const AUDIO_EXTENSIONS = new Set(['mp3', 'wav', 'wma', 'aac', 'm4a', 'aiff', 'aif', 'ogg', 'oga', 'flac', 'mid', 'midi']);
const VIDEO_EXTENSIONS = new Set(['mp4', 'm4v', 'mov', 'qt', 'avi', 'wmv', 'mpeg', 'mpg', 'webm', 'ogv', '3gp']);
const ANIMATION_ELEMENTS = new Set(['anim', 'animClr', 'animEffect', 'animMotion', 'animRot', 'animScale', 'cmd', 'set']);
const PRIVATE_USE = /[\uE000-\uF8FF\u{F0000}-\u{FFFFD}\u{100000}-\u{10FFFD}]/u;
const SPECIAL_GLYPHS = /[\u{1F000}-\u{1FAFF}\uFE0F\u200D]/u;

function descendants(node) {
  return Array.from(node.getElementsByTagName('*'));
}

function children(node) {
  return Array.from(node.childNodes || []).filter(child => child.nodeType === 1);
}

function isDrawing(node, localName) {
  return DRAWING_NS.has(node.namespaceURI) && (!localName || node.localName === localName);
}

function isPresentation(node, localName) {
  return PRESENTATION_NS.has(node.namespaceURI) && (!localName || node.localName === localName);
}

function drawingChild(node, localName) {
  return children(node).find(child => isDrawing(child, localName));
}

function locationFor(part) {
  const slide = /^ppt\/slides\/slide(\d+)\.xml$/.exec(part);
  if (slide) return `幻灯片部件 ${slide[1]}（不一定是播放顺序）`;
  if (part.startsWith('ppt/slideMasters/')) return `幻灯片母版 · ${part.split('/').pop()}`;
  if (part.startsWith('ppt/slideLayouts/')) return `幻灯片版式 · ${part.split('/').pop()}`;
  if (part.startsWith('ppt/theme/')) return `主题 · ${part.split('/').pop()}`;
  if (part.startsWith('ppt/charts/')) return `图表 · ${part.split('/').pop()}`;
  if (part.startsWith('ppt/notes')) return `备注 · ${part.split('/').pop()}`;
  return part;
}

function relationshipSource(part) {
  return part.replace(/\/_rels\/([^/]+)\.rels$/, '/$1');
}

function extension(part) {
  return part.split('.').pop().toLowerCase();
}

function numericAttribute(node, name) {
  if (!node || !node.hasAttribute(name)) return null;
  const value = Number(node.getAttribute(name));
  return Number.isFinite(value) ? value : null;
}

function contentTypes(xml) {
  const doc = xml.get('[Content_Types].xml');
  const defaults = new Map();
  const overrides = new Map();
  if (doc) {
    for (const node of descendants(doc)) {
      if (node.namespaceURI !== NS.ct) continue;
      if (node.localName === 'Default') defaults.set(attribute(node, 'Extension').toLowerCase(), attribute(node, 'ContentType'));
      if (node.localName === 'Override') overrides.set(attribute(node, 'PartName').replace(/^\//, ''), attribute(node, 'ContentType'));
    }
  }
  return part => overrides.get(part) || defaults.get(extension(part)) || '';
}

function mediaKind(part, contentType) {
  if (contentType.startsWith('audio/')) return 'audio';
  if (contentType.startsWith('video/')) return 'video';
  if (contentType.startsWith('image/')) return 'image';
  if (AUDIO_EXTENSIONS.has(extension(part))) return 'audio';
  if (VIDEO_EXTENSIONS.has(extension(part))) return 'video';
  return 'image-or-other';
}

function externalKind(type) {
  const kind = type.split('/').pop();
  if (kind === 'hyperlink') return 'hyperlink';
  if (['audio', 'video', 'media', 'image'].includes(kind)) return kind;
  return 'data-or-object';
}

function hasAncestor(node, predicate) {
  for (let parent = node.parentNode; parent && parent.nodeType === 1; parent = parent.parentNode) {
    if (predicate(parent)) return true;
  }
  return false;
}

// This is a lower-bound estimate from explicit hard lines and font sizes, not a
// renderer. Inherited sizes, wrapping, glyph metrics, kerning and layout-master
// inheritance are deliberately not guessed.
function explicitTextHeight(textBody) {
  let height = 0;
  for (const paragraph of children(textBody).filter(node => isDrawing(node, 'p'))) {
    const runs = children(paragraph).filter(node => isDrawing(node, 'r') || isDrawing(node, 'fld'));
    if (!runs.length) continue;
    const paragraphProperties = drawingChild(paragraph, 'pPr');
    const defaultRun = paragraphProperties && drawingChild(paragraphProperties, 'defRPr');
    const defaultSize = numericAttribute(defaultRun, 'sz');
    const sizes = runs.map(run => numericAttribute(drawingChild(run, 'rPr'), 'sz') ?? defaultSize);
    if (sizes.some(size => size === null || size <= 0)) return null;
    const hardLines = 1 + children(paragraph).filter(node => isDrawing(node, 'br')).length;
    // Explicit paragraph line spacing can make lines closer than the font size.
    // Skip it instead of claiming that tightly spaced text overflows.
    if (paragraphProperties && drawingChild(paragraphProperties, 'lnSpc')) return null;
    height += Math.max(...sizes) * 127 * hardLines;
  }
  return height || null;
}

function inspectTextBounds(doc, part, slideSize, addIssue, checks) {
  for (const shape of descendants(doc).filter(node => isPresentation(node, 'sp'))) {
    const textBody = children(shape).find(node => isPresentation(node, 'txBody'));
    if (!textBody || !descendants(textBody).some(node => isDrawing(node, 't') && node.textContent.trim())) continue;
    checks.textBoxesFound += 1;
    const shapeProperties = children(shape).find(node => isPresentation(node, 'spPr'));
    const transform = shapeProperties && drawingChild(shapeProperties, 'xfrm');
    const offset = transform && drawingChild(transform, 'off');
    const extent = transform && drawingChild(transform, 'ext');
    const x = numericAttribute(offset, 'x');
    const y = numericAttribute(offset, 'y');
    const width = numericAttribute(extent, 'cx');
    const height = numericAttribute(extent, 'cy');
    const bodyProperties = drawingChild(textBody, 'bodyPr');
    if ([x, y, width, height].some(value => value === null) || width <= 0 || height <= 0
      || hasAncestor(shape, node => isPresentation(node, 'grpSp'))
      || (numericAttribute(transform, 'rot') || 0) !== 0
      || (numericAttribute(bodyProperties, 'rot') || 0) !== 0
      || (bodyProperties && !['', 'horz'].includes(attribute(bodyProperties, 'vert')))) {
      checks.skippedTextBoxes += 1;
      continue;
    }
    checks.coordinateChecks += 1;
    const nameNode = descendants(shape).find(node => isPresentation(node, 'cNvPr'));
    const name = nameNode ? attribute(nameNode, 'name', '未命名文本框') : '未命名文本框';
    const shapeId = nameNode ? attribute(nameNode, 'id') : '';
    // Small rounding differences at the slide edge are not useful warnings.
    const tolerance = 12700;
    if (slideSize && (x < -tolerance || y < -tolerance || x + width > slideSize.width + tolerance || y + height > slideSize.height + tolerance)) {
      addIssue('PPTX_TEXT_OUTSIDE_SLIDE', part, {
        severity: 'info', title: '文本框边界超出幻灯片画布（启发式）',
        description: '根据未旋转、未分组文本框的显式坐标发现越界；这可能是有意排版，不代表文字实际被裁切。',
        evidence: `${name}（对象 ${shapeId || '未提供编号'}）；x=${x}, y=${y}, cx=${width}, cy=${height} EMU`,
        recommendation: '在目标软件中检查画布边缘；工具没有移动或缩放对象。',
      });
    }
    if (!bodyProperties || !drawingChild(bodyProperties, 'noAutofit')) continue;
    const estimatedHeight = explicitTextHeight(textBody);
    if (estimatedHeight === null) continue;
    checks.explicitLineHeightChecks += 1;
    const topMargin = numericAttribute(bodyProperties, 'tIns') ?? 45720;
    const bottomMargin = numericAttribute(bodyProperties, 'bIns') ?? 45720;
    const available = height - topMargin - bottomMargin;
    if (available > 0 && estimatedHeight > available * 1.25) {
      addIssue('PPTX_TEXT_OVERFLOW_HEURISTIC', part, {
        severity: 'warning', title: '显式行数与字号提示文本框可能过矮（启发式）',
        description: '此文本框关闭了自动适配；仅按显式换行和字号估算的高度已超过可用高度。未测量真实字形、字体替换、自动换行或目标软件排版，因此不能确认实际溢出。',
        evidence: `${name}；估算行高合计 ${(estimatedHeight / 12700).toFixed(1)} pt，可用高度 ${(available / 12700).toFixed(1)} pt`,
        recommendation: '在目标 PowerPoint 中检查文字边界。需要时手动调整文本框或字号，并保留可编辑文字。',
      });
    }
  }
}

/** Read-only inspection. It does not render a slide or assert target compatibility. */
export function scanPptx(context) {
  const { entries = new Map(), xml = new Map(), options = {} } = context;
  const issues = [];
  const counters = new Map();
  const fontInventory = new Map();
  const inventory = {
    slideCount: 0, fonts: [], externalLinks: [], media: [], embeddedObjects: [],
    embeddedFiles: [], embeddedFontFamilies: [], animations: [], chartCount: 0,
    heuristicChecks: { textBoxesFound: 0, coordinateChecks: 0, explicitLineHeightChecks: 0, skippedTextBoxes: 0 },
    notes: [
      '字体清单包含主题、母版、版式和图表中的声明；声明存在不代表每个字体都用于可见文字。',
      '未读取目标电脑的字体、字形覆盖或媒体解码器；没有进行 Windows Office 渲染或播放验证。',
      '文本边界检查仅覆盖有显式坐标的未分组、未旋转文本框；未完整解析继承、自动换行和字形度量。',
    ],
  };
  function addIssue(code, part, details) {
    const key = `${code}:${part}`;
    const index = (counters.get(key) || 0) + 1;
    counters.set(key, index);
    issues.push({
      id: `${key}:${index}`, code, severity: 'info', location: locationFor(part), part,
      fixable: false, defaultSelected: false, ...details,
    });
  }

  const presentation = xml.get('ppt/presentation.xml');
  const slideSizeNode = presentation && descendants(presentation).find(node => isPresentation(node, 'sldSz'));
  const slideWidth = numericAttribute(slideSizeNode, 'cx');
  const slideHeight = numericAttribute(slideSizeNode, 'cy');
  const slideSize = slideWidth > 0 && slideHeight > 0 ? { width: slideWidth, height: slideHeight } : null;
  if (presentation) {
    inventory.embeddedFontFamilies = descendants(presentation)
      .filter(node => isPresentation(node, 'font') && hasAncestor(node, ancestor => isPresentation(ancestor, 'embeddedFont')))
      .map(node => attribute(node, 'typeface')).filter(Boolean);
    if (inventory.embeddedFontFamilies.length) {
      addIssue('PPTX_EMBEDDED_FONTS', 'ppt/presentation.xml', {
        title: '文稿声明了嵌入字体',
        description: '检测到嵌入字体声明；未验证字体二进制、字符子集、嵌入授权或目标软件是否会采用它，不能据此保证字体一致。',
        evidence: inventory.embeddedFontFamilies.join('、'),
        recommendation: '在目标电脑中实际打开并编辑一段文字，确认缺失字符及新增文字都能正常显示。',
      });
    }
  }

  const parts = [...xml.entries()].filter(([part]) => part.startsWith('ppt/')).sort(([a], [b]) => a.localeCompare(b, 'en'));
  for (const [part, doc] of parts) {
    const nodes = descendants(doc);
    if (part.endsWith('.rels')) {
      for (const node of nodes.filter(node => RELATIONSHIP_NS.has(node.namespaceURI) && node.localName === 'Relationship')) {
        if (attribute(node, 'TargetMode') !== 'External') continue;
        const target = attribute(node, 'Target');
        const type = attribute(node, 'Type');
        const kind = externalKind(type);
        const isWebLink = kind === 'hyperlink' && /^(?:https?:|mailto:|tel:)/i.test(target);
        const link = { part, sourcePart: relationshipSource(part), id: attribute(node, 'Id'), type, target, kind };
        inventory.externalLinks.push(link);
        addIssue(isWebLink ? 'PPTX_WEB_HYPERLINK' : 'PPTX_EXTERNAL_DEPENDENCY', part, {
          severity: isWebLink ? 'info' : 'warning',
          title: isWebLink ? '包含需要外部访问的超链接' : '包含文稿之外的文件或媒体依赖',
          description: isWebLink
            ? '普通网页、邮件或电话链接本身不是排版错误；使用时仍取决于目标电脑的网络、权限及可用应用。'
            : '此关系指向文稿包以外的资源。更换电脑后，原路径、共享位置、访问权限或外部应用可能不可用。',
          evidence: `${link.id} → ${target}`,
          recommendation: isWebLink ? '需要使用链接时，在目标环境验证它是否可访问。' : '在 PowerPoint 中确认资源来源；需要离线播放时用应用自身功能嵌入可嵌入的媒体，再检查新副本。工具不会下载或改写链接。',
        });
      }
      continue;
    }

    const perPartFonts = new Map();
    for (const node of nodes.filter(node => isDrawing(node) && node.hasAttribute('typeface'))) {
      // Preserve the exact attribute for fix.from: the repair engine makes an
      // exact, minimally invasive match rather than normalizing source XML.
      const name = attribute(node, 'typeface');
      if (!name.trim()) continue;
      const supplemental = node.localName === 'font' && node.hasAttribute('script')
        && hasAncestor(node, ancestor => isDrawing(ancestor, 'majorFont') || isDrawing(ancestor, 'minorFont'));
      const symbol = node.localName === 'sym' || numericAttribute(node, 'charset') === 2;
      const key = `${node.namespaceURI}\u0000${name}`;
      const existing = perPartFonts.get(key);
      if (existing) {
        existing.count += 1;
        existing.symbol ||= symbol;
        existing.supplementalOnly &&= supplemental;
      } else perPartFonts.set(key, { name, namespace: node.namespaceURI, count: 1, symbol, supplementalOnly: supplemental });
    }
    const unverifiedSupplementalFonts = [];
    for (const { name, namespace, count, symbol, supplementalOnly } of perPartFonts.values()) {
      const themeReference = /^\+(?:mj|mn)-(?:lt|ea|cs)$/.test(name.trim());
      const isSymbol = symbol || isProtectedFont(name);
      const advice = themeReference || isSymbol ? null : fontAdvice(name, options);
      const status = themeReference ? 'theme-reference' : isSymbol ? 'special-glyph-font' : advice ? 'suggested-replacement' : isKnownFont(name, options) ? 'common-target-font' : 'unverified';
      const known = fontInventory.get(name) || { name, parts: [], status, declarations: 0 };
      if (!known.parts.includes(part)) known.parts.push(part);
      known.declarations += count;
      fontInventory.set(name, known);
      if (themeReference) continue;
      if (isSymbol) {
        addIssue('PPTX_SYMBOL_FONT', part, {
          title: '符号、公式或图标字体需要保留字形含义',
          description: '此字体可能用专用字形表达符号、公式、表情或图标。换成普通正文字体可能改变含义，因此不会自动提供字体替换。',
          evidence: `${name}；${count} 处字体声明`,
          recommendation: '确认目标电脑具有该字体或可用的嵌入字体，并逐项核对符号。',
        });
      } else if (advice && advice.to && advice.to !== name) {
        addIssue('PPTX_FONT_REPLACEMENT', part, {
          severity: 'warning', title: `字体可能需要替换：${name}`,
          description: `${advice.reason} 替换可能改变字宽、换行和版面；这里只提供需要主动选择的候选方案。`,
          evidence: `${name}；${count} 处字体声明`,
          recommendation: `可考虑替换为 ${advice.to}，随后在目标软件中检查文字、符号和换行。`,
          fixable: true, defaultSelected: false,
          fix: { kind: 'font-replace', from: name, to: advice.to, scope: namespace },
        });
      } else if (!isKnownFont(name, options) && supplementalOnly) {
        unverifiedSupplementalFonts.push(name);
      } else if (!isKnownFont(name, options)) {
        addIssue('PPTX_FONT_UNVERIFIED', part, {
          title: `目标字体尚未确认：${name}`,
          description: '此字体不在当前目标环境的常见字体清单内。工具无法得知收件人的实际安装情况，也不将未知字体判定为不兼容。',
          evidence: `${name}；${count} 处字体声明`,
          recommendation: '确认收件人具有该字体；如需替换，请先选择能覆盖原文字和符号的字体并检查版面。',
        });
      }
    }
    if (unverifiedSupplementalFonts.length) {
      addIssue('PPTX_THEME_FONTS_UNVERIFIED', part, {
        title: '主题中有尚未确认的备用脚本字体',
        description: '主题为多种语言或文字系统预设了备用字体。这些声明可能没有用于实际内容，不能据此认定缺字或排版错误；已合并显示，避免把每项预设都当成独立风险。',
        evidence: [...new Set(unverifiedSupplementalFonts)].join('、'),
        recommendation: '若文稿确实使用这些语言或脚本，再确认对应字体及目标软件的实际显示。完整声明保留在字体清单中。',
      });
    }

    const texts = nodes.filter(node => isDrawing(node, 't') || isDrawing(node, 'buChar'))
      .map(node => node.localName === 'buChar' ? attribute(node, 'char') : node.textContent || '');
    if (texts.some(text => PRIVATE_USE.test(text))) {
      addIssue('PPTX_PRIVATE_USE_GLYPHS', part, {
        severity: 'warning', title: '存在依赖特定字体的私用区字符',
        description: '私用区字符没有统一的字形含义，常见于图标字体。即使文字编码没有损坏，换字体也可能出现方框或不同图标。',
        recommendation: '保留并核对原字体及对应图标；工具没有检查字体文件的字形覆盖，也不会猜测替代字符。',
      });
    }
    if (texts.some(text => text.includes('\uFFFD'))) {
      addIssue('PPTX_REPLACEMENT_CHARACTER', part, {
        severity: 'warning', title: '文字中含 Unicode 替代字符 �',
        description: '这可能表示内容在保存之前已发生解码或字符丢失，也可能是作者有意输入；仅凭当前文稿不能恢复原字符。',
        recommendation: '与原始文本逐处核对；修改编码声明或更换字体无法可靠恢复已经丢失的内容。',
      });
    }
    if (texts.some(text => SPECIAL_GLYPHS.test(text))) {
      addIssue('PPTX_COMPLEX_GLYPHS', part, {
        title: '表情或组合字符的显示可能依赖字体与软件',
        description: '检测到表情区字符、变体选择符或连接符。它们在不同字体和软件中可能以彩色、黑白或拆分字形显示；这不是已经确认的缺字。',
        recommendation: '在目标软件中核对这些文字，特别是组合表情及其显示宽度。',
      });
    }

    const oleObjects = nodes.filter(node => isPresentation(node, 'oleObj') || isPresentation(node, 'control'));
    for (const node of oleObjects) {
      inventory.embeddedObjects.push({ part, kind: node.localName, programId: attribute(node, 'progId'), name: attribute(node, 'name') });
    }
    if (oleObjects.length) {
      addIssue('PPTX_EMBEDDED_OBJECT', part, {
        severity: 'warning', title: '包含依赖外部应用的嵌入对象或控件',
        description: 'OLE 对象或控件的预览与实际编辑、运行是不同能力。目标系统可能需要对应应用，并受平台或安全设置限制。',
        evidence: `${oleObjects.length} 个 OLE 对象或控件`,
        recommendation: '分别测试显示、双击编辑与控件功能；保留嵌入数据，工具不会将对象转成图片或执行它。',
      });
    }

    const animationCount = nodes.filter(node => isPresentation(node) && ANIMATION_ELEMENTS.has(node.localName)).length;
    const transitions = nodes.filter(node => isPresentation(node, 'transition'));
    if (animationCount || transitions.length) {
      inventory.animations.push({ part, animationNodes: animationCount, transitions: transitions.length });
      addIssue('PPTX_ANIMATION_REVIEW', part, {
        title: '包含需要实际播放核对的动画或切换',
        description: '检测到动画行为或页面切换；检测其存在不意味着不兼容。时序、触发器和视觉效果需要在目标软件中播放验证。',
        evidence: `${animationCount} 个动画行为节点；${transitions.length} 个切换定义（节点数不等于界面中的动画条目数）`,
        recommendation: '在目标电脑中放映，核对顺序、触发器与切换效果。工具保留所有动画定义。',
      });
    }

    const special = new Set();
    for (const node of nodes) {
      if (isDrawing(node) && ['scene3d', 'sp3d'].includes(node.localName)
        && !hasAncestor(node, ancestor => isDrawing(ancestor, 'fmtScheme'))) special.add('三维效果');
      if (isDrawing(node, 'prstTxWarp') && attribute(node, 'prst') !== 'textNoShape') special.add('变形艺术字');
      if (isDrawing(node, 'graphicData') && /\/diagram$/.test(attribute(node, 'uri'))) special.add('SmartArt');
      if (node.localName === 'svgBlip' && node.namespaceURI === 'http://schemas.microsoft.com/office/drawing/2016/SVG/main') special.add('SVG 图形');
      if (node.localName === 'model3d' && /schemas\.microsoft\.com\/office\/drawing\//.test(node.namespaceURI || '')) special.add('三维模型');
    }
    if (special.size) {
      addIssue('PPTX_ADVANCED_GRAPHICS', part, {
        title: '包含需要核对的图形或文字效果',
        description: '这些内容依赖相应的渲染与编辑功能。现代 PowerPoint 可支持其中多种功能；旧版本或其他演示软件可能有不同表现，不能仅凭元素类型判定故障。',
        evidence: [...special].join('、'),
        recommendation: '在目标软件中核对效果及可编辑性。工具保留原对象和已有的备用表示。',
      });
    }
    if (/^ppt\/slides\/[^/]+\.xml$/.test(part) && isPresentation(doc.documentElement, 'sld')) {
      inventory.slideCount += 1;
      inspectTextBounds(doc, part, slideSize, addIssue, inventory.heuristicChecks);
    }
    if (/^ppt\/charts\/[^/]+\.xml$/.test(part) && doc.documentElement.localName === 'chartSpace') inventory.chartCount += 1;
  }

  const typeFor = contentTypes(xml);
  for (const part of [...entries.keys()].sort()) {
    if (part.startsWith('ppt/embeddings/') && !part.endsWith('/')) inventory.embeddedFiles.push(part);
    if (!part.startsWith('ppt/media/') || part.endsWith('/')) continue;
    const contentType = typeFor(part);
    const kind = mediaKind(part, contentType);
    inventory.media.push({ part, contentType, kind, bytes: entries.get(part).byteLength });
    if (['audio', 'video'].includes(kind)) {
      addIssue('PPTX_EMBEDDED_MEDIA', part, {
        title: `包含嵌入${kind === 'audio' ? '音频' : '视频'}，需要播放验证`,
        description: '嵌入资源随文稿保存，但能否播放仍取决于媒体编解码器、PowerPoint 版本和目标系统。扩展名不能证明内部编码兼容。',
        evidence: `${part}；${contentType || '包内未声明可识别的媒体类型'}`,
        recommendation: '在目标电脑中完整播放；如遇问题，使用可信媒体工具或 PowerPoint 的媒体兼容性功能处理副本，再核对时长、声音和画面。',
      });
    }
    if (['svg', 'emf', 'wmf'].includes(extension(part))) {
      addIssue('PPTX_VECTOR_MEDIA', part, {
        title: '包含矢量图片，需要核对目标软件渲染',
        description: '矢量图片通常能保持清晰度，但不同软件及版本的支持范围、字体和备用图使用方式可能不同。文件存在不代表该资源一定显示在幻灯片中。',
        evidence: part,
        recommendation: '核对实际显示效果；工具不将矢量图转成位图，也不改写原资源。',
      });
    }
  }
  inventory.fonts = [...fontInventory.values()].sort((a, b) => a.name.localeCompare(b.name, 'en'));
  return { issues, inventory };
}
