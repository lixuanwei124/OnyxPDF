import { removeTextOperations } from './pdf-content-stream.js';

const FONT_FILES = {
  regular: {
    '微软雅黑': 'C:/Windows/Fonts/msyh.ttc',
    '等线': 'C:/Windows/Fonts/Deng.ttf',
    '宋体': 'C:/Windows/Fonts/simsun.ttc',
    '黑体': 'C:/Windows/Fonts/simhei.ttf',
    '楷体': 'C:/Windows/Fonts/simkai.ttf',
    '仿宋': 'C:/Windows/Fonts/simfang.ttf',
    '幼圆': 'C:/Windows/Fonts/SIMYOU.TTF',
    arial: 'C:/Windows/Fonts/arial.ttf',
    'times new roman': 'C:/Windows/Fonts/times.ttf',
    calibri: 'C:/Windows/Fonts/calibri.ttf',
    'segoe ui': 'C:/Windows/Fonts/segoeui.ttf',
    tahoma: 'C:/Windows/Fonts/tahoma.ttf',
    verdana: 'C:/Windows/Fonts/verdana.ttf',
    georgia: 'C:/Windows/Fonts/georgia.ttf',
    'courier new': 'C:/Windows/Fonts/cour.ttf'
    ,
    symbol: 'C:/Windows/Fonts/symbol.ttf'
  },
  bold: {
    '微软雅黑': 'C:/Windows/Fonts/msyhbd.ttc',
    '等线': 'C:/Windows/Fonts/Dengb.ttf',
    '宋体': 'C:/Windows/Fonts/simsun.ttc',
    '黑体': 'C:/Windows/Fonts/simhei.ttf',
    '楷体': 'C:/Windows/Fonts/simkai.ttf',
    '仿宋': 'C:/Windows/Fonts/simfang.ttf',
    '幼圆': 'C:/Windows/Fonts/SIMYOU.TTF',
    arial: 'C:/Windows/Fonts/arialbd.ttf',
    'times new roman': 'C:/Windows/Fonts/timesbd.ttf',
    calibri: 'C:/Windows/Fonts/calibrib.ttf',
    'segoe ui': 'C:/Windows/Fonts/segoeuib.ttf',
    tahoma: 'C:/Windows/Fonts/tahomabd.ttf',
    verdana: 'C:/Windows/Fonts/verdanab.ttf',
    georgia: 'C:/Windows/Fonts/georgiab.ttf',
    'courier new': 'C:/Windows/Fonts/courbd.ttf'
    ,
    symbol: 'C:/Windows/Fonts/symbol.ttf'
  }
};

const FONT_ALIASES = {
  regular: {
    '微软雅黑': ['microsoft yahei', 'msyh', 'msyh.ttf'],
    '等线': ['dengxian', 'dengxian light'],
    '宋体': ['simsun', 'nsimsun', 'songti', 'song'],
    '黑体': ['simhei', 'hei', 'heit'],
    '楷体': ['kaiti', 'kaiti gb2312', 'simkai'],
    '仿宋': ['fangsong', 'fangsong gb2312', 'simfang', 'fs'],
    '幼圆': ['youyuan', 'simyou'],
    arial: ['arial', 'arialmt'],
    'times new roman': ['times new roman', 'timesnewroman', 'timesnewromanpsmt'],
    calibri: ['calibri', 'calibri light'],
    'segoe ui': ['segoe ui', 'segoeui'],
    tahoma: ['tahoma'],
    verdana: ['verdana'],
    georgia: ['georgia'],
    'courier new': ['courier new', 'couriernew', 'couriernewpsmt'],
    symbol: ['symbol', 'symbolmt']
  }
};

const fontBytes = new Map();
const embeddedFonts = new WeakMap();
const subsetFonts = new WeakMap();

function normalizeFont(value) {
  return String(value || '微软雅黑').split(',')[0].replace(/["']/g, '').trim().toLowerCase() || '微软雅黑';
}

function canonicalFont(value) {
  const name = normalizeFont(value);
  for (const [canonical, aliases] of Object.entries(FONT_ALIASES.regular)) {
    if (name === canonical || aliases.includes(name)) return canonical;
  }
  return name;
}

function fontPath(value, bold) {
  const name = canonicalFont(value);
  const table = bold ? FONT_FILES.bold : FONT_FILES.regular;
  if (table[name]) return table[name];
  if (/(arial|helvetica|times|calibri|segoe|tahoma|verdana|georgia|courier)/.test(name)) {
    return table.arial;
  }
  return table['微软雅黑'];
}

async function loadFont(path) {
  if (!fontBytes.has(path)) {
    const result = await window.onyx.readFontFile(path);
    if (!result?.ok) throw new Error(result?.error || `无法读取字体：${path}`);
    fontBytes.set(path, new Uint8Array(result.bytes));
  }
  return fontBytes.get(path);
}

async function embedFont(pdfDocument, run) {
  const pdfLib = window.PDFLib;
  if (!pdfLib || !window.fontkit) throw new Error('PDF 原生文字组件未加载');
  const family = normalizeFont(run.font);
  const bold = !!run.bold;
  const key = run.fontData ? `original:${run.fontKey || run.font}:${bold}` : `${family}:${bold}`;
  const cache = embeddedFonts.get(pdfDocument) || new Map();
  if (cache.has(key)) return cache.get(key);

  const bytes = run.fontData instanceof Uint8Array ? run.fontData : await loadFont(fontPath(family, bold));
  pdfDocument.registerFontkit({
    create: data => {
      const parsed = window.fontkit.create(data);
      if (typeof parsed.createSubset === 'function') return parsed;
      const fonts = parsed.fonts || [];
      return fonts.find(font => typeof font.createSubset === 'function') ||
        fonts.find(font => /YaHei|SimSun|UI/i.test(font.postscriptName || '')) ||
        fonts[0];
    }
  });
  const font = await pdfDocument.embedFont(bytes, { subset: true });
  cache.set(key, font);
  embeddedFonts.set(pdfDocument, cache);
  return font;
}

function originalSubsetFont(run) {
  if (!(run.fontData instanceof Uint8Array)) return null;
  if (subsetFonts.has(run.fontData)) return subsetFonts.get(run.fontData);
  let font = null;
  try {
    const parsed = window.fontkit.create(run.fontData);
    font = typeof parsed.createSubset === 'function' ? parsed :
      (parsed.fonts || []).find(item => typeof item.createSubset === 'function') ||
      parsed.fonts?.[0] || parsed;
  } catch (_) {
    font = null;
  }
  subsetFonts.set(run.fontData, font);
  return font;
}

export function originalFontSupportsText(run) {
  if (!(run.fontData instanceof Uint8Array)) return false;
  const font = originalSubsetFont(run);
  if (!font || typeof font.hasGlyphForCodePoint !== 'function') return false;
  return Array.from(run.text).every(character => {
    const codePoint = character.codePointAt(0);
    return character === '\n' || character === '\r' ||
      font.hasGlyphForCodePoint(codePoint) !== false;
  });
}

export function splitRunBySupportedGlyphs(run) {
  if (!run.fontData || !run.text) return [run];
  const font = originalSubsetFont(run);
  if (!font || typeof font.hasGlyphForCodePoint !== 'function') {
    return [{ ...run, fontKey: null, fontData: null, loadedName: null }];
  }

  const parts = [];
  for (const character of Array.from(run.text)) {
    const supported = character === '\n' || character === '\r' ||
      font.hasGlyphForCodePoint(character.codePointAt(0)) !== false;
    const last = parts[parts.length - 1];
    if (last?.supported === supported) last.text += character;
    else parts.push({ text: character, supported });
  }

  return parts.map(part => part.supported ? { ...run, text: part.text } : {
    ...run,
    text: part.text,
    fontKey: null,
    fontData: null,
    loadedName: null
  });
}
function runColor(value) {
  const source = String(value || '#000000').trim();
  const rgb = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(source);
  if (rgb) {
    return window.PDFLib.rgb(
      Number(rgb[1]) / 255,
      Number(rgb[2]) / 255,
      Number(rgb[3]) / 255
    );
  }
  const hex = source.replace('#', '').padEnd(6, '0');
  return window.PDFLib.rgb(
    parseInt(hex.slice(0, 2), 16) / 255 || 0,
    parseInt(hex.slice(2, 4), 16) / 255 || 0,
    parseInt(hex.slice(4, 6), 16) / 255 || 0
  );
}

function contentAtoms(runs) {
  const atoms = [];
  runs.forEach((run, runIndex) => {
    const source = String(run.text ?? '');
    let index = 0;
    while (index < source.length) {
      if (source.startsWith('\r\n', index)) {
        atoms.push({ runIndex, text: '\n', newline: true });
        index += 2;
      } else if (source[index] === '\n' || source[index] === '\r') {
        atoms.push({ runIndex, text: '\n', newline: true });
        index++;
      } else if (/\s/u.test(source[index])) {
        let end = index + 1;
        while (end < source.length && /\s/u.test(source[end]) && source[end] !== '\n' && source[end] !== '\r') end++;
        atoms.push({ runIndex, text: source.slice(index, end) });
        index = end;
      } else if (/[\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/u.test(source[index])) {
        atoms.push({ runIndex, text: source[index] });
        index++;
      } else {
        let end = index + 1;
        while (end < source.length && !/\s/u.test(source[end]) &&
          !/[\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/u.test(source[end])) end++;
        atoms.push({ runIndex, text: source.slice(index, end) });
        index = end;
      }
    }
  });
  return atoms;
}

const NO_LINE_START = /[!%),\.:;\?\]\}¢°·'"”’›〕〗〞﹚﹜！％＇），．：；？］｝、。：；？！]/u;
const NO_LINE_END = /[\(\[\{£¥'"“‘‹〔〖﹙﹛（《“‘【〔]/u;

function clusterContentAtoms(atoms) {
  const clusters = [];
  for (let index = 0; index < atoms.length; index++) {
    const atom = atoms[index];
    if (!atom.newline && NO_LINE_END.test(Array.from(atom.text).pop() || '')) {
      const next = atoms[index + 1];
      if (next && !next.newline && next.runIndex === atom.runIndex) {
        clusters.push({ ...atom, text: atom.text + next.text });
        index++;
        continue;
      }
    }
    if (!atom.newline && NO_LINE_START.test(Array.from(atom.text)[0] || '')) {
      const previous = clusters[clusters.length - 1];
      if (previous && !previous.newline && previous.runIndex === atom.runIndex) {
        previous.text += atom.text;
        continue;
      }
    }
    clusters.push({ ...atom });
  }
  return clusters;
}

export function visualRows(runs, maxWidth, metricTolerance = 0) {
  const rows = [[]];
  let rowWidth = 0;
  const textWidth = item => item.run.font.widthOfTextAtSize(item.text, item.run.size);
  const flush = () => {
    if (rows[rows.length - 1].length) rows.push([]);
    rowWidth = 0;
  };

  for (const atom of clusterContentAtoms(contentAtoms(runs))) {
    if (atom.newline) {
      flush();
      continue;
    }
    const run = runs[atom.runIndex];
    const item = { run, text: atom.text };
    const width = textWidth(item);
    if ((rowWidth + width <= maxWidth + metricTolerance || !rowWidth) && width <= maxWidth) {
      rows[rows.length - 1].push(item);
      rowWidth += width;
      continue;
    }

    flush();
    const characters = Array.from(atom.text);
    for (let start = 0; start < characters.length;) {
      let end = characters.length;
      while (end > start + 1) {
        const text = characters.slice(start, end).join('');
        if (run.font.widthOfTextAtSize(text, run.size) <= maxWidth) break;
        end--;
      }
      const text = characters.slice(start, end).join('');
      const piece = { run, text };
      rows[rows.length - 1].push(piece);
      rowWidth = textWidth(piece);
      start = end;
      if (start < characters.length) flush();
    }
  }
  while (rows.length > 1 && !rows[rows.length - 1].length) rows.pop();

  const canStartRow = item => !NO_LINE_START.test(Array.from(item.text)[0] || '');
  const canEndRow = item => !NO_LINE_END.test(Array.from(item.text).pop() || '');
  const rowWidths = rows.map(row => row.reduce((width, item) => width + textWidth(item), 0));
  for (let pass = 0; pass < 4; pass++) {
    let changed = false;
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      const previous = rows[rowIndex - 1];
      const current = rows[rowIndex];
      while (previous.length > 1 && current.length && !canStartRow(current[0])) {
        const moved = previous.pop();
        current.unshift(moved);
        rowWidths[rowIndex - 1] -= textWidth(moved);
        rowWidths[rowIndex] += textWidth(moved);
        changed = true;
      }
      while (previous.length > 1 && current.length && !canEndRow(previous[previous.length - 1])) {
        const moved = previous.pop();
        current.unshift(moved);
        rowWidths[rowIndex - 1] -= textWidth(moved);
        rowWidths[rowIndex] += textWidth(moved);
        changed = true;
      }
    }
    if (!changed) break;
  }
  return rows.map(row => row.map(item => ({ ...item.run, text: item.text })));
}

export async function drawContentBlock(page, block) {
  if (block.patch && Number(block.patch.width) > 0 && Number(block.patch.height) > 0) {
    page.drawRectangle({
      x: Number(block.patch.x) || 0,
      y: Number(block.patch.y) || 0,
      width: Number(block.patch.width),
      height: Number(block.patch.height),
      color: runColor(block.patchColor || '#ffffff')
    });
  }
  const preparedRuns = [];
  for (const run of block.runs || []) {
    for (const supportedRun of splitRunBySupportedGlyphs(run)) {
      if (!supportedRun.text) continue;
      const family = normalizeFont(supportedRun.font);
      const syntheticBold = !!supportedRun.bold && !(supportedRun.fontData instanceof Uint8Array) &&
        fontPath(family, true) === fontPath(family, false);
      preparedRuns.push({
        ...supportedRun,
        syntheticBold,
        font: await embedFont(page.doc, supportedRun),
        size: Number(supportedRun.size) || block.size || 12
      });
    }
  }

  const rotation = Math.abs(page.getRotation().angle) % 360;
  const pageSize = page.getSize();
  const landscape = [90, 270].includes(rotation);
  const pageWidth = landscape ? pageSize.height : pageSize.width;
  const pageHeight = landscape ? pageSize.width : pageSize.height;
  const savedWidth = Number(block.maxWidth);
  const availableWidth = Math.max(pageWidth - block.x, 1);
  const fixedWidth = block.manualWidth === true || savedWidth > 0 || Number(block.lineCount) > 1;
  const metricTolerance = 0;
  const maxWidth = fixedWidth && savedWidth > 0
    ? Math.max(Math.min(savedWidth, availableWidth), 1)
    : availableWidth;
  const rows = visualRows(preparedRuns, maxWidth + (fixedWidth ? metricTolerance : 0), metricTolerance);
  const rowHeight = block.lineHeight || block.size * 1.22 || 12;
  const bottomPadding = rowHeight * 0.2;
  const minimumStart = (rows.length - 1) * rowHeight + bottomPadding;
  const maximumStart = pageHeight - bottomPadding;
  const startBaseline = Math.min(Math.max(block.baseline, minimumStart), Math.max(minimumStart, maximumStart));
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const rowWidth = row.reduce((width, run) => width + run.font.widthOfTextAtSize(run.text, run.size), 0);
    let cursor = block.x;
    const frameWidth = Number(block.width) || savedWidth || 0;
    if (block.alignment === 'center' && frameWidth > 0) {
      cursor += (frameWidth - rowWidth) / 2;
    } else if (block.alignment === 'right' && frameWidth > 0) {
      cursor += Math.max(0, frameWidth - rowWidth);
    }
    const baseline = startBaseline - rowIndex * rowHeight;
    for (const run of row) {
      if (!run.text) continue;
      if (run.syntheticBold) {
        const pdfLib = window.PDFLib;
        const strokeColor = runColor(run.color);
        page.pushOperators(
          pdfLib.pushGraphicsState(),
          pdfLib.setLineWidth(Math.max(0.08, run.size * 0.022)),
          pdfLib.setStrokingColor(strokeColor),
          pdfLib.setTextRenderingMode(2)
        );
        page.drawText(run.text, {
          x: cursor,
          y: baseline,
          size: run.size,
          font: run.font,
          color: strokeColor
        });
        page.pushOperators(pdfLib.popGraphicsState());
      } else {
        page.drawText(run.text, {
          x: cursor,
          y: baseline,
          size: run.size,
          font: run.font,
          color: runColor(run.color)
        });
      }
      cursor += run.font.widthOfTextAtSize(run.text, run.size);
    }
  }
}

export async function applyContentEdits(pdfDocument, edits) {
  const byPage = new Map();
  for (const edit of edits || []) {
    if (edit.engine !== 2) continue;
    if (!byPage.has(edit.page)) byPage.set(edit.page, new Set());
    for (const operation of edit.sourceOps || []) {
      if (Number.isInteger(operation)) byPage.get(edit.page).add(operation);
    }
  }
  for (const [pageNumber, sourceOperations] of byPage) {
    try {
      removeTextOperations(pdfDocument, pageNumber - 1, [...sourceOperations], { all: false });
    } catch (error) {
      if (!/PDF 原文移除不完整/.test(error?.message || '')) throw error;
      removeTextOperations(pdfDocument, pageNumber - 1, [], { all: true });
    }
  }
  for (const edit of edits || []) {
    const page = pdfDocument.getPage(edit.page - 1);
    await drawContentBlock(page, edit);
  }
  return byPage.size > 0;
}
