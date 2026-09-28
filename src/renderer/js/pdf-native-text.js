/**
 * Native PDF text writer for the content editor.
 * Edited blocks are removed from the source content stream and redrawn as
 * vector text with embedded system fonts. No canvas tiles or cover layers.
 */

const FONT_FILES = {
  regular: {
    '微软雅黑': 'C:/Windows/Fonts/msyh.ttc',
    '等线': 'C:/Windows/Fonts/Deng.ttf',
    '宋体': 'C:/Windows/Fonts/simsun.ttc',
    '黑体': 'C:/Windows/Fonts/simhei.ttf',
    '楷体': 'C:/Windows/Fonts/simkai.ttf',
    '仿宋': 'C:/Windows/Fonts/simfang.ttf',
    '幼圆': 'C:/Windows/Fonts/SIMYOU.TTF',
    'arial': 'C:/Windows/Fonts/arial.ttf',
    'times new roman': 'C:/Windows/Fonts/times.ttf',
    'calibri': 'C:/Windows/Fonts/calibri.ttf',
    'segoe ui': 'C:/Windows/Fonts/segoeui.ttf',
    'tahoma': 'C:/Windows/Fonts/tahoma.ttf',
    'verdana': 'C:/Windows/Fonts/verdana.ttf',
    'georgia': 'C:/Windows/Fonts/georgia.ttf',
    'courier new': 'C:/Windows/Fonts/cour.ttf'
  },
  bold: {
    '微软雅黑': 'C:/Windows/Fonts/msyhbd.ttc',
    '等线': 'C:/Windows/Fonts/Dengb.ttf',
    '宋体': 'C:/Windows/Fonts/simsun.ttc',
    '黑体': 'C:/Windows/Fonts/simhei.ttf',
    '楷体': 'C:/Windows/Fonts/simkai.ttf',
    '仿宋': 'C:/Windows/Fonts/simfang.ttf',
    '幼圆': 'C:/Windows/Fonts/SIMYOU.TTF',
    'arial': 'C:/Windows/Fonts/arialbd.ttf',
    'times new roman': 'C:/Windows/Fonts/timesbd.ttf',
    'calibri': 'C:/Windows/Fonts/calibrib.ttf',
    'segoe ui': 'C:/Windows/Fonts/segoeuib.ttf',
    'tahoma': 'C:/Windows/Fonts/tahomabd.ttf',
    'verdana': 'C:/Windows/Fonts/verdanab.ttf',
    'georgia': 'C:/Windows/Fonts/georgiab.ttf',
    'courier new': 'C:/Windows/Fonts/courbd.ttf'
  }
};

const fontBytesCache = new Map();
const embeddedFontCache = new WeakMap();

function normalizeFontName(value) {
  return String(value || '').split(',')[0].replace(/["']/g, '').trim().toLowerCase() || '微软雅黑';
}

function isLatinFont(name) {
  return /^(arial|times new roman|calibri|segoe ui|tahoma|verdana|georgia|courier new)$/.test(name);
}

function fallbackFont(name) {
  return isLatinFont(name) ? 'Arial' : '微软雅黑';
}

function fontFilePath(name, bold) {
  const table = bold ? FONT_FILES.bold : FONT_FILES.regular;
  return table[name] || table[fallbackFont(name)];
}

async function loadFontBytes(path) {
  if (!fontBytesCache.has(path)) {
    const result = await window.onyx.readFontFile(path);
    if (!result?.ok) throw new Error(result?.error || `无法读取字体：${path}`);
    fontBytesCache.set(path, new Uint8Array(result.bytes));
  }
  return fontBytesCache.get(path);
}

async function embedFont(pdfDoc, run) {
  const pdfLib = window.PDFLib;
  if (!pdfLib || !window.fontkit) throw new Error('原生文字引擎未加载（pdf-lib/fontkit）');
  const family = normalizeFontName(run.font);
  const bold = !!run.bold;
  const cacheKey = `${family}:${bold}`;
  const cache = embeddedFontCache.get(pdfDoc) || new Map();
  if (cache.has(cacheKey)) return cache.get(cacheKey);
  const bytes = await loadFontBytes(fontFilePath(family, bold));
  const familyFontkit = {
    create: (data) => {
      const parsed = window.fontkit.create(data);
      if (typeof parsed.createSubset === 'function') return parsed;
      const fonts = parsed.fonts || [];
      return fonts.find(font => typeof font.createSubset === 'function') ||
        fonts.find(font => /YaHei|SimSun|UI/i.test(font.postscriptName || '')) ||
        fonts[0];
    }
  };
  pdfDoc.registerFontkit(familyFontkit);
  pdfDoc.context.fontkitRegistered = true;
  const font = await pdfDoc.embedFont(bytes, { subset: true });
  cache.set(cacheKey, font);
  embeddedFontCache.set(pdfDoc, cache);
  return font;
}

function hexToRgb(value) {
  const hex = String(value || '#000000').replace('#', '').padEnd(6, '0');
  return window.PDFLib.rgb(
    parseInt(hex.slice(0, 2), 16) / 255,
    parseInt(hex.slice(2, 4), 16) / 255,
    parseInt(hex.slice(4, 6), 16) / 255
  );
}

function splitHardLines(runs) {
  const rows = [];
  let row = [];
  for (const run of runs) {
    const parts = String(run.text ?? '').split('\n');
    parts.forEach((part, index) => {
      if (index) {
        rows.push(row);
        row = [];
      }
      if (part) row.push({ ...run, text: part });
    });
  }
  rows.push(row);
  return rows;
}

function canBreakBefore(text, index) {
  if (index <= 0) return true;
  const previous = text[index - 1];
  const current = text[index];
  return /[\s\p{P}]/u.test(previous) || /[\u4e00-\u9fff\u3040-\u30ff]/u.test(current);
}

function wrapRuns(rows, font, size, maxWidth) {
  if (!Number.isFinite(maxWidth) || maxWidth <= 0) {
    return rows.map(row => ({ runs: row, width: row.reduce((sum, run) => sum + font.widthOfTextAtSize(run.text, size), 0) }));
  }
  const result = [];
  for (const row of rows) {
    let current = [];
    let width = 0;
    for (const run of row) {
      let text = '';
      let textWidth = 0;
      for (const char of run.text) {
        const charWidth = font.widthOfTextAtSize(char, size);
        if (current.length && width + textWidth + charWidth > maxWidth && canBreakBefore(text, text.length)) {
          current.push({ ...run, text });
          result.push({ runs: current, width: width + textWidth });
          current = [];
          width = 0;
          text = char;
          textWidth = charWidth;
        } else {
          text += char;
          textWidth += charWidth;
        }
      }
      if (text) {
        current.push({ ...run, text });
        width += textWidth;
      }
    }
    result.push({ runs: current, width });
  }
  return result;
}

export async function drawNativeTextRuns(page, x, yTop, width, lineHeight, runs, options = {}) {
  const visible = (runs || []).filter(run => run.text && String(run.text).length);
  if (!visible.length) return 0;
  const sizes = visible.map(run => Number(run.size) || options.size || 12);
  const size = Math.max(...sizes);
  const rows = splitHardLines(visible);
  const firstFont = await embedFont(page.doc, visible[0]);
  const wrapped = wrapRuns(rows, firstFont, size, width);
  const height = page.getHeight();
  const baseLineHeight = Number(lineHeight) || size * 1.45;
  let rowIndex = 0;
  for (const row of wrapped) {
    let cursor = x;
    if (options.align === 'center') cursor += Math.max(0, (width - row.width) / 2);
    if (options.align === 'right') cursor += Math.max(0, width - row.width);
    const baseline = height - yTop - rowIndex * baseLineHeight - size * 0.82;
    for (const run of row.runs) {
      if (!run.text) continue;
      const font = await embedFont(page.doc, run);
      const runSize = Number(run.size) || size;
      const runBaseline = height - yTop - rowIndex * baseLineHeight - runSize * 0.82;
      page.drawText(run.text, {
        x: cursor,
        y: runBaseline,
        size: runSize,
        font,
        color: hexToRgb(run.color || options.color || '#000000')
      });
      cursor += font.widthOfTextAtSize(run.text, runSize);
    }
    rowIndex++;
  }
  return rowIndex;
}



