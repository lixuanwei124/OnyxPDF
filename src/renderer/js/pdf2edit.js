import * as pdfjsLib from '../vendor/pdf.min.mjs';

/**
 * OnyxPDF - PDF 行级文本提取（可编辑导出与编辑模式共用）
 * 输出坐标：PDF 用户空间，y 轴已换算为「顶部原点向下」便于排版映射。
 */
function normalizeForAlignment(text) {
  return String(text || '').replace(/\s+/g, '');
}

function glyphArgsToText(args) {
  return (args || []).map(glyph => {
    if (typeof glyph === 'number') return ' ';
    return glyph?.unicode ?? glyph?.fontChar ?? '';
  }).join('');
}

function rgbHex(color) {
  return '#' + color.map(value => Math.max(0, Math.min(255, value)).toString(16).padStart(2, '0')).join('');
}

function normalizeComponent(value) {
  const number = Number(value) || 0;
  return number > 1 ? number : number * 255;
}

function mapTextItemsToShowOps(items, showTexts) {
  const mapped = [];
  const targetIndexes = new Map();
  (items || []).forEach((item, index) => {
    if (normalizeForAlignment(item.str)) targetIndexes.set(mapped.length, index);
    mapped.push([]);
  });
  const targets = (items || []).filter(item => normalizeForAlignment(item.str));
  let showIndex = 0;
  let showCursor = 0;
  for (const [targetIndex, item] of targets.entries()) {
    const target = normalizeForAlignment(item.str);
    const sourceOps = [];
    let consumed = 0;
    while (target && consumed < target.length && showIndex < showTexts.length) {
      const opText = normalizeForAlignment(showTexts[showIndex]);
      if (!opText) {
        showIndex++;
        showCursor = 0;
        continue;
      }
      if (showCursor >= opText.length) {
        showIndex++;
        showCursor = 0;
        continue;
      }
      const take = Math.min(opText.length - showCursor, target.length - consumed);
      if (take > 0) {
        consumed += take;
        sourceOps.push(showIndex);
        showCursor += take;
      }
      if (showCursor >= opText.length) {
        showIndex++;
        showCursor = 0;
      }
    }
    mapped[targetIndexes.get(targetIndex)] = sourceOps;
  }
  return mapped;
}

function normalizeRunText(value, isFirst) {
  let text = String(value ?? '').replace(/\s+/g, ' ');
  if (isFirst) text = text.replace(/^\s+/, '');
  return text;
}

function buildRowRuns(row) {
  const runs = [];
  let previousEnd = null;
  for (const fragment of row) {
    if (previousEnd !== null && fragment.x - previousEnd > fragment.size * 0.22) {
      const previous = runs[runs.length - 1];
      if (previous && /\s$/.test(previous.text)) previous.text += ' ';
      else runs.push({
        text: ' ', font: fragment.font || '微软雅黑', size: fragment.size,
        color: fragment.color || '#000000', bold: !!fragment.bold, sourceOps: []
      });
    }
    const text = normalizeRunText(fragment.text, !runs.length);
    if (text) {
      runs.push({
        text,
        font: fragment.font || '微软雅黑',
        size: Math.round(fragment.size * 10) / 10,
        color: fragment.color || '#000000',
        bold: !!fragment.bold,
        sourceOps: [...(fragment.sourceOps || [])].filter(Number.isInteger)
      });
    }
    previousEnd = fragment.x + fragment.w;
  }
  while (runs.length && !runs[runs.length - 1].text.trim()) runs.pop();
  return runs;
}

function normalizeRuns(runs) {
  const output = [];
  for (const run of runs || []) {
    const text = String(run.text ?? '').replace(/\u00a0/g, ' ');
    if (!text) continue;
    const style = JSON.stringify([run.font || '微软雅黑', run.size, run.color || '#000000', !!run.bold]);
    const previous = output[output.length - 1];
    if (previous && previous.__style === style) previous.text += text;
    else output.push({
      text,
      font: run.font || '微软雅黑',
      size: Number(run.size) || 12,
      color: run.color || '#000000',
      bold: !!run.bold,
      sourceOps: [...(run.sourceOps || [])].filter(Number.isInteger),
      __style: style
    });
  }
  output.forEach(run => delete run.__style);
  return output;
}


/**
 * 提取一页的文本行
 * @returns {Promise<{lines: Array<{x,y,w,h,size,text}>, width, height}>}
 *  x,y = 行左上角（PDF 空间，y 向下）；size ≈ 字号 pt
 */
export async function extractPageLines(doc, pageNum) {
  const page = await doc.getPage(pageNum);
  const vp1 = page.getViewport({ scale: 1 });
  const tc = await page.getTextContent();

  // 文字颜色与 sourceOp 都必须按原子显示操作对齐；一个 text item 可能包含多个 Tj/TJ。
  const showColors = [];
  const showTexts = [];
  try {
    const OPSL = pdfjsLib.OPS;
    const opList = await page.getOperatorList();
    let fill = [0, 0, 0];
    for (let i = 0; i < opList.fnArray.length; i++) {
      const fn = opList.fnArray[i];
      const a = opList.argsArray[i];
      if (fn === OPSL.setFillRGBColor) fill = Array.from(a, normalizeComponent).map(v => Math.round(v));
      else if (fn === OPSL.setFillCMYKColor) {
        const [cc, mm, yy, kk] = a;
        fill = [(1 - cc) * (1 - kk), (1 - mm) * (1 - kk), (1 - yy) * (1 - kk)].map(v => Math.round(v * 255));
      } else if (fn === OPSL.setFillGray) {
        const g = Math.round(a[0] * 255); fill = [g, g, g];
      } else if (fn === OPSL.showText || fn === OPSL.showSpacedText) {
        showColors.push(rgbHex(fill));
        showTexts.push(glyphArgsToText(opList.argsArray[i]?.[0]));
      }
    }
  } catch (_) { /* 颜色对齐失败回退黑色 */ }
  const itemSourceOps = mapTextItemsToShowOps(tc.items, showTexts);
  // 字体名解析：fontName(id) → 原始字体名 + 粗体标记（V12.3.7：粗体信息供提交继承，改字不再丢粗体）
  const fontNameOf = (id) => {
    try {
      const f = page.commonObjs.get(id);
      let name = (f && (f.name || f.fallbackName)) || '';
      const bold = /(?:Bold|Heavy|Black|Semib?old)|(?:^|[-_ ])Bd(?: |$|-)/i.test(name);
      name = name.replace(/^[A-Z]{6}\+/, '').replace(/-(Bold|Regular|Italic|Light|Medium)$/i, '').trim();
      return { name: name || null, bold };
    } catch (_) { return { name: null, bold: false }; }
  };

  // 1) item → 字符块（PDF 空间，y 顶部原点）
  const frags = [];
  for (let itemIndex = 0; itemIndex < tc.items.length; itemIndex++) {
    const it = tc.items[itemIndex];
    const itemColor = itemSourceOps[itemIndex]
      .map(index => showColors[index])
      .find(Boolean) || '#000000';
    if (!it.str || !it.str.trim()) {
      continue;
    }
    const [, , , d, e, f] = it.transform; // d=字号近似（缩放 y），e=x，f=基线 y
    const size = Math.hypot(it.transform[2], it.transform[3]) || Math.abs(it.transform[3]) || 10;
    const w = it.width || 0;
    const h = it.height || size;
    const fn = fontNameOf(it.fontName);
    frags.push({
      x: e,
      yTop: vp1.height - f - h,   // 基线转顶部原点
      w, h, size,
      font: fn.name,
      bold: fn.bold,
      color: itemColor,
      sourceOps: itemSourceOps[itemIndex],
      text: it.str
    });
  }
  if (!frags.length) return { lines: [], width: vp1.width, height: vp1.height };

  // 2) 按 y 聚类成行（容差 = 字号 40%，且字号相近、高度相近才合并——避免大标题与小字混行）
  frags.sort((a, b) => a.yTop - b.yTop || a.x - b.x);
  const rows = [];
  let cur = [frags[0]];
  for (let i = 1; i < frags.length; i++) {
    const f = frags[i];
    const ref = cur[0];
    const yOk = Math.abs(f.yTop - ref.yTop) < Math.min(ref.size, f.size) * 0.5;
    const sizeOk = Math.abs(f.size - ref.size) < Math.max(ref.size, f.size) * 0.35;
    if (yOk && sizeOk) {
      cur.push(f);
    } else { rows.push(cur); cur = [f]; }
  }
  rows.push(cur);

  // 2.5) 行内水平大间隙拆分：同一视觉行里若两段文字间隙巨大（左右分栏/独立元素），
  //      拆成两行 —— 防止左段落与右侧大字/标签被并成一行导致导出错位。
  const splitRows = [];
  for (const row of rows) {
    row.sort((a, b) => a.x - b.x);
    let cur2 = [row[0]];
    for (let i = 1; i < row.length; i++) {
      const prev = cur2[cur2.length - 1];
      const gap = row[i].x - (prev.x + prev.w);
      if (gap > Math.max(prev.size, row[i].size) * 1.2) {
        splitRows.push(cur2);
        cur2 = [row[i]];
      } else {
        cur2.push(row[i]);
      }
    }
    splitRows.push(cur2);
  }

  // 3) 行内按 x 排序拼接；相邻 frag 间加空格（有间隙时）
  const lines = splitRows.map(row => {
    row.sort((a, b) => a.x - b.x);
    // V12.3.2：字符宽度表（PDF 真实度量，pt）——断行排版用，逼近嵌入字体引擎
    const cw = new Map();
    const addW = (ch, w) => {
      if (!ch.trim()) return;
      const e = cw.get(ch) || { s: 0, n: 0 };
      e.s += w; e.n++;
      cw.set(ch, e);
    };
    for (const f of row) {
      const per = f.w / Math.max(f.text.length, 1);
      for (const ch of f.text) addW(ch, per);
    }
    const charWidthMap = {};
    for (const [ch, e] of cw) charWidthMap[ch] = Math.round(e.s / e.n * 100) / 100;
    const x = Math.min(...row.map(f => f.x));
    const yTop = Math.min(...row.map(f => f.yTop));
    const right = Math.max(...row.map(f => f.x + f.w));
    const size = Math.max(...row.map(f => f.size));
    const font = (row.find(f => f.font) || {}).font || null;
    const color = (row.find(f => f.color) || {}).color || '#000000';
    const bold = row.filter(f => f.font).some(f => f.bold); // V12.3.7：行级粗体（多数有字体的 frag 为粗体即粗体）
    const runs = buildRowRuns(row);
    const text = runs.map(run => run.text).join('').replace(/\s+/g, ' ').trim();
    return {
      x, y: yTop,
      w: right - x,
      h: Math.max(...row.map(f => f.h)),
      size: Math.round(size * 10) / 10,
      font, color, bold,
      sourceOps: [...new Set(row.flatMap(f => f.sourceOps || []).filter(Number.isInteger))].sort((a, b) => a - b),
      charWidthMap,
      runs: normalizeRuns(runs.map(run => ({ ...run, text: run.text.trim() === '' ? ' ' : run.text }))),
      text
    };
  }).filter(l => l.text);

  const paras = buildParagraphs(lines);
  return { lines, paras, width: vp1.width, height: vp1.height };
}

/* ══════════ 段落识别（行 → 段落，用于整块编辑与重排） ══════════ */
/**
 * 将行聚类为段落：字号/颜色/字体一致、行距合适、水平范围有交集。
 */
export function buildParagraphs(lines) {
  const paras = [];
  let cur = null;
  const overlaps = (a, b) => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) > Math.min(a.w, b.w) * 0.25;
  const order = lines.map((l, i) => i).sort((a, b) => lines[a].y - lines[b].y || lines[a].x - lines[b].x);
  for (const i of order) {
    const l = lines[i];
    const sameStyle = cur &&
      Math.abs(l.size - cur.size) < Math.max(cur.size, l.size) * 0.12 &&
      (l.color || '#000') === (cur.color || '#000') &&
      ((l.font || '') === (cur.font || ''));
    const vGap = cur ? l.y - (cur.y + cur.h) : 1e9;
    const flowOk = sameStyle && vGap > -cur.size * 0.3 && vGap < cur.size * 1.05;
    const xOk = cur ? overlaps(cur, l) || Math.abs(l.x - cur.x) < cur.size * 2.2 : false;
    if (flowOk && xOk) {
      const right = Math.max(cur.x + cur.w, l.x + l.w);
      cur.lineIdxs.push(i);
      cur.x = Math.min(cur.x, l.x);
      cur.w = right - cur.x;
      cur.h = (l.y + l.h) - cur.y;
      cur.runs = normalizeRuns([...(cur.runs || []), { text: ' ', font: l.font || '微软雅黑', size: l.size, color: l.color || '#000000', bold: !!l.bold }, ...(l.runs || [])]);
      cur.text += l.text;
    } else {
      if (cur) paras.push(cur);
      cur = {
        lineIdxs: [i],
        x: l.x, y: l.y, w: l.w, h: l.h,
        size: l.size, color: l.color || '#000000', font: l.font,
        bold: !!l.bold, // V12.3.7
        runs: normalizeRuns(l.runs || [{ text: l.text, font: l.font || '微软雅黑', size: l.size, color: l.color || '#000000', bold: !!l.bold }]),
        sourceOps: [...(l.sourceOps || [])],
        text: l.text
      };
    }
  }
  if (cur) paras.push(cur);
  return paras;
}

/**
 * V12.2 极光式编辑块划分：进入编辑模式后全文自动划分为逻辑文本块。
 * 规则（对齐极光 PDF 观感）：
 *  - 编号条目（"1." "2." "①" 等）自成一块起点，下一个编号出现即切块
 *  - 只有紧排段落（行距紧、字号相同、左对齐/重叠）才合并
 *  - 其余每行独立成块（标题、公司名、日期等）
 */
export function buildEditBlocks(lines) {
  const blocks = [];
  let cur = null;
  const order = lines.map((l, i) => i).sort((a, b) => lines[a].y - lines[b].y || lines[a].x - lines[b].x);
  const numRe = /^\s*(\d+\s*[.、）)]|\(\d+\)|[①②③④⑤⑥⑦⑧⑨⑩]|第\s*[一二三四五六七八九十]+\s*[条章节,.、])/;
  for (const i of order) {
    const l = lines[i];
    const text = l.text || '';
    const startsNumbered = numRe.test(text);
    const vGap = cur ? l.y - (cur.y + cur.h) : 1e9;
    const sameSize = cur && Math.abs(l.size - cur.size) < Math.max(cur.size, l.size) * 0.1;
    const leftAligned = cur && Math.abs(l.x - cur.x) < Math.max(cur.size, l.size) * 1.2;
    const cA = cur ? cur.x + cur.w / 2 : 0, cB = l.x + l.w / 2;
    const centerAligned = cur && Math.abs(cA - cB) < Math.max(cur.size, l.size) * 2;
    const sameFont = cur && (l.font || '') === (cur.font || '');
    const sameColor = cur && (l.color || '#000') === (cur.color || '#000');
    // 紧排合并：字号/字体/颜色一致 + 行距紧（<0.6 倍字号）+ 左对齐或重叠 + 当前行不是新编号条目
    const labelEnds = cur && /[：:]\s*$/.test(cur.text); // 标签行（以冒号结尾）独立
    const merge = cur && !startsNumbered && !labelEnds && sameSize && sameFont && sameColor &&
      vGap > -cur.size * 0.4 && vGap < Math.max(cur.size, l.size) * 1.0 &&
      (leftAligned || centerAligned);
    if (merge) {
      const right = Math.max(cur.x + cur.w, l.x + l.w);
      cur.lineIdxs.push(i);
      cur.x = Math.min(cur.x, l.x);
      cur.w = right - cur.x;
      cur.h = (l.y + l.h) - cur.y;
      const nextRuns = l.runs || [{ text, font: l.font || '微软雅黑', size: l.size, color: l.color || '#000000', bold: !!l.bold }];
      cur.runs = normalizeRuns([...(cur.runs || []), { text: ' ', font: l.font || '微软雅黑', size: l.size, color: l.color || '#000000', bold: !!l.bold }, ...nextRuns]);
      cur.sourceOps = [...new Set([...(cur.sourceOps || []), ...(l.sourceOps || [])])].sort((a, b) => a - b);
      if (l.charWidthMap) cur.charWidthMap = Object.assign(cur.charWidthMap || {}, l.charWidthMap);
      cur.text += (/[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]$/.test(cur.text) && /^[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(text)) ? text : ' ' + text;
    } else {
      if (cur) blocks.push(cur);
      cur = {
        lineIdxs: [i],
        x: l.x, y: l.y, w: l.w, h: l.h,
        size: l.size, color: l.color || '#000000', font: l.font,
        charWidthMap: l.charWidthMap || {},
        bold: !!l.bold, // V12.3.7：块继承首行粗体
        runs: normalizeRuns(l.runs || [{ text, font: l.font || '微软雅黑', size: l.size, color: l.color || '#000000', bold: !!l.bold }]),
        sourceOps: [...(l.sourceOps || [])],
        text
      };
    }
  }
  if (cur) blocks.push(cur);
  return blocks;
}

/**
 * 按宽度重排文本（canvas 测量，逐字填充，中文稳妥）。
 */
export function wrapTextToWidth(text, maxWidthPt, size, font, bold, charWidthMap) {
  const canvas = document.createElement('canvas');
  const ctx2d = canvas.getContext('2d');
  const quoted = '"' + String(font || 'Microsoft YaHei').replace(/["']/g, '') + '"';
  ctx2d.font = (bold ? '700 ' : '') + size + 'px ' + quoted + ', "Microsoft YaHei", sans-serif';
  // V12.3.2：统一 pt 度量。原字符优先用 PDF 真实宽度（charWidthMap），新字符用 canvas 度量换算
  const maxPt = maxWidthPt;
  const canvasPt = (s2) => ctx2d.measureText(s2).width * 72 / 96;
  const measure = (ch) => (charWidthMap && charWidthMap[ch]) || canvasPt(ch);
  const out = [];
  for (const hardLine of String(text).split('\n')) {
    if (!hardLine) { out.push(''); continue; }
    let line = '';
    let lineW = 0;
    for (const ch of hardLine) {
      const chW = measure(ch);
      if (line && lineW + chW > maxPt) {
        out.push(line);
        line = ch === ' ' ? '' : ch;
        lineW = ch === ' ' ? 0 : chW;
      } else {
        line += ch;
        lineW += chW;
      }
    }
    out.push(line);
  }
  return out;
}


/* ══════════ 内嵌图片提取 ══════════ */
/**
 * 提取页面内嵌位图及其位置（PDF 顶部原点坐标）。
 * @returns {Promise<Array<{x,y,w,h,dataUrl}>>}
 */
export async function extractPageImages(doc, pageNum) {
  const page = await doc.getPage(pageNum);
  const vp1 = page.getViewport({ scale: 1 });
  let opList;
  try {
    opList = await page.getOperatorList();
  } catch (_) { return []; }

  const OPS = pdfjsLib.OPS || {};
  const c = document.createElement('canvas');
  const ctx2d = c.getContext('2d');
  const images = [];
  let ctm = [1, 0, 0, 1, 0, 0]; // 当前变换（PDF 空间）

  const matMul = (m1, m2) => [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5]
  ];

  for (let i = 0; i < opList.fnArray.length; i++) {
    const fn = opList.fnArray[i];
    const args = opList.argsArray[i];
    if (fn === OPS.save) {
      // 压栈：记录深度即可（简化：不完整还原嵌套，只在顶层收集）
      continue;
    }
    if (fn === OPS.transform) {
      try { ctm = matMul(ctm, args); } catch (_) { /* noop */ }
      continue;
    }
    if (fn === OPS.paintImageXObject || fn === OPS.paintJpegXObject ||
        (OPS.paintImageXObjectRepeat && fn === OPS.paintImageXObjectRepeat)) {
      const objId = args[0];
      let obj = null;
      try {
        obj = page.objs && page.objs.has(objId) ? page.objs.get(objId)
          : await new Promise((res) => { try { page.objs.on(objId, res); setTimeout(() => res(null), 3000); } catch (_) { res(null); } });
      } catch (_) { continue; }
      if (!obj || !obj.width || !obj.data) continue;

      // 组装 PNG
      const { width, height, kind, data } = obj;
      c.width = width; c.height = height;
      const imageData = ctx2d.createImageData(width, height);
      const out = imageData.data;
      try {
        if (kind === 3) { out.set(data); }
        else if (kind === 2) {
          for (let k = 0, j = 0; k < data.length; k += 3, j += 4) {
            out[j] = data[k]; out[j + 1] = data[k + 1]; out[j + 2] = data[k + 2]; out[j + 3] = 255;
          }
        } else if (kind === 1) {
          // 1bpp 灰度：每字节 8 像素
          const rowBytes = (width + 7) >> 3;
          for (let y = 0; y < height; y++) {
            for (let x = 0; x < width; x++) {
              const bit = (data[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1;
              const v = bit ? 255 : 0;
              const j = (y * width + x) * 4;
              out[j] = out[j + 1] = out[j + 2] = v; out[j + 3] = 255;
            }
          }
        } else { continue; }
      } catch (_) { continue; }
      ctx2d.putImageData(imageData, 0, 0);

      // 位置：单位方阵经 ctm
      const [a, b, , d, e, f] = ctm;
      const w = Math.abs(a) || width;
      const h = Math.abs(d) || height;
      const x = e;
      const yTopPdf = vp1.height - f - h; // PDF y 向上 → 顶部原点
      images.push({
        x, y: yTopPdf, w, h,
        dataUrl: c.toDataURL('image/png')
      });
      continue;
    }
  }
  return images.filter(im => im.w > 8 && im.h > 8);
}


/* ══════════ 矢量矩形/色块提取（第二步：语义识别 - 形状） ══════════ */
/**
 * 提取页面矢量矩形色块（PDF 顶部原点坐标，CMYK/Gray 已转 RGB）。
 * 过滤：线框（stroke-only）、整页背景外的微小噪点（<2pt）。
 * @returns {Promise<Array<{x,y,w,h,fill:'#RRGGBB',z}>>}
 */
export async function extractPageRects(doc, pageNum) {
  const page = await doc.getPage(pageNum);
  const vp1 = page.getViewport({ scale: 1 });
  let opList;
  try { opList = await page.getOperatorList(); } catch (_) { return []; }

  const OPSL = pdfjsLib.OPS || {};
  const rects = [];
  let ctm = [1, 0, 0, 1, 0, 0];
  const ctmStack = [];
  let curFill = [0, 0, 0];
  let pendingRect = null; // {x,y,w,h}（PDF 底部原点）
  let z = 0;

  const matMul = (m1, m2) => [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5]
  ];
  const rgbOf = () => {
    let [r, g, b] = curFill;
    if (r > 1 || g > 1 || b > 1) { r /= 255; g /= 255; b /= 255; } // 兼容
    const hx = (v) => Math.max(0, Math.min(255, Math.round(v * 255))).toString(16).padStart(2, '0');
    return '#' + hx(r) + hx(g) + hx(b);
  };

  for (let i = 0; i < opList.fnArray.length; i++) {
    const fn = opList.fnArray[i];
    const a = opList.argsArray[i];
    if (fn === OPSL.save) { ctmStack.push(ctm); continue; }
    if (fn === OPSL.restore) { ctm = ctmStack.pop() || ctm; continue; }
    if (fn === OPSL.transform) { try { ctm = matMul(ctm, a); } catch (_) {} continue; }

    if (fn === OPSL.setFillRGBColor) curFill = a.map(normalizeComponent).map(v => v / 255);
    else if (fn === OPSL.setFillCMYKColor) {
      const [cc, mm, yy, kk] = a;
      curFill = [(1 - cc) * (1 - kk), (1 - mm) * (1 - kk), (1 - yy) * (1 - kk)];
    } else if (fn === OPSL.setFillGray) {
      curFill = [a[0], a[0], a[0]];
    }

    if (fn === OPSL.constructPath) {
      // args = [ops, coords, minMax?]
      const pathOps = a[0] || [];
      const hasRect = pathOps.some(op2 => op2 === OPSL.rectangle);
      const hasCurve = pathOps.some(op2 => op2 === OPSL.bezierCurveTo);
      if (hasRect && !hasCurve) {
        // 取 rectangle 坐标（coords 中 rectangle 的 4 元组）
        let rectArgs = null;
        let ci = 0;
        for (const op2 of pathOps) {
          if (op2 === OPSL.rectangle) { rectArgs = coords4(a[1], ci); break; }
          else if (op2 === OPSL.moveTo || op2 === OPSL.lineTo) ci += 2;
          else if (op2 === OPSL.bezierCurveTo) ci += 6;
        }
        if (rectArgs) {
          const [rx, ry, rw, rh] = rectArgs;
          // 应用 ctm 到原点与尺寸（轴对齐近似）
          const p0 = matMul(ctm, [1, 0, 0, 1, rx, ry]);
          const px = p0[4], pyP = p0[5];
          const w = Math.abs(rw * ctm[0]) || rw;
          const h = Math.abs(rh * ctm[3]) || rh;
          pendingRect = { x: px, y: pyP, w, h };
        }
      }
      continue;
    }

    // 填充操作完成 → 记录 pendingRect
    if (pendingRect && (fn === OPSL.fill || fn === OPSL.eoFill || fn === OPSL.fillStroke || fn === OPSL.eoFillStroke)) {
      const { x, y, w, h } = pendingRect;
      const yTop = vp1.height - y - h;
      if (w > 2 && h > 2 && w < vp1.width * 1.02 && h < vp1.height * 1.02) {
        rects.push({ x, y: yTop, w, h, fill: rgbOf(), z: ++z });
      }
      pendingRect = null;
      continue;
    }
    if (fn === OPSL.stroke || fn === OPSL.closeStroke || fn === OPSL.endPath) {
      pendingRect = null; // stroke-only 线框本轮忽略
    }
  }
  return rects;
}

// 从 coords 数组按序取 rectangle 的 4 元组
function coords4(all, ci) {
  return [all[ci], all[ci + 1], all[ci + 2], all[ci + 3]];
}
