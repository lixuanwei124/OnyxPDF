/**
 * OnyxPDF - 应用入口 / 全局编排（V3：多标签页 + 扩展导出 + 页面编辑）
 */
import * as pdfjsLib from '../vendor/pdf.min.mjs';
import { Viewer } from './viewer.js';
import { Annotations } from './annotations.js';
import { ContentEditor } from './content-editor.js';
import { Search } from './search.js';
import { Sidebar } from './sidebar.js';
import { exportAnnotatedPdf } from './exporter.js';
import { initMenus, I } from './menus.js';
import { openModal } from './ui.js';
import * as tools from './tools.js';
import { buildDocxParagraphLayout } from './docx.js';
import { zipStore } from './docx.js';
import { buildPptx, buildPptxEditable } from './pptx.js';
import { buildXlsxLayout } from './xlsx.js';
import { clamp, toast, escapeHtml, formatBytes, timeAgo } from './utils.js';
import { extractPageLines, extractPageImages, extractPageRects } from './pdf2edit.js';
import { extractContentPage, applySavedEdits } from './content-model.js';
import { removeTextOperations } from './pdf-content-stream.js';
import { extractEmbeddedFonts } from './embedded-fonts.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = '../vendor/pdf.worker.min.mjs';

/* ══════════ 全局状态 ══════════ */
const state = {
  filePath: null,
  fileName: '',
  fileSize: 0,
  numPages: 0,
  originalBytes: null
};
let fileKey = '';
let docSeq = 0;

/* ══════════ 标签页 ══════════ */
const tabs = [];            // { id, path, name, bytes? }
let activeTabId = null;
let tabSeq = 0;

const viewer = new Viewer({
  pdfjsLib,
  fileKey: '',
  onStateChange: onViewerEvent
});

const annotations = new Annotations({
  get viewer() { return viewer; },
  get fileKey() { return fileKey; },
  getOriginalBytes() { return state.originalBytes; }
});

const contentEditor = new ContentEditor({
  get viewer() { return viewer; },
  getOriginalBytes() { return state.originalBytes; }
});

const search = new Search({
  get viewer() { return viewer; }
});

const sidebar = new Sidebar({
  get viewer() { return viewer; },
  onReorderApply: applyPageReorder
});

function onViewerEvent(ev) {
  if (ev.type === 'page') {
    document.getElementById('st-page').textContent = `${viewer.currentPage} / ${state.numPages}`;
    sidebar.markCurrentThumb(viewer.currentPage);
  } else if (ev.type === 'zoom') {
    document.getElementById('st-zoom').textContent = `${Math.round(viewer.scale * 100)}%`;
  } else if (ev.type === 'page-rendered') {
    viewer.renderTextLayer(ev.page, ev.viewport);
    if (contentEditor.enabled) contentEditor.renderPage(ev.page);
    annotations.redrawPage(ev.page);
    if (annotations.editMode && !annotations._previewBusy) annotations.renderElementOverlay(ev.page);
  }
}

/* ────── 标签页管理 ────── */
function renderTabs() {
  const bar = document.getElementById('tabbar');
  bar.innerHTML = '';
  document.body.classList.toggle('has-tabs', tabs.length > 0);
  for (const t of tabs) {
    const el = document.createElement('div');
    el.className = 'tab' + (t.id === activeTabId ? ' active' : '');
    el.title = t.path || t.name;
    el.innerHTML = `<span class="t-name">${escapeHtml(t.name)}</span><button class="t-close" data-id="${t.id}">✕</button>`;
    el.addEventListener('click', (e) => {
      if (e.target.closest('.t-close')) return;
      activateTab(t.id);
    });
    el.querySelector('.t-close').addEventListener('click', (e) => {
      e.stopPropagation();
      closeTab(t.id);
    });
    bar.appendChild(el);
  }
  const add = document.createElement('button');
  add.className = 'tab-add';
  add.title = '打开 PDF（新标签页）';
  add.textContent = '+';
  add.addEventListener('click', openViaDialog);
  bar.appendChild(add);
}

async function activateTab(id) {
  if (id === activeTabId) return;
  activeTabId = id;
  renderTabs();
  const t = tabs.find(x => x.id === id);
  if (!t) return;
  if (t.bytes) await loadDocFromBytes(t.bytes, t.name);
  else await loadDocFromPath(t.path);
}

async function closeTab(id) {
  const idx = tabs.findIndex(t => t.id === id);
  if (idx < 0) return;
  tabs.splice(idx, 1);
  if (activeTabId === id) {
    if (tabs.length) {
      await activateTab(tabs[Math.max(0, idx - 1)].id);
      return;
    }
    activeTabId = null;
    // 无标签页：回到欢迎页
    viewer.ctx.fileKey = '';
    await viewer.close();
    fileKey = '';
    Object.assign(state, { filePath: null, fileName: '', numPages: 0, originalBytes: null });
    document.getElementById('tb-file-title').textContent = '未打开文档';
    document.getElementById('st-page').textContent = '– / –';
    document.getElementById('st-file').textContent = '';
    search.reset();
  }
  renderTabs();
}

/* ────── 文档装载 ────── */
async function loadDocFromPath(p) {
  const res = await window.onyx.readPdf(p);
  if (!res.ok) { toast(`读取失败：${res.error}`, true); return; }
  const name = p.replace(/^.*[\\/]/, '');
  const buf = res.data;
  state.filePath = p;
  state.fileName = name;
  state.originalBytes = buf;
  state.fileSize = buf.byteLength;
  state.modified = false;

  viewer.ctx.fileKey = '';
  // pdf.js 会 detach 传入的 ArrayBuffer，必须传副本以保留原始字节供导出使用
  const info = await viewer.open(p, buf.slice(0), { path: p, name, size: buf.byteLength });
  state.numPages = info.numPages;
  fileKey = `${name}::${buf.byteLength}`;
  viewer.ctx.fileKey = fileKey;
  contentEditor.load(fileKey);
  applyDocChrome(name);
  viewer.goToPage(1, false);
  annotations.loadFor(fileKey);
  await contentEditor.showCommitted().catch(error => console.warn('恢复内容编辑失败:', error));
  if (annotations.count()) await annotations._refreshPreview().catch(error => console.warn('恢复编辑预览失败:', error));
  search.reset();
  sidebar.reset();
  sidebar.buildThumbs();
  sidebar.buildOutline();
  viewer.restorePosition();
}

async function loadDocFromBytes(masterBytes, name) {
  const buf = masterBytes.slice(0);
  state.filePath = null;
  state.fileName = name;
  state.originalBytes = buf;
  state.fileSize = buf.byteLength;
  state.modified = false;

  viewer.ctx.fileKey = '';
  const info = await viewer.open(name, buf.slice(0), { path: null, name, size: buf.byteLength });
  state.numPages = info.numPages;
  fileKey = `${name}::${buf.byteLength}`;
  viewer.ctx.fileKey = fileKey;
  contentEditor.load(fileKey);
  applyDocChrome(name);
  viewer.goToPage(1, false);
  annotations.loadFor(fileKey);
  await contentEditor.showCommitted().catch(error => console.warn('恢复内容编辑失败:', error));
  if (annotations.count()) await annotations._refreshPreview().catch(error => console.warn('恢复编辑预览失败:', error));
  search.reset();
  sidebar.reset();
  sidebar.buildThumbs();
  sidebar.buildOutline();
}

function applyDocChrome(name) {
  document.getElementById('tb-file-title').textContent = name;
  document.getElementById('st-page').textContent = `1 / ${state.numPages}`;
  document.getElementById('st-zoom').textContent = `${Math.round(viewer.scale * 100)}%`;
  document.getElementById('st-file').textContent = `${name} · ${formatBytes(state.fileSize)}`;
}

/* 打开入口：同路径聚焦标签，否则新建标签 */
async function openFromPath(p) {
  if (!p || !/\.pdf$/i.test(p)) { toast('仅支持 PDF 文件', true); return; }
  try {
    const existing = tabs.find(t => t.path === p);
    if (existing) {
      if (existing.id === activeTabId) {
        await loadDocFromPath(p); // 重新加载（外部变更场景）
      } else {
        await activateTab(existing.id);
      }
      return;
    }
    const t = { id: `t${++tabSeq}`, path: p, name: p.replace(/^.*[\\/]/, '') };
    tabs.push(t);
    activeTabId = t.id;
    renderTabs();
    await loadDocFromPath(p);
    await window.onyx.recentsAdd({
      path: p, name: t.name,
      size: state.fileSize, time: Date.now()
    });
  } catch (err) {
    console.error('打开失败:', err);
    toast(`打开失败：${err && err.message || err}`, true);
  }
}

async function openFromBytes(bytes, name) {
  // 同名虚拟标签直接替换内容
  let t = tabs.find(x => !x.path && x.name === name);
  if (!t) {
    t = { id: `t${++tabSeq}`, path: null, name, bytes };
    tabs.push(t);
  } else {
    t.bytes = bytes;
  }
  activeTabId = t.id;
  renderTabs();
  await loadDocFromBytes(bytes, name);
}

async function openViaDialog() {
  const p = await window.onyx.openDialog();
  if (p) await openFromPath(p);
}

/* ══════════ 保存 / 另存为 / 导出 ══════════ */
function collectAnnotations() {
  try { if (contentEditor.enabled) contentEditor.commitDom(); } catch (_) { /* noop */ }
  const all = [];
  for (const list of annotations.map.values()) all.push(...list);
  all.push(...contentEditor.allImages());
  return all;
}

async function buildAnnotatedBytes() {
  const all = collectAnnotations();
  const contentEdits = contentEditor.allEdits();
  if (!all.length && !contentEdits.length) return null;
  const { bytes } = await exportAnnotatedPdf(
    state.originalBytes, all,
    state.fileName.replace(/\.pdf$/i, '') + '-annotated.pdf', contentEdits);
  return bytes;
}

/** 保存（Ctrl+S）：有路径则覆写原文件（含批注）；无路径转另存为 */
async function doSave() {
  if (contentEditor.enabled) contentEditor.commitDom();
  if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
  if (!state.filePath) { await doSaveAs(); return; }
  if (!collectAnnotations().length && !contentEditor.allEdits().length) { toast('文档未做修改，无需保存'); return; }
  try {
    const bytes = await buildAnnotatedBytes();
    const r = await window.onyx.writePath({ path: state.filePath, bytes });
    if (r.ok) {
      toast('已保存到原文件（含批注）');
      const p = state.filePath;
      await openFromPath(p);
    } else {
      toast(`原文件写入失败（${r.error}），转另存为`, true);
      await doSaveAs();
    }
  } catch (err) {
    toast(`保存失败：${err && err.message || err}`, true);
  }
}

/** 另存为（Ctrl+Shift+S）：含批注 PDF */
async function doSaveAs() {
  if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
  try {
    const bytes = await buildAnnotatedBytes();
    if (!bytes) { toast('当前文档没有批注，原样另存'); }
    const payload = bytes || new Uint8Array(state.originalBytes);
    const r = await window.onyx.exportPdf({
      bytes: payload,
      suggestedName: state.fileName.replace(/\.pdf$/i, '') + (bytes ? '-annotated.pdf' : '-copy.pdf'),
      title: '另存为'
    });
    if (r.ok) toast(`已另存：${r.path}`);
    else if (!r.canceled) toast(`另存失败：${r.error}`, true);
  } catch (err) {
    console.error('另存失败:', err);
    toast(`另存失败：${err && err.message || err}`, true);
  }
}

/** 导出为…（Acrobat 式：保存类型下拉选择） */
async function doExportAs() {
  if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
  const formats = [
    { id: 'pdf-annot', label: 'PDF 文档（含批注）(.pdf)', ext: 'pdf', filter: 'PDF 文档' },
    { id: 'pdf-plain', label: 'PDF 文档（原样）(.pdf)', ext: 'pdf', filter: 'PDF 文档' },
    { id: 'docx', label: '导出 Word', ext: 'docx', filter: 'Word 文档' },
    { id: 'xlsx', label: '导出 Excel', ext: 'xlsx', filter: 'Excel 工作簿' },
    { id: 'pptx', label: '导出 PPT', ext: 'pptx', filter: 'PowerPoint 演示文稿' },
    { id: 'txt', label: '纯文本 (.txt)', ext: 'txt', filter: '文本文件' },
    { id: 'png', label: 'PNG 图片（全部页）(.zip)', ext: 'zip', filter: 'PNG 图片包' },
    { id: 'jpg', label: 'JPG 图片（全部页）(.zip)', ext: 'zip', filter: 'JPG 图片包' }
  ];
  openModal({
    title: '导出为',
    width: 400,
    bodyHTML: `
      <div class="field">
        <label>保存类型</label>
        <select class="export-format" data-role="fmt">
          ${formats.map(f => `<option value="${f.id}">${f.label}</option>`).join('')}
        </select>
      </div>
      <p class="hint">Word/PPT/Excel 会按原位置还原可编辑段落和原版式背景；批注与编辑内容会一并写入 PDF 导出。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '导出', primary: true,
        onClick: async (m) => {
          const id = m.query('[data-role="fmt"]').value;
          const fmt = formats.find(f => f.id === id);
          m.close();
          try {
            let bytes, suggestedName;
            const base = state.fileName.replace(/\.pdf$/i, '');
            toast(`正在生成 ${fmt.label}…`);
            switch (id) {
              case 'pdf-annot': {
                const merged = await buildAnnotatedBytes();
                bytes = merged || new Uint8Array(state.originalBytes);
                suggestedName = base + (merged ? '-annotated.pdf' : '-copy.pdf');
                break;
              }
              case 'pdf-plain': bytes = new Uint8Array(state.originalBytes); suggestedName = base + '-copy.pdf'; break;
              case 'docx': ({ bytes, suggestedName } = await buildWordLayoutExport()); break;
              case 'xlsx': ({ bytes, suggestedName } = await buildXlsxExport()); break;
              case 'pptx': ({ bytes, suggestedName } = await buildPptxEditableExportV2()); break;
              case 'txt': ({ bytes, suggestedName } = await buildTxtExport()); break;
              case 'png': ({ bytes, suggestedName } = await buildPageImageExport('png')); break;
              case 'jpg': ({ bytes, suggestedName } = await buildPageImageExport('jpg')); break;
            }
            const r = await window.onyx.exportPdf({
              bytes, suggestedName,
              title: '导出为 ' + fmt.label,
              filters: [{ name: fmt.filter, extensions: [fmt.ext] }]
            });
            if (r.ok) toast(`已导出：${r.path}`);
            else if (!r.canceled) toast(`导出失败：${r.error}`, true);
          } catch (err) {
            console.error('导出失败:', err);
            toast(`导出失败：${err && err.message || err}`, true);
          }
        }
      }
    ]
  });
}

/** 导出（工具栏按钮 / 菜单）：有批注导出含批注 PDF；无批注原样导出 */
async function doExport() {
  if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
  const hasChanges = collectAnnotations().length || contentEditor.allEdits().length;
  if (!hasChanges) {
    // 无批注：原样导出当前 PDF
    const r = await window.onyx.exportPdf({
      bytes: new Uint8Array(state.originalBytes),
      suggestedName: state.fileName.replace(/\.pdf$/i, '') + '-copy.pdf',
      title: '导出 PDF'
    });
    if (r.ok) toast(`已导出：${r.path}`);
    else if (!r.canceled) toast(`导出失败：${r.error}`, true);
    return;
  }
  try {
    const { bytes, drawn } = await exportAnnotatedPdf(
      state.originalBytes, collectAnnotations(), state.fileName.replace(/\.pdf$/i, '') + '-annotated.pdf',
      contentEditor.allEdits());
    const r = await window.onyx.exportPdf({
      bytes,
      suggestedName: state.fileName.replace(/\.pdf$/i, '') + '-annotated.pdf',
      title: '导出含批注的 PDF'
    });
    if (r.ok) toast(`已导出 ${drawn} 处批注 → ${r.path}`);
    else if (!r.canceled) toast(`导出失败：${r.error}`, true);
  } catch (err) {
    console.error('导出失败:', err);
    toast(`导出失败：${err && err.message || err}`, true);
  }
}

/* ────── 文本采集（TXT/Word/Excel 共用） ────── */
/**
 * 结构化导出必须读取原始 PDF。
 * 内容编辑模式下 viewer.doc 是抹除文字后的背景文档，直接读取会导致导出内容为空。
 */
async function getExportDocument() {
  if (!viewer.doc) throw new Error('请先打开 PDF 文档');
  if (!state.originalBytes) throw new Error('原始 PDF 数据不可用');
  if (!contentEditor.enabled) return viewer.doc;
  contentEditor.commitDom();
  return contentEditor.getOriginalDoc();
}

async function collectPageLines() {
  const exportDocument = await getExportDocument();
  const pages = [];
  for (let p = 1; p <= state.numPages; p++) {
    const page = await exportDocument.getPage(p);
    const tc = await page.getTextContent();
    const lines = [];
    let cur = '';
    for (const it of tc.items) {
      cur += it.str;
      if (it.hasEOL) { lines.push(cur); cur = ''; }
    }
    if (cur.trim()) lines.push(cur);
    pages.push(lines);
  }
  return pages;
}

/** 构建文本导出 */
async function buildTxtExport() {
  const pages = await collectPageLines();
  const text = pages.map((lines, i) => `── 第 ${i + 1} 页 ──\n${lines.join('\n')}`).join('\n\n');
  return {
    bytes: new TextEncoder().encode(text),
    suggestedName: state.fileName.replace(/\.pdf$/i, '') + '.txt'
  };
}

/** 导出文本为 TXT */
async function doExportText() {
  if (!viewer.doc) { toast('请先打开 PDF 文档', true); return; }
  try {
    const { bytes, suggestedName } = await buildTxtExport();
    const r = await window.onyx.exportPdf({
      bytes, suggestedName,
      title: '导出文本',
      filters: [{ name: '文本文件', extensions: ['txt'] }]
    });
    if (r.ok) toast(`已导出文本：${r.path}`);
  } catch (err) {
    toast(`导出失败：${err && err.message || err}`, true);
  }
}

/**
 * 乱码行检测：PDF 缺 ToUnicode 时 pdf.js 返回 PUA/替换符 —— 无法还原，导出时保留原样视觉。
 */
function isGarbled(text) {
  const t = String(text);
  if (!t) return false;
  let bad = 0;
  for (const ch of t) {
    const c = ch.codePointAt(0);
    if (c === 0xFFFD || (c >= 0xE000 && c <= 0xF8FF) || (c >= 0x00C0 && c <= 0x00FF && !/[À-ÿ]/.test(ch))) bad++;
  }
  return bad / [...t].length > 0.12;
}

/** PDF 字体名 → 系统可用中文字体（导出/编辑共用） */
const FONT_FALLBACK_MAP = {
  'AlibabaPuHuiTi': '微软雅黑', 'AlibabaPuHuiTi-Bold': '微软雅黑', 'PuHuiTi': '微软雅黑',
  'SimSun': '宋体', 'SimHei': '黑体', 'KaiTi': '楷体', 'FangSong': '仿宋',
  'MicrosoftYaHei': '微软雅黑', 'MicrosoftYaHei-Bold': '微软雅黑',
  'PingFangSC': '微软雅黑', 'STZhongsong': '华文中宋', 'STXihei': '华文细黑',
  'STKaiti': '华文楷体', 'STSong': '华文宋体', 'DengXian': '等线',
  'SourceHanSansCN': '思源黑体', 'NotoSansSC': '思源黑体', 'MSung': '宋体'
};
/**
 * 跨页重复文字检测：同一文字出现在同位置的多页（PPT 母版残留，如被覆盖的"目录"）→ 标记 recurring。
 * recurring 行：不导出文本框、不抹除背景（原样保留视觉）。
 */
function markRecurringLines(allLines) {
  const N = allLines.length;
  if (N < 4) return;
  const count = new Map();
  for (const lines of allLines) {
    for (const l of lines) {
      const k = l.text + '@' + Math.round(l.x) + ',' + Math.round(l.y);
      count.set(k, (count.get(k) || 0) + 1);
    }
  }
  const threshold = Math.max(3, Math.ceil(N * 0.3));
  for (const lines of allLines) {
    for (const l of lines) {
      const k = l.text + '@' + Math.round(l.x) + ',' + Math.round(l.y);
      if ((count.get(k) || 0) >= threshold) l.recurring = true;
    }
  }
}

let _measureCtx = null;
/**
 * 字号自适应：映射字体通常比 PDF 嵌入字体宽，实测收缩以保证单行宽度 ≤ 原行宽（不换行不截断）。
 * @returns 缩放后的 pt 字号
 */
function fitFontSize(text, maxWpt, size, fontCssName, bold) {
  if (!_measureCtx) _measureCtx = document.createElement('canvas').getContext('2d');
  const px = (pt) => pt * 96 / 72;
  _measureCtx.font = `${bold ? '700 ' : ''}${px(size)}px "${fontCssName}", "Microsoft YaHei", sans-serif`;
  const w = _measureCtx.measureText(String(text)).width;
  if (w > px(maxWpt) && w > 0) {
    return Math.max(6, size * (px(maxWpt) / w));
  }
  return size;
}

function mapFontName(f) {
  const main = String(f || '').split(',')[0].replace(/["']/g, '').trim();
  if (!main) return '微软雅黑';
  if (FONT_FALLBACK_MAP[main]) return FONT_FALLBACK_MAP[main];
  // 去子集前缀再试
  const stripped = main.replace(/^[A-Z]{6}\+/, '');
  if (FONT_FALLBACK_MAP[stripped]) return FONT_FALLBACK_MAP[stripped];
  if (/^[A-Za-z0-9_\-]+$/.test(stripped) && /[\u4e00-\u9fff]/.test('')) return stripped;
  return stripped || '微软雅黑';
}

/** 渲染一页为 PNG dataURL（指定倍率） */
async function renderPageToDataUrl(p, scale = 2) {
  const page = await viewer.doc.getPage(p);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  return { dataUrl: canvas.toDataURL('image/png'), w: canvas.width, h: canvas.height };
}

/**
 * 渲染一页并抹除原文字（V10）：整页 2x 渲染后，
 * 对每个文本行框用「框外上下采样带的中位数背景色」填充 —— 背景图形完整保留、原文字消失，
 * 文字由可编辑文本框承载 → 保真 + 可编辑 + 无双重文字。
 */
async function renderPageErased(p, scale = 2, preLines = null) {
  const page = await viewer.doc.getPage(p);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  await page.render({ canvasContext: ctx, viewport }).promise;
  const ed = await extractPageLines(viewer.doc, p);
  if (preLines) { ed.lines = preLines; }
  // V12.1：整页单次 ImageData，纯数组操作
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const D = img.data, W = canvas.width, H = canvas.height;
  const getPx = (x, y) => {
    const i = (y * W + x) * 4;
    return [D[i], D[i + 1], D[i + 2]];
  };
  const setPx = (x, y, c) => {
    const i = (y * W + x) * 4;
    D[i] = c[0]; D[i + 1] = c[1]; D[i + 2] = c[2]; D[i + 3] = 255;
  };
  const median3 = (arr) => {
    if (!arr.length) return [255, 255, 255];
    const m = (k) => {
      const v = arr.map(a => a[k]).sort((a, b) => a - b);
      return v[Math.floor(v.length / 2)];
    };
    return [m(0), m(1), m(2)];
  };

  for (const line of ed.lines) {
    if (isGarbled(line.text) || line.recurring) continue; // 乱码/母版残留：原样保留
    const lineH = Math.max(Math.min(line.h, line.size * 1.3), line.size * 0.85);
    let x0 = Math.max(0, Math.floor((line.x - 1.5) * scale));
    let x1 = Math.min(W - 1, Math.ceil((line.x + line.w + 1.5) * scale));
    const yT = Math.max(0, Math.floor((line.y - 1.5) * scale));
    const yB = Math.min(H - 1, Math.ceil((line.y + lineH + 1.5) * scale));
    // V12.3.7：墨迹延伸——提取行宽可能小于实际渲染文字宽（原文行尾残留的根源）。
    // 从提取矩形边界向外按列扫描连通墨迹，扩展擦除范围；词间隙超过 14pt 或达 30pt 上限即止（不误伤同行其他元素）。
    {
      const bh0 = yB - yT + 1;
      const yS = [yT + Math.round(bh0 * 0.3), yT + Math.round(bh0 * 0.55), yT + Math.round(bh0 * 0.8)];
      const refX = x0 > 1 ? Math.max(0, x0 - 3) : Math.min(W - 1, x1 + 3);
      const ref = getPx(refX, Math.max(0, Math.min(H - 1, yT + Math.round(bh0 / 2))));
      const hasInk = (x) => {
        for (const yy of yS) {
          const p = getPx(x, Math.max(0, Math.min(H - 1, yy)));
          if ((p[0] - ref[0]) ** 2 + (p[1] - ref[1]) ** 2 + (p[2] - ref[2]) ** 2 > 2500) return true;
        }
        return false;
      };
      const lim = Math.round(30 * scale), gapMax = Math.round(14 * scale);
      let gap = 0, xr = x1;
      for (let x = x1 + 2; x <= Math.min(W - 1, x1 + lim); x++) {
        if (hasInk(x)) { xr = x; gap = 0; } else if (++gap > gapMax) break;
      }
      x1 = Math.min(W - 1, xr + Math.round(1.5 * scale));
      gap = 0; let xl = x0;
      for (let x = x0 - 2; x >= Math.max(0, x0 - lim); x--) {
        if (hasInk(x)) { xl = x; gap = 0; } else if (++gap > gapMax) break;
      }
      x0 = Math.max(0, xl - Math.round(1.5 * scale));
    }
    const bw = x1 - x0 + 1, bh = yB - yT + 1;
    if (bw < 2 || bh < 2) continue;

    // 1) 每列上下背景采样（中值抗字形）→ 背景上下色
    const colAbove = new Array(bw), colBelow = new Array(bw);
    const yA0 = Math.max(0, yT - Math.round(5 * scale)), yA1 = Math.max(0, yT - Math.round(1.5 * scale));
    const yB0 = Math.min(H - 1, yB + Math.round(1.5 * scale)), yB1 = Math.min(H - 1, yB + Math.round(5 * scale));
    for (let x = 0; x < bw; x++) {
      const sa = [], sb = [];
      for (let y = yA0; y <= yA1; y += 1) sa.push(getPx(x0 + x, y));
      for (let y = yB0; y <= yB1; y += 1) sb.push(getPx(x0 + x, y));
      colAbove[x] = median3(sa);
      colBelow[x] = median3(sb);
    }
    const dist2 = (a, b) => (a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2;

    // 2) 墨迹分析：与该列背景（上下插值）差异 > 60 的像素 = 墨迹
    //    → 主色（量化众数）= 文字真实颜色；墨迹 x 范围 = 实际文字范围
    let inkCount = 0;
    const hist = new Map();
    let inkMinX = bw, inkMaxX = -1;
    for (let y = yT; y <= yB; y++) {
      for (let x = 0; x < bw; x++) {
        const p = getPx(x0 + x, y);
        const t = (y - yT) / Math.max(1, bh - 1);
        const bgp = [
          colAbove[x][0] + (colBelow[x][0] - colAbove[x][0]) * t,
          colAbove[x][1] + (colBelow[x][1] - colAbove[x][1]) * t,
          colAbove[x][2] + (colBelow[x][2] - colAbove[x][2]) * t
        ];
        if (dist2(p, bgp) > 3600) { // 距离 > 60
          inkCount++;
          if (x < inkMinX) inkMinX = x;
          if (x > inkMaxX) inkMaxX = x;
          const key = ((p[0] >> 4) << 8) | ((p[1] >> 4) << 4) | (p[2] >> 4);
          hist.set(key, (hist.get(key) || 0) + 1);
        }
      }
    }
    line.hidden = inkCount < bw * bh * 0.004; // 几乎无墨迹 = 被覆盖不可见
    if (line.hidden) continue;
    // 墨迹主色 → 文字真实颜色（修复白字/彩字被提黑的问题）
    if (inkCount > 8 && hist.size) {
      let bestKey = null, bestN = 0;
      for (const [k, n] of hist) if (n > bestN) { bestN = n; bestKey = k; }
      if (bestKey !== null && bestN > inkCount * 0.3) {
        const r = ((bestKey >> 8) & 15) * 17, g = ((bestKey >> 4) & 15) * 17, b = (bestKey & 15) * 17;
        line.inkColor = '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
      }
    }
    // 3) 逐列渐变插值填充（保色块渐变），水平收紧到墨迹范围（防破坏旁侧装饰）
    const ix0 = inkMaxX >= 0 ? Math.max(0, inkMinX - Math.round(1.5 * scale)) : 0;
    const ix1 = inkMaxX >= 0 ? Math.min(bw - 1, inkMaxX + Math.round(1.5 * scale)) : bw - 1;
    for (let x = ix0; x <= ix1; x++) {
      const ca = colAbove[x], cb = colBelow[x];
      for (let y = yT; y <= yB; y++) {
        const t = (y - yT) / Math.max(1, bh - 1);
        setPx(x0 + x, y, [
          ca[0] + (cb[0] - ca[0]) * t,
          ca[1] + (cb[1] - ca[1]) * t,
          ca[2] + (cb[2] - ca[2]) * t
        ]);
      }
    }
    // 墨迹几何（pt）→ 文本框定位/宽度基准
    line.inkX = line.x + inkMinX / scale;
    line.inkW = (inkMaxX - inkMinX + 1) / scale;
  }
  ctx.putImageData(img, 0, 0);
  return { dataUrl: canvas.toDataURL('image/png'), lines: ed.lines, paras: ed.paras, width: ed.width, height: ed.height };
}

/** 构建 PowerPoint 导出（.pptx，每页整图幻灯片 → 版式 100% 保留） */
async function buildPptxExport() {
  const pagePngs = [];
  for (let p = 1; p <= state.numPages; p++) {
    toast(`渲染页面 ${p}/${state.numPages}…`);
    pagePngs.push(await renderPageToDataUrl(p, 2));
  }
  return {
    bytes: buildPptx(pagePngs),
    suggestedName: state.fileName.replace(/\.pdf$/i, '') + '.pptx'
  };
}

async function doExportPptx() {
  if (!viewer.doc) { toast('请先打开 PDF 文档', true); return; }
  try {
    const { bytes, suggestedName } = await buildPptxExport();
    const r = await window.onyx.exportPdf({
      bytes, suggestedName,
      title: '导出 PowerPoint 演示文稿',
      filters: [{ name: 'PowerPoint 演示文稿', extensions: ['pptx'] }]
    });
    if (r.ok) toast(`已导出 PPT：${r.path}`);
  } catch (err) {
    toast(`导出失败：${err && err.message || err}`, true);
  }
}

/** 构建 Excel 导出（.xlsx，页/行文本表格） */
async function buildXlsxExport() {
  return {
    bytes: buildXlsxLayout(await buildParagraphExportPages()),
    suggestedName: state.fileName.replace(/\.pdf$/i, '') + '.xlsx'
  };
}

async function doExportXlsx() {
  if (!viewer.doc) { toast('请先打开 PDF 文档', true); return; }
  try {
    const { bytes, suggestedName } = await buildXlsxExport();
    const r = await window.onyx.exportPdf({
      bytes, suggestedName,
      title: '导出 Excel 工作簿',
      filters: [{ name: 'Excel 工作簿', extensions: ['xlsx'] }]
    });
    if (r.ok) toast(`已导出 Excel：${r.path}`);
  } catch (err) {
    toast(`导出失败：${err && err.message || err}`, true);
  }
}

/** 收集全部页面的行级文本（可编辑导出共用），并应用元素编辑 */
async function collectEditPages() {
  const exportDocument = await getExportDocument();
  const pages = [];
  for (let p = 1; p <= state.numPages; p++) {
    toast(`解析页面 ${p}/${state.numPages}…`);
    pages.push(await extractPageLines(exportDocument, p));
  }
  // 元素编辑（移动/删除）应用到导出数据
  const withObjs = await Promise.all(pages.map(async (pg, i) => {
    let images = [];
    try { images = await extractPageImages(exportDocument, i + 1); } catch (_) { /* noop */ }
    let rects = [];
    try { rects = await extractPageRects(exportDocument, i + 1); } catch (_) { /* noop */ }
    return { lines: pg.lines, images, rects, width: pg.width, height: pg.height };
  }));
  return withObjs;
}

/** PPT 可编辑（.pptx，行→文本框，位置/字号按 PDF 实测） */
/** PPT 可编辑 v2（.pptx，段落文本框 + 原文图形背景，版式还原） */
async function buildPptxEditableExportV2() {
  const pages = await buildParagraphExportPages();
  await flattenParagraphExportBackgrounds(pages);
  return {
    bytes: buildPptxEditable(pages),
    suggestedName: state.fileName.replace(/\.pdf$/i, '') + '.pptx'
  };
}

async function flattenParagraphExportBackgrounds(pages) {
  for (const page of pages) {
    if (!page.bgDataUrl) continue;
    const canvas = document.createElement('canvas');
    const background = new Image();
    const patch = new Image();
    await Promise.all([
      new Promise((resolve, reject) => {
        background.onload = resolve;
        background.onerror = () => reject(new Error('PPTX 背景解码失败'));
        background.src = page.bgDataUrl;
      }),
      page.patchDataUrl ? new Promise((resolve, reject) => {
        patch.onload = resolve;
        patch.onerror = () => reject(new Error('PPTX 文字补丁解码失败'));
        patch.src = page.patchDataUrl;
      }) : Promise.resolve()
    ]);
    canvas.width = background.naturalWidth;
    canvas.height = background.naturalHeight;
    const context = canvas.getContext('2d');
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(background, 0, 0);
    if (page.patchDataUrl) context.drawImage(patch, 0, 0);
    if (!page.patchDataUrl && page.coverShapes?.length) {
      const backgroundScale = canvas.width / Number(page.width);
      for (const cover of page.coverShapes) {
        context.fillStyle = cover.color || '#FFFFFF';
        context.fillRect(
          Number(cover.x) * backgroundScale,
          Number(cover.y) * backgroundScale,
          Number(cover.width) * backgroundScale,
          Number(cover.height) * backgroundScale
        );
      }
    }
    page.bgDataUrl = canvas.toDataURL('image/png');
    delete page.patchDataUrl;
    delete page.coverShapes;
  }
}

async function buildPptxEditableExport() {
  const pageData = await collectEditPages();
  return {
    bytes: buildPptxEditable(pageData),
    suggestedName: state.fileName.replace(/\.pdf$/i, '') + '.pptx'
  };
}

/** Word 版式还原 + 可编辑合一：原文图形背景 + 段落框架文本（字体/字号按 PDF 实测） */
async function buildWordLayoutExport() {
  const pages = await buildParagraphExportPages();
  return {
    bytes: buildDocxParagraphLayout(pages),
    suggestedName: state.fileName.replace(/\.pdf$/i, '') + '.docx'
  };
}

/** 与内容编辑器共享同一个段落模型：原文只删文字流，不渲染矩形涂抹层。 */
async function buildParagraphExportPages() {
  if (!viewer.doc) throw new Error('请先打开 PDF 文档');
  if (!state.originalBytes) throw new Error('原始 PDF 数据不可用');

  if (contentEditor.enabled) contentEditor.commitDom();
  const editsByPage = new Map();
  for (const edit of contentEditor.allEdits()) {
    if (!editsByPage.has(edit.page)) editsByPage.set(edit.page, []);
    editsByPage.get(edit.page).push(edit);
  }

  const sourceDocument = await pdfjsLib.getDocument({
    data: state.originalBytes.slice(0),
    cMapUrl: '../vendor/cmaps/',
    cMapPacked: true,
    enableXfa: false,
    fontExtraProperties: true
  }).promise;
  const embeddedFonts = await extractEmbeddedFonts(
    (await contentEditor.editableSourceBytes()).slice(0), window.PDFLib, window.fontkit
  );
  const libDocument = await window.PDFLib.PDFDocument.load(
    await contentEditor.editableSourceBytes(), { ignoreEncryption: true, updateMetadata: false }
  );
  const pages = [];
  let textFreeDocument = null;

  try {
    for (let pageNumber = 1; pageNumber <= state.numPages; pageNumber++) {
      toast(`解析段落 ${pageNumber}/${state.numPages}…`);
      const model = await extractContentPage(sourceDocument, pageNumber, { libDocument });
      const originalExportLines = new Map(model.blocks.map(block => [
        block.id,
        (block.exportLines || []).map(line => ({ ...line }))
      ]));
      applySavedEdits(model, editsByPage.get(pageNumber) || []);
      const paragraphs = model.blocks.filter(block => block.runs?.some(run => run.text));

      const imageList = contentEditor.images.get(pageNumber) || [];
      pages.push({
        paragraphs: paragraphs.map(block => ({
          ...block,
          runs: (block.runs || []).map(run => ({ ...run })),
          sourceInpaintLines: originalExportLines.get(block.id) || [],
          sourceOps: [...(block.sourceOps || [])],
          top: paragraphTopPt(model, block),
          frameHeight: Math.max(
            Number(block.height) || 0,
            (((Number(block.lineCount) || 1) - 1) * (Number(block.lineHeight) || (Number(block.size) || 12) * 1.22)) +
              (Number(block.size) || 12) * 1.22,
            Number(block.lineHeight) || 0
          )
        })),
        images: imageList.map(image => ({ ...image })),
        width: model.width,
        height: model.height,
        embeddedFonts,
        pageNumber
      });
    }

    const sourceOperationsByPage = new Map(pages.map(page => [
      page.pageNumber,
      [...new Set(page.paragraphs.flatMap(paragraph => paragraph.sourceOps || []))]
        .filter(Number.isInteger)
        .sort((left, right) => left - right)
    ]));
    textFreeDocument = await pdfjsLib.getDocument({
      data: await createTextFreeDocumentBytes(
        state.originalBytes.slice(0), sourceOperationsByPage
      ),
      cMapUrl: '../vendor/cmaps/',
      cMapPacked: true,
      enableXfa: false,
      fontExtraProperties: true
    }).promise;

    for (const page of pages) {
      toast(`生成背景 ${page.pageNumber}/${state.numPages}…`);
      page.bgDataUrl = await renderAndInpaintPage(
        sourceDocument, page.pageNumber, 2, page.paragraphs, textFreeDocument
      );
      page.patchDataUrl = null;
      page.coverShapes = [];
    }
  } finally {
    await sourceDocument.cleanup?.();
    await textFreeDocument.cleanup?.();
  }

  return pages;
}

/**
 * 原文渲染保留图形和轮廓字；可见正文由局部擦除移除。
 */
async function createTextFreeDocumentBytes(sourceBytes, sourceOperationsByPage) {
  const pdfDocument = await window.PDFLib.PDFDocument.load(sourceBytes, {
    ignoreEncryption: true,
    updateMetadata: false
  });
  for (let pageIndex = 0; pageIndex < pdfDocument.getPageCount(); pageIndex++) {
    removeTextOperations(
      pdfDocument,
      pageIndex,
      sourceOperationsByPage.get(pageIndex + 1) || []
    );
  }
  return pdfDocument.save({ useObjectStreams: false });
}

/**
 * 渲染原文背景，并只把“原文与无文字渲染”差异处的背景像素写入透明补丁。
 * 相比矩形色块，这种方法不会破坏渐变、红底、表格和窄标签。
 */
async function renderPageBackgroundPatch(
  originalDocument, textFreeDocument, pageNumber, scale = 2, paragraphs = []
) {
  const renderCanvas = async (pdfDocument) => {
    const page = await pdfDocument.getPage(pageNumber);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.floor(viewport.width));
    canvas.height = Math.max(1, Math.floor(viewport.height));
    const context = canvas.getContext('2d', { willReadFrequently: true });
    await page.render({ canvasContext: context, viewport }).promise;
    return { canvas, context, viewport };
  };

  const textFreeRender = await renderCanvas(textFreeDocument);
  const originalRender = await renderCanvas(originalDocument);
  const canvas = originalRender.canvas;
  const viewport = originalRender.viewport;
  const imageWidth = canvas.width;
  const imageHeight = canvas.height;
  if (
    textFreeRender.canvas.width !== imageWidth ||
    textFreeRender.canvas.height !== imageHeight
  ) {
    throw new Error('PDF 原文与背景渲染尺寸不一致');
  }

  const originalImage = originalRender.context.getImageData(0, 0, imageWidth, imageHeight);
  const textFreeImage = textFreeRender.context.getImageData(0, 0, imageWidth, imageHeight);
  const originalPixelsForCheck = originalImage.data;
  const textFreePixelsForCheck = textFreeImage.data;
  const sourceLinesFor = paragraph => Array.isArray(paragraph.sourceInpaintLines) && paragraph.sourceInpaintLines.length
    ? paragraph.sourceInpaintLines
    : paragraph.exportLines || [];
  const parseRunColor = run => {
    const hex = String(run?.color || '#000000').replace('#', '');
    const full = hex.length === 3 ? hex.split('').map(part => part + part).join('') : hex.padEnd(6, '0');
    return [0, 2, 4].map(offset => parseInt(full.slice(offset, offset + 2), 16) || 0);
  };

  for (const paragraph of paragraphs || []) {
    let hasMeasuredLine = false;
    for (const line of sourceLinesFor(paragraph)) {
      const runs = (line.runs || []).filter(run => String(run.text ?? '').trim());
      const validRuns = runs.filter(run =>
        [run.x, run.baseline, run.width, run.size].every(value => Number.isFinite(Number(value))));
      if (!validRuns.length) continue;

      const left = Math.min(...validRuns.map(run => Number(run.x)));
      const right = Math.max(...validRuns.map(run => Number(run.x) + Number(run.width)));
      const size = Math.max(...validRuns.map(run => Number(run.size) || 12));
      const baseline = validRuns[0].baseline;
      const horizontalPadding = Math.max(1, size * 0.18);
      const ascent = Math.max(1, size * 1.28);
      const descent = Math.max(1, size * 0.48);
      const x0 = Math.max(0, Math.floor((left - horizontalPadding) * scale));
      const x1 = Math.min(imageWidth - 1, Math.ceil((right + horizontalPadding) * scale));
      const y0 = Math.max(0, Math.floor(viewport.height - (baseline + ascent) * scale));
      const y1 = Math.min(imageHeight - 1, Math.ceil(viewport.height - (baseline - descent) * scale));
      if (x1 <= x0 || y1 <= y0) continue;

      let diffCount = 0;
      let matchedColorCount = 0;
      let minX = Number.MAX_SAFE_INTEGER;
      let minY = Number.MAX_SAFE_INTEGER;
      let maxX = -1;
      let maxY = -1;
      const declaredColors = validRuns.map(parseRunColor);
      for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
          const offset = (y * imageWidth + x) * 4;
          const difference = Math.max(
            Math.abs(originalPixelsForCheck[offset] - textFreePixelsForCheck[offset]),
            Math.abs(originalPixelsForCheck[offset + 1] - textFreePixelsForCheck[offset + 1]),
            Math.abs(originalPixelsForCheck[offset + 2] - textFreePixelsForCheck[offset + 2])
          );
          if (difference <= 10) continue;
          diffCount++;
          minX = Math.min(minX, x);
          minY = Math.min(minY, y);
          maxX = Math.max(maxX, x);
          maxY = Math.max(maxY, y);
          if (difference >= 28 && declaredColors.some(color => Math.hypot(
            originalPixelsForCheck[offset] - color[0],
            originalPixelsForCheck[offset + 1] - color[1],
            originalPixelsForCheck[offset + 2] - color[2]
          ) <= 96)) {
            matchedColorCount++;
          }
        }
      }

      const characterCount = Math.max(1, validRuns.reduce((total, run) =>
        total + String(run.text || '').replace(/\s+/g, '').length, 0));
      const expectedArea = characterCount * size * size * scale * scale * 0.16;
      const boundsMargin = Math.max(2, size * 0.22 * scale);
      let lineReliable = diffCount > 0 &&
        diffCount >= expectedArea * 0.015 &&
        diffCount <= expectedArea * 16 &&
        matchedColorCount >= diffCount * 0.03 &&
        minX >= x0 - boundsMargin &&
        minY >= y0 - boundsMargin &&
        maxX <= x1 + boundsMargin &&
        maxY <= y1 + boundsMargin;
      paragraph.reliable = hasMeasuredLine ? (paragraph.reliable && lineReliable) : lineReliable;
      paragraph.renderedSourcePixelCount =
        (paragraph.renderedSourcePixelCount || 0) + diffCount;
      hasMeasuredLine = true;
    }
    if (!hasMeasuredLine) paragraph.reliable = false;
  }

  if (!paragraphs.some(paragraph => paragraph.reliable === false)) {
    return {
      dataUrl: textFreeRender.canvas.toDataURL('image/png'),
      patchDataUrl: null,
      covers: [],
      changedPixelCount: 0
    };
  }

  const inpaintMask = new Uint8Array(imageWidth * imageHeight);
  let inpaintBoxCount = 0;
  for (const paragraph of paragraphs || []) {
    if (paragraph.reliable === false) continue;
    for (const line of sourceLinesFor(paragraph)) {
      const runs = (line.runs || []).filter(run => String(run.text ?? '').trim());
      const validRuns = runs.filter(run =>
        [run.x, run.baseline, run.width, run.size].every(value => Number.isFinite(Number(value))));
      if (!validRuns.length) continue;

      let left = Math.min(...validRuns.map(run => Number(run.x)));
      let right = Math.max(...validRuns.map(run => Number(run.x) + Number(run.width)));
      if ([line.x, line.width].every(value => Number.isFinite(Number(value)))) {
        left = Math.min(left, Number(line.x));
        right = Math.max(right, Number(line.x) + Number(line.width));
      }
      const size = Math.max(...validRuns.map(run => Number(run.size) || 12));
      const baseline = validRuns[0].baseline;
      const ascent = size * 1.05;
      const descent = size * 0.32;
      const horizontalPadding = Math.max(1, size * 0.14);
      const x0 = Math.max(0, Math.floor((left - horizontalPadding) * scale));
      const x1 = Math.min(imageWidth - 1, Math.ceil((right + horizontalPadding) * scale));
      const y0 = Math.max(0, Math.floor(viewport.height - (baseline + ascent) * scale));
      const y1 = Math.min(imageHeight - 1, Math.ceil(viewport.height - (baseline - descent) * scale));
      if (x1 <= x0 || y1 <= y0) continue;

      for (let y = y0; y <= y1; y++) {
        inpaintMask.fill(1, y * imageWidth + x0, y * imageWidth + x1 + 1);
      }
      inpaintBoxCount++;
    }
  }

  const patchCanvas = document.createElement('canvas');
  patchCanvas.width = imageWidth;
  patchCanvas.height = imageHeight;
  const patchContext = patchCanvas.getContext('2d', { willReadFrequently: true });
  const patchImage = patchContext.createImageData(imageWidth, imageHeight);
  const originalPixels = originalImage.data;
  const textFreePixels = textFreeImage.data;
  const patchPixels = patchImage.data;
  let changedPixelCount = 0;

  for (let offset = 0; offset < originalPixels.length; offset += 4) {
    if (!inpaintMask[offset / 4]) continue;
    const difference = Math.max(
      Math.abs(originalPixels[offset] - textFreePixels[offset]),
      Math.abs(originalPixels[offset + 1] - textFreePixels[offset + 1]),
      Math.abs(originalPixels[offset + 2] - textFreePixels[offset + 2])
    );
    if (difference <= 6) continue;
    patchPixels[offset + 3] = 255;
    changedPixelCount++;
  }
  const seedAlpha = new Uint8Array(imageWidth * imageHeight);
  for (let offset = 0, pixelIndex = 0; offset < patchPixels.length; offset += 4, pixelIndex++) {
    seedAlpha[pixelIndex] = patchPixels[offset + 3];
  }
  const patchRadius = Math.max(6, Math.round(scale * 2.5));
  for (let y = 0; y < imageHeight; y++) {
    for (let x = 0; x < imageWidth; x++) {
      const pixelIndex = y * imageWidth + x;
      let shouldExpand = false;
      for (let deltaY = -patchRadius; deltaY <= patchRadius && !shouldExpand; deltaY++) {
        const sourceY = y + deltaY;
        if (sourceY < 0 || sourceY >= imageHeight) continue;
        for (let deltaX = -patchRadius; deltaX <= patchRadius; deltaX++) {
          const sourceX = x + deltaX;
          if (sourceX < 0 || sourceX >= imageWidth) continue;
          if (seedAlpha[sourceY * imageWidth + sourceX]) {
            shouldExpand = true;
            break;
          }
        }
      }
      if (!shouldExpand) continue;
      if (!inpaintMask[pixelIndex]) continue;
      const offset = (y * imageWidth + x) * 4;
      patchPixels[offset + 3] = 255;
    }
  }

  const patchMask = new Uint8Array(imageWidth * imageHeight);
  const packedColors = new Uint32Array(imageWidth * imageHeight);
  for (let pixelIndex = 0; pixelIndex < patchMask.length; pixelIndex++) {
    patchMask[pixelIndex] = patchPixels[pixelIndex * 4 + 3] ? 1 : 0;
    const offset = pixelIndex * 4;
    packedColors[pixelIndex] =
      (255 << 24) |
      (originalPixels[offset + 2] << 16) |
      (originalPixels[offset + 1] << 8) |
      originalPixels[offset];
  }
  const createNearestField = () => ({
    colors: new Uint32Array(imageWidth * imageHeight),
    distances: new Uint16Array(imageWidth * imageHeight).fill(65535)
  });
  const nearestLeft = createNearestField();
  const nearestRight = createNearestField();
  const nearestTop = createNearestField();
  const nearestBottom = createNearestField();
  for (let y = 0; y < imageHeight; y++) {
    const row = y * imageWidth;
    for (let x = 0; x < imageWidth; x++) {
      const pixelIndex = row + x;
      const previous = x ? pixelIndex - 1 : -1;
      if (!patchMask[pixelIndex]) {
        nearestLeft.colors[pixelIndex] = packedColors[pixelIndex];
        nearestLeft.distances[pixelIndex] = 0;
      } else if (previous >= 0) {
        nearestLeft.colors[pixelIndex] = nearestLeft.colors[previous];
        nearestLeft.distances[pixelIndex] = Math.min(65535, nearestLeft.distances[previous] + 1);
      }
      const rightIndex = row + (imageWidth - 1 - x);
      const rightPrevious = x ? rightIndex + 1 : -1;
      if (!patchMask[rightIndex]) {
        nearestRight.colors[rightIndex] = packedColors[rightIndex];
        nearestRight.distances[rightIndex] = 0;
      } else if (rightPrevious < row + imageWidth) {
        nearestRight.colors[rightIndex] = nearestRight.colors[rightPrevious];
        nearestRight.distances[rightIndex] = Math.min(65535, nearestRight.distances[rightPrevious] + 1);
      }
    }
  }
  for (let x = 0; x < imageWidth; x++) {
    for (let y = 0; y < imageHeight; y++) {
      const pixelIndex = y * imageWidth + x;
      const previous = y ? pixelIndex - imageWidth : -1;
      if (!patchMask[pixelIndex]) {
        nearestTop.colors[pixelIndex] = packedColors[pixelIndex];
        nearestTop.distances[pixelIndex] = 0;
      } else if (previous >= 0) {
        nearestTop.colors[pixelIndex] = nearestTop.colors[previous];
        nearestTop.distances[pixelIndex] = Math.min(65535, nearestTop.distances[previous] + 1);
      }
      const bottomIndex = (imageHeight - 1 - y) * imageWidth + x;
      const bottomPrevious = y ? bottomIndex + imageWidth : -1;
      if (!patchMask[bottomIndex]) {
        nearestBottom.colors[bottomIndex] = packedColors[bottomIndex];
        nearestBottom.distances[bottomIndex] = 0;
      } else if (bottomPrevious < imageWidth * imageHeight) {
        nearestBottom.colors[bottomIndex] = nearestBottom.colors[bottomPrevious];
        nearestBottom.distances[bottomIndex] = Math.min(65535, nearestBottom.distances[bottomPrevious] + 1);
      }
    }
  }
  for (let pixelIndex = 0; pixelIndex < patchMask.length; pixelIndex++) {
    if (!patchMask[pixelIndex]) continue;
    const candidates = [
      { color: nearestLeft.colors[pixelIndex], distance: nearestLeft.distances[pixelIndex] },
      { color: nearestRight.colors[pixelIndex], distance: nearestRight.distances[pixelIndex] },
      { color: nearestTop.colors[pixelIndex], distance: nearestTop.distances[pixelIndex] },
      { color: nearestBottom.colors[pixelIndex], distance: nearestBottom.distances[pixelIndex] }
    ].filter(candidate => candidate.distance < 65535);
    if (!candidates.length) continue;
    const nearestDistance = Math.min(...candidates.map(candidate => candidate.distance));
    const matched = candidates.filter(candidate => candidate.distance <= nearestDistance + 1);
    let red = 0;
    let green = 0;
    let blue = 0;
    let weightTotal = 0;
    for (const candidate of matched) {
      const weight = 1 / (candidate.distance + 1);
      red += (candidate.color & 255) * weight;
      green += ((candidate.color >> 8) & 255) * weight;
      blue += ((candidate.color >> 16) & 255) * weight;
      weightTotal += weight;
    }
    const offset = pixelIndex * 4;
    patchPixels[offset] = red / weightTotal;
    patchPixels[offset + 1] = green / weightTotal;
    patchPixels[offset + 2] = blue / weightTotal;
  }
  patchContext.putImageData(patchImage, 0, 0);

  return {
    dataUrl: canvas.toDataURL('image/png'),
    patchDataUrl: patchCanvas.toDataURL('image/png'),
    covers: [],
    changedPixelCount
  };
}

/**
 * 渲染不含 PDF 文本操作的页面背景，保留矢量图形、渐变和图片。
 */
async function renderAndInpaintPage(originalDocument, pageNumber, scale = 2, paragraphs = [], cleanDocument = null) {
  const page = await originalDocument.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const context = canvas.getContext('2d', { willReadFrequently: true });
  await page.render({ canvasContext: context, viewport }).promise;

  const exportRegions = [];
  for (const paragraph of paragraphs || []) {
    const sourceLines = Array.isArray(paragraph.sourceInpaintLines) && paragraph.sourceInpaintLines.length
      ? paragraph.sourceInpaintLines
      : paragraph.exportLines || [];
    for (const line of sourceLines) {
      if (!line?.runs?.some(run => String(run.text ?? '').trim())) continue;
      const trimmedRuns = line.runs.filter(run => String(run.text ?? '').trim());
      const validRuns = trimmedRuns.filter(run =>
        [run.x, run.baseline, run.width, run.size].every(value => Number.isFinite(Number(value))));
      if (validRuns.length === trimmedRuns.length) {
        for (const run of validRuns) {
          exportRegions.push({
            x: Number(run.x),
            baseline: Number(run.baseline),
            width: Math.max(Number(run.width), 1),
            size: Number(run.size),
            color: run.color || '#000000',
            sourceRun: run,
            text: run.text || ''
          });
        }
      } else {
        const fallbackRun = trimmedRuns[0];
        exportRegions.push({
          x: Number(line.x ?? paragraph.x) || 0,
          baseline: Number(line.baseline ?? paragraph.baseline) || 0,
          width: Number(line.width ?? paragraph.width) || 12,
          size: Number(line.size) || Number(paragraph.size) || Number(fallbackRun?.size) || 12,
          color: fallbackRun?.color || paragraph.runs?.[0]?.color || '#000000',
          sourceRun: fallbackRun,
          text: line.text || fallbackRun?.text || ''
        });
      }
    }
  }

  const inpaintStats = [];
  let inpaintRegions = 0;
  if (cleanDocument && exportRegions.length) {
    try {
      const cleanPage = await cleanDocument.getPage(pageNumber);
      const cleanViewport = cleanPage.getViewport({ scale });
      const cleanCanvas = document.createElement('canvas');
      cleanCanvas.width = canvas.width;
      cleanCanvas.height = canvas.height;
      const cleanContext = cleanCanvas.getContext('2d', { willReadFrequently: true });
      await cleanPage.render({ canvasContext: cleanContext, viewport: cleanViewport }).promise;

      for (const region of exportRegions) {
        const size = Math.max(4, Number(region.size) || 12);
        const x0 = Math.max(0, Math.floor(region.x * scale - size * 0.08 * scale));
        const x1 = Math.min(canvas.width - 1, Math.ceil((region.x + region.width) * scale + size * 0.08 * scale));
        const y0 = Math.max(0, Math.floor(viewport.height - (region.baseline + size * 0.98) * scale));
        const y1 = Math.min(canvas.height - 1, Math.ceil(viewport.height - (region.baseline - size * 0.26) * scale));
        const boxWidth = x1 - x0 + 1;
        const boxHeight = y1 - y0 + 1;
        if (boxWidth <= 2 || boxHeight <= 2) continue;
        context.drawImage(cleanCanvas, x0, y0, boxWidth, boxHeight, x0, y0, boxWidth, boxHeight);
      }
      return canvas.toDataURL('image/jpeg', 0.88);
    } catch (error) {
      console.warn('矢量背景重建失败，改用局部像素修复。', error);
    }
  }

  if (exportRegions.length) {
    const width = canvas.width;
    const height = canvas.height;
    const sourceImage = context.getImageData(0, 0, width, height);
    const sourcePixel = (x, y) => {
      const offset = (y * width + x) * 4;
      return [
        sourceImage.data[offset],
        sourceImage.data[offset + 1],
        sourceImage.data[offset + 2]
      ];
    };
    const dominantColor = (samples) => {
      if (!samples.length) return [255, 255, 255];
      const buckets = new Map();
      samples.forEach(color => {
        const key = ((color[0] >> 4) << 8) | ((color[1] >> 4) << 4) | (color[2] >> 4);
        const bucket = buckets.get(key) || { count: 0, sum: [0, 0, 0] };
        bucket.count++;
        bucket.sum[0] += color[0];
        bucket.sum[1] += color[1];
        bucket.sum[2] += color[2];
        buckets.set(key, bucket);
      });
      const center = [...buckets.values()].sort((left, right) => right.count - left.count)[0];
      const matched = samples.filter(color =>
        Math.hypot(color[0] - center.sum[0] / center.count,
          color[1] - center.sum[1] / center.count,
          color[2] - center.sum[2] / center.count) <= 64);
      const source = matched.length ? matched : samples;
      return [0, 1, 2].map(index => source.reduce((sum, color) => sum + color[index], 0) / source.length);
    };

    const targetColor = (value) => {
      const hex = String(value || '#000000').replace('#', '');
      const full = hex.length === 3 ? hex.split('').map(part => part + part).join('') : hex.padEnd(6, '0');
      return [0, 2, 4].map(offset => parseInt(full.slice(offset, offset + 2), 16) || 0);
    };
    const colorHex = (value) => value.map(channel =>
      Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, '0')
    ).join('').toUpperCase();

    const nearestKnown = (data, mask, x, y, deltaX, deltaY, boxWidth, boxHeight, x0, y0) => {
      let currentX = x + deltaX;
      let currentY = y + deltaY;
      let distance = 1;
      while (currentX >= 0 && currentX < width && currentY >= 0 && currentY < height) {
        const insideX = currentX >= 0 && currentX < boxWidth;
        const insideY = currentY >= 0 && currentY < boxHeight;
        if (!insideX || !insideY || !mask[currentY * boxWidth + currentX]) {
          const globalX = x0 + currentX;
          const globalY = y0 + currentY;
          const color = insideX && insideY
            ? null
            : sourcePixel(globalX, globalY);
          const offset = insideX && insideY ? (currentY * boxWidth + currentX) * 4 : 0;
          const resolved = color || [data[offset], data[offset + 1], data[offset + 2]];
          return {
            color: resolved,
            distance
          };
        }
        currentX += deltaX;
        currentY += deltaY;
        distance++;
      }
      return null;
    };

    const eraseRun = (region) => {
      const size = Math.max(4, Number(region.size) || 12);
      const ascent = size * 0.98;
      const descent = size * 0.26;
      const padding = size * 0.08 * scale;
      const x0 = Math.max(0, Math.floor(region.x * scale - padding));
      const x1 = Math.min(width - 1, Math.ceil((region.x + region.width) * scale + padding));
      const y0 = Math.max(0, Math.floor(viewport.height - (region.baseline + ascent) * scale));
      const y1 = Math.min(height - 1, Math.ceil(viewport.height - (region.baseline - descent) * scale));
      if (x1 <= x0 || y1 <= y0) return;
      inpaintRegions++;

      const boxWidth = x1 - x0 + 1;
      const boxHeight = y1 - y0 + 1;
      const patch = context.getImageData(x0, y0, boxWidth, boxHeight);
      const data = patch.data;
      const mask = new Uint8Array(boxWidth * boxHeight);
      const borderColors = [];
      const border = [];
      const localBackgroundSamples = [];
      for (let x = x0 - 2; x <= x1 + 2; x += 2) {
        for (const y of [y0 - 2, y1 + 2]) {
          if (x >= 0 && x < width && y >= 0 && y < height) {
            border.push(sourcePixel(x, y));
          }
        }
      }
      for (let y = y0; y <= y1; y += 2) {
        for (const x of [x0 - 2, x1 + 2]) {
          if (x >= 0 && x < width && y >= 0 && y < height) {
            border.push(sourcePixel(x, y));
          }
        }
      }
      borderColors.push(...border);
      const sortedBackground = [...borderColors].sort((left, right) =>
        (left[0] + left[1] + left[2]) - (right[0] + right[1] + right[2]));
      const background = sortedBackground.length
        ? sortedBackground[Math.floor(sortedBackground.length / 2)]
        : [255, 255, 255];

      const area = boxWidth * boxHeight;
      const buckets = new Map();
      for (let y = 0; y < boxHeight; y += 1) {
        for (let x = 0; x < boxWidth; x += 1) {
          const offset = (y * boxWidth + x) * 4;
          const color = [data[offset], data[offset + 1], data[offset + 2]];
          const key = ((color[0] >> 4) << 8) | ((color[1] >> 4) << 4) | (color[2] >> 4);
          const bucket = buckets.get(key) || { count: 0, sum: [0, 0, 0] };
          bucket.count++;
          bucket.sum[0] += color[0];
          bucket.sum[1] += color[1];
          bucket.sum[2] += color[2];
          buckets.set(key, bucket);
        }
      }

      const declaredTarget = targetColor(region.color);
      const declaredDistance = Math.hypot(
        declaredTarget[0] - background[0],
        declaredTarget[1] - background[1],
        declaredTarget[2] - background[2]
      );
      const candidates = [...buckets.values()]
        .map(bucket => ({
          count: bucket.count,
          color: bucket.sum.map(value => value / bucket.count),
          distance: Math.hypot(
            bucket.sum[0] / bucket.count - background[0],
            bucket.sum[1] / bucket.count - background[1],
            bucket.sum[2] / bucket.count - background[2]
          )
        }))
        .filter(candidate => candidate.count >= Math.max(2, area * 0.002) &&
          candidate.distance >= Math.max(36, declaredDistance * 0.55));
      if (declaredDistance >= 48) {
        let declaredCount = 0;
        for (let y = 0; y < boxHeight; y += 1) {
          for (let x = 0; x < boxWidth; x += 1) {
            const offset = (y * boxWidth + x) * 4;
            if (Math.hypot(
              data[offset] - declaredTarget[0],
              data[offset + 1] - declaredTarget[1],
              data[offset + 2] - declaredTarget[2]
            ) <= Math.max(42, declaredDistance * 0.24)) {
              declaredCount++;
            }
          }
        }
        if (declaredCount >= Math.max(2, area * 0.002)) {
          candidates.push({ count: declaredCount, color: declaredTarget, distance: declaredDistance });
        }
      }
      const foreground = candidates
        .map(candidate => ({
          ...candidate,
          score: candidate.distance * Math.log(candidate.count + 1)
        }))
        .sort((left, right) => right.score - left.score)[0];
      if (!foreground) {
        inpaintStats.push({
          text: region.text || '',
          x0, y0, boxWidth, boxHeight,
          target: declaredTarget,
          background,
          targetLimit: 0,
          maskCount: 0,
          dilatedCount: 0,
          writes: 0,
          skipped: true
        });
        return;
      }

      const declaredToForegroundDistance = Math.hypot(
        declaredTarget[0] - foreground.color[0],
        declaredTarget[1] - foreground.color[1],
        declaredTarget[2] - foreground.color[2]
      );
      if (region.sourceRun && declaredDistance < foreground.distance * 0.72 &&
        declaredToForegroundDistance > Math.max(65, foreground.distance * 0.72)) {
        const measuredColor = colorHex(foreground.color);
        region.color = '#' + measuredColor;
        region.sourceRun.color = region.color;
      }
      const measuredTarget = targetColor(region.color);
      const measuredDeclaredDistance = Math.hypot(
        measuredTarget[0] - background[0],
        measuredTarget[1] - background[1],
        measuredTarget[2] - background[2]
      );
      if (measuredDeclaredDistance < foreground.distance * 0.72) {
        inpaintStats.push({
          text: region.text || '',
          x0, y0, boxWidth, boxHeight,
          target: measuredTarget,
          background,
          targetLimit: 0,
          maskCount: 0,
          dilatedCount: 0,
          writes: 0,
          skipped: true
        });
        return;
      }

      const target = foreground.color;
      const targetBackgroundDistance = foreground.distance;
      const targetLimit = Math.max(42, targetBackgroundDistance * 0.86);
      const perpendicularLimit = Math.max(34, targetBackgroundDistance * 0.48);

      for (let y = 0; y < boxHeight; y++) {
        for (let x = 0; x < boxWidth; x++) {
          const offset = (y * boxWidth + x) * 4;
          const targetDistance = Math.hypot(
            data[offset] - target[0],
            data[offset + 1] - target[1],
            data[offset + 2] - target[2]
          );
          const backgroundDistance = Math.hypot(
            data[offset] - background[0],
            data[offset + 1] - background[1],
            data[offset + 2] - background[2]
          );
          const blendX = data[offset] - background[0];
          const blendY = data[offset + 1] - background[1];
          const blendZ = data[offset + 2] - background[2];
          const targetX = target[0] - background[0];
          const targetY = target[1] - background[1];
          const targetZ = target[2] - background[2];
          const targetLengthSquared = targetX * targetX + targetY * targetY + targetZ * targetZ;
          const blendRatio = targetLengthSquared
            ? (blendX * targetX + blendY * targetY + blendZ * targetZ) / targetLengthSquared
            : 0;
          const segmentX = blendX - blendRatio * targetX;
          const segmentY = blendY - blendRatio * targetY;
          const segmentZ = blendZ - blendRatio * targetZ;
          const blendDistance = Math.hypot(segmentX, segmentY, segmentZ);
          if (targetDistance <= targetLimit ||
            (blendRatio >= 0.16 && blendRatio <= 1.08 && blendDistance <= perpendicularLimit)) {
            mask[y * boxWidth + x] = 1;
          }
        }
      }

      const dilated = new Uint8Array(mask);
      for (let y = 0; y < boxHeight; y++) {
        for (let x = 0; x < boxWidth; x++) {
          if (!mask[y * boxWidth + x]) continue;
          for (let offsetY = -1; offsetY <= 1; offsetY++) {
            for (let offsetX = -1; offsetX <= 1; offsetX++) {
              const nearbyY = y + offsetY;
              const nearbyX = x + offsetX;
              if (nearbyY >= 0 && nearbyY < boxHeight && nearbyX >= 0 && nearbyX < boxWidth) {
                dilated[nearbyY * boxWidth + nearbyX] = 1;
              }
            }
          }
        }
      }
      inpaintStats.push({
        text: region.text || '',
        x0, y0, boxWidth, boxHeight,
        target,
        background,
        targetLimit,
        maskCount: mask.reduce((sum, value) => sum + value, 0),
        dilatedCount: dilated.reduce((sum, value) => sum + value, 0),
        writes: 0
      });
      const stat = inpaintStats[inpaintStats.length - 1];

      for (let y = 0; y < boxHeight; y++) {
        for (let x = 0; x < boxWidth; x++) {
          if (!dilated[y * boxWidth + x]) continue;
          const candidates = [
            nearestKnown(data, dilated, x, y, -1, 0, boxWidth, boxHeight, x0, y0),
            nearestKnown(data, dilated, x, y, 1, 0, boxWidth, boxHeight, x0, y0),
            nearestKnown(data, dilated, x, y, 0, -1, boxWidth, boxHeight, x0, y0),
            nearestKnown(data, dilated, x, y, 0, 1, boxWidth, boxHeight, x0, y0)
          ].filter(Boolean);
          if (!candidates.length) continue;
          let red = 0;
          let green = 0;
          let blue = 0;
          let totalWeight = 0;
          for (const candidate of candidates) {
            const weight = 1 / (candidate.distance * candidate.distance);
            red += candidate.color[0] * weight;
            green += candidate.color[1] * weight;
            blue += candidate.color[2] * weight;
            totalWeight += weight;
          }
          const offset = (y * boxWidth + x) * 4;
          data[offset] = Math.round(red / totalWeight);
          data[offset + 1] = Math.round(green / totalWeight);
          data[offset + 2] = Math.round(blue / totalWeight);
          data[offset + 3] = 255;
          stat.writes++;
        }
      }
      context.putImageData(patch, x0, y0);
    };

    for (const region of exportRegions) eraseRun(region);
    void inpaintRegions;
  }

  if (cleanDocument && exportRegions.length) {
    const cleanPage = await cleanDocument.getPage(pageNumber);
    const cleanViewport = cleanPage.getViewport({ scale });
    const cleanCanvas = document.createElement('canvas');
    cleanCanvas.width = canvas.width;
    cleanCanvas.height = canvas.height;
    const cleanContext = cleanCanvas.getContext('2d', { willReadFrequently: true });
    await cleanPage.render({ canvasContext: cleanContext, viewport: cleanViewport }).promise;

    for (const region of exportRegions) {
      const size = Math.max(4, Number(region.size) || 12);
      const x0 = Math.max(0, Math.floor(region.x * scale - size * 0.08 * scale));
      const x1 = Math.min(canvas.width - 1, Math.ceil((region.x + region.width) * scale + size * 0.08 * scale));
      const y0 = Math.max(0, Math.floor(viewport.height - (region.baseline + size * 0.98) * scale));
      const y1 = Math.min(canvas.height - 1, Math.ceil(viewport.height - (region.baseline - size * 0.26) * scale));
      const boxWidth = x1 - x0 + 1;
      const boxHeight = y1 - y0 + 1;
      if (boxWidth <= 2 || boxHeight <= 2) continue;

      const pad = 4;
      const sampleX = Math.max(0, x0 - pad);
      const sampleY = Math.max(0, y0 - pad);
      const sampleWidth = Math.min(canvas.width, x1 + pad + 1) - sampleX;
      const sampleHeight = Math.min(canvas.height, y1 + pad + 1) - sampleY;
      if (sampleWidth <= 0 || sampleHeight <= 0) continue;
      const originalSamples = context.getImageData(sampleX, sampleY, sampleWidth, sampleHeight).data;
      const cleanSamples = cleanContext.getImageData(sampleX, sampleY, sampleWidth, sampleHeight).data;
      const pixelAt = (samples, x, y) => {
        const offset = ((y - sampleY) * sampleWidth + (x - sampleX)) * 4;
        return [samples[offset], samples[offset + 1], samples[offset + 2]];
      };
      const ringAverage = (samples, inside) => {
        const ring = Math.max(2, Math.round(Math.min(boxWidth, boxHeight) * 0.08));
        const points = [];
        for (let x = x0; x <= x1; x += 2) {
          points.push([x, inside ? y0 + ring : y0 - 2]);
          points.push([x, inside ? y1 - ring : y1 + 2]);
        }
        for (let y = y0; y <= y1; y += 2) {
          points.push([inside ? x0 + ring : x0 - 2, y]);
          points.push([inside ? x1 - ring : x1 + 2, y]);
        }
        const validSamples = points.filter(([x, y]) =>
          x >= sampleX && x < sampleX + sampleWidth && y >= sampleY && y < sampleY + sampleHeight
        ).map(([x, y]) => pixelAt(samples, x, y));
        if (!validSamples.length) return null;
        return [0, 1, 2].map(index =>
          validSamples.reduce((sum, color) => sum + color[index], 0) / validSamples.length);
      };

      const outsideAverage = ringAverage(originalSamples, false);
      const cleanBorderAverage = ringAverage(cleanSamples, true);
      if (!outsideAverage || !cleanBorderAverage) continue;
      const borderDistance = Math.hypot(
        outsideAverage[0] - cleanBorderAverage[0],
        outsideAverage[1] - cleanBorderAverage[1],
        outsideAverage[2] - cleanBorderAverage[2]
      );
      context.drawImage(cleanCanvas, x0, y0, boxWidth, boxHeight, x0, y0, boxWidth, boxHeight);
    }
  }

  return canvas.toDataURL('image/jpeg', 0.88);
}

function paragraphTopPt(model, block) {
  const size = Number(block.size) || Number(block.runs?.[0]?.size) || 12;
  const lineHeight = Number(block.lineHeight) || size * 1.22;
  const ascent = (Number(block.ascent) || 0.82) * size;
  const leading = (lineHeight - size * 1.22) / 2;
  return Math.max(0, model.height - block.baseline - ascent - leading);
}

async function renderDocumentPageToDataUrl(pdfDocument, pageNumber, scale = 2, format = 'png') {
  const page = await pdfDocument.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  return canvas.toDataURL(format === 'jpg' ? 'image/jpeg' : 'image/png', 0.92);
}

/* ────── 导出图片 ────── */
function dataUrlToBytes(dataUrl) {
  const b64 = dataUrl.split(',')[1];
  const bin = atob(b64);
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8;
}

/** 捕获当前页画布为图片字节 */
async function buildPageImageExport(format) {
  if (!viewer.doc) throw new Error('请先打开 PDF 文档');
  if (!state.originalBytes) throw new Error('原始 PDF 数据不可用');
  contentEditor.commitDom();
  const annotatedBytes = await buildAnnotatedBytes();
  const sourceDocument = await pdfjsLib.getDocument({
    data: (annotatedBytes || state.originalBytes).slice(0),
    cMapUrl: '../vendor/cmaps/',
    cMapPacked: true,
    enableXfa: false,
    fontExtraProperties: true
  }).promise;
  const images = [];
  try {
    for (let pageNumber = 1; pageNumber <= state.numPages; pageNumber++) {
      toast(`渲染图片 ${pageNumber}/${state.numPages}…`);
      const dataUrl = await renderDocumentPageToDataUrl(sourceDocument, pageNumber, 2, format);
      images.push({ name: `page-${pageNumber}.${format}`, data: dataUrlToBytes(dataUrl) });
    }
  } finally {
    await sourceDocument.cleanup?.();
  }
  const suffix = format === 'png' ? 'PNG' : 'JPG';
  return {
    bytes: zipStore(images),
    suggestedName: `${state.fileName.replace(/\.pdf$/i, '')}-${suffix}-images.zip`
  };
}

function capturePageImage(format) {
  const box = viewer.pageBoxes.get(viewer.currentPage);
  const canvas = box && box.querySelector('canvas');
  if (!canvas) throw new Error('页面尚未渲染完成');
  const bytes = dataUrlToBytes(canvas.toDataURL(format === 'png' ? 'image/png' : 'image/jpeg', 0.92));
  return { bytes, suggestedName: `${state.fileName.replace(/\.pdf$/i, '')}-p${viewer.currentPage}.${format}` };
}

/** 导出当前页为图片 */
async function doExportImage(format) {
  if (!viewer.doc) { toast('请先打开 PDF 文档', true); return; }
  try {
    const { bytes, suggestedName } = capturePageImage(format);
    const r = await window.onyx.exportPdf({
      bytes, suggestedName,
      title: `导出图片（${format.toUpperCase()}）`,
      filters: [{ name: `${format.toUpperCase()} 图片`, extensions: [format] }]
    });
    if (r.ok) toast(`已导出图片：${r.path}`);
  } catch (err) {
    toast(`导出失败：${err && err.message || err}`, true);
  }
}

/** 导出全部页面为图片（渲染倍率 2x） */
async function doExportAllImages(format) {
  if (!viewer.doc) { toast('请先打开 PDF 文档', true); return; }
  const dir = await window.onyx.openDirDialog();
  if (!dir) return;
  const base = state.fileName.replace(/\.pdf$/i, '');
  try {
    for (let p = 1; p <= state.numPages; p++) {
      toast(`渲染页面 ${p}/${state.numPages}…`);
      const page = await viewer.doc.getPage(p);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      const bytes = dataUrlToBytes(canvas.toDataURL(format === 'png' ? 'image/png' : 'image/jpeg', 0.92));
      const r = await window.onyx.writePath({ path: `${dir}\\${base}-p${p}.${format}`, bytes });
      if (!r.ok) throw new Error(r.error);
    }
    toast(`已导出 ${state.numPages} 张图片到 ${dir}`);
  } catch (err) {
    toast(`导出失败：${err && err.message || err}`, true);
  }
}

/* ══════════ 打印 ══════════ */
async function doBatchPrint() {
  const paths = await window.onyx.openMultiDialog();
  if (!paths) return;
  for (let i = 0; i < paths.length; i++) {
    const name = paths[i].replace(/^.*[\\/]/, '');
    toast(`打印 ${i + 1}/${paths.length}：${name}`);
    await openFromPath(paths[i]);
    await new Promise(r => setTimeout(r, 600));
    const res = await window.onyx.print();
    if (res && res.canceled) break;
  }
  toast('批量打印完成');
}

/* ══════════ 新建 PDF ══════════ */
async function doNewBlank() {
  openModal({
    title: '新建空白 PDF',
    bodyHTML: `
      <div class="field">
        <label>页数</label>
        <input type="number" data-role="pages" value="1" min="1" max="500" />
      </div>
      <p class="hint">创建 A4 纵向空白文档，可在阅读器中查看后另存。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '创建', primary: true,
        onClick: async (m) => {
          const n = clamp(parseInt(m.query('[data-role="pages"]').value, 10) || 1, 1, 500);
          const bytes = await tools.toolNewBlank(n);
          docSeq++;
          await openFromBytes(new Uint8Array(bytes), `未命名-${docSeq}.pdf`);
          m.close();
          toast(`已创建 ${n} 页空白文档`);
        }
      }
    ]
  });
}

async function doNewFromImages() {
  try {
    const doc = await tools.buildPdfFromImages();
    if (!doc) return;
    const bytes = await doc.save();
    docSeq++;
    await openFromBytes(new Uint8Array(bytes), `图片文档-${docSeq}.pdf`);
    toast('已从图片创建 PDF');
  } catch (err) {
    toast(err.message, true);
  }
}

/* ══════════ 页面编辑 ══════════ */
async function applyPageReorder(order) {
  if (!state.originalBytes || !order || !order.length) return;
  try {
    const bytes = await tools.reorderPagesBytes(state.originalBytes, order);
    const name = `${state.fileName.replace(/\.pdf$/i, '')}-重排.pdf`;
    await openFromBytes(new Uint8Array(bytes), name);
    toast('页面顺序已调整（新文档，可另存保存）');
  } catch (err) {
    toast(`排序失败：${err && err.message || err}`, true);
  }
}

/** 插入图片页（作为整页插入当前文档，产出新文档） */
async function doInsertImagePages() {
  if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
  try {
    const result = await tools.toolInsertImagePages(toolCtx, viewer.currentPage);
    if (!result) return;
    const name = state.fileName.replace(/\.pdf$/i, '') + result.suffix;
    await openFromBytes(new Uint8Array(result.bytes), name);
    toast('已插入图片页（新文档，可另存保存）');
  } catch (err) {
    toast(err.message, true);
  }
}

/* ══════════ 关于 ══════════ */
async function doAbout() {
  let ver = '';
  try { ver = await window.onyx.appVersion(); } catch (_) { /* noop */ }
  openModal({
    title: '关于 OnyxPDF',
    width: 360,
    bodyHTML: `
      <div class="about-body">
        <div class="about-logo">
          <svg viewBox="0 0 24 24"><polygon points="12,2 22,20 2,20" fill="none" stroke="currentColor" stroke-width="1.4"/><circle cx="12" cy="14.5" r="3.2" fill="currentColor"/></svg>
        </div>
        <h2>ONYX<em>PDF</em></h2>
        <div class="a-ver">版本 ${ver || '1.0.0'}</div>
        <div class="a-desc">
          OnyxPDF，国产一站式PDF阅览工具，支持文档编辑、多格式导出，高效处理PDF文档。
        </div>
        <div class="a-author">作者：Friday（周五王）</div>
      </div>
    `,
    buttons: [{ label: '确定', primary: true, onClick: (m) => m.close() }]
  });
}

/* ══════════ 最近文件 ══════════ */
let _recentsCache = [];
async function renderRecents() {
  _recentsCache = await window.onyx.recentsGet();
  const list = _recentsCache;
  const ul = document.getElementById('recents-list');
  if (!list.length) {
    ul.innerHTML = '<li class="recents-empty" style="cursor:default;border:0">暂无记录</li>';
    return;
  }
  ul.innerHTML = '';
  for (const it of list) {
    const li = document.createElement('li');
    li.innerHTML = `<span class="r-dot"></span><span class="r-name" title="${escapeHtml(it.path)}">${escapeHtml(it.name)}</span><span class="r-time">${timeAgo(it.time)}</span>`;
    li.addEventListener('click', () => openFromPath(it.path));
    ul.appendChild(li);
  }
}

/* ══════════ 菜单定义 ══════════ */
const toolCtx = {
  hasDoc: () => !!state.originalBytes,
  currentPath: () => state.filePath,
  pageCount: () => state.numPages,
  getBytes: () => state.originalBytes,
  getName: () => state.fileName
};

function setupMenus() {
  const ctx = {
    menuDefs: {
      file: () => {
        const items = [
          { id: 'open', label: '打开', key: 'Ctrl+O', icon: I.open, action: openViaDialog },
          {
            id: 'new', label: '新建 PDF', icon: I.newdoc,
            children: [
              { id: 'new-blank', label: '空白文档', icon: I.blank, action: doNewBlank },
              { id: 'new-img', label: '从图片创建', icon: I.image, action: doNewFromImages }
            ]
          },
          {
            id: 'recents', label: '打开最近的文件', icon: I.clock,
            children: _recentsCache.length
              ? _recentsCache.map(r => ({
                  id: `recent-${r.path}`,
                  label: r.name,
                  icon: I.clock,
                  action: () => openFromPath(r.path)
                }))
              : [{ id: 'no-recent', label: '（暂无记录）', enabled: () => false }]
          },
          { sep: true },
          { id: 'save', label: '保存', key: 'Ctrl+S', icon: I.save, enabled: () => !!state.originalBytes, action: doSave },
          { id: 'saveas', label: '另存为', key: 'Ctrl+Shift+S', icon: I.save, enabled: () => !!state.originalBytes, action: doSaveAs },
          { sep: true },
          { id: 'export', label: '导出为…', key: '', icon: I.exportd, enabled: () => !!state.originalBytes, action: doExportAs },
          { id: 'export-quick', label: '快速导出（含批注 PDF）', icon: I.exportd, enabled: () => !!state.originalBytes, action: doExport },
          { sep: true },
          { id: 'print', label: '打印', key: 'Ctrl+P', icon: I.print, enabled: () => !!state.originalBytes, action: () => window.onyx.print() },
          { id: 'batchprint', label: '批量打印', icon: I.print, action: doBatchPrint },
          { sep: true },
          { id: 'about', label: '关于', icon: I.info, action: doAbout },
          { sep: true },
          { id: 'quit', label: '退出', key: 'Ctrl+Q', icon: I.quit, action: () => window.onyx.winClose() }
        ];
        return items;
      },
      tools: () => [
        { id: 'merge', label: '合并 PDF…', icon: I.merge, action: tools.toolMerge },
        { id: 'split', label: '拆分 PDF…', icon: I.split, action: () => tools.toolSplit(toolCtx) },
        { id: 'extract', label: '提取页面…', icon: I.extract, action: () => tools.toolExtract(toolCtx) },
        { id: 'delete-pages', label: '删除页面…', icon: I.trash, action: () => tools.toolDeletePages(toolCtx) },
        { id: 'rotate-pages', label: '旋转页面…', icon: I.rotate, action: () => tools.toolRotatePages(toolCtx) },
        { id: 'insert-blank', label: '插入空白页…', icon: I.blank, action: () => tools.toolInsertBlank(toolCtx) },
        { id: 'insert-image-page', label: '插入图片页…', icon: I.image, action: doInsertImagePages },
        { id: 'insert-text', label: '编辑内容（点击页面文字直接修改）…', icon: I.newdoc, action: () => {
            if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
            if (!annotations.editMode) document.getElementById('btn-edit-mode').click();
            toast('编辑内容已开启：点击页面文字直接修改，选中部分文字可单独设置格式');
          } },
        { sep: true },
        { id: 'watermark', label: '文本水印…', icon: I.water, action: () => tools.toolWatermark(toolCtx) },
        { id: 'pagenum', label: '添加页码…', icon: I.hash, action: () => tools.toolPageNumbers(toolCtx) },
        { sep: true },
        { id: 'img2pdf', label: '图片转 PDF…', icon: I.image, action: tools.toolImagesToPdf },
        { id: 'optimize', label: '清理优化…', icon: I.broom, action: () => tools.toolOptimize(toolCtx) },
        { sep: true },
        { id: 'reorder-hint', label: '页面排序：侧栏拖动缩略图', icon: I.rotate, enabled: () => !!state.originalBytes, action: () => {
          const el = document.getElementById('btn-sidebar');
          if (el && document.getElementById('sidebar').classList.contains('hidden')) el.click();
          toast('在缩略图面板中拖动页面到目标位置，点击“应用排序”');
        } }
      ],
      help: () => [
        { id: 'about', label: '关于 OnyxPDF', icon: I.info, action: doAbout }
      ]
    }
  };
  initMenus(ctx);
}

/* ══════════ 工具栏 / 窗口 / 快捷键 ══════════ */
function bindChrome() {
  document.getElementById('btn-open').addEventListener('click', openViaDialog);
  document.getElementById('btn-open-welcome').addEventListener('click', openViaDialog);
  document.getElementById('btn-export').addEventListener('click', doExportAs);
  document.getElementById('btn-edit-mode').addEventListener('click', () => {
    contentEditor.setEnabled(!contentEditor.enabled).then(() => {
      toast(contentEditor.enabled
        ? '编辑内容已开启：点击页面文字直接修改；选中部分文字可单独设置格式'
        : '编辑内容已关闭');
    }).catch(error => {
      console.error(error);
      toast('内容编辑切换失败：' + (error?.message || error), true);
    });
  });

  // 编辑模式插入图片：选文件 → 点击页面放置
  const imgInput = document.getElementById('img-file-input');
  document.getElementById('btn-insert-image').addEventListener('click', () => {
    if (!state.originalBytes) { toast('请先打开 PDF 文档', true); return; }
    if (contentEditor.enabled) {
      imgInput.value = '';
      imgInput.click();
      return;
    }
    annotations.pendingImage = true;
    imgInput.value = '';
    imgInput.click();
  });
  imgInput.addEventListener('change', () => {
    const f = imgInput.files && imgInput.files[0];
    if (!f) return;
    if (contentEditor.enabled) {
      contentEditor.prepareImageFile(f).catch(error => toast(error.message, true));
      return;
    }
    if (!annotations.pendingImage) return;
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const im = new Image();
      im.onload = () => {
        // 按 96dpi 换算 pt，并限制最大 300pt 宽
        let wPt = im.naturalWidth * 0.75, hPt = im.naturalHeight * 0.75;
        if (wPt > 300) { hPt *= 300 / wPt; wPt = 300; }
        annotations._pendingImageData = { dataUrl, wPt, hPt };
        toast('在页面上点击位置放置图片');
      };
      im.src = dataUrl;
    };
    reader.readAsDataURL(f);
  });
  document.getElementById('btn-print').addEventListener('click', () => window.onyx.print());

  document.getElementById('btn-zoom-in').addEventListener('click', () => viewer.zoomStep(1));
  document.getElementById('btn-zoom-out').addEventListener('click', () => viewer.zoomStep(-1));
  document.getElementById('btn-fit-width').addEventListener('click', () => viewer.setScaleMode('fit-width'));
  document.getElementById('btn-fit-page').addEventListener('click', () => viewer.setScaleMode('fit-page'));
  document.getElementById('btn-rotate').addEventListener('click', () => viewer.rotate());

  const sidebarEl = document.getElementById('sidebar');
  document.getElementById('btn-sidebar').addEventListener('click', () => {
    sidebarEl.classList.toggle('hidden');
    if (!sidebarEl.classList.contains('hidden') && viewer.doc) sidebar.buildThumbs();
  });

  document.getElementById('btn-search').addEventListener('click', () => {
    sidebarEl.classList.remove('hidden');
    document.querySelectorAll('.side-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === 'search'));
    document.querySelectorAll('.side-panel').forEach(p => p.classList.add('hidden'));
    document.getElementById('panel-search').classList.remove('hidden');
    document.getElementById('search-input').focus();
  });

  // 窗口控制
  document.getElementById('btn-win-min').addEventListener('click', () => window.onyx.winMinimize());
  document.getElementById('btn-win-max').addEventListener('click', () => window.onyx.winMaximize());
  document.getElementById('btn-win-close').addEventListener('click', () => window.onyx.winClose());
  window.onyx.isMaximized().then(v => document.body.classList.toggle('is-max', v));
  window.onyx.onWinState(v => document.body.classList.toggle('is-max', v));

  // 最近文件清空
  document.getElementById('btn-recents-clear').addEventListener('click', async () => {
    await window.onyx.recentsClear();
    _recentsCache = [];
    renderRecents();
  });

  // 快捷键
  window.addEventListener('keydown', (e) => {
    const ctrl = e.ctrlKey || e.metaKey;
    const editableTarget = e.target instanceof Element && e.target.closest('input, textarea, select, [contenteditable="true"], [contenteditable=""]');
    if (editableTarget && !(ctrl && (e.key.toLowerCase() === 's' || e.key.toLowerCase() === 'q'))) return;
    if (ctrl && e.key.toLowerCase() === 'o') { e.preventDefault(); openViaDialog(); }
    else if (ctrl && e.shiftKey && e.key.toLowerCase() === 's') { e.preventDefault(); doSaveAs(); }
    else if (ctrl && e.key.toLowerCase() === 's') { e.preventDefault(); doSave(); }
    else if (ctrl && e.key.toLowerCase() === 'q') { e.preventDefault(); window.onyx.winClose(); }
    else if (ctrl && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      document.getElementById('btn-search').click();
    }
    else if (ctrl && e.key.toLowerCase() === 'p') { e.preventDefault(); window.onyx.print(); }
    else if (ctrl && (e.key === '=' || e.key === '+')) { e.preventDefault(); viewer.zoomStep(1); }
    else if (ctrl && e.key === '-') { e.preventDefault(); viewer.zoomStep(-1); }
    else if (ctrl && e.key === '0') { e.preventDefault(); viewer.setScaleMode('fit-width'); }
    else if (ctrl && e.key.toLowerCase() === 'v' && annotations.editMode && !e.target.matches('input,textarea,[contenteditable]')) {
      e.preventDefault();
      annotations.pasteImageFromClipboard(viewer.currentPage || 1);
    }
    else if (e.key === 'Delete' && annotations.editMode && !e.target.matches('input,textarea,[contenteditable]')) {
      e.preventDefault();
      annotations.toggleDeleteSelected();
    }
    else if (ctrl && e.key === '\t') { e.preventDefault(); if (tabs.length) { const i = tabs.findIndex(t => t.id === activeTabId); awaitSwitchTab((i + 1) % tabs.length); } }
    else if (!ctrl && viewer.doc) {
      if (e.key === 'ArrowRight' || e.key === 'PageDown') {
        e.preventDefault(); viewer.goToPage(viewer.currentPage + 1);
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault(); viewer.goToPage(viewer.currentPage - 1);
      } else if (e.key === 'Home') { e.preventDefault(); viewer.goToPage(1); }
      else if (e.key === 'End') { e.preventDefault(); viewer.goToPage(state.numPages); }
      else if (e.key.toLowerCase() === 'v') annotations.setTool('select');
      else if (e.key.toLowerCase() === 'h') annotations.setTool('highlight');
      else if (e.key.toLowerCase() === 'u') annotations.setTool('underline');
      else if (e.key.toLowerCase() === 'r') annotations.setTool('rect');
      else if (e.key.toLowerCase() === 'n') annotations.setTool('note');
      else if (e.key.toLowerCase() === 's' && !e.target.matches('input,textarea')) {
        document.getElementById('btn-sidebar').click();
      }
      else if (e.key.toLowerCase() === 'w' && !e.target.matches('input,textarea')) viewer.setScaleMode('fit-width');
    }
  });

  // 拖放打开
  const mask = document.getElementById('drag-mask');
  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => { e.preventDefault(); dragDepth++; mask.classList.remove('hidden'); });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; mask.classList.add('hidden'); } });
  window.addEventListener('drop', async (e) => {
    e.preventDefault();
    dragDepth = 0;
    mask.classList.add('hidden');
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    const p = window.onyx.pathForFile(f);
    if (p) await openFromPath(p);
  });

  // 主进程投递的启动参数（命令行 / 双击关联）
  window.onyx.onOpenPath(p => openFromPath(p));
}

async function awaitSwitchTab(i) {
  const t = tabs[i];
  if (t) await activateTab(t.id);
}

/* ══════════ 启动 ══════════ */
(async function boot() {
  // 布局保险：任何意外导致 document 滚动（顶部缺失）时立即归位
  const clampScroll = () => {
    if (document.documentElement.scrollTop || document.documentElement.scrollLeft ||
        document.body.scrollTop || document.body.scrollLeft) {
      document.documentElement.scrollTop = 0;
      document.documentElement.scrollLeft = 0;
      document.body.scrollTop = 0;
      document.body.scrollLeft = 0;
      window.scrollTo(0, 0);
    }
  };
  window.addEventListener('scroll', clampScroll, true);
  new MutationObserver(clampScroll).observe(document.body, { childList: true, subtree: true });
  clampScroll();

  bindChrome();
  annotations.setTool('select');
  setupMenus();
  await renderRecents();
  // 启动参数直开
  try {
    const args = window.process?.argv || [];
    const pdfArg = args.slice(1).find(a => /\.pdf$/i.test(a));
    if (pdfArg) await openFromPath(pdfArg);
  } catch (_) { /* noop */ }
  // 供自动化测试使用的钩子（仅内部）
  // V12.3.5：品牌区显示版本号（用户可自查是否新版）
  try {
    window.onyx.version().then(v => {
      const bn = document.querySelector('.brand-name');
      if (bn) bn.innerHTML = 'ONYX<em>PDF</em> <span style=\"font-size:11px;opacity:.75;font-weight:400\">v' + v + '</span>';
    }).catch(() => {});
  } catch (_) {}
  window.onyxRender = { renderPageErased: async (p) => renderPageErased(p, 2) };
  window.__onyx_test = {
    state, viewer, annotations, contentEditor, search, tabs,
    openFromPath, doExport, doSave, doSaveAs, doBatchPrint, doExportText,
    doExportImage, doExportAllImages,
    buildPptxExport, buildXlsxExport, buildPptxEditableExport, buildWordLayoutExport, buildPptxEditableExportV2, buildParagraphExportPages, flattenParagraphExportBackgrounds, mapFontName, isGarbled, markRecurringLines, buildTxtExport, capturePageImage, buildPageImageExport, collectPageLines, doExportAs, createTextFreeDocumentBytes, renderAndInpaintPage, renderPageBackgroundPatch,
    tools, openFromBytes, doNewBlank, doNewFromImages, applyPageReorder, listTestExport: () => window.onyx.testExportList(), renderPageErased: async (p) => renderPageErased(p, 2), renderPageErasedDebug: async (p) => { const r = await renderPageErased(p, 2); return { lines: r.lines.filter(l => /\u76ee\u5f55|\u4e91\u4ea7\u54c1/.test(l.text)).map(l => ({ text: l.text.slice(0,20), x: Math.round(l.x), y: Math.round(l.y), size: l.size, hidden: !!l.hidden, garbled: !!l.garbled })) }; }, extractContentPage,
    pdfjsLib,
    activateTab, closeTab
  };
})();
