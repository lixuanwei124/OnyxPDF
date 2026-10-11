import * as pdfjsLib from '../vendor/pdf.min.mjs';
import { clamp, toast, debounce } from './utils.js';
import { extractContentPage, applySavedEdits, buildContentViews } from './content-model.js';
import { applyContentEdits, originalFontSupportsText, splitRunBySupportedGlyphs } from './pdf-text-writer.js';
import { removeTextOperations } from './pdf-content-stream.js';
import { exportAnnotatedPdf } from './exporter.js';

const MODE_LABELS = {
  paragraph: '段落模式',
  line: '行模式'
};

function lineHeightRatio(block) {
  const base = Number(block?.originalLineHeight) ||
    Number(block?.lineHeight) ||
    (Number(block?.size) || 12) * 1.22;
  return Number(((Number(block?.lineHeight) || base) / base).toFixed(2));
}

function lineHeightFromRatio(ratio, block) {
  const base = Number(block?.originalLineHeight) ||
    Number(block?.lineHeight) ||
    (Number(block?.size) || 12) * 1.22;
  return clamp(ratio * base, 4, 192);
}

function quantile(values, ratio) {
  if (!values.length) return 0;
  const position = clamp(ratio, 0, 1) * (values.length - 1);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  const weight = position - lower;
  return values[lower] * (1 - weight) + values[upper] * weight;
}

function fontFamilyName(value) {
  return String(value || '').split(',')[0].replace(/["']/g, '').trim() || 'PDF Original';
}

function toolbarFontFamilyName(value) {
  return fontFamilyName(value)
    .replace(/[-_ ]?(?:Regular|Bold|Heavy|Black|Semib?old|Light|Medium|Italic|Oblique)$/i, '')
    .trim() || fontFamilyName(value);
}

function cssFontSizePt(value, scale = 1) {
  if (typeof value === 'number') return value;
  const raw = String(value || '').trim();
  const number = parseFloat(raw);
  if (!Number.isFinite(number)) return 0;
  return raw.endsWith('pt') ? number * 0.75 : number / (scale || 1);
}

function runStyle(run) {
  return [
    `font-family:${run.cssFontFamily || `"${fontFamilyName(run.font)}","Microsoft YaHei",sans-serif`}`,
    `font-size:${(Number(run.size) || 12) * (run.scale || 1)}px`,
    `color:${run.color || '#000000'}`,
    `font-weight:${run.bold ? '700' : '400'}`,
    `font-style:${run.italic ? 'italic' : 'normal'}`,
    `text-decoration:${run.underline ? 'underline' : 'none'}`
  ].join(';');
}

function globalOffset(editor, node, offset) {
  if (node === editor) return offset === 0 ? 0 : editor.textContent.length;
  let total = 0;
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const current = walker.currentNode;
    if (current === node) return total + offset;
    total += current.textContent.length;
  }
  return total;
}

function textNodes(editor) {
  const nodes = [];
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) nodes.push(walker.currentNode);
  return nodes;
}

function normalizeFormattedSpans(editor) {
  for (const span of [...editor.querySelectorAll('span')]) {
    if (span.classList.contains('content-frame-handle')) continue;
    const parent = span.parentElement;
    if (parent === editor || parent?.tagName !== 'SPAN' ||
      parent.classList.contains('content-frame-handle') ||
      parent.childNodes.length !== 1 || parent.firstElementChild !== span) continue;
    parent.replaceWith(span);
  }
}

function selectionRange(editor) {
  const selection = window.getSelection();
  if (!selection.rangeCount) return null;
  const range = selection.getRangeAt(0);
  return editor.contains(range.commonAncestorContainer) ? range : null;
}

function restoreSelection(editor, saved) {
  if (!saved || saved.element !== editor || saved.end <= saved.start) return false;
  const nodes = textNodes(editor);
  let cursor = 0;
  let startNode = null;
  let endNode = null;
  for (const node of nodes) {
    const next = cursor + node.textContent.length;
    if (!startNode && saved.start < next) startNode = node;
    if (saved.end >= cursor && saved.end <= next) endNode = node;
    cursor = next;
  }
  if (!startNode || !endNode) return false;
  const range = document.createRange();
  range.setStart(startNode, saved.start - globalOffset(editor, startNode, 0));
  range.setEnd(endNode, saved.end - globalOffset(editor, endNode, 0));
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  return true;
}

function collectRuns(editor, fallback, scale = 1) {
  const runs = [];
  const push = (text, owner) => {
    if (!text) return;
    const style = owner.style || {};
    const systemFont = owner._onyxSystemFont === true;
    const next = {
      text,
      font: systemFont
        ? fontFamilyName(style.fontFamily || fallback.font)
        : owner._onyxFontName ?? fontFamilyName(style.fontFamily || fallback.font),
      fontKey: owner._onyxFontKey ?? fallback.fontKey ?? null,
      loadedName: owner._onyxCssFont ?? fallback.loadedName ?? null,
      fontData: systemFont ? null : owner._onyxFontData ?? fallback.fontData ?? null,
      size: Math.round((cssFontSizePt(style.fontSize, scale) || fallback.size) * 10) / 10,
      color: style.color || fallback.color || '#000000',
      bold: Number(style.fontWeight || (fallback.bold ? '700' : '400')) >= 600,
      italic: String(style.fontStyle || (fallback.italic ? 'italic' : 'normal')) === 'italic',
      underline: String(style.textDecoration || '').includes('underline') || !!fallback.underline
    };
    const previous = runs[runs.length - 1];
    if (previous && previous.fontKey === next.fontKey && previous.font === next.font &&
      previous.loadedName === next.loadedName &&
      previous.fontData === next.fontData &&
      Math.abs(previous.size - next.size) < 0.05 && previous.color === next.color &&
      previous.bold === next.bold && previous.italic === next.italic &&
      previous.underline === next.underline) {
      previous.text += next.text;
      return;
    }
    runs.push(next);
  };

  const walk = (element, inherited = {}) => {
    const current = {
      ...inherited,
      fontFamily: element.style?.fontFamily || inherited.fontFamily || fallback.font,
      fontSize: element.style?.fontSize || inherited.fontSize || `${fallback.size * scale}px`,
      color: element.style?.color || inherited.color || fallback.color,
      fontWeight: element.style?.fontWeight || inherited.fontWeight || (fallback.bold ? '700' : '400'),
      fontStyle: element.style?.fontStyle || inherited.fontStyle || (fallback.italic ? 'italic' : 'normal'),
      textDecoration: element.style?.textDecoration || inherited.textDecoration || (fallback.underline ? 'underline' : 'none')
    };
    for (const node of element.childNodes) {
      const owner = {
        style: current,
        _onyxFontName: element._onyxFontName,
        _onyxFontKey: element._onyxFontKey,
        _onyxFontData: element._onyxFontData,
        _onyxCssFont: element._onyxCssFont,
        _onyxSystemFont: element._onyxSystemFont
      };
      if (node.nodeType === Node.TEXT_NODE) push(node.textContent, owner);
      else if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'BR') push('\n', owner);
        else walk(node, current);
      }
    }
  };
  walk(editor);
  return runs;
}

function applyFormatToSelection(editor, patch, fallback) {
  const liveRange = selectionRange(editor);
  let range = liveRange && !liveRange.collapsed ? liveRange : null;
  if (!range && editor._onyxSelectedRange?.end > editor._onyxSelectedRange.start) {
    if (!restoreSelection(editor, editor._onyxSelectedRange)) return false;
    range = selectionRange(editor);
  }
  if (!range || range.collapsed) {
    range = document.createRange();
    range.selectNodeContents(editor);
  }
  const startOffset = globalOffset(editor, range.startContainer, range.startOffset);
  const endOffset = globalOffset(editor, range.endContainer, range.endOffset);
  if (endOffset <= startOffset) return false;
  const scale = fallback.scale || 1;
  let cursor = 0;
  let changed = false;
  for (const node of textNodes(editor)) {
    const nodeStart = cursor;
    const nodeEnd = cursor + node.textContent.length;
    cursor = nodeEnd;
    const overlapStart = Math.max(startOffset, nodeStart);
    const overlapEnd = Math.min(endOffset, nodeEnd);
    if (overlapStart >= overlapEnd) continue;

    const parent = node.parentElement || editor;
    const current = {
      font: parent._onyxFontName ?? fontFamilyName(parent.style.fontFamily || fallback.font),
      fontKey: parent._onyxFontKey ?? fallback.fontKey ?? null,
      loadedName: parent._onyxCssFont ?? fallback.loadedName ?? null,
      fontData: parent._onyxFontData ?? fallback.fontData ?? null,
      size: cssFontSizePt(parent.style.fontSize, scale) || fallback.size,
      color: parent.style.color || fallback.color,
      bold: Number(parent.style.fontWeight || (fallback.bold ? '700' : '400')) >= 600,
      italic: String(parent.style.fontStyle || (fallback.italic ? 'italic' : 'normal')) === 'italic',
      underline: String(parent.style.textDecoration || '').includes('underline') || !!fallback.underline
    };
    const systemFont = parent._onyxSystemFont === true;
    const makeSpan = (
      text, style,
      fontKey = parent._onyxFontKey ?? null,
      fontData = parent._onyxFontData ?? null,
      fontName = parent._onyxFontName ?? null,
      cssFont = parent._onyxCssFont ?? null,
      explicitSystemFont = systemFont
    ) => {
      const span = document.createElement('span');
      span.style.cssText = style;
      span._onyxFontName = fontName;
      span._onyxFontKey = fontKey;
      span._onyxFontData = fontData;
      span._onyxCssFont = cssFont;
      if (explicitSystemFont) span._onyxSystemFont = true;
      span.textContent = text;
      return span;
    };
    const fragment = document.createDocumentFragment();
    if (overlapStart > nodeStart) {
      fragment.appendChild(makeSpan(node.textContent.slice(0, overlapStart - nodeStart), runStyle({ ...current, scale })));
    }
    if (overlapEnd > overlapStart) {
      const patched = { ...current, ...patch, scale };
      const preserveFont = !Object.prototype.hasOwnProperty.call(patch, 'font') &&
        !Object.prototype.hasOwnProperty.call(patch, 'bold');
      fragment.appendChild(makeSpan(
        node.textContent.slice(overlapStart - nodeStart, overlapEnd - nodeStart),
        runStyle(patched),
        preserveFont ? current.fontKey : null,
        preserveFont ? current.fontData : null,
        preserveFont ? current.font : null,
        preserveFont ? current.loadedName : null,
        preserveFont ? systemFont : true
      ));
    }
    if (nodeEnd > overlapEnd) {
      fragment.appendChild(makeSpan(node.textContent.slice(overlapEnd - nodeStart), runStyle({ ...current, scale })));
    }
    node.replaceWith(fragment);
    changed = true;
  }

  if (changed) {
    normalizeFormattedSpans(editor);
    editor._onyxSelectedRange = { element: editor, start: startOffset, end: endOffset };
    restoreSelection(editor, editor._onyxSelectedRange);
  }
  return changed;
}

export class ContentEditor {
  constructor(context) {
    this.ctx = context;
    this.enabled = false;
    this.fileKey = '';
    this.models = new Map();
    this.edits = new Map();
    this.images = new Map();
    this.mode = 'paragraph';
    this.showBorders = true;
    this.selected = null;
    this.selectedRange = null;
    this.selectedImage = null;
    this.pendingImage = null;
    this.committing = false;
    this.originalDoc = null;
    this.cachedSourceLibDocument = null;
    this.sourceLibPromise = null;
    this.toolbar = null;
    this.modelPromises = new Map();
    this.fontPromises = new Map();
    this.lineMetrics = new Map();
    this.erasedBytes = null;
    this.erasedPromise = null;
    this._expandSingleLinePreview = debounce(element => this.expandSingleLinePreviewNow(element), 140);
    this._syncToolbarIdle = debounce(() => this.syncToolbar(), 100);
  }

  get storageKey() {
    return `onyx:content-editor-v4:${this.fileKey}`;
  }

  load(fileKey) {
    this.fileKey = fileKey;
    this.enabled = false;
    this.models.clear();
    this.originalDoc = null;
    this.cachedSourceLibDocument = null;
    this.sourceLibPromise = null;
    this.modelPromises.clear();
    this.lineMetrics.clear();
    this.erasedBytes = null;
    this.erasedPromise = null;
    this.edits.clear();
    this.images.clear();
    this.mode = 'paragraph';
    this.selected = null;
    this.selectedRange = null;
    this.selectedImage = null;
    this.pendingImage = null;
    this.removeLayer();
    this.hideToolbar();
  }

  allEdits() {
    return [...this.edits.values()].filter(edit => edit.engine === 2);
  }

  allImages() {
    return [...this.images.values()].flat();
  }

  persist() {
    return;
  }

  async getOriginalDoc() {
    if (this.originalDoc) return this.originalDoc;
    const bytes = this.ctx.getOriginalBytes();
    if (!bytes) throw new Error('原始 PDF 不可用');
    this.originalDoc = await pdfjsLib.getDocument({
      data: bytes.slice(0),
      cMapUrl: '../vendor/cmaps/',
      cMapPacked: true,
      enableXfa: false,
      fontExtraProperties: true
    }).promise;
    return this.originalDoc;
  }

  async buildErasedBytes() {
    if (this.erasedBytes) return this.erasedBytes.slice(0);
    if (this.erasedPromise) return await this.erasedPromise;
    this.erasedPromise = (async () => {
      const { PDFDocument } = window.PDFLib;
      const sourceBytes = await this.editableSourceBytes();
      const pdfDocument = await PDFDocument.load(sourceBytes, {
        ignoreEncryption: true,
        updateMetadata: false
      });
      const originalDoc = await this.getOriginalDoc();
      for (let pageNumber = 1; pageNumber <= originalDoc.numPages; pageNumber++) {
        removeTextOperations(pdfDocument, pageNumber - 1, [], { all: true, protectUnreliable: true });
      }
      const bytes = await pdfDocument.save();
      this.erasedBytes = bytes.slice(0);
      return bytes.slice(0);
    })();
    try {
      return await this.erasedPromise;
    } catch (error) {
      this.erasedPromise = null;
      throw error;
    }
  }

  async editableSourceBytes() {
    const bytes = this.ctx.getOriginalBytes();
    if (!window.onyx?.decryptPdf) return bytes;
    const result = await window.onyx.decryptPdf(bytes);
    if (!result?.ok) throw new Error(result?.error || 'PDF 解密失败');
    return result.encrypted ? result.bytes : bytes;
  }

  async sourceLibDocument() {
    if (this.cachedSourceLibDocument) return this.cachedSourceLibDocument;
    if (!this.sourceLibPromise) {
      const sourceBytes = await this.editableSourceBytes();
      this.sourceLibPromise = window.PDFLib.PDFDocument.load(sourceBytes, {
        ignoreEncryption: true,
        updateMetadata: false
      });
    }
    try {
      this.cachedSourceLibDocument = await this.sourceLibPromise;
      return this.cachedSourceLibDocument;
    } catch (error) {
      this.sourceLibPromise = null;
      throw error;
    }
  }

  previewLineMetrics(run, cssFont, scale = 1) {
    const size = Math.max(1, Number(run.size) || 12);
    const pixels = size * scale;
    const key = `${cssFont}|${run.bold ? 'bold' : 'regular'}|${Math.round(pixels * 100)}`;
    if (this.lineMetrics.has(key)) return this.lineMetrics.get(key);
    const context = document.createElement('canvas').getContext('2d');
    context.font = `${run.bold ? '700' : '400'} ${pixels}px ${cssFont}`;
    const metrics = context.measureText('汉字Ag');
    const ascent = (metrics.fontBoundingBoxAscent || pixels * 0.82) / scale;
    const descent = (metrics.fontBoundingBoxDescent || pixels * 0.4) / scale;
    const value = {
      ascent: Math.max(ascent, size * 0.5),
      contentHeight: Math.max(ascent + descent, size)
    };
    this.lineMetrics.set(key, value);
    return value;
  }

  measureNaturalWidth(element) {
    if (!element.textContent) return 0;
    const probe = element.cloneNode(true);
    probe.querySelectorAll('.content-frame-handle').forEach(handle => handle.remove());
    probe.removeAttribute('data-changed');
    probe.removeAttribute('data-resized');
    probe.removeAttribute('id');
    probe.style.cssText = `${element.style.cssText};position:absolute;left:-99999px;top:0;width:max-content;max-width:none;white-space:pre;visibility:hidden;pointer-events:none;`;
    document.body.appendChild(probe);
    const width = probe.getBoundingClientRect().width / (element._scale || this.ctx.viewer.scale || 1);
    probe.remove();
    return width;
  }

  expandSingleLinePreview(element) {
    this._expandSingleLinePreview(element);
  }

  expandSingleLinePreviewNow(element) {
    if (!element?.isConnected || element.dataset.changed !== '1') return;
    const pageNumber = Number(element.dataset.page);
    const model = this.models.get(pageNumber);
    const block = this.blockFor(element);
    if (!model || !block || block.lineCount !== 1 || element.dataset.resized === '1') return;
    if (element.querySelector('br[data-onyx-source-break]')) return;
    const naturalWidth = this.measureNaturalWidth(element);
    if (naturalWidth <= 0) return;
    const availableWidth = Math.max(1, model.width - block.x);
    const width = Math.min(Math.max(naturalWidth + 1, block.width), availableWidth);
    if (width <= block.width + 0.01) return;
    block.width = width;
    this.applyBlockGeometry(element, block, model, element._scale || this.ctx.viewer.scale);
  }

  async swapPdfDocument(bytes, currentPage = 1) {
    return this.ctx.viewer.replaceDocument(bytes, currentPage);
  }

  async setEnabled(enabled) {
    if (this.committing) return;
    if (!this.ctx.viewer.doc) return;
    if (enabled === this.enabled) return;
    if (enabled) await this.enable();
    else await this.complete();
  }

  async enable() {
    const viewer = this.ctx.viewer;
    const currentPage = viewer.currentPage || 1;
    const erasedPromise = this.buildErasedBytes();
    const modelPromise = this.ensurePageModel(currentPage);
    const erased = await erasedPromise;
    const model = await modelPromise;
    if (model?.unreliableTextCount) {
      toast(`第 ${currentPage} 页有 ${model.unreliableTextCount} 处字体缺少 Unicode 映射，已保留原文不可编辑`);
    }
    this.enabled = true;
    document.body.classList.add('content-editing');
    document.getElementById('btn-edit-mode')?.setAttribute('aria-pressed', 'true');
    await this.swapPdfDocument(erased, viewer.currentPage || 1);
    await this.renderPage(viewer.currentPage || 1);
    document.getElementById('viewer')?.removeEventListener('pointerdown', this._onViewerPointerDown);
    document.getElementById('viewer')?.addEventListener('pointerdown', this._onViewerPointerDown);
    this.showToolbar();
  }

  async ensurePageModel(pageNumber) {
    if (this.models.has(pageNumber)) return this.models.get(pageNumber);
    if (this.modelPromises.has(pageNumber)) return this.modelPromises.get(pageNumber);
    const promise = (async () => {
      const originalDoc = await this.getOriginalDoc();
      if (pageNumber < 1 || pageNumber > originalDoc.numPages) return null;
      const libDocument = await this.sourceLibDocument();
      const model = await extractContentPage(originalDoc, pageNumber, { libDocument });
      model.mode = this.mode;
      const savedEdits = this.allEdits().filter(edit => edit.page === pageNumber);
      this.models.set(pageNumber, applySavedEdits(model, savedEdits));
      return model;
    })();
    this.modelPromises.set(pageNumber, promise);
    try {
      const model = await promise;
      this.modelPromises.delete(pageNumber);
      return model;
    } catch (error) {
      this.modelPromises.delete(pageNumber);
      throw error;
    }
  }

  activeBlocks(pageNumber) {
    const model = this.models.get(pageNumber);
    return model?.activeBlocks || model?.blocks || [];
  }

  async setMode(mode) {
    if (!MODE_LABELS[mode]) return;
    if (document.querySelector('.content-block[data-changed="1"]')) {
      toast('请先完成后切换模式', true);
      this.syncToolbar();
      return;
    }
    this.mode = mode;
    this.selected = null;
    this.selectedRange = null;
    for (const model of this.models.values()) {
      model.mode = mode;
      buildContentViews(model);
    }
    this.removeLayer();
    await this.renderPage(this.ctx.viewer.currentPage || 1);
    this.syncToolbar();
  }

  setBorders(showBorders) {
    this.showBorders = !!showBorders;
    document.body.classList.toggle('content-borders-hidden', !this.showBorders);
  }

  async renderPage(pageNumber, model = this.models.get(pageNumber)) {
    if (!this.enabled) return;
    model = await this.ensurePageModel(pageNumber);
    if (!model) return;
    const box = this.ctx.viewer.pageBoxes.get(pageNumber);
    if (!box) return;
    box.querySelector('.content-editor-layer')?.remove();
    const layer = document.createElement('div');
    layer.className = 'content-editor-layer';
    layer.dataset.page = String(pageNumber);
    for (const block of this.activeBlocks(pageNumber)) {
      layer.appendChild(await this.createBlock(block, pageNumber, this.ctx.viewer.scale));
    }
    for (const image of this.images.get(pageNumber) || []) {
      layer.appendChild(this.createImageElement(image, pageNumber, this.ctx.viewer.scale));
    }
    box.appendChild(layer);
  }

  async registerEmbeddedFont(run) {
    if (run.fontData && !originalFontSupportsText(run)) return run.font;
    if (run.loadedName && document.fonts.check(`12px "${run.loadedName}"`)) return run.loadedName;
    if (!run.fontData || !run.fontKey) return run.font;
    const name = run.loadedName || `onyx-${run.fontKey}`;
    if (this.fontPromises.has(name)) return await this.fontPromises.get(name);
    const promise = (async () => {
      try {
        const face = new FontFace(name, run.fontData instanceof Uint8Array
          ? run.fontData.buffer.slice(run.fontData.byteOffset, run.fontData.byteOffset + run.fontData.byteLength)
          : run.fontData);
        await face.load();
        document.fonts.add(face);
        return name;
      } catch (_) {
        return run.font;
      }
    })();
    this.fontPromises.set(name, promise);
    return await promise;
  }

  async createBlock(block, pageNumber, scale) {
    const element = document.createElement('div');
    const model = this.models.get(pageNumber);
    const availableWidth = model.width - block.x;
    let width = clamp(block.width || availableWidth, 1, Math.max(availableWidth, 1));
    const height = block.height || block.lineHeight || block.size * 1.22;
    element.className = 'content-block';
    element.contentEditable = 'true';
    element.spellcheck = false;
    element.dataset.page = String(pageNumber);
    element.dataset.blockKey = block.id;
    element.dataset.kind = block.kind || 'paragraph';
    element.style.cssText = [
      `left:${block.x * scale}px`,
      `top:${this.blockTopPt(block, model) * scale}px`,
      `width:${width * scale}px`,
      `min-height:${height * scale}px`,
      `line-height:${block.lineHeight * scale}px`,
      `text-align:${block.alignment || 'left'}`,
      block.patch ? `background:${block.patchColor || '#ffffff'}` : ''
    ].join(';');

    const sourceLines = (block.mode === 'paragraph' || (!block.mode && this.mode === 'paragraph')) &&
        (!Array.isArray(block.sourceLineRuns) || !block.sourceLineRuns.length)
      ? [block.runs]
      : (Array.isArray(block.sourceLineRuns) && block.sourceLineRuns.length
        ? block.sourceLineRuns
        : [block.runs]);
    const runs = [];
    let firstLineAscent = 0;
    let maxContentHeight = 0;
    for (const [lineIndex, lineRuns] of sourceLines.entries()) {
      if (lineIndex) {
        const lineBreak = document.createElement('br');
        lineBreak.dataset.onyxSourceBreak = '1';
        runs.push(lineBreak);
      }
      for (const sourceRun of lineRuns) {
        for (const run of splitRunBySupportedGlyphs(sourceRun)) {
          const cssFont = await this.registerEmbeddedFont(run);
          const metrics = this.previewLineMetrics(run, `"${cssFont}","${fontFamilyName(run.font)}","Microsoft YaHei",sans-serif`, scale);
          if (lineIndex === 0) firstLineAscent = Math.max(firstLineAscent, metrics.ascent);
          maxContentHeight = Math.max(maxContentHeight, metrics.contentHeight);
          runs.push(await this.createRunSpan(run, scale, cssFont));
        }
      }
    }
    if (runs.length) element.append(...runs);
    else element.textContent = block.text || '';

    const naturalWidth = this.measureNaturalWidth(element);
    if (naturalWidth + 2 > width + 0.05) {
      width = clamp(naturalWidth + 2, width, availableWidth);
      block.width = width;
      element.style.width = `${width * scale}px`;
      element._widthPt = width;
    }

    if (firstLineAscent) block.previewAscent = firstLineAscent;
    if (maxContentHeight) block.previewContentHeight = maxContentHeight;

    element.addEventListener('pointerdown', () => this.selectBlock(element));
    element.addEventListener('keydown', event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.selectBlock(null);
      }
    });
    ['keyup', 'mouseup'].forEach(type => element.addEventListener(type, () => this.storeSelection(element)));
    element.addEventListener('focus', () => {
      this.selectBlock(element, false);
      this.storeSelection(element);
    });
    element.addEventListener('beforeinput', () => {
      element.querySelectorAll('br[data-onyx-source-break]').forEach(lineBreak => lineBreak.remove());
    });
    element.addEventListener('input', () => {
      element.dataset.changed = '1';
      this._syncToolbarIdle();
    });
    element.addEventListener('beforeinput', event => {
      if (event.inputType !== 'insertText' || !event.data) return;
      const selection = window.getSelection();
      if (!selection.rangeCount) return;
      const range = selection.getRangeAt(0);
      if (!element.contains(range.commonAncestorContainer)) return;
      const span = this.caretRunSpan(element, range);
      if (!span) return;
      event.preventDefault();
      if (!range.collapsed) range.deleteContents();
      const selectionAfterInsert = window.getSelection();
      const currentRange = selection.rangeCount ? selection.getRangeAt(0) : range;
      let textNode = currentRange.startContainer;
      let offset = currentRange.startOffset;
      if (textNode.nodeType !== Node.TEXT_NODE || !span.contains(textNode)) {
        textNode = document.createTextNode('');
        span.appendChild(textNode);
        offset = 0;
      }
      textNode.insertData(offset, event.data);
      const nextRange = document.createRange();
      nextRange.setStart(textNode, offset + event.data.length);
      nextRange.collapse(true);
      selectionAfterInsert.removeAllRanges();
      selectionAfterInsert.addRange(nextRange);
      element.dataset.changed = '1';
      this._syncToolbarIdle();
    });
    element.addEventListener('paste', event => {
      event.preventDefault();
      document.execCommand('insertText', false, event.clipboardData?.getData('text/plain') || '');
    });
    element._scale = scale;
    element._widthPt = width;
    element._heightPt = height;
    if (this.selected?.dataset.blockKey === block.id && Number(this.selected.dataset.page) === pageNumber) {
      this.selected = element;
      element.classList.add('selected');
      this.addFrame(element, block, pageNumber, scale);
    }
    return element;
  }

  async createRunSpan(run, scale, cssFont) {
    const span = document.createElement('span');
    span.style.cssText = runStyle({ ...run, scale, cssFontFamily: `"${cssFont}","${fontFamilyName(run.font)}","Microsoft YaHei",sans-serif` });
    span._onyxFontName = run.font || null;
    span._onyxFontKey = run.fontKey || null;
    span._onyxFontData = run.fontData || null;
    span._onyxCssFont = run.loadedName || cssFont || null;
    span._onyxSystemFont = !run.fontData && !run.fontKey;
    span.textContent = run.text;
    return span;
  }

  caretRunSpan(element, range) {
    if (range.startContainer.nodeType === Node.TEXT_NODE) {
      const span = range.startContainer.parentElement;
      return span && span !== element && element.contains(span) &&
        !span.classList.contains('content-frame-handle') ? span : null;
    }
    if (range.startContainer !== element) return null;
    const previous = range.startContainer.childNodes[range.startOffset - 1];
    const next = range.startContainer.childNodes[range.startOffset];
    for (const node of [previous, next]) {
      if (node?.nodeType === Node.ELEMENT_NODE && node.tagName === 'SPAN' &&
        !node.classList.contains('content-frame-handle')) return node;
    }
    return null;
  }

  selectBlock(element, focus = true) {
    document.querySelectorAll('.content-block.selected').forEach(item => {
      item.classList.remove('selected');
      item.querySelectorAll('.content-frame-handle').forEach(handle => handle.remove());
    });
    this.selected = element;
    if (!element) {
      this.selectedRange = null;
      this.syncToolbar();
      return;
    }
    element.classList.add('selected');
    const block = this.blockFor(element);
    if (block) this.addFrame(element, block, Number(element.dataset.page), element._scale || this.ctx.viewer.scale);
    if (focus) element.focus({ preventScroll: true });
    this.syncToolbar();
  }

  addFrame(element, block, pageNumber, scale) {
    const model = this.models.get(pageNumber);
    const moveHandle = document.createElement('span');
    moveHandle.className = 'content-frame-handle move';
    moveHandle.dataset.frameAction = 'move';
    moveHandle.title = '移动文本框';
    moveHandle.contentEditable = 'false';
    moveHandle.addEventListener('pointerdown', event => this.startFrameDrag(event, element, block, model, 'move', scale));
    element.appendChild(moveHandle);
    for (const action of ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se']) {
      const handle = document.createElement('span');
      handle.className = `content-frame-handle ${action}`;
      handle.dataset.frameAction = action;
      handle.contentEditable = 'false';
      handle.addEventListener('pointerdown', event => this.startFrameDrag(event, element, block, model, action, scale));
      element.appendChild(handle);
    }
  }

  startFrameDrag(event, element, block, model, action, scale) {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startY = event.clientY;
    const original = {
      x: block.x,
      baseline: block.baseline,
      width: block.width,
      height: block.height || block.lineHeight || block.size * 1.22
    };
    const onMove = moveEvent => {
      const dx = (moveEvent.clientX - startX) / scale;
      const dy = (moveEvent.clientY - startY) / scale;
      let left = original.x;
      let right = original.x + original.width;
      let baseline = original.baseline;
      let height = original.height;
      if (action === 'move') {
        const maximumLeft = Math.max(0, model.width - original.width);
        left = clamp(original.x + dx, 0, maximumLeft);
        baseline = clamp(original.baseline - dy, 6, model.height - 6);
        right = left + original.width;
      } else {
        if (action.includes('w')) left = clamp(original.x + dx, 0, original.x + original.width - 12);
        if (action.includes('e')) right = Math.max(left + 12, original.x + original.width + dx);
        if (action.includes('n')) {
          baseline = clamp(original.baseline - dy, 6, model.height - 6);
          height = Math.max(10, original.height - dy);
        }
        if (action.includes('s')) height = Math.max(10, original.height + dy);
      }
      block.x = left;
      block.baseline = baseline;
      block.width = clamp(right - left, 12, Math.max(12, model.width - left));
      block.height = height;
      element.dataset.changed = '1';
      if (action.includes('e') || action.includes('w')) element.dataset.resized = '1';
      this.applyBlockGeometry(element, block, model, scale);
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  }

  applyBlockGeometry(element, block, model, scale = element._scale || this.ctx.viewer.scale) {
    const width = clamp(block.width, 1, Math.max(1, model.width - block.x));
    const height = block.height || block.lineHeight || block.size * 1.22;
    element.style.left = `${block.x * scale}px`;
    element.style.top = `${this.blockTopPt(block, model) * scale}px`;
    element.style.width = `${width * scale}px`;
    element.style.minHeight = `${height * scale}px`;
    element._widthPt = width;
    element._heightPt = height;
  }

  storeSelection(element) {
    const range = selectionRange(element);
    if (range) {
      this.selectedRange = {
        element,
        start: globalOffset(element, range.startContainer, range.startOffset),
        end: globalOffset(element, range.endContainer, range.endOffset)
      };
      element._onyxSelectedRange = this.selectedRange;
    }
    this.selectBlock(element, false);
  }

  removeLayer() {
    document.querySelectorAll('.content-editor-layer').forEach(layer => layer.remove());
  }

  _onViewerPointerDown = event => {
    if (event.target.closest('.content-block, .content-image, #content-editor-bar')) return;
    if (!this.pendingImage || event.button !== 0) return;
    const outer = event.target.closest('.page-outer');
    const canvas = outer?.querySelector('canvas');
    if (!outer || !canvas) return;
    event.preventDefault();
    event.stopPropagation();
    const rect = canvas.getBoundingClientRect();
    this.placeImage(Number(outer.dataset.page), event.clientX - rect.left, event.clientY - rect.top, this.pendingImage);
    this.pendingImage = null;
  };

  async prepareImageFile(file) {
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error('图片读取失败'));
      reader.readAsDataURL(file);
    });
    const image = new Image();
    await new Promise((resolve, reject) => {
      image.onload = resolve;
      image.onerror = () => reject(new Error('图片格式不支持'));
      image.src = dataUrl;
    });
    let width = image.naturalWidth * 0.75;
    let height = image.naturalHeight * 0.75;
    if (width > 300) {
      height *= 300 / width;
      width = 300;
    }
    this.pendingImage = { dataUrl, width, height };
    toast('在页面点击位置插入图片');
  }

  async placeImage(pageNumber, cssX, cssY, pending) {
    const page = await this.ctx.viewer.doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: this.ctx.viewer.scale, rotation: this.ctx.viewer.rotation });
    const point = viewport.convertToPdfPoint(cssX, cssY);
    const image = {
      type: 'image',
      id: `image-${Date.now()}-${Math.random().toString(16).slice(2)}`,
      page: pageNumber,
      x: point[0],
      y: point[1] - pending.height,
      w: pending.width,
      h: pending.height,
      dataUrl: pending.dataUrl,
      created: Date.now()
    };
    if (!this.images.has(pageNumber)) this.images.set(pageNumber, []);
    this.images.get(pageNumber).push(image);
    this.selectedImage = image;
    this.renderPage(pageNumber);
  }

  createImageElement(image, pageNumber, scale) {
    const model = this.models.get(pageNumber);
    const item = document.createElement('div');
    item.className = 'content-image' + (this.selectedImage?.id === image.id ? ' selected' : '');
    item.dataset.page = String(pageNumber);
    item.style.cssText = [
      `left:${image.x * scale}px`,
      `top:${(model.height - image.y - image.h) * scale}px`,
      `width:${image.w * scale}px`,
      `height:${image.h * scale}px`
    ].join(';');
    const picture = document.createElement('img');
    picture.src = image.dataUrl;
    item.appendChild(picture);
    for (const direction of ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se']) {
      const handle = document.createElement('span');
      handle.className = 'image-handle';
      handle.dataset.direction = direction;
      item.appendChild(handle);
    }
    item.addEventListener('pointerdown', event => {
      event.preventDefault();
      event.stopPropagation();
      this.selectedImage = image;
      const direction = event.target.dataset.direction;
      let shiftPressed = event.shiftKey;
      const syncShift = keyEvent => { shiftPressed = keyEvent.shiftKey; };
      const startX = event.clientX;
      const startY = event.clientY;
      const original = { x: image.x, y: image.y, w: image.w, h: image.h };
      const onMove = moveEvent => {
        const dx = (moveEvent.clientX - startX) / scale;
        const dy = (moveEvent.clientY - startY) / scale;
        const preserveRatio = (moveEvent.shiftKey || shiftPressed) && direction;
        if (!direction) {
          image.x = original.x + dx;
          image.y = original.y - dy;
        } else {
          let left = original.x;
          let right = original.x + original.w;
          let top = original.y + original.h;
          let bottom = original.y;
          if (direction.includes('w')) left = original.x + dx;
          if (direction.includes('e')) right = original.x + original.w + dx;
          if (direction.includes('n')) top = original.y + original.h - dy;
          if (direction.includes('s')) bottom = Math.min(original.y - dy, top - 12);
          let width = Math.max(12, right - left);
          let height = Math.max(12, top - bottom);
          if (preserveRatio) {
            const ratio = original.w / original.h;
            if (direction.includes('n') || direction.includes('s')) {
              width = Math.max(12, height * ratio);
              if (direction.includes('w')) left = right - width;
            } else {
              height = Math.max(12, width / ratio);
              if (direction.includes('n')) bottom = top - height;
            }
          }
          image.x = left;
          image.w = width;
          image.h = height;
          image.y = bottom;
        }
        this.applyImageGeometry(item, image, scale);
      };
      const onUp = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
      window.removeEventListener('keydown', syncShift, true);
      window.removeEventListener('keyup', syncShift, true);
      };
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
    window.addEventListener('keydown', syncShift, true);
    window.addEventListener('keyup', syncShift, true);
    });
    item.addEventListener('keydown', event => {
      if (event.key === 'Delete' && this.selectedImage === image) {
        event.preventDefault();
        this.images.set(pageNumber, (this.images.get(pageNumber) || []).filter(entry => entry.id !== image.id));
        this.selectedImage = null;
        this.renderPage(pageNumber);
      }
    });
    item.tabIndex = 0;
    item._scale = scale;
    return item;
  }

  applyImageGeometry(element, image, scale = element._scale || this.ctx.viewer.scale) {
    const pageNumber = Number(element.dataset.page);
    const model = this.models.get(pageNumber);
    if (!model) return;
    element.style.left = `${image.x * scale}px`;
    element.style.top = `${(model.height - image.y - image.h) * scale}px`;
    element.style.width = `${image.w * scale}px`;
    element.style.height = `${image.h * scale}px`;
  }

  showToolbar() {
    this.hideToolbar();
    const bar = document.createElement('div');
    bar.id = 'content-editor-bar';
    bar.innerHTML = `
      <select data-f="mode" title="编辑模式">
        ${Object.entries(MODE_LABELS).map(([value, label]) => `<option value="${value}">${label}</option>`).join('')}
      </select>
      <label class="border-toggle"><input type="checkbox" data-f="borders" checked>内容边框</label>
      <span class="separator"></span>
      <select data-f="font" title="字体">
        ${['微软雅黑', '宋体', '黑体', '楷体', '仿宋', '等线', 'Arial', 'Times New Roman', 'Calibri', 'Segoe UI'].map(font => `<option value="${font}">${font}</option>`).join('')}
      </select>
      <input data-f="size" type="number" min="6" max="96" step="0.1" title="字号">
      <label class="number-field" title="行距（倍数）"><span>行距</span><input data-f="line-height" type="number" min="0.5" max="4" step="0.05"></label>
      <button data-f="bold" title="粗体"><b>B</b></button>
      <button data-f="alignment" data-alignment="left" title="居左">⯇</button>
      <button data-f="alignment" data-alignment="center" title="居中">≡</button>
      <button data-f="alignment" data-alignment="right" title="居右">⯈</button>
      <input data-f="color" type="color" value="#000000" title="颜色">
      <span class="separator"></span>
      <button data-f="image" title="插入图片">插入图片</button>
      <button data-f="done" class="primary">完成</button>`;
    document.body.appendChild(bar);
    bar.style.top = `${Math.max(8, (document.getElementById('toolbar')?.getBoundingClientRect().bottom || 48) + 6)}px`;
    bar.style.left = `${Math.max(8, (window.innerWidth - bar.offsetWidth) / 2)}px`;
    bar.querySelector('[data-f="mode"]').value = this.mode;
    bar.querySelector('[data-f="mode"]').addEventListener('change', event => this.setMode(event.target.value));
    bar.querySelector('[data-f="borders"]').addEventListener('change', event => this.setBorders(event.target.checked));
    const font = bar.querySelector('[data-f="font"]');
    font.addEventListener('change', event => this.formatSelection({ font: event.target.value }));
    const size = bar.querySelector('[data-f="size"]');
    const applySize = () => {
      const parsed = Number.parseFloat(size.value);
      if (!Number.isFinite(parsed)) return;
      const value = clamp(parsed, 6, 96);
      size.value = String(value);
      this.formatSelection({ size: value });
    };
    size.addEventListener('change', applySize);
    size.addEventListener('keydown', event => {
      if (event.key === 'Enter') {
        event.preventDefault();
        applySize();
      }
    });
    const lineHeight = bar.querySelector('[data-f="line-height"]');
    const applyLineHeight = () => {
      const element = this.activeElement();
      if (!element) return;
      const rawValue = parseFloat(lineHeight.value);
      const block = this.blockFor(element);
      if (!Number.isFinite(rawValue)) {
        lineHeight.value = String(lineHeightRatio(block));
        return;
      }
      const value = clamp(rawValue, 0.5, 4);
      lineHeight.value = String(value);
      this.formatSelection({ lineHeight: lineHeightFromRatio(value, block) });
    };
    lineHeight.addEventListener('change', applyLineHeight);
    lineHeight.addEventListener('keydown', event => {
      if (event.key === 'Enter') {
        event.preventDefault();
        applyLineHeight();
      }
    });
    bar.querySelector('[data-f="bold"]').addEventListener('click', () => {
      const style = this.selectionStyle();
      this.formatSelection({ bold: !(style.bold) });
    });
    bar.querySelectorAll('[data-f="alignment"]').forEach(button => button.addEventListener('click', () => {
      this.formatSelection({ alignment: button.dataset.alignment });
    }));
    bar.querySelector('[data-f="color"]').addEventListener('input', event => this.formatSelection({ color: event.target.value }));
    bar.querySelector('[data-f="image"]').addEventListener('click', () => {
      const input = document.getElementById('img-file-input');
      input.value = '';
      input.click();
    });
    bar.querySelector('[data-f="done"]').addEventListener('click', () => this.setEnabled(false));
    this.toolbar = bar;
    this.syncToolbar();
  }

  hideToolbar() {
    document.getElementById('content-editor-bar')?.remove();
    this.toolbar = null;
  }

  detachViewerEvents() {
    document.getElementById('viewer')?.removeEventListener('pointerdown', this._onViewerPointerDown);
  }

  activeElement() {
    return this.selected?.isConnected ? this.selected : document.querySelector('.content-block:focus');
  }

  blockFor(element) {
    const model = this.models.get(Number(element.dataset.page));
    return this.activeBlocks(Number(element.dataset.page)).find(block => block.id === element.dataset.blockKey) || null;
  }

  blockTopPt(block, model) {
    const size = block.size || 12;
    const lineHeight = block.lineHeight || size * 1.22;
    const contentHeight = block.previewContentHeight || size * 1.22;
    const ascent = block.previewAscent || (block.ascent || 0.82) * size;
    const leading = (lineHeight - contentHeight) / 2;
    return model.height - block.baseline - ascent - leading;
  }

  contentAlignmentFrame(block, pageNumber) {
    const model = this.models.get(pageNumber);
    if (!model) return null;
    const blocks = this.activeBlocks(pageNumber).filter(candidate =>
      candidate.text && Number(candidate.width) >= 12 && !candidate.patch);
    let left;
    let right;
    if (blocks.length >= 5) {
      left = quantile(blocks.map(candidate => candidate.x).sort((a, b) => a - b), 0.12);
      right = quantile(blocks.map(candidate => candidate.x + candidate.width)
        .sort((a, b) => a - b), 0.88);
    } else {
      left = model.width * 0.06;
      right = model.width * 0.94;
    }
    left = Math.min(left, block.x);
    right = Math.max(right, block.x + block.width);
    left = clamp(left, 0, Math.max(0, model.width - 12));
    right = clamp(right, left + Math.max(block.width, 12), model.width);
    return { x: left, width: right - left };
  }

  formatSelection(patch) {
    const element = this.activeElement();
    if (!element) {
      toast('请先点击文字块', true);
      return;
    }
    if (patch.alignment) {
      const block = this.blockFor(element);
      if (block) {
        block.alignment = patch.alignment;
        if (element.dataset.resized !== '1' && block.manualWidth !== true) {
          const frame = this.contentAlignmentFrame(block, Number(element.dataset.page));
          if (frame) {
            block.x = frame.x;
            block.width = frame.width;
            this.applyBlockGeometry(element, block, this.models.get(Number(element.dataset.page)),
              element._scale || this.ctx.viewer.scale);
          }
        }
      }
      element.style.textAlign = patch.alignment;
      element.dataset.changed = '1';
      element.focus({ preventScroll: true });
      if (this.selectedRange?.element === element && this.selectedRange.end > this.selectedRange.start) {
        restoreSelection(element, this.selectedRange);
      }
      this.syncToolbar();
      return;
    }
    if (Object.prototype.hasOwnProperty.call(patch, 'lineHeight')) {
      const block = this.blockFor(element);
      if (block) {
        block.lineHeight = clamp(patch.lineHeight, 4, 192);
        element.style.lineHeight = `${block.lineHeight * (element._scale || this.ctx.viewer.scale)}px`;
        element.style.top = `${this.blockTopPt(block, this.models.get(Number(element.dataset.page))) * (element._scale || this.ctx.viewer.scale)}px`;
        element.dataset.changed = '1';
      }
      element.focus({ preventScroll: true });
      this.syncToolbar();
      return;
     }
    if (!this.selectedRange || this.selectedRange.element !== element) this.storeSelection(element);
    const block = this.blockFor(element) || {};
    const fallback = block.runs?.[0] || block;
    const changed = applyFormatToSelection(element, patch, { ...fallback, scale: element._scale || this.ctx.viewer.scale });
    if (changed) {
      element.dataset.changed = '1';
    }
    element.focus({ preventScroll: true });
    if (this.selectedRange?.element === element && this.selectedRange.end > this.selectedRange.start) {
      restoreSelection(element, this.selectedRange);
    }
    this.syncToolbar();
  }

  selectionStyle() {
    const element = this.activeElement();
    if (!element) return {};
    const range = this.selectedRange?.element === element && this.selectedRange.end > this.selectedRange.start
      ? this.selectedRange
      : null;
    let nodes;
    if (range) {
      nodes = textNodes(element).filter(node => {
        const start = globalOffset(element, node, 0);
        const end = start + node.textContent.length;
        return end > range.start && start < range.end;
      });
    } else {
      nodes = [textNodes(element)[0]].filter(Boolean);
    }
    if (!nodes.length) return {};
    const styles = nodes.map(node => {
      const style = node.parentElement?.style || element.style;
      return {
        font: toolbarFontFamilyName(node.parentElement?._onyxFontName ||
          style.fontFamily || getComputedStyle(node.parentElement || element).fontFamily),
        size: Math.round((cssFontSizePt(style.fontSize, element._scale || this.ctx.viewer.scale) || 12) * 10) / 10,
        color: style.color || '',
        bold: Number(style.fontWeight || 400) >= 600
      };
    });
    const same = key => styles.every(style => style[key] === styles[0][key]);
    const result = {
      font: same('font') ? styles[0].font : '',
      size: same('size') ? styles[0].size : '',
      color: same('color') ? styles[0].color : '',
      bold: styles[0].bold
    };
    if (range) this.selectedRange = range;
    return result;
  }

  syncToolbar() {
    const bar = this.toolbar || document.getElementById('content-editor-bar');
    if (!bar) return;
    bar.querySelector('[data-f="mode"]').value = this.mode;
    bar.querySelector('[data-f="borders"]').checked = this.showBorders;
    const element = this.activeElement();
    if (!element) return;
    const style = this.selectionStyle();
    const fontSelector = bar.querySelector('[data-f="font"]');
    if (style.font) {
      if (![...fontSelector.options].some(option => option.value === style.font)) {
        const option = document.createElement('option');
        option.value = style.font;
        option.textContent = style.font;
        fontSelector.prepend(option);
      }
      fontSelector.value = style.font;
    } else {
      fontSelector.value = '';
    }
    const sizeControl = bar.querySelector('[data-f="size"]');
    sizeControl.value = style.size || '';
    if (style.color.startsWith('#')) bar.querySelector('[data-f="color"]').value = style.color;
    else if (style.color) {
      const rgb = style.color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
      if (rgb) bar.querySelector('[data-f="color"]').value = '#' + rgb.slice(1).map(value => Number(value).toString(16).padStart(2, '0')).join('');
    }
    bar.querySelector('[data-f="bold"]').style.background = style.bold ? '#e34d3f' : '';
    const alignment = this.blockFor(element)?.alignment || element.style.textAlign || 'left';
    bar.querySelectorAll('[data-f="alignment"]').forEach(button => {
      button.style.background = button.dataset.alignment === alignment ? '#e34d3f' : '';
    });
    bar.querySelector('[data-f="line-height"]').value = String(lineHeightRatio(this.blockFor(element)));
  }

  commitDom() {
    let changed = 0;
    for (const element of document.querySelectorAll('.content-block[data-changed="1"]')) {
      const pageNumber = Number(element.dataset.page);
      const model = this.models.get(pageNumber);
      const block = this.blockFor(element);
      if (!model || !block) continue;
      const collectedRuns = collectRuns(element, {
        font: block.runs[0]?.font || block.font || '微软雅黑',
        fontKey: block.runs[0]?.fontKey || block.fontKey || null,
        loadedName: block.runs[0]?.loadedName || block.loadedName || null,
        fontData: block.runs[0]?.fontData || block.fontData || null,
        size: block.runs[0]?.size || block.size || 12,
        color: block.runs[0]?.color || block.color || '#000000',
        bold: block.runs[0]?.bold ?? false,
        italic: block.runs[0]?.italic ?? false,
        underline: block.runs[0]?.underline ?? false
      }, element._scale || this.ctx.viewer.scale);
      const runs = collectedRuns.flatMap(run => splitRunBySupportedGlyphs(run));
      if (!runs.some(run => run.text)) continue;
      block.runs = runs.map(run => ({ ...run }));
      block.text = runs.map(run => run.text).join('');
      block.sourceLineRuns = [block.runs.map(run => ({ ...run }))];
      block.width = element._widthPt || block.width;
      block.height = element._heightPt || block.height;
      const manualWidth = element.dataset.resized === '1';
      const saved = this.edits.get(block.id);
      this.edits.set(block.id, {
        engine: 2,
        page: pageNumber,
        blockKey: block.id,
        x: block.x,
        baseline: block.baseline,
        lineHeight: block.lineHeight,
        patch: block.patch ? { ...block.patch } : null,
        patchColor: block.patchColor || '#ffffff',
        alignment: block.alignment || 'left',
        size: block.size,
        height: block.height,
        maxWidth: block.width,
        manualWidth,
        lineCount: block.lineCount,
        runs: block.runs.map(run => ({ ...run })),
        sourceOps: [...new Set([...(saved?.sourceOps || []), ...block.sourceOps])]
      });
      delete element.dataset.changed;
      changed++;
    }
    if (changed) this.persist();
    return changed;
  }

  buildPageReplacementEdits() {
    return this.allEdits().filter(edit => edit.engine === 2);
  }

  async buildChangedBytes(pageEdits = this.allEdits()) {
    const { bytes } = await exportAnnotatedPdf(
      this.ctx.getOriginalBytes(),
      this.allImages(),
      'onyx-content-edited.pdf',
      pageEdits
    );
    return bytes;
  }

  finishUi() {
    this.enabled = false;
    document.body.classList.remove('content-editing');
    document.getElementById('btn-edit-mode')?.setAttribute('aria-pressed', 'false');
    this.hideToolbar();
    this.removeLayer();
    this.detachViewerEvents();
    this.selected = null;
    this.selectedRange = null;
  }

  async complete() {
    if (this.committing) return;
    this.committing = true;
    const currentPage = this.ctx.viewer.currentPage || 1;
    try {
      this.commitDom();
      const changedCount = this.allEdits().length;
      const pageEdits = this.buildPageReplacementEdits();
      const hasChanges = !!pageEdits.length || !!this.allImages().length;
      if (!hasChanges) {
        await this.swapPdfDocument(new Uint8Array(this.ctx.getOriginalBytes()), currentPage);
        this.finishUi();
        return;
      }
      const bytes = await this.buildChangedBytes(pageEdits);
      await this.swapPdfDocument(bytes, currentPage);
      await this.ctx.viewer.doc.cleanup();
      this.edits = new Map(pageEdits.map(edit => [edit.blockKey, edit]));
      this.models.clear();
      this.modelPromises.clear();
      this.finishUi();
      toast(`已完成 ${changedCount + this.allImages().length} 个内容修改`);
    } catch (error) {
      console.error('完成内容编辑失败:', error);
      toast('完成编辑失败：' + (error?.message || error), true);
      await this.swapPdfDocument(new Uint8Array(this.ctx.getOriginalBytes()), currentPage);
      this.finishUi();
    } finally {
      this.committing = false;
    }
  }

  async showCommitted() {
    if (!this.edits.size || !this.ctx.viewer.doc) return;
    await this.swapPdfDocument(await this.buildChangedBytes(this.buildPageReplacementEdits()), this.ctx.viewer.currentPage || 1);
  }
}
