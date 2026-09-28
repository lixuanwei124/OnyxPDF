import { ANNOT_COLORS, clamp, toast } from './utils.js';
import { openModal } from './ui.js';
import * as pdfjsLib from '../vendor/pdf.min.mjs';
import { buildEditBlocks } from './pdf2edit.js';
import { exportAnnotatedPdf } from './exporter.js';

function uid(prefix = 'a') {
  return `${prefix}${Date.now()}${Math.floor(Math.random() * 1000000)}`;
}

function fontFamilyName(value) {
  return String(value || '微软雅黑').split(',')[0].replace(/["']/g, '').trim() || '微软雅黑';
}

function normalizeCssColor(value) {
  const text = String(value || '').trim();
  if (/^#[0-9a-f]{3,8}$/i.test(text)) return text;
  const match = text.match(/rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (match) return '#' + match.slice(1).map(item => Number(item).toString(16).padStart(2, '0')).join('');
  return '#000000';
}

function runCss(run, scale) {
  return [
    `font-family:"${fontFamilyName(run.font)}","Microsoft YaHei",sans-serif`,
    `font-size:${(Number(run.size) || 12) * scale}px`,
    `color:${normalizeCssColor(run.color)}`,
    `font-weight:${run.bold ? '700' : '400'}`
  ].join(';') + ';';
}

function renderRunHtml(runs, scale) {
  return (runs || []).map(run => {
    const text = String(run.text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const css = runCss(run, scale).replace(/"/g, '&quot;');
    return `<span style="${css}">${text}</span>`;
  }).join('');
}

function walkTextNodes(root, visitor) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  nodes.forEach(visitor);
}

function collectTextRuns(root, scale, fallback) {
  const runs = [];
  const readNode = (node, inherited = {}) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent.replace(/\u00a0/g, ' ');
      if (text) runs.push({ ...inherited, text });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    if (node.tagName === 'BR') {
      runs.push({ ...inherited, text: '\n' });
      return;
    }
    const style = node.style || {};
    const next = {
      font: fontFamilyName(style.fontFamily || inherited.font || fallback.font),
      size: parseFloat(style.fontSize) / scale || inherited.size || fallback.size,
      color: normalizeCssColor(style.color || inherited.color || fallback.color),
      bold: style.fontWeight ? parseInt(style.fontWeight, 10) >= 600 : !!inherited.bold
    };
    for (const child of node.childNodes) readNode(child, next);
  };
  for (const child of root.childNodes) readNode(child, {});

  const output = [];
  for (const run of runs) {
    const signature = JSON.stringify([run.font, Math.round(run.size * 10) / 10, run.color, !!run.bold]);
    const previous = output[output.length - 1];
    if (previous && previous.__signature === signature) previous.text += run.text;
    else output.push({ ...run, __signature: signature });
  }
  output.forEach(run => delete run.__signature);
  return output;
}

function selectionOffsets(root) {
  const selection = window.getSelection();
  if (!selection.rangeCount) return null;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return null;
  const position = (container, offset, toEnd) => {
    const probe = document.createRange();
    probe.selectNodeContents(root);
    if (toEnd) probe.setEnd(container, offset);
    else probe.setStart(container, offset);
    return probe.toString().length;
  };
  return {
    start: position(range.startContainer, range.startOffset, false),
    end: position(range.endContainer, range.endOffset, true)
  };
}

function setTextSelection(root, start, end) {
  let range = document.createRange();
  let before = 0;
  let placedStart = false;
  let placedEnd = false;
  walkTextNodes(root, node => {
    const length = node.textContent.length;
    if (!placedStart && before + length >= start) {
      range.setStart(node, Math.max(0, start - before));
      placedStart = true;
    }
    if (!placedEnd && before + length >= end) {
      range.setEnd(node, Math.max(0, end - before));
      placedEnd = true;
    }
    before += length;
  });
  if (!placedStart || !placedEnd) {
    range = document.createRange();
    range.selectNodeContents(root);
  }
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

function splitTextRange(range) {
  if (range.startContainer.nodeType === Node.TEXT_NODE && range.startOffset > 0) {
    const node = range.startContainer.splitText(range.startOffset);
    range.setStart(node, 0);
  }
  if (range.endContainer.nodeType === Node.TEXT_NODE && range.endOffset < range.endContainer.textContent.length) {
    range.endContainer.splitText(range.endOffset);
  }
}

function styleRange(editor, range, patch, fallback, scale) {
  splitTextRange(range);
  const nodes = [];
  const collect = root => walkTextNodes(root, node => {
    if (node.isConnected && range.intersectsNode(node) && !nodes.includes(node)) nodes.push(node);
  });
  collect(range.commonAncestorContainer);
  collect(editor);
  for (const node of nodes) {
    if (!node.textContent) continue;
    const computed = getComputedStyle(node.parentElement || editor);
    const current = {
      font: fontFamilyName(computed.fontFamily),
      size: (parseFloat(computed.fontSize) || fallback.size * scale) / scale,
      color: normalizeCssColor(computed.color),
      bold: parseInt(computed.fontWeight, 10) >= 600
    };
    const parent = node.parentElement;
    if (parent && parent !== editor && parent.style && parent.childNodes.length === 1 && parent.firstChild === node) {
      parent.style.cssText = runCss({ ...current, ...patch }, scale);
    } else {
      const span = document.createElement('span');
      span.style.cssText = runCss({ ...current, ...patch }, scale);
      node.parentNode.insertBefore(span, node);
      span.appendChild(node);
    }
  }
  return nodes.length;
}

function applyFormatToRuns(editor, offsets, patch, fallback, scale) {
  const sourceRuns = collectTextRuns(editor, scale, fallback);
  const selectedStart = Math.max(0, Math.min(offsets.start, offsets.end));
  const selectedEnd = Math.min(editor.textContent.length, Math.max(offsets.start, offsets.end));
  const output = [];
  let cursor = 0;

  for (const run of sourceRuns) {
    const start = cursor;
    const end = cursor + run.text.length;
    cursor = end;
    const beforeText = run.text.slice(0, Math.max(0, selectedStart - start));
    const selectedText = run.text.slice(Math.max(0, selectedStart - start), Math.max(0, selectedEnd - start));
    const afterText = run.text.slice(Math.max(0, selectedEnd - start));
    if (beforeText) output.push({ ...run, text: beforeText });
    if (selectedText) output.push({ ...run, ...patch, text: selectedText });
    if (afterText) output.push({ ...run, text: afterText });
  }

  const normalized = output.filter(run => run.text);
  editor.innerHTML = renderRunHtml(normalized, scale);
  return normalized;
}

export class Annotations {
  constructor(ctx) {
    this.ctx = ctx;
    this.editMode = false;
    this.map = new Map();
    this.tool = 'select';
    this.elementSnapshots = new Map();
    this._lineCache = new Map();
    this._originalDoc = null;
    this._originalDocPromise = null;
    this._paraFmt = new Map();
    this._blockMoves = new Map();
    this._overlaySignatures = new Map();
    this._overlayGenerations = new Map();
    this._bindTools();
    this._bindPointer();
    this._bindNoteEditor();
  }

  loadFor(fileKey) {
    if (this.editMode) {
      this.editMode = false;
      this._hideToolbar();
      this.removeElementOverlay();
      document.body.classList.remove('edit-mode');
      document.getElementById('btn-edit-mode')?.setAttribute('aria-pressed', 'false');
    }
    this.selectedElement = null;
    this._lastSelection = null;
    this._selectedRange = null;
    this.map.clear();
    this.elementSnapshots.clear();
    this._lineCache.clear();
    this._destroyOriginalDoc();
    this._paraFmt.clear();
    this._blockMoves.clear();
    try {
      const annotations = JSON.parse(localStorage.getItem(`onyx:annot:${fileKey}`) || '[]');
      for (const annotation of annotations) {
        if (annotation.type === 'textedit' || annotation.type === 'objectedit') continue;
        if (annotation.type === 'content-edit') continue;
        if (!this.map.has(annotation.page)) this.map.set(annotation.page, []);
        this.map.get(annotation.page).push(annotation);
      }
    } catch (_) { /* 忽略损坏历史数据 */ }
    this.redrawAll();
  }

  persist() {
    try {
      const all = [];
      for (const list of this.map.values()) all.push(...list);
      localStorage.setItem(`onyx:annot:${this.ctx.fileKey}`, JSON.stringify(all));
    } catch (_) { /* 存储不可用时保持内存状态 */ }
  }

  count() {
    let total = 0;
    for (const list of this.map.values()) total += list.length;
    return total;
  }

  _bindTools() {
    document.querySelectorAll('.tool-toggle').forEach(button => {
      button.addEventListener('click', () => this.setTool(button.dataset.tool));
    });
    document.getElementById('btn-clear-annot').addEventListener('click', () => {
      const page = this.ctx.viewer.currentPage;
      const list = this.map.get(page);
      if (!list?.length) {
        toast('本页没有批注', true);
        return;
      }
      this.map.delete(page);
      this.persist();
      this.redrawPage(page);
      toast(`已清除第 ${page} 页的 ${list.length} 条批注`);
    });
  }

  setTool(tool) {
    this.tool = tool;
    document.body.dataset.tool = tool;
    document.querySelectorAll('.tool-toggle').forEach(button => {
      button.setAttribute('aria-pressed', String(button.dataset.tool === tool));
    });
    document.getElementById('st-tool').textContent = tool.toUpperCase();
  }

  async redrawPage(pageNumber) {
    const box = this.ctx.viewer.pageBoxes.get(pageNumber);
    if (!box) return;
    const layer = box.querySelector('.annot-layer');
    layer.innerHTML = '';
    const list = this.map.get(pageNumber) || [];
    if (!list.length) return;
    const page = await this.ctx.viewer.doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: this.ctx.viewer.scale, rotation: this.ctx.viewer.rotation });
    for (const annotation of list) {
      if (annotation.type === 'highlight' || annotation.type === 'underline') {
        const colorClass = ANNOT_COLORS[annotation.color]?.cls || '';
        for (const rect of annotation.rects) {
          const view = this.ctx.viewer.pdfRectToViewport(pageNumber, rect, viewport);
          const div = document.createElement('div');
          div.className = `annot annot-${annotation.type} ${colorClass}`;
          div.style.left = `${view.x}px`;
          div.style.top = `${view.y}px`;
          div.style.width = `${view.w}px`;
          div.style.height = `${annotation.type === 'underline' ? 0 : view.h}px`;
          if (annotation.type === 'underline') div.style.top = `${view.y + view.h - 1}px`;
          layer.appendChild(div);
        }
      } else if (annotation.type === 'rect') {
        const view = this.ctx.viewer.pdfRectToViewport(pageNumber, annotation.rect, viewport);
        const div = document.createElement('div');
        div.className = 'annot annot-rect';
        div.style.left = `${view.x}px`;
        div.style.top = `${view.y}px`;
        div.style.width = `${view.w}px`;
        div.style.height = `${view.h}px`;
        div.dataset.annotId = annotation.id;
        this._bindRectDrag(div, annotation, pageNumber, viewport);
        layer.appendChild(div);
      } else if (annotation.type === 'note') {
        const view = this.ctx.viewer.pdfRectToViewport(pageNumber, {
          x: annotation.x, y: annotation.y, x2: annotation.x + 14, y2: annotation.y + 14
        }, viewport);
        const div = document.createElement('div');
        div.className = `annot annot-note-icon${annotation.content ? ' has-content' : ''}`;
        div.textContent = 'N';
        div.style.left = `${view.x}px`;
        div.style.top = `${view.y}px`;
        div.dataset.annotId = annotation.id;
        div.addEventListener('click', event => {
          event.stopPropagation();
          this._openNoteEditor(annotation, div);
        });
        layer.appendChild(div);
      } else if (annotation.type === 'image') {
        const start = viewport.convertToViewportPoint(annotation.x, annotation.y);
        const end = viewport.convertToViewportPoint(annotation.x + annotation.w, annotation.y - annotation.h);
        const div = document.createElement('div');
        div.className = 'annot annot-image';
        div.style.left = `${Math.min(start[0], end[0])}px`;
        div.style.top = `${Math.min(start[1], end[1])}px`;
        div.style.width = `${Math.abs(end[0] - start[0])}px`;
        div.style.height = `${Math.abs(end[1] - start[1])}px`;
        const image = document.createElement('img');
        image.src = annotation.dataUrl;
        div.appendChild(image);
        this._bindImageDrag(div, annotation, pageNumber, viewport);
        layer.appendChild(div);
      } else if (annotation.type === 'text') {
        const point = viewport.convertToViewportPoint(annotation.x, annotation.y);
        const div = document.createElement('div');
        div.className = 'annot annot-textbox';
        div.textContent = annotation.text || ' ';
        div.style.left = `${point[0]}px`;
        div.style.top = `${point[1]}px`;
        div.style.fontSize = `${annotation.size * this.ctx.viewer.scale}px`;
        div.style.color = annotation.color || '#d64545';
        div.style.fontWeight = annotation.bold ? '700' : '400';
        div.dataset.annotId = annotation.id;
        this._bindTextboxDrag(div, annotation, pageNumber, viewport);
        div.addEventListener('dblclick', event => {
          event.stopPropagation();
          this._openTextboxEditor(annotation);
        });
        layer.appendChild(div);
      }
    }
  }

  redrawAll() {
    for (const pageNumber of this.ctx.viewer.pageBoxes.keys()) this.redrawPage(pageNumber);
  }

  async setEditMode(enabled) {
    this.editMode = !!enabled;
    document.body.classList.toggle('edit-mode', this.editMode);
    document.getElementById('btn-edit-mode')?.setAttribute('aria-pressed', String(this.editMode));
    if (this._committing) return;
    if (this.editMode) {
      this.setTool('select');
      this._showToolbar();
      this._overlaySignatures?.clear();
      await this.renderElementOverlay(this.ctx.viewer.currentPage || 1);
      return;
    }
    const changed = this._commitAllParagraphEdits();
    if (changed > 0 || this._previewDirty) {
      this._committing = true;
      try {
        const annotations = [];
        for (const list of this.map.values()) annotations.push(...list);
        const originalBytes = this.ctx.getOriginalBytes?.();
        if (!originalBytes) throw new Error('原始 PDF 不可用');
        const { bytes } = await exportAnnotatedPdf(originalBytes, annotations, 'onyx-edit.pdf');
        await this.ctx.viewer.replaceDocument(bytes, this.ctx.viewer.currentPage || 1);
        this._previewDirty = false;
      } catch (error) {
        console.error('完成编辑失败:', error);
        toast('完成编辑失败：' + (error?.message || error), true);
      } finally {
        this._committing = false;
      }
    }
    this._hideToolbar();
    this.removeElementOverlay();
    this.selectedElement = null;
    this._selectedRange = null;
    if (changed > 0) toast(`已完成 ${changed} 处文字修改`);
  }

  async _getLines(pageNumber) {
    if (!this._lineCache.has(pageNumber)) {
      const { extractPageLines } = await import('./pdf2edit.js');
      const originalDoc = await this._getOriginalDoc();
      this._lineCache.set(pageNumber, await extractPageLines(originalDoc, pageNumber));
    }
    return this._lineCache.get(pageNumber);
  }

  async _getOriginalDoc() {
    if (this._originalDoc) return this._originalDoc;
    if (!this._originalDocPromise) {
      const originalBytes = this.ctx.getOriginalBytes?.();
      if (!originalBytes) throw new Error('原始 PDF 不可用');
      this._originalDocPromise = pdfjsLib.getDocument({
        data: originalBytes.slice(0),
        cMapUrl: '../vendor/cmaps/',
        cMapPacked: true,
        enableXfa: false
      }).promise.then(doc => {
        this._originalDoc = doc;
        return doc;
      }).catch(error => {
        this._originalDocPromise = null;
        throw error;
      });
    }
    return this._originalDocPromise;
  }

  _destroyOriginalDoc() {
    const originalDoc = this._originalDoc;
    this._originalDoc = null;
    this._originalDocPromise = null;
    originalDoc?.destroy?.().catch?.(() => {});
  }

  async _ensureSnapshot(pageNumber) {
    if (this.elementSnapshots.has(pageNumber)) return this.elementSnapshots.get(pageNumber);
    const extracted = await this._getLines(pageNumber);
    const paras = buildEditBlocks(extracted.lines);
    const annotations = this.map.get(pageNumber) || [];
    paras.forEach((para, index) => {
      para.runs = (para.runs || []).map(run => ({
        ...run,
        font: run.font || para.font || '微软雅黑',
        size: Number(run.size) || para.size,
        color: run.color || para.color || '#000000',
        bold: !!run.bold
      }));
      if (!para.runs.length) {
        para.runs = [{ text: para.text || '', font: para.font || '微软雅黑', size: para.size, color: para.color || '#000000', bold: !!para.bold }];
      }
      const committed = annotations.find(item => item.type === 'content-edit' && item.blockKey === `${pageNumber}:${index}`);
      if (committed?.objectEdit === true) {
        para.sourceOps = [...(committed.sourceOps || [])];
        para.runs = (committed.runs || para.runs).map(run => ({ ...run }));
        para.text = para.runs.map(run => run.text).join('');
        para.align = committed.align || para.align;
        para.committed = true;
      }
    });
    const snapshot = { lines: extracted.lines, paras, width: extracted.width, height: extracted.height };
    this.elementSnapshots.set(pageNumber, snapshot);
    return snapshot;
  }

  async renderElementOverlay(pageNumber) {
    if (!this.editMode || !this.ctx.viewer.doc || this._previewBusy) return;
    const box = this.ctx.viewer.pageBoxes.get(pageNumber);
    if (!box) return;
    const signature = [
      pageNumber, this.ctx.viewer.scale, this.ctx.viewer.rotation,
      Math.round(box.getBoundingClientRect().width),
      Math.round(box.getBoundingClientRect().height)
    ].join('|');
    this._overlaySignatures = this._overlaySignatures || new Map();
    if (!this._forceOverlayRebuild &&
        this._overlaySignatures.get(pageNumber) === signature &&
        box.querySelector('.element-layer')) return;
    this._forceOverlayRebuild = false;
    const generation = (this._overlayGenerations.get(pageNumber) || 0) + 1;
    this._overlayGenerations.set(pageNumber, generation);
    box.querySelector('.element-layer')?.remove();
    this._overlaySignatures.delete(pageNumber);
    if (!this.editMode || !this.ctx.viewer.doc || this._previewBusy) return;
    const snapshot = await this._ensureSnapshot(pageNumber);
    if (!this.editMode || !this.ctx.viewer.doc || this._previewBusy ||
        generation !== this._overlayGenerations.get(pageNumber)) return;
    const layer = box.querySelector('.element-layer') || document.createElement('div');
    layer.className = 'element-layer';
    layer.innerHTML = '';
    if (!layer.isConnected) box.appendChild(layer);
    const scale = this.ctx.viewer.scale;
    snapshot.paras.forEach((para, index) => {
      const move = this._blockMoves.get(`${pageNumber}:${index}`) || { dx: 0, dy: 0 };
      const block = document.createElement('div');
      block.className = 'elem-hit';
      block.dataset.key = `P${index}`;
      block.dataset.page = String(pageNumber);
      block.style.left = `${(para.x + move.dx) * scale}px`;
      block.style.top = `${(para.y + move.dy) * scale}px`;
      block.style.width = `${Math.max(para.w, 16) * scale}px`;
      block.style.minHeight = `${Math.max(para.h, 10) * scale}px`;
      block.title = '点击编辑；拖选文字后可单独设置格式';
      block.addEventListener('mousedown', event => {
        if (!this.editMode) return;
        const editor = block.querySelector('.elem-edit');
        const inEditor = !!event.target.closest('.elem-edit');
        if (inEditor && !event.altKey) {
          const rect = block.getBoundingClientRect();
          const nearEdge = event.clientX < rect.left + 6 || event.clientX > rect.right - 6 ||
            event.clientY < rect.top + 6 || event.clientY > rect.bottom - 6;
          if (!nearEdge) return;
        } else if (inEditor) {
          event.stopPropagation();
          return;
        }
        event.stopPropagation();
        this._activateBlock(pageNumber, index);
        const startX = event.clientX;
        const startY = event.clientY;
        let moved = false;
        const onMove = mouseEvent => {
          const dx = (mouseEvent.clientX - startX) / scale;
          const dy = (mouseEvent.clientY - startY) / scale;
          if (Math.abs(dx) > 1 || Math.abs(dy) > 1) moved = true;
          if (!moved) return;
          block.style.left = `${(para.x + move.dx + dx) * scale}px`;
          block.style.top = `${(para.y + move.dy + dy) * scale}px`;
        };
        const onUp = mouseEvent => {
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          if (!moved) return;
          const dx = (mouseEvent.clientX - startX) / scale;
          const dy = (mouseEvent.clientY - startY) / scale;
          const key = `${pageNumber}:${index}`;
          const previous = this._blockMoves.get(key) || { dx: 0, dy: 0 };
          this._blockMoves.set(key, { dx: previous.dx + dx, dy: previous.dy + dy });
          const activeEditor = block.querySelector('.elem-edit');
          if (activeEditor) {
            activeEditor.dataset.dirty = '1';
          }
          toast('已移动，点击完成时写入 PDF');
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
      layer.appendChild(block);
      if (para.committed) this._mountEditor(block, pageNumber, index, false);
    });
    this._overlaySignatures.set(pageNumber, signature);
    if (this.selectedElement?.pageNumber === pageNumber) {
      this._activateBlock(pageNumber, this.selectedElement.index);
    }
  }

  removeElementOverlay() {
    document.querySelectorAll('.element-layer').forEach(layer => layer.remove());
    this._overlaySignatures.clear();
    this._overlayGenerations.clear();
  }

  _activateBlock(pageNumber, index) {
    const snapshot = this.elementSnapshots.get(pageNumber);
    const para = snapshot?.paras[index];
    const block = document.querySelector(`.elem-hit[data-key="P${index}"][data-page="${pageNumber}"]`);
    if (!block || !para) return;
    this.selectedElement = { pageNumber, index };
    this._lastSelection = { pageNumber, index };
    document.querySelectorAll('.elem-hit.elem-selected').forEach(item => {
      item.classList.remove('elem-selected');
      const previousIndex = Number(item.querySelector('.elem-edit')?.dataset.index);
      const previousPage = Number(item.dataset.page);
      const previousPara = this.elementSnapshots.get(previousPage)?.paras[previousIndex];
      if (!item.querySelector('.elem-edit[data-dirty="1"]') && !previousPara?.committed) {
        item.querySelector('.block-inpaint')?.remove();
        item.querySelector('.elem-edit')?.remove();
        item.querySelectorAll('.rz-handle').forEach(handle => handle.remove());
      }
    });
    block.classList.add('elem-selected');
    const existing = block.querySelector('.elem-edit');
    if (existing) {
      existing.focus();
      this._syncToolbar();
      return;
    }
    const scale = this.ctx.viewer.scale;
    const format = this._paraFmt.get(`${pageNumber}:${index}`) || {};
    this._mountEditor(block, pageNumber, index, true);
  }

  _mountEditor(block, pageNumber, index, focus) {
    const snapshot = this.elementSnapshots.get(pageNumber);
    const para = snapshot?.paras[index];
    if (!block || !para || block.querySelector('.elem-edit')) return;
    const scale = this.ctx.viewer.scale;
    const format = this._paraFmt.get(`${pageNumber}:${index}`) || {};
    const editor = document.createElement('div');
    editor.className = 'elem-edit';
    editor.contentEditable = 'true';
    editor.spellcheck = false;
    editor.dataset.page = String(pageNumber);
    editor.dataset.index = String(index);
    const lineHeight = Math.max(para.h / Math.max(para.lineIdxs.length, 1), para.size * 1.45) * scale;
    editor.style.cssText = [
      `width:${Math.max(format.w || para.w, 16) * scale}px`,
      `min-height:${Math.max(para.h, 10) * scale}px`,
      `line-height:${lineHeight}px`,
      `text-align:${format.align || para.align || 'left'}`,
      `font-family:"${fontFamilyName(format.font || para.font)}","Microsoft YaHei",sans-serif`,
      `font-size:${(format.size || para.size || 12) * scale}px`,
      `color:${normalizeCssColor(format.color || para.color || '#000000')}`,
      `font-weight:${(format.bold ?? para.bold) ? '700' : '400'}`,
      'white-space:pre-wrap'
    ].join(';');
    editor.innerHTML = renderRunHtml(para.runs, scale);
    const storeSelection = () => {
      if (document.activeElement !== editor) return;
      const offsets = selectionOffsets(editor);
      if (!offsets) return;
      this._selectedRange = { editor, ...offsets };
    };
    editor.addEventListener('input', () => {
      editor.dataset.dirty = '1';
      this._previewDirty = true;
      this._selectedRange = null;
      this._syncToolbar();
    });
    ['keyup', 'mouseup'].forEach(eventType => editor.addEventListener(eventType, storeSelection));
    editor.addEventListener('paste', event => {
      event.preventDefault();
      document.execCommand('insertText', false, event.clipboardData?.getData('text/plain') || '');
    });
    editor.addEventListener('mousedown', () => {
      requestAnimationFrame(() => {
        if (this.editMode && focus && document.body.contains(editor)) editor.focus();
      });
    });
    editor.addEventListener('blur', () => {
      storeSelection();
    });
    editor.addEventListener('mousedown', event => event.stopPropagation());
    this._paintBlockBackground(block);
    block.appendChild(editor);
    this._attachResizeHandles(block, editor, pageNumber, index, scale);
    if (!focus) return;
    setTimeout(() => {
      if (!this.editMode || !document.body.contains(editor)) return;
      editor.focus();
      setTextSelection(editor, editor.textContent.length, editor.textContent.length);
      storeSelection();
      this._syncToolbar();
    }, 0);
    this._syncToolbar();
  }

  _paintBlockBackground(block) {
    if (block.querySelector('.block-inpaint')) return;
    const canvas = block.closest('.page-outer')?.querySelector('canvas');
    if (!canvas) return;
    const pageRect = canvas.getBoundingClientRect();
    const blockRect = block.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const left = Math.max(0, Math.floor((blockRect.left - pageRect.left) * dpr) - 2);
    const top = Math.max(0, Math.floor((blockRect.top - pageRect.top) * dpr) - 2);
    const width = Math.min(canvas.width - left, Math.ceil(blockRect.width * dpr) + 4);
    const height = Math.min(canvas.height - top, Math.ceil(blockRect.height * dpr) + 4);
    if (width <= 0 || height <= 0) return;
    const source = canvas.getContext('2d', { willReadFrequently: true });
    const patch = document.createElement('canvas');
    patch.className = 'block-inpaint';
    patch.width = width;
    patch.height = height;
    patch.style.width = `${width / dpr}px`;
    patch.style.height = `${height / dpr}px`;
    const context = patch.getContext('2d');
    let pixels;
    try {
      pixels = source.getImageData(left, top, width, height);
    } catch (_) {
      return;
    }
    const output = context.createImageData(width, height);
    output.data.set(pixels.data);
    for (let x = 0; x < width; x++) {
      const first = (top > 0 ? 2 : 0) * width + x;
      const last = (height - (top + height < canvas.height ? 3 : 1)) * width + x;
      const startR = pixels.data[first];
      const startG = pixels.data[first + 1];
      const startB = pixels.data[first + 2];
      const endR = pixels.data[last];
      const endG = pixels.data[last + 1];
      const endB = pixels.data[last + 2];
      for (let y = 2; y < height - 2; y++) {
        const ratio = (y - 2) / Math.max(1, height - 5);
        const offset = (y * width + x) * 4;
        output.data[offset] = startR + (endR - startR) * ratio;
        output.data[offset + 1] = startG + (endG - startG) * ratio;
        output.data[offset + 2] = startB + (endB - startB) * ratio;
        output.data[offset + 3] = 255;
      }
    }
    context.putImageData(output, 0, 0);
    block.insertBefore(patch, block.firstChild);
  }

  _attachResizeHandles(block, editor, pageNumber, index, scale) {
    ['e', 's', 'se'].forEach(direction => {
      const handle = document.createElement('div');
      handle.className = `rz-handle rz-${direction}`;
      handle.addEventListener('mousedown', event => {
        event.stopPropagation();
        event.preventDefault();
        const startX = event.clientX;
        const startY = event.clientY;
        const startWidth = parseFloat(editor.style.width) || block.offsetWidth;
        const startHeight = block.offsetHeight;
        const onMove = moveEvent => {
          if (direction.includes('e')) {
            const width = Math.max(24, startWidth + moveEvent.clientX - startX);
            editor.style.width = `${width}px`;
            block.style.width = `${width}px`;
          }
          if (direction.includes('s')) {
            const height = Math.max(14, startHeight + moveEvent.clientY - startY);
            block.style.minHeight = `${height}px`;
            editor.style.minHeight = `${height}px`;
          }
        };
        const onUp = () => {
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          this._paraFmt.set(`${pageNumber}:${index}`, {
            ...this._paraFmt.get(`${pageNumber}:${index}`),
            w: (parseFloat(editor.style.width) || block.offsetWidth) / scale,
            h: block.offsetHeight / scale
          });
          editor.dataset.dirty = '1';
          this._previewDirty = true;
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
      block.appendChild(handle);
    });
  }

  _applyFmtToSelected(patch) {
    if (!this.selectedElement && this._lastSelection) {
      this._activateBlock(this._lastSelection.pageNumber, this._lastSelection.index);
    }
    const editor = document.querySelector('.elem-hit.elem-selected .elem-edit');
    if (!editor) {
      toast('请先点击一个文字块', true);
      return;
    }
    const stored = this._selectedRange?.editor === editor ? this._selectedRange : null;
    const live = selectionOffsets(editor);
    const offsets = (stored && stored.end > stored.start)
      ? stored
      : (live && live.end > live.start)
        ? live
        : { start: 0, end: editor.textContent.length };
    const selectedStart = Math.min(offsets.start, offsets.end);
    const selectedEnd = Math.max(offsets.start, offsets.end);
    const pageNumber = Number(editor.dataset.page);
    const index = Number(editor.dataset.index);
    const para = this.elementSnapshots.get(pageNumber)?.paras[index];
    const fallback = {
      font: para?.font || '微软雅黑',
      size: para?.size || 12,
      color: para?.color || '#000000',
      bold: !!para?.bold
    };
    let bold = patch.bold;
    const toggleBold = bold === 'toggle';
    if (bold === undefined || toggleBold) {
      let boldLength = 0;
      let totalLength = 0;
      let boldCursor = 0;
      collectTextRuns(editor, this.ctx.viewer.scale, fallback).forEach(run => {
        const runStart = boldCursor;
        const runEnd = runStart + run.text.length;
        boldCursor = runEnd;
        const overlap = Math.min(runEnd, selectedEnd) - Math.max(runStart, selectedStart);
        if (overlap <= 0) return;
        totalLength += overlap;
        if (run.bold) boldLength += overlap;
      });
      bold = toggleBold ? boldLength * 2 < totalLength : boldLength * 2 >= totalLength;
    }
    const normalizedPatch = { ...patch, bold };
    if (normalizedPatch.font) normalizedPatch.font = fontFamilyName(normalizedPatch.font);
    applyFormatToRuns(editor, offsets, normalizedPatch, fallback, this.ctx.viewer.scale);
    setTextSelection(editor, offsets.start, offsets.end);
    this._selectedRange = { editor, ...offsets };
    if (normalizedPatch.align) {
      editor.style.textAlign = normalizedPatch.align;
      this._paraFmt.set(`${pageNumber}:${index}`, {
        ...this._paraFmt.get(`${pageNumber}:${index}`), align: normalizedPatch.align
      });
    }
    editor.dataset.dirty = '1';
    this._previewDirty = true;
    this._syncToolbar();
  }

  _showToolbar() {
    this._hideToolbar();
    const bar = document.createElement('div');
    bar.id = 'fmt-toolbar';
    bar.innerHTML = `
      <span class="ft-label">文字格式</span>
      <select data-f="font" title="字体"></select>
      <input data-f="size" type="number" min="6" max="96" title="字号 (pt)" />
      <button data-f="bold" title="加粗"><b>B</b></button>
      <input data-f="color" type="color" title="文字颜色" />
      <button data-f="align-left" title="居左">左</button>
      <button data-f="align-center" title="居中">中</button>
      <button data-f="align-right" title="居右">右</button>
      <button data-f="done" class="ft-done" title="完成并写入修改">完成</button>`;
    document.body.appendChild(bar);
    bar.addEventListener('mousedown', event => {
      if (event.target.closest('button')) event.preventDefault();
    });
    const fontSelector = bar.querySelector('[data-f="font"]');
    fontSelector.innerHTML = ['微软雅黑', '宋体', '黑体', '楷体', '仿宋', '等线', 'Arial', 'Times New Roman', 'Calibri', 'Segoe UI']
      .map(font => `<option value="${font}">${font}</option>`).join('');
    fontSelector.addEventListener('change', () => this._applyFmtToSelected({ font: fontSelector.value }));
    const sizeInput = bar.querySelector('[data-f="size"]');
    sizeInput.addEventListener('change', () => {
      const size = clamp(parseFloat(sizeInput.value) || 12, 6, 96);
      this._applyFmtToSelected({ size });
      sizeInput.value = String(size);
    });
    sizeInput.addEventListener('keydown', event => {
      if (event.key === 'Enter') sizeInput.dispatchEvent(new Event('change', { bubbles: true }));
    });
    bar.querySelector('[data-f="bold"]').addEventListener('click', () => this._applyFmtToSelected({ bold: 'toggle' }));
    bar.querySelector('[data-f="color"]').addEventListener('input', event => this._applyFmtToSelected({ color: event.target.value }));
    bar.querySelector('[data-f="align-left"]').addEventListener('click', () => this._applyFmtToSelected({ align: 'left' }));
    bar.querySelector('[data-f="align-center"]').addEventListener('click', () => this._applyFmtToSelected({ align: 'center' }));
    bar.querySelector('[data-f="align-right"]').addEventListener('click', () => this._applyFmtToSelected({ align: 'right' }));
    bar.querySelector('[data-f="done"]').addEventListener('click', () => this.setEditMode(false));
    this._syncToolbar();
  }

  _hideToolbar() {
    document.getElementById('fmt-toolbar')?.remove();
  }

  _syncToolbar() {
    const bar = document.getElementById('fmt-toolbar');
    const editor = document.querySelector('.elem-hit.elem-selected .elem-edit');
    if (!bar || !editor) return;
    if (bar.contains(document.activeElement) && document.activeElement.matches('input, select')) return;
    const offsets = this._selectedRange?.editor === editor ? this._selectedRange : null;
    let sample = editor;
    if (offsets && offsets.end > offsets.start) {
      const selection = window.getSelection();
      if (selection.rangeCount && editor.contains(selection.getRangeAt(0).startContainer)) {
        const temporary = document.createElement('div');
        temporary.appendChild(selection.getRangeAt(0).cloneContents());
        sample = temporary;
      }
    }
    let sampleElement = null;
    walkTextNodes(sample, node => {
      if (!sampleElement && node.textContent.trim()) sampleElement = node.parentElement || editor;
    });
    const styleElement = sampleElement || editor;
    const inlineStyle = styleElement.style || {};
    const style = getComputedStyle(styleElement);
    const fontSelector = bar.querySelector('[data-f="font"]');
    const fontName = fontFamilyName(inlineStyle.fontFamily || style.fontFamily);
    if ([...fontSelector.options].some(option => option.value === fontName)) fontSelector.value = fontName;
    const cssSize = parseFloat(inlineStyle.fontSize || style.fontSize) / this.ctx.viewer.scale;
    const sizeControl = bar.querySelector('[data-f="size"]');
    if (Number.isFinite(cssSize)) {
      sizeControl.value = Math.round(cssSize * 10) / 10;
    } else if (document.activeElement !== sizeControl && !sizeControl.value) {
      const para = this.elementSnapshots.get(Number(editor.dataset.page))?.paras[Number(editor.dataset.index)];
      sizeControl.value = Math.round((para?.size || 12) * 10) / 10;
    }
    bar.querySelector('[data-f="color"]').value = this._toHexColor(inlineStyle.color || style.color);
    bar.querySelector('[data-f="bold"]').style.background = parseInt(inlineStyle.fontWeight || style.fontWeight, 10) >= 600 ? '#e34d3f' : '';
    const block = editor.closest('.elem-hit');
    if (block) {
      const blockRect = block.getBoundingClientRect();
      const width = bar.offsetWidth || 460;
      const left = clamp(blockRect.left + blockRect.width / 2 - width / 2, 12, window.innerWidth - width - 12);
      const top = clamp(blockRect.top - 42, 70, window.innerHeight - 52);
      bar.style.left = `${left}px`;
      bar.style.top = `${top}px`;
      bar.style.transform = 'none';
    }
  }

  _toHexColor(value) {
    const match = String(value || '').match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
    return match ? '#' + match.slice(1).map(item => Number(item).toString(16).padStart(2, '0')).join('') : '#000000';
  }

  _schedulePreview(delay = 500) {
    clearTimeout(this._previewTimer);
    this._previewTimer = setTimeout(() => {
      this._refreshPreview().catch(error => {
        console.error('编辑预览失败:', error);
        toast('编辑预览失败：' + (error?.message || error), true);
      });
    }, delay);
  }

  _showCommitVeil() {
    void 0;
  }

  _hideCommitVeil() {
    void 0;
  }

  async _refreshPreview() {
    if (this._previewBusy) {
      this._previewQueued = true;
      return;
    }
    this._previewBusy = true;
    const restorePage = this.ctx.viewer.currentPage || 1;
    const selected = this.selectedElement;
    const activeEditor = document.querySelector('.elem-hit.elem-selected .elem-edit');
    const caret = activeEditor ? this._caretOffset(activeEditor) : null;
    try {
      this._commitAllParagraphEdits();
      const annotations = [];
      for (const list of this.map.values()) annotations.push(...list);
      const originalBytes = this.ctx.getOriginalBytes?.();
      if (!annotations.some(item => item.type === 'content-edit') || !originalBytes) return;
      const { bytes } = await exportAnnotatedPdf(originalBytes, annotations, 'onyx-preview.pdf');
      const oldScaleMode = this.ctx.viewer.scaleMode;
      this.ctx.viewer.scaleMode = oldScaleMode;
      this._forceOverlayRebuild = true;
      this._overlaySignatures.clear();
      this._overlayGenerations.clear();
      await this.ctx.viewer.replaceDocument(bytes, restorePage);
      if (this.editMode) {
        await this.renderElementOverlay(restorePage);
        if (selected) {
          this._activateBlock(selected.pageNumber, selected.index);
          const nextEditor = document.querySelector('.elem-hit.elem-selected .elem-edit');
          if (nextEditor && caret !== null) this._setCaretOffset(nextEditor, Math.min(caret, nextEditor.textContent.length));
        }
      }
      this._previewDirty = false;
    } finally {
      this._previewBusy = false;
      if (this._previewQueued) {
        this._previewQueued = false;
        this._schedulePreview(20);
      }
    }
  }

  _caretOffset(root) {
    const selection = window.getSelection();
    if (!selection.rangeCount) return null;
    const current = selection.getRangeAt(0);
    const range = current.cloneRange();
    range.selectNodeContents(root);
    range.setEnd(current.endContainer, current.endOffset);
    return range.toString().length;
  }

  _setCaretOffset(root, offset) {
    setTextSelection(root, offset, offset);
  }

  _commitAllParagraphEdits() {
    let changed = 0;
    const allEditors = document.querySelectorAll('.elem-edit');
    document.querySelectorAll('.elem-edit[data-dirty="1"]').forEach(editor => {
      const block = editor.closest('.elem-hit');
      const pageNumber = Number(block?.dataset.page);
      const index = Number(editor.dataset.index);
      const snapshot = this.elementSnapshots.get(pageNumber);
      const para = snapshot?.paras[index];
      if (!block || !para) return;
      const format = this._paraFmt.get(`${pageNumber}:${index}`) || {};
      const runs = collectTextRuns(editor, this.ctx.viewer.scale, {
        font: para.font || '微软雅黑', size: para.size || 12, color: para.color || '#000000', bold: !!para.bold
      }).map(run => ({
        text: run.text,
        font: run.font || para.font || '微软雅黑',
        size: Math.round((Number(run.size) || para.size || 12) * 10) / 10,
        color: run.color || para.color || '#000000',
        bold: !!run.bold
      }));
      const move = this._blockMoves.get(`${pageNumber}:${index}`) || { dx: 0, dy: 0 };
      const oldSignature = JSON.stringify([para.runs?.map(run => [run.text, run.font, run.size, run.color, run.bold]), move]);
      const newSignature = JSON.stringify([runs.map(run => [run.text, run.font, run.size, run.color, run.bold]), move]);
      if (oldSignature === newSignature) {
        editor.dataset.dirty = '';
        return;
      }
      const existing = (this.map.get(pageNumber) || []).find(item => item.type === 'content-edit' && item.blockKey === `${pageNumber}:${index}`);
      const sourceOps = [...new Set(existing?.sourceOps || para.sourceOps || para.lineIdxs.flatMap(lineIndex =>
        snapshot?.lines[lineIndex]?.sourceOps || []
      ))].filter(Number.isInteger).sort((a, b) => a - b);
      const annotation = {
        id: existing?.id || uid('ce'),
        type: 'content-edit',
        page: pageNumber,
        blockKey: `${pageNumber}:${index}`,
        objectEdit: true,
        sourceOps,
        x: para.x + move.dx,
        y: (snapshot?.height || 792) - (para.y + move.dy),
        w: Math.max(format.w || para.w, 16),
        h: Math.max(para.h, 10),
        lineHeight: Math.max(para.h / Math.max(para.lineIdxs.length, 1), Math.max(...runs.map(run => run.size || para.size || 12), 1) * 1.45),
        align: format.align || para.align || 'left',
        runs,
        created: Date.now()
      };
      if (!this.map.has(pageNumber)) this.map.set(pageNumber, []);
      this.map.set(pageNumber, this.map.get(pageNumber).filter(item => item.blockKey !== annotation.blockKey || item.type !== 'content-edit'));
      this.map.get(pageNumber).push(annotation);
      para.sourceOps = sourceOps;
      para.runs = runs;
      para.text = runs.map(run => run.text).join('');
      para.align = annotation.align;
      editor.dataset.dirty = '';
      changed++;
    });
    if (changed) this.persist();
    return changed;
  }

  async pasteImageFromClipboard(pageNumber) {
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const type = item.types.find(value => value.startsWith('image/'));
        if (!type) continue;
        const blob = await item.getType(type);
        const dataUrl = await new Promise(resolve => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result);
          reader.readAsDataURL(blob);
        });
        const image = new Image();
        await new Promise((resolve, reject) => { image.onload = resolve; image.onerror = reject; image.src = dataUrl; });
        let width = image.naturalWidth * 0.75;
        let height = image.naturalHeight * 0.75;
        if (width > 300) {
          height *= 300 / width;
          width = 300;
        }
        const rect = this.ctx.viewer.pageBoxes.get(pageNumber)?.getBoundingClientRect();
        if (rect) await this.placeImage(pageNumber, rect.width / 2, rect.height / 2, dataUrl, width, height);
        return true;
      }
    } catch (_) { /* 无剪贴板权限 */ }
    toast('剪贴板中没有图片', true);
    return false;
  }

  async placeImage(pageNumber, cssX, cssY, dataUrl, widthPt, heightPt) {
    const page = await this.ctx.viewer.doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: this.ctx.viewer.scale, rotation: this.ctx.viewer.rotation });
    const point = viewport.convertToPdfPoint(cssX, cssY);
    const annotation = {
      id: uid(), type: 'image', page: pageNumber,
      x: point[0], y: point[1] + heightPt, w: widthPt, h: heightPt,
      dataUrl, created: Date.now()
    };
    if (!this.map.has(pageNumber)) this.map.set(pageNumber, []);
    this.map.get(pageNumber).push(annotation);
    this.persist();
    await this.redrawPage(pageNumber);
    toast('图片已插入（可拖动位置）');
  }

  _bindPointer() {
    const viewerElement = document.getElementById('viewer');
    document.addEventListener('mouseup', async () => {
      if (this.tool !== 'highlight' && this.tool !== 'underline') return;
      const page = this.ctx.viewer.currentPage;
      const selection = window.getSelection();
      if (!page || !selection || selection.isCollapsed) return;
      const result = await this.ctx.viewer.selectionToPdfRects(page);
      if (!result) return;
      const annotation = {
        id: uid(), type: this.tool, page,
        rects: result.rects, text: result.text.slice(0, 500),
        color: 'default', created: Date.now()
      };
      if (!this.map.has(page)) this.map.set(page, []);
      this.map.get(page).push(annotation);
      this.persist();
      await this.redrawPage(page);
      toast(`已添加${this.tool === 'highlight' ? '高亮' : '下划线'}`);
    });

    viewerElement.addEventListener('mousedown', event => {
      if (!this.editMode && this.tool !== 'rect' && this.tool !== 'note' && this.tool !== 'text') return;
      if (event.target.closest?.('.annot') || event.target.closest?.('.elem-hit')) return;
      const outer = event.target.closest?.('.page-outer');
      if (!outer) return;
      const pageNumber = Number(outer.dataset.page);
      const canvas = outer.querySelector('canvas');
      if (!canvas) {
        toast('页面尚未渲染完成', true);
        return;
      }
      const canvasRect = canvas.getBoundingClientRect();
      const startX = event.clientX;
      const startY = event.clientY;
      if (this.editMode) {
        event.preventDefault();
        window.getSelection().removeAllRanges();
        return;
      }
      if (this.tool === 'note') {
        this._createNote(pageNumber, startX - canvasRect.left, startY - canvasRect.top, canvasRect);
        return;
      }
      if (this.tool === 'text') {
        event.preventDefault();
        this._createTextbox(pageNumber, startX - canvasRect.left, startY - canvasRect.top, canvasRect);
        return;
      }
      const ghost = document.createElement('div');
      ghost.className = 'annot-ghost';
      outer.querySelector('.annot-layer').appendChild(ghost);
      const onMove = moveEvent => {
        const x = clamp(Math.min(startX, moveEvent.clientX), canvasRect.left, canvasRect.right);
        const y = clamp(Math.min(startY, moveEvent.clientY), canvasRect.top, canvasRect.bottom);
        ghost.style.left = `${x - canvasRect.left}px`;
        ghost.style.top = `${y - canvasRect.top}px`;
        ghost.style.width = `${Math.abs(moveEvent.clientX - startX)}px`;
        ghost.style.height = `${Math.abs(moveEvent.clientY - startY)}px`;
      };
      const onUp = async moveEvent => {
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
        ghost.remove();
        if (Math.abs(moveEvent.clientX - startX) < 4 || Math.abs(moveEvent.clientY - startY) < 4) return;
        const page = await this.ctx.viewer.doc.getPage(pageNumber);
        const viewport = page.getViewport({ scale: this.ctx.viewer.scale, rotation: this.ctx.viewer.rotation });
        const point1 = viewport.convertToPdfPoint(startX - canvasRect.left, startY - canvasRect.top);
        const point2 = viewport.convertToPdfPoint(moveEvent.clientX - canvasRect.left, moveEvent.clientY - canvasRect.top);
        const annotation = {
          id: uid(), type: 'rect', page: pageNumber,
          rect: {
            x: Math.min(point1[0], point2[0]), y: Math.min(point1[1], point2[1]),
            x2: Math.max(point1[0], point2[0]), y2: Math.max(point1[1], point2[1])
          },
          created: Date.now()
        };
        if (!this.map.has(pageNumber)) this.map.set(pageNumber, []);
        this.map.get(pageNumber).push(annotation);
        this.persist();
        await this.redrawPage(pageNumber);
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    });
  }

  async _createNote(pageNumber, cssX, cssY) {
    const page = await this.ctx.viewer.doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: this.ctx.viewer.scale, rotation: this.ctx.viewer.rotation });
    const point = viewport.convertToPdfPoint(cssX, cssY);
    const annotation = { id: uid(), type: 'note', page: pageNumber, x: point[0], y: point[1], content: '', created: Date.now() };
    if (!this.map.has(pageNumber)) this.map.set(pageNumber, []);
    this.map.get(pageNumber).push(annotation);
    this.persist();
    await this.redrawPage(pageNumber);
    const icon = this.ctx.viewer.pageBoxes.get(pageNumber)?.querySelector(`.annot-note-icon[data-annot-id="${annotation.id}"]`);
    if (icon) this._openNoteEditor(annotation, icon);
  }

  async _createTextbox(pageNumber, cssX, cssY) {
    const page = await this.ctx.viewer.doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: this.ctx.viewer.scale, rotation: this.ctx.viewer.rotation });
    const point = viewport.convertToPdfPoint(cssX, cssY);
    const annotation = {
      id: uid(), type: 'text', page: pageNumber, x: point[0], y: point[1] + 8,
      text: '', size: 14, color: '#d64545', bold: false, created: Date.now()
    };
    if (!this.map.has(pageNumber)) this.map.set(pageNumber, []);
    this.map.get(pageNumber).push(annotation);
    this.persist();
    await this.redrawPage(pageNumber);
    this._openTextboxEditor(annotation);
  }

  _openTextboxEditor(annotation) {
    openModal({
      title: '文本框',
      width: 360,
      bodyHTML: `
        <div class="field"><label>文字内容</label><textarea data-role="txt" rows="3">${annotation.text || ''}</textarea></div>
        <div class="field-row">
          <div class="field"><label>字号 (pt)</label><input type="number" data-role="size" value="${annotation.size}" min="6" max="96" /></div>
          <div class="field"><label>颜色</label><input type="color" data-role="color" value="${annotation.color || '#d64545'}" /></div>
          <div class="field"><label><input type="checkbox" data-role="bold" ${annotation.bold ? 'checked' : ''} /> 加粗</label></div>
        </div>`,
      buttons: [
        {
          label: '删除', onClick: modal => {
            const list = this.map.get(annotation.page) || [];
            const index = list.findIndex(item => item.id === annotation.id);
            if (index >= 0) list.splice(index, 1);
            this.persist();
            this.redrawPage(annotation.page);
            modal.close();
          }
        },
        { label: '取消', onClick: modal => modal.close() },
        {
          label: '保存', primary: true, onClick: modal => {
            const text = modal.query('[data-role="txt"]').value;
            if (!text.trim()) {
              toast('文字内容不能为空', true);
              return;
            }
            annotation.text = text;
            annotation.size = clamp(parseInt(modal.query('[data-role="size"]').value, 10) || 14, 6, 96);
            annotation.color = modal.query('[data-role="color"]').value;
            annotation.bold = modal.query('[data-role="bold"]').checked;
            this.persist();
            this.redrawPage(annotation.page);
            modal.close();
            toast('文本框已保存');
          }
        }
      ]
    });
  }

  _bindImageDrag(div, annotation) {
    div.addEventListener('mousedown', event => {
      event.stopPropagation();
      const startX = event.clientX;
      const startY = event.clientY;
      const originalX = annotation.x;
      const originalY = annotation.y;
      const onMove = moveEvent => {
        annotation.x = originalX + (moveEvent.clientX - startX) / this.ctx.viewer.scale;
        annotation.y = originalY - (moveEvent.clientY - startY) / this.ctx.viewer.scale;
      };
      const onUp = () => {
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
        this.redrawPage(annotation.page);
        this.persist();
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    });
  }

  _bindTextboxDrag(div, annotation, pageNumber, viewport) {
    div.addEventListener('mousedown', event => {
      event.stopPropagation();
      const canvasRect = div.closest('.page-outer').querySelector('canvas').getBoundingClientRect();
      const startX = event.clientX;
      const startY = event.clientY;
      const originalX = annotation.x;
      const originalY = annotation.y;
      const onMove = moveEvent => {
        const dx = (moveEvent.clientX - startX) * (viewport.width / canvasRect.width);
        const dy = (moveEvent.clientY - startY) * (viewport.height / canvasRect.height);
        annotation.x = originalX + dx;
        annotation.y = originalY - dy;
        const point = viewport.convertToViewportPoint(annotation.x, annotation.y);
        div.style.left = `${point[0]}px`;
        div.style.top = `${point[1]}px`;
      };
      const onUp = () => {
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
        void pageNumber;
        this.persist();
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    });
  }

  _bindRectDrag(div, annotation, pageNumber, viewport) {
    div.addEventListener('mousedown', event => {
      if (this.tool === 'rect') event.stopPropagation();
      const canvasRect = div.closest('.page-outer').querySelector('canvas').getBoundingClientRect();
      const startX = event.clientX;
      const startY = event.clientY;
      const original = { ...annotation.rect };
      const onMove = moveEvent => {
        const dx = (moveEvent.clientX - startX) * (viewport.width / canvasRect.width);
        const dy = (moveEvent.clientY - startY) * (viewport.height / canvasRect.height);
        annotation.rect = { x: original.x + dx, y: original.y + dy, x2: original.x2 + dx, y2: original.y2 + dy };
        const view = this.ctx.viewer.pdfRectToViewport(pageNumber, annotation.rect, viewport);
        div.style.left = `${view.x}px`;
        div.style.top = `${view.y}px`;
      };
      const onUp = () => {
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
        this.persist();
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    });
  }

  _bindNoteEditor() {
    this._noteTarget = null;
    document.getElementById('note-save').addEventListener('click', () => {
      if (!this._noteTarget) return;
      this._noteTarget.content = document.getElementById('note-text').value;
      this.persist();
      this.redrawPage(this._noteTarget.page);
      this._closeNoteEditor();
      toast('便签已保存');
    });
    document.getElementById('note-cancel').addEventListener('click', () => this._closeNoteEditor());
    document.getElementById('note-del').addEventListener('click', () => {
      if (!this._noteTarget) return;
      const list = this.map.get(this._noteTarget.page) || [];
      const index = list.findIndex(item => item.id === this._noteTarget.id);
      if (index >= 0) list.splice(index, 1);
      const page = this._noteTarget.page;
      this.persist();
      this._closeNoteEditor();
      this.redrawPage(page);
      toast('便签已删除');
    });
  }

  _openNoteEditor(annotation, anchorElement) {
    this._noteTarget = annotation;
    const editor = document.getElementById('note-editor');
    document.getElementById('note-text').value = annotation.content || '';
    editor.classList.remove('hidden');
    const rect = anchorElement.getBoundingClientRect();
    editor.style.left = `${clamp(rect.right + 6, 8, window.innerWidth - 248)}px`;
    editor.style.top = `${clamp(rect.top, 8, window.innerHeight - 208)}px`;
    document.getElementById('note-text').focus();
  }

  _closeNoteEditor() {
    this._noteTarget = null;
    document.getElementById('note-editor').classList.add('hidden');
  }
}
