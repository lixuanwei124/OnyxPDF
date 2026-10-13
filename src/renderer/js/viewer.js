/**
 * OnyxPDF - Viewer 核心
 * 职责：PDF 加载、页面渲染（懒加载）、缩放/适配/旋转、翻页定位、文本层、事件分发
 */
import { clamp, debounce } from './utils.js';

export class Viewer {
  constructor(ctx) {
    this.ctx = ctx;            // { pdfjsLib, onStateChange }
    this.doc = null;
    this.fileInfo = null;      // { path, name, size }
    this.scale = 1;            // 当前缩放
    this.scaleMode = 'fit-width'; // 'fit-width' | 'fit-page' | 'custom'
    this.rotation = 0;
    this.currentPage = 0;
    this.renderTasks = new Map();  // pageNum -> renderTask
    this.pageBoxes = new Map();    // pageNum -> page-outer 元素
    this.renderedPages = new Set();
    this.baseDims = new Map();     // pageNum -> {w,h} scale=1 rotation=0 的原始尺寸
    this.textCache = new Map();    // pageNum -> textContent
    this.observer = null;
    this.lastScaleForFits = 1;

    this.elViewer = document.getElementById('viewer');
    this._bindObserver();
    this._bindScroll();
    this._bindResize();
    this._bindWheel();
  }

  /* Ctrl + 滚轮缩放（锚定光标附近页面） */
  _bindWheel() {
    this.elViewer.addEventListener('wheel', (e) => {
      if (!e.ctrlKey || !this.doc) return;
      e.preventDefault();
      const mid = this.elViewer.scrollTop + this.elViewer.clientHeight / 2;
      let anchorPage = this.currentPage || 1;
      let bestDist = Infinity;
      for (const [num, box] of this.pageBoxes) {
        const d = Math.abs(box.offsetTop + box.offsetHeight / 2 - mid);
        if (d < bestDist) { bestDist = d; anchorPage = num; }
      }
      const next = clamp(this.scale * (e.deltaY < 0 ? 1.1 : 1 / 1.1), 0.25, 5);
      this.setScaleMode('custom', next);
      requestAnimationFrame(() => {
        const box = this.pageBoxes.get(anchorPage);
        if (box) box.scrollIntoView({ block: 'center', behavior: 'auto' });
      });
    }, { passive: false });
  }

  /* ══════════ 文档加载 ══════════ */
  async open(path, data, fileInfo) {
    await this.close();
    this.fileInfo = fileInfo;
    const task = this.ctx.pdfjsLib.getDocument({
      data,
      cMapUrl: './vendor/cmaps/',
      standardFontDataUrl: './vendor/standard_fonts/',
      cMapPacked: true,
      enableXfa: false,
      fontExtraProperties: true
    });
    this.doc = await task.promise;

    // 预取原始页面尺寸（scale=1, rotation=0）
    const first = await this.doc.getPage(1);
    const base = first.getViewport({ scale: 1 });
    this.baseDims.set(1, { w: base.width, h: base.height });
    this.baseViewport0 = base;

    // 先显示视图再计算缩放（display:none 时 clientWidth 为 0，会导致负缩放）
    this.elViewer.classList.remove('hidden');
    document.getElementById('welcome').classList.add('hidden');

    this._buildPages();
    this._computeInitialScale();
    this._relayout(true);

    return { numPages: this.doc.numPages };
  }

  async replaceDocument(data, restorePage = 1) {
    const task = this.ctx.pdfjsLib.getDocument({
      data: data.slice(0),
      cMapUrl: './vendor/cmaps/',
      standardFontDataUrl: './vendor/standard_fonts/',
      cMapPacked: true,
      enableXfa: false,
      fontExtraProperties: true
    });
    const nextDoc = await task.promise;
    if (!nextDoc.numPages) throw new Error('预览文档没有页面');

    const oldDoc = this.doc;
    const oldBoxes = new Map(this.pageBoxes);
    const oldTasks = new Map(this.renderTasks);
    const restore = clamp(Math.round(restorePage) || 1, 1, nextDoc.numPages);

    const nextBoxes = new Map();
    const nextRendered = new Set();
    const nextBaseDims = new Map();
    const fragment = document.createDocumentFragment();
    for (let i = 1; i <= nextDoc.numPages; i++) {
      const outer = document.createElement('div');
      outer.className = 'page-outer';
      outer.dataset.page = String(i);
      outer.innerHTML = '<div class="annot-layer"></div><div class="text-layer"></div>';
      fragment.appendChild(outer);
      nextBoxes.set(i, outer);
    }

    const first = await nextDoc.getPage(1);
    const nextBaseViewport0 = first.getViewport({ scale: 1 });
    nextBaseDims.set(1, { w: nextBaseViewport0.width, h: nextBaseViewport0.height });
    const page = await nextDoc.getPage(restore);
    const viewport = page.getViewport({ scale: this.scale, rotation: this.rotation });
    const box = nextBoxes.get(restore);
    box.style.width = `${viewport.width}px`;
    box.style.height = `${viewport.height}px`;
    const canvas = document.createElement('canvas');
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;
    box.insertBefore(canvas, box.firstChild);
    await page.render({
      canvasContext: canvas.getContext('2d'),
      viewport,
      transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null
    }).promise;
    nextRendered.add(restore);
    box.dataset.rendered = '1';

    oldTasks.forEach(item => { try { item.cancel(); } catch (_) { /* noop */ } });
    for (const item of oldBoxes.values()) this.observer.unobserve(item);
    if (oldDoc) { try { oldDoc.destroy(); } catch (_) { /* noop */ } }

    this.doc = nextDoc;
    this.pageBoxes = nextBoxes;
    this.renderTasks.clear();
    this.renderedPages = nextRendered;
    this.baseDims = nextBaseDims;
    this.baseViewport0 = nextBaseViewport0;
    this.textCache.clear();
    this.elViewer.innerHTML = '';
    this.elViewer.appendChild(fragment);
    this._relayout(false);
    for (const item of nextBoxes.values()) this.observer.observe(item);
    this.goToPage(restore, false);
    this.renderTextLayer(restore, viewport);
    return { numPages: nextDoc.numPages };
  }

  async close() {
    if (this.doc) {
      try { this.doc.destroy(); } catch (_) { /* noop */ }
    }
    this.doc = null;
    this.renderTasks.forEach(t => { try { t.cancel(); } catch (_) { /* noop */ } });
    this.renderTasks.clear();
    this.pageBoxes.clear();
    this.renderedPages.clear();
    this.baseDims.clear();
    this.textCache.clear();
    this.currentPage = 0;
    this.rotation = 0;
    this.elViewer.innerHTML = '';
    this.elViewer.classList.add('hidden');
    document.getElementById('welcome').classList.remove('hidden');
  }

  _buildPages() {
    this.elViewer.innerHTML = '';
    this.pageBoxes.clear();
    this.renderedPages.clear();
    const frag = document.createDocumentFragment();
    for (let i = 1; i <= this.doc.numPages; i++) {
      const outer = document.createElement('div');
      outer.className = 'page-outer';
      outer.dataset.page = String(i);
      outer.innerHTML =
        '<div class="annot-layer"></div>' +
        '<div class="text-layer"></div>';
      frag.appendChild(outer);
      this.pageBoxes.set(i, outer);
      this.observer.observe(outer);
    }
    this.elViewer.appendChild(frag);
  }

  /* ══════════ 缩放 / 适配 ══════════ */
  _fitWidthScale() {
    const availW = this.elViewer.clientWidth - 48;
    return availW / (this.baseViewport0 ? this.baseViewport0.width : 612);
  }
  _fitPageScale() {
    const availW = this.elViewer.clientWidth - 48;
    const availH = this.elViewer.clientHeight - 44;
    const v0 = this.baseViewport0 || { width: 612, height: 792 };
    return Math.min(availW / v0.width, availH / v0.height);
  }
  _computeInitialScale() {
    this.scale = this.scaleMode === 'fit-page' ? this._fitPageScale() : this._fitWidthScale();
  }

  setScaleMode(mode, customScale) {
    this.scaleMode = mode;
    if (mode === 'custom' && customScale) this.scale = customScale;
    else this._computeInitialScale();
    this._relayout(true);
  }

  zoomStep(dir) {
    const next = clamp(this.scale * (dir > 0 ? 1.2 : 1 / 1.2), 0.25, 5);
    this.setScaleMode('custom', next);
  }

  rotate() {
    this.rotation = (this.rotation + 90) % 360;
    this._relayout(true);
  }

  /* ══════════ 布局与渲染 ══════════ */
  async _relayout(rerenderAll = false) {
    if (!this.doc) return;
    if (rerenderAll) {
      this.renderTasks.forEach(t => { try { t.cancel(); } catch (_) { /* noop */ } });
      this.renderTasks.clear();
      this.renderedPages.clear();
    }
    // 逐页设置容器尺寸（以第 1 页基准尺寸近似所有页；不同尺寸页在渲染时按实际视口修正）
    for (const [num, box] of this.pageBoxes) {
      let dims = this.baseDims.get(num);
      if (!dims) {
        // 惰性获取并缓存真实尺寸
        try {
          const page = await this.doc.getPage(num);
          const v = page.getViewport({ scale: 1 });
          dims = { w: v.width, h: v.height };
          this.baseDims.set(num, dims);
        } catch (_) { dims = this.baseDims.get(1); }
      }
      const swapped = this.rotation % 180 !== 0;
      const w = (swapped ? dims.h : dims.w) * this.scale;
      const h = (swapped ? dims.w : dims.h) * this.scale;
      box.style.width = `${w}px`;
      box.style.height = `${h}px`;
      box.style.setProperty('--scale-factor', String(this.scale));
      if (rerenderAll) {
        const canvas = box.querySelector('canvas');
        if (canvas) canvas.remove();
        const tl = box.querySelector('.text-layer');
        if (tl) tl.innerHTML = '';
        this.textCache.delete(`tl:${num}`);
      }
    }
    if (rerenderAll) {
      // 重新 observe 触发 IntersectionObserver 对当前可见页的回调，恢复懒渲染
      for (const [, box] of this.pageBoxes) {
        this.observer.unobserve(box);
        this.observer.observe(box);
      }
    }
    this.ctx.onStateChange({ type: 'zoom' });
  }

  /** 由 IntersectionObserver 触发的懒渲染 */
  async renderPage(num) {
    if (!this.doc || this.renderedPages.has(num) || this.renderTasks.has(num)) return;
    const box = this.pageBoxes.get(num);
    if (!box) return;
    try {
      const page = await this.doc.getPage(num);
      const viewport = page.getViewport({ scale: this.scale, rotation: this.rotation });
      // 修正容器尺寸为真实尺寸
      box.style.width = `${viewport.width}px`;
      box.style.height = `${viewport.height}px`;

      const canvas = document.createElement('canvas');
      const dpr = window.devicePixelRatio || 1;
      canvas.width = Math.floor(viewport.width * dpr);
      canvas.height = Math.floor(viewport.height * dpr);
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;
      box.insertBefore(canvas, box.firstChild);

      const rc = page.render({
        canvasContext: canvas.getContext('2d'),
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null
      });
      this.renderTasks.set(num, rc);
      await rc.promise;
      this.renderedPages.add(num);
      box.dataset.rendered = '1';
      this.ctx.onStateChange({ type: 'page-rendered', page: num, viewport });
    } catch (err) {
      if (!/cancel/i.test(String(err && err.message))) {
        console.error(`渲染第 ${num} 页失败:`, err);
      }
    } finally {
      this.renderTasks.delete(num);
    }
  }

  async renderTextLayer(num, viewport) {
    if (!this.doc || this.textCache.has(`tl:${num}`)) return;
    const box = this.pageBoxes.get(num);
    if (!box) return;
    const tlEl = box.querySelector('.text-layer');
    try {
      const page = await this.doc.getPage(num);
      const tc = await page.getTextContent();
      this.textCache.set(num, tc);
      const { TextLayer } = this.ctx.pdfjsLib;
      if (!TextLayer) return;
      const tl = new TextLayer({ textContentSource: tc, container: tlEl, viewport });
      await tl.render();
      this.textCache.set(`tl:${num}`, true);
    } catch (err) {
      if (!/cancel|abort/i.test(String(err && err.message))) {
        console.error(`文本层 ${num} 失败:`, err);
      }
    }
  }

  /* ══════════ 定位 ══════════ */
  goToPage(num, center = true) {
    num = clamp(Math.round(num) || 1, 1, this.doc ? this.doc.numPages : 1);
    const box = this.pageBoxes.get(num);
    if (!box) return;
    box.scrollIntoView({ block: center ? 'center' : 'start', behavior: 'auto' });
    this._setCurrentPage(num);
  }

  _setCurrentPage(num) {
    if (num === this.currentPage) return;
    this.currentPage = num;
    this.ctx.onStateChange({ type: 'page', page: num });
    for (const [n, box] of this.pageBoxes) {
      box.classList.toggle('current', n === num);
    }
    this._savePosition();
  }

  _savePosition = debounce(() => {
    if (!this.fileInfo || !this.currentPage) return;
    try {
      localStorage.setItem(`onyx:pos:${this.ctx.fileKey}`, String(this.currentPage));
    } catch (_) { /* noop */ }
  }, 400);

  restorePosition() {
    try {
      const v = localStorage.getItem(`onyx:pos:${this.ctx.fileKey}`);
      if (v) this.goToPage(parseInt(v, 10));
    } catch (_) { /* noop */ }
  }

  /* ══════════ 事件 ══════════ */
  _bindObserver() {
    this.observer = new IntersectionObserver((entries) => {
      for (const en of entries) {
        if (en.isIntersecting) {
          const num = parseInt(en.target.dataset.page, 10);
          this.renderPage(num);
        }
      }
    }, { root: this.elViewer, rootMargin: '700px 0px' });
  }

  _bindScroll() {
    this.elViewer.addEventListener('scroll', debounce(() => {
      if (!this.doc || !this.pageBoxes.size) return;
      const mid = this.elViewer.scrollTop + this.elViewer.clientHeight / 2;
      let best = 1, bestDist = Infinity;
      for (const [num, box] of this.pageBoxes) {
        const top = box.offsetTop, h = box.offsetHeight;
        const d = Math.abs(top + h / 2 - mid);
        if (d < bestDist) { bestDist = d; best = num; }
      }
      this._setCurrentPage(best);
    }, 120));
  }

  _bindResize() {
    window.addEventListener('resize', debounce(() => {
      if (!this.doc) return;
      if (this.scaleMode !== 'custom') this._computeInitialScale();
      this._relayout(false);
    }, 180));
  }

  /* 文本选择 → 归一化区域（供批注使用），返回 PDF 用户空间矩形数组 */
  async selectionToPdfRects(pageNum) {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || sel.rangeCount === 0) return null;
    const box = this.pageBoxes.get(pageNum);
    if (!box || !box.contains(sel.anchorNode)) return null;
    const canvas = box.querySelector('canvas');
    if (!canvas) return null;

    const page = await this.doc.getPage(pageNum);
    const viewport = page.getViewport({ scale: this.scale, rotation: this.rotation });
    const cRect = canvas.getBoundingClientRect();
    const rects = [];
    for (let i = 0; i < sel.rangeCount; i++) {
      const r = sel.getRangeAt(i).getBoundingClientRect();
      if (r.width < 1 || r.height < 1) continue;
      // 可能把矩形切到跨行多块：用 client rects
      if (typeof r.width !== 'undefined' && r.left >= cRect.left - 2 && r.right <= cRect.right + 2) {
        rects.push(r);
      }
    }
    if (!rects.length) return null;

    const text = sel.toString();
    sel.removeAllRanges();

    const out = [];
    for (const r of rects) {
      const x1 = Math.max(r.left - cRect.left, 0);
      const y1 = Math.max(r.top - cRect.top, 0);
      const x2 = Math.min(r.right - cRect.left, cRect.width);
      const y2 = Math.min(r.bottom - cRect.top, cRect.height);
      const p1 = viewport.convertToPdfPoint(x1, y1);
      const p2 = viewport.convertToPdfPoint(x2, y2);
      out.push({
        x: Math.min(p1[0], p2[0]), y: Math.min(p1[1], p2[1]),
        x2: Math.max(p1[0], p2[0]), y2: Math.max(p1[1], p2[1])
      });
    }
    return { rects: out, text };
  }

  /* PDF 用户空间矩形 → 视口 CSS 像素 */
  pdfRectToViewport(pageNum, pdfRect, viewport) {
    const [x1, y1] = viewport.convertToViewportPoint(pdfRect.x, pdfRect.y);
    const [x2, y2] = viewport.convertToViewportPoint(pdfRect.x2, pdfRect.y2);
    return {
      x: Math.min(x1, x2), y: Math.min(y1, y2),
      w: Math.abs(x2 - x1), h: Math.abs(y2 - y1)
    };
  }
}
