/**
 * OnyxPDF - 侧栏（缩略图 / 书签）
 */
import { escapeHtml } from './utils.js';

export class Sidebar {
  constructor(ctx) {
    this.ctx = ctx;
    this.thumbObserver = null;
    this._bindTabs();
  }

  _bindTabs() {
    document.querySelectorAll('.side-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.side-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const target = tab.dataset.tab;
        document.querySelectorAll('.side-panel').forEach(p => p.classList.add('hidden'));
        const panel = document.getElementById(`panel-${target}`);
        panel.classList.remove('hidden');
        if (target === 'thumbs' && this.ctx.viewer.doc) this.buildThumbs();
        if (target === 'outline' && this.ctx.viewer.doc) this.buildOutline();
      });
    });
  }

  reset() {
    document.getElementById('panel-thumbs').innerHTML = '';
    document.getElementById('panel-outline').innerHTML = '';
    document.getElementById('panel-search').classList.add('hidden');
    if (this.thumbObserver) this.thumbObserver.disconnect();
  }

  /* ══════════ 缩略图 ══════════ */
  buildThumbs() {
    const panel = document.getElementById('panel-thumbs');
    panel.innerHTML = '';
    if (this.thumbObserver) this.thumbObserver.disconnect();
    this.thumbObserver = new IntersectionObserver((entries) => {
      for (const en of entries) {
        if (en.isIntersecting) {
          this._renderThumb(parseInt(en.target.dataset.page, 10), en.target);
          this.thumbObserver.unobserve(en.target);
        }
      }
    }, { root: panel, rootMargin: '300px 0px' });

    // 页面排序栏（拖拽缩略图后出现）
    const bar = document.createElement('div');
    bar.className = 'reorder-bar hidden';
    bar.innerHTML = '<span>拖动缩略图调整顺序</span><button class="btn" data-role="apply">应用排序</button>';
    bar.querySelector('[data-role="apply"]').addEventListener('click', () => {
      const order = [...panel.querySelectorAll('.thumb-item')]
        .map(el => parseInt(el.dataset.page, 10) - 1);
      bar.classList.add('hidden');
      if (this.ctx.onReorderApply) this.ctx.onReorderApply(order);
    });
    panel.appendChild(bar);

    for (let i = 1; i <= this.ctx.viewer.doc.numPages; i++) {
      const item = document.createElement('div');
      item.className = 'thumb-item';
      item.dataset.page = String(i);
      item.draggable = true;
      item.innerHTML = `<div class="thumb-canvas"></div><div class="thumb-page">${i}</div>`;
      item.addEventListener('click', () => {
        this.ctx.viewer.goToPage(i, false);
      });
      // 拖拽排序
      item.addEventListener('dragstart', (e) => {
        item.classList.add('dragging');
        e.dataTransfer.setData('text/plain', String(i));
        e.dataTransfer.effectAllowed = 'move';
      });
      item.addEventListener('dragend', () => item.classList.remove('dragging'));
      item.addEventListener('dragover', (e) => {
        e.preventDefault();
        item.classList.add('drag-over');
      });
      item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
      item.addEventListener('drop', (e) => {
        e.preventDefault();
        item.classList.remove('drag-over');
        const dragged = panel.querySelector('.thumb-item.dragging');
        if (!dragged || dragged === item) return;
        const r = item.getBoundingClientRect();
        const after = (e.clientY - r.top) > r.height / 2;
        panel.insertBefore(dragged, after ? item.nextSibling : item);
        bar.classList.remove('hidden');
      });
      panel.appendChild(item);
      this.thumbObserver.observe(item);
    }
    this._markCurrent();
  }

  async _renderThumb(num, itemEl) {
    if (itemEl.dataset.done) return;
    itemEl.dataset.done = '1';
    try {
      const page = await this.ctx.viewer.doc.getPage(num);
      const v0 = page.getViewport({ scale: 1 });
      const scale = 190 / v0.width;
      const viewport = page.getViewport({ scale, rotation: this.ctx.viewer.rotation });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      const holder = itemEl.querySelector('.thumb-canvas');
      holder.innerHTML = '';
      holder.appendChild(canvas);
    } catch (err) {
      console.error(`缩略图 ${num} 失败:`, err);
    }
  }

  markCurrentThumb(page) {
    document.querySelectorAll('.thumb-item').forEach(el => {
      el.classList.toggle('active', parseInt(el.dataset.page, 10) === page);
    });
    const active = document.querySelector('.thumb-item.active');
    if (active && active.getBoundingClientRect) {
      const pr = active.parentElement.getBoundingClientRect();
      const r = active.getBoundingClientRect();
      if (r.top < pr.top || r.bottom > pr.bottom) {
        active.scrollIntoView({ block: 'nearest' });
      }
    }
  }

  _markCurrent() {
    this.markCurrentThumb(this.ctx.viewer.currentPage || 1);
  }

  /* ══════════ 书签 ══════════ */
  async buildOutline() {
    const panel = document.getElementById('panel-outline');
    panel.innerHTML = '';
    let outline = null;
    try { outline = await this.ctx.viewer.doc.getOutline(); } catch (_) { /* noop */ }
    if (!outline || !outline.length) {
      panel.innerHTML = '<div class="outline-empty">本文档没有内嵌书签目录</div>';
      return;
    }
    const frag = document.createDocumentFragment();
    const walk = async (items, depth) => {
      for (const it of items) {
        const btn = document.createElement('button');
        btn.className = 'outline-item' + (depth === 0 ? ' bold' : '');
        btn.style.paddingLeft = `${12 + depth * 14}px`;
        btn.textContent = it.title || '(未命名)';
        btn.addEventListener('click', () => this._gotoDest(it.dest, btn));
        frag.appendChild(btn);
        if (it.items && it.items.length) await walk(it.items, depth + 1);
      }
    };
    await walk(outline, 0);
    panel.appendChild(frag);
  }

  async _gotoDest(dest, btn) {
    try {
      const doc = this.ctx.viewer.doc;
      let d = dest;
      if (typeof d === 'string') d = await doc.getDestination(d);
      if (!d || !d.length) return;
      const pageIndex = await doc.getPageIndex(d[0]);
      this.ctx.viewer.goToPage(pageIndex + 1, false);
    } catch (err) {
      console.error('书签跳转失败:', err);
    }
  }
}
