/**
 * OnyxPDF - 全文检索
 * 逐页提取文本 → 匹配关键词 → 视口坐标高亮 + 结果列表导航
 * 说明：V1 按文本项内匹配（跨行拼接匹配列入二期）
 */
import { debounce, escapeHtml } from './utils.js';

export class Search {
  constructor(ctx) {
    this.ctx = ctx;
    this.query = '';
    this.hits = [];        // { page, indexInPage, rects:[viewportRect], text }
    this.cursor = -1;
    this.searching = false;
    this._bind();
  }

  reset() {
    this.query = '';
    this.hits = [];
    this.cursor = -1;
    this._clearHighlights();
    this._renderResults();
    this._updateCount();
  }

  _clearHighlights() {
    document.querySelectorAll('.search-hl').forEach(el => el.remove());
  }

  _bind() {
    const input = document.getElementById('search-input');
    input.addEventListener('input', debounce(() => {
      this.run(input.value.trim());
    }, 300));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.jump(e.shiftKey ? -1 : 1);
      }
    });
    document.getElementById('search-prev').addEventListener('click', () => this.jump(-1));
    document.getElementById('search-next').addEventListener('click', () => this.jump(1));
  }

  async run(query) {
    this.query = query;
    this._clearHighlights();
    this.hits = [];
    this.cursor = -1;
    if (!query || !this.ctx.viewer.doc) {
      this._renderResults();
      this._updateCount();
      return;
    }
    this.searching = true;
    this._renderResults('检索中…');

    const v = this.ctx.viewer;
    const q = query.toLowerCase();
    const found = [];
    try {
      for (let p = 1; p <= v.doc.numPages; p++) {
        let tc = v.textCache.get(p);
        const page = await v.doc.getPage(p);
        if (!tc) {
          tc = await page.getTextContent();
          v.textCache.set(p, tc);
        }
        const viewport = page.getViewport({ scale: v.scale, rotation: v.rotation });
        let idx = 0;
        for (const item of tc.items) {
          const s = item.str.toLowerCase();
          let pos = s.indexOf(q);
          while (pos >= 0) {
            // 按字符宽度比例近似匹配区域
            const ratio = item.width / Math.max(item.str.length, 1);
            const x1 = item.transform[4] + pos * ratio;
            const w = q.length * ratio;
            const h = item.height || Math.abs(item.transform[3]) || 10;
            const vr = v.pdfRectToViewport(p,
              { x: x1, y: item.transform[5], x2: x1 + w, y2: item.transform[5] + h }, viewport);
            found.push({ page: p, idx: idx++, rect: vr, text: item.str });
            pos = s.indexOf(q, pos + 1);
          }
        }
        if (found.length > 2000) break; // 防御：超长结果截断
      }
    } catch (err) {
      console.error('检索失败:', err);
    }
    if (this.query !== query) return; // 用户已改关键词
    this.searching = false;
    this.hits = found;
    this._paintHighlights();
    this._renderResults();
    this._updateCount();
    if (found.length) this.jump(1);
  }

  _paintHighlights() {
    this._clearHighlights();
    for (let i = 0; i < this.hits.length; i++) {
      const h = this.hits[i];
      const box = this.ctx.viewer.pageBoxes.get(h.page);
      if (!box) continue;
      const div = document.createElement('div');
      div.className = 'search-hl' + (i === this.cursor ? ' current' : '');
      div.style.left = `${h.rect.x}px`; div.style.top = `${h.rect.y}px`;
      div.style.width = `${h.rect.w}px`; div.style.height = `${h.rect.h}px`;
      box.appendChild(div);
    }
  }

  jump(dir) {
    if (!this.hits.length) return;
    this.cursor = (this.cursor + dir + this.hits.length) % this.hits.length;
    const h = this.hits[this.cursor];
    this.ctx.viewer.goToPage(h.page);
    document.querySelectorAll('.search-hl.current').forEach(el => el.classList.remove('current'));
    const boxes = document.querySelectorAll(`.page-outer[data-page="${h.page}"] .search-hl`);
    // 当前命中对应每页高亮中的第 idx 个
    const pageHits = this.hits.filter(x => x.page === h.page);
    const localIdx = pageHits.indexOf(h);
    if (boxes[localIdx]) {
      boxes[localIdx].classList.add('current');
      boxes[localIdx].scrollIntoView({ block: 'nearest' });
    }
    this._renderResults();
    this._updateCount();
  }

  _updateCount() {
    const el = document.getElementById('search-count');
    el.textContent = this.hits.length
      ? `${this.cursor + 1}/${this.hits.length}`
      : (this.query ? '0/0' : '0/0');
  }

  _renderResults(placeholder) {
    const panel = document.getElementById('search-results');
    if (placeholder) {
      panel.innerHTML = `<div class="search-none">${escapeHtml(placeholder)}</div>`;
      return;
    }
    if (!this.query) {
      panel.innerHTML = '<div class="search-none">输入关键词开始全文检索</div>';
      return;
    }
    if (!this.hits.length) {
      panel.innerHTML = '<div class="search-none">未找到匹配内容</div>';
      return;
    }
    const q = this.query;
    const frag = document.createDocumentFragment();
    this.hits.forEach((h, i) => {
      const div = document.createElement('div');
      div.className = 'search-hit' + (i === this.cursor ? ' active' : '');
      const esc = escapeHtml(h.text);
      const pos = h.text.toLowerCase().indexOf(q.toLowerCase());
      let html;
      if (pos >= 0) {
        html = `${escapeHtml(h.text.slice(0, pos))}<mark>${escapeHtml(h.text.slice(pos, pos + q.length))}</mark>${escapeHtml(h.text.slice(pos + q.length))}`;
      } else {
        html = esc;
      }
      div.innerHTML = `<div class="sh-page">P.${h.page}</div><div class="sh-text">${html}</div>`;
      div.addEventListener('click', () => {
        this.cursor = i;
        this.ctx.viewer.goToPage(h.page);
        this._renderResults();
        this._updateCount();
      });
      frag.appendChild(div);
    });
    panel.innerHTML = '';
    panel.appendChild(frag);
  }
}
