/**
 * OnyxPDF - 工具集（移植 PDFPatcher 核心功能）
 * 合并 / 拆分 / 提取页面 / 删除页面 / 旋转页面 / 插入空白页 /
 * 文本水印 / 添加页码 / 图片转 PDF / 清理优化
 * 全部基于 pdf-lib 在本地完成，无网络依赖。
 */
import { openModal } from './ui.js';
import { toast } from './utils.js';

/* ══════════ 基础工具 ══════════ */

/** 解析页码范围字符串 "1-3,5,8-"（1-based，含端点），返回去重升序 0-based 数组 */
export function parseRanges(str, numPages) {
  const s = String(str || '').trim();
  if (!s) throw new Error('页码范围不能为空');
  const out = new Set();
  for (const part of s.split(/[,，]/)) {
    const p = part.trim();
    if (!p) continue;
    let a, b;
    const m = /^(\d+)?\s*-\s*(\d+)?$/.exec(p);
    if (m) {
      a = m[1] ? parseInt(m[1], 10) : 1;
      b = m[2] ? parseInt(m[2], 10) : numPages;
    } else if (/^\d+$/.test(p)) {
      a = b = parseInt(p, 10);
    } else {
      throw new Error(`无法识别的页码范围："${p}"`);
    }
    a = Math.max(1, a); b = Math.min(numPages, b);
    if (a > b) throw new Error(`无效的页码范围："${p}"`);
    for (let i = a; i <= b; i++) out.add(i - 1);
  }
  if (!out.size) throw new Error('页码范围为空');
  return [...out].sort((x, y) => x - y);
}

/** 解析页码范围字符串为独立分组，例如 "1-3,3-5,7-" 保持三个输出范围。 */
export function parseRangeGroups(str, numPages) {
  const s = String(str || '').trim();
  if (!s) throw new Error('页码范围不能为空');
  const groups = [];
  for (const part of s.split(/[,，]/)) {
    const p = part.trim();
    if (!p) continue;
    let a, b;
    const m = /^(\d+)?\s*-\s*(\d+)?$/.exec(p);
    if (m) {
      a = m[1] ? parseInt(m[1], 10) : 1;
      b = m[2] ? parseInt(m[2], 10) : numPages;
    } else if (/^\d+$/.test(p)) {
      a = b = parseInt(p, 10);
    } else {
      throw new Error(`无法识别的页码范围："${p}"`);
    }
    a = Math.max(1, a);
    b = Math.min(numPages, b);
    if (a > b) throw new Error(`无效的页码范围："${p}"`);
    groups.push(Array.from({ length: b - a + 1 }, (_, index) => a + index - 1));
  }
  if (!groups.length) throw new Error('页码范围为空');
  return groups;
}

/**
 * 加载操作源：优先当前打开文档（内存字节，支持虚拟文档），否则弹文件选择。
 * 返回 { doc, base }（base = 不带扩展名的文件名，用于产物命名）。
 */
async function loadSourceDoc(ctx) {
  let bytes = null;
  let base = 'document';
  if (ctx.getBytes && ctx.getBytes()) {
    bytes = ctx.getBytes();
    base = (ctx.getName() || 'document').replace(/\.pdf$/i, '');
  } else {
    const paths = await window.onyx.openMultiDialog();
    if (!paths || !paths.length) return null;
    const res = await window.onyx.readPdf(paths[0]);
    if (!res.ok) throw new Error(`读取失败：${paths[0]}（${res.error}）`);
    bytes = res.data;
    base = nameWithoutExt(paths[0]);
  }
  const doc = await window.PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
  return { doc, base };
}

async function readPdfDoc(path) {
  const res = await window.onyx.readPdf(path);
  if (!res.ok) throw new Error(`读取失败：${path}（${res.error}）`);
  return window.PDFLib.PDFDocument.load(res.data);
}

async function saveDoc(doc, suggestedName, title = '另存为') {
  console.error('[saveDoc] start', suggestedName);
  const bytes = await doc.save();
  console.error('[saveDoc] saved bytes', bytes.length);
  const r = await window.onyx.exportPdf({ bytes, suggestedName, title });
  if (r.ok) toast(`已保存：${r.path}`);
  return r;
}

function nameWithoutExt(p) {
  return p.replace(/^.*[\\/]/, '').replace(/\.pdf$/i, '');
}

/* ══════════ 1. 合并 PDF ══════════ */
export async function toolMerge() {
  const paths = await window.onyx.openMultiDialog();
  if (!paths) return;
  if (paths.length < 2) { toast('请至少选择 2 个 PDF 文件', true); return; }
  try {
    const out = await window.PDFLib.PDFDocument.create();
    for (const p of paths) {
      const src = await readPdfDoc(p);
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach(pg => out.addPage(pg));
    }
    await saveDoc(out, '合并文档.pdf', '合并 PDF - 另存为');
  } catch (err) {
    toast(err.message, true);
  }
}

/* ══════════ 2. 拆分 PDF ══════════ */
export function toolSplit(ctx) {
  return new Promise(resolve => {
    openModal({
      title: '拆分 PDF',
      bodyHTML: `
      <div class="field">
        <label>拆分方式</label>
        <select data-role="mode">
          <option value="every">每 N 页拆为一个文件</option>
          <option value="range">按页码范围拆分（逗号分隔多个范围）</option>
          <option value="single">每页拆为一个文件</option>
        </select>
      </div>
      <div class="field" data-role="wrap-n">
        <label>N（页数）</label>
        <input type="number" data-role="n" value="1" min="1" />
      </div>
      <div class="field hidden" data-role="wrap-ranges">
        <label>页码范围（如 1-3,4-6,7-）</label>
        <input type="text" data-role="ranges" placeholder="1-3,4-6,7-" />
      </div>
      <p class="hint">拆分结果将写入你选择的输出文件夹，文件名自动加序号。</p>
    `,
      buttons: [
        { label: '取消', onClick: (m) => { m.close(); resolve(null); } },
        {
          label: '选择文件夹并拆分', primary: true,
          onClick: async (m) => {
            const dir = await window.onyx.openDirDialog();
            if (!dir) return;
            const mode = m.query('[data-role="mode"]').value;
            try {
              const source = await loadSourceDoc(ctx);
              if (!source) return;
              const src = source.doc;
              const srcPath = source.base;
              const n = src.getPageCount();
              const groups = [];
              if (mode === 'single') {
                for (let i = 0; i < n; i++) groups.push([i]);
              } else if (mode === 'every') {
                const size = Math.max(1, parseInt(m.query('[data-role="n"]').value, 10) || 1);
                for (let i = 0; i < n; i += size) {
                  groups.push(Array.from({ length: Math.min(size, n - i) }, (_, k) => i + k));
                }
              } else {
                groups.push(...parseRangeGroups(m.query('[data-role="ranges"]').value, n));
              }
              const base = nameWithoutExt(srcPath);
              let idx = 0;
              for (const g of groups) {
                const out = await window.PDFLib.PDFDocument.create();
                const pages = await out.copyPages(src, g);
                pages.forEach(pg => out.addPage(pg));
                const bytes = await out.save();
                const r = await window.onyx.writePath({
                  path: `${dir}\\${base}_part${++idx}.pdf`, bytes
                });
                if (!r.ok) throw new Error(r.error);
              }
              toast(`拆分完成：${idx} 个文件已写入输出文件夹`);
              resolve(true);
              m.close();
            } catch (err) {
              toast(err.message, true);
            }
          }
        }
      ],
      onClose: () => resolve(null)
    });
    const modal = document.querySelector('#modal-mask .modal');
    modal.querySelector('[data-role="mode"]').addEventListener('change', event => {
      const value = event.target.value;
      modal.querySelector('[data-role="wrap-n"]').classList.toggle('hidden', value !== 'every');
      modal.querySelector('[data-role="wrap-ranges"]').classList.toggle('hidden', value !== 'range');
    });
  });
}

/* ══════════ 3. 提取页面 ══════════ */
export function toolExtract(ctx) {
  const total = ctx.pageCount() || 0;
  openModal({
    title: '提取页面',
    bodyHTML: `
      <div class="field">
        <label>要提取的页码范围${total ? `（共 ${total} 页）` : ''}</label>
        <input type="text" data-role="ranges" placeholder="如 1-3,5,8-" />
      </div>
      <p class="hint">提取的页面将组成一个新的 PDF 文件。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '提取并另存', primary: true,
        onClick: async (m) => {
          try {
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const src = ld.doc;
            const srcPath = ld.base;
            const indices = parseRanges(m.query('[data-role="ranges"]').value, src.getPageCount());
            const out = await window.PDFLib.PDFDocument.create();
            const pages = await out.copyPages(src, indices);
            pages.forEach(pg => out.addPage(pg));
            await saveDoc(out, `${nameWithoutExt(srcPath)}_提取.pdf`, '提取页面 - 另存为');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/* ══════════ 4. 删除页面 ══════════ */
export function toolDeletePages(ctx) {
  const total = ctx.pageCount() || 0;
  openModal({
    title: '删除页面',
    bodyHTML: `
      <div class="field">
        <label>要删除的页码范围${total ? `（共 ${total} 页）` : ''}</label>
        <input type="text" data-role="ranges" placeholder="如 2,5-7" />
      </div>
      <p class="hint">剩余页面将另存为新的 PDF 文件（原文件不修改）。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '删除并另存', primary: true,
        onClick: async (m) => {
          try {
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const src = ld.doc;
            const srcPath = ld.base;
            const n = src.getPageCount();
            const del = new Set(parseRanges(m.query('[data-role="ranges"]').value, n));
            const keep = [];
            for (let i = 0; i < n; i++) if (!del.has(i)) keep.push(i);
            if (!keep.length) throw new Error('不能删除全部页面');
            const out = await window.PDFLib.PDFDocument.create();
            const pages = await out.copyPages(src, keep);
            pages.forEach(pg => out.addPage(pg));
            await saveDoc(out, `${nameWithoutExt(srcPath)}_删除后.pdf`, '删除页面 - 另存为');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/* ══════════ 5. 旋转页面 ══════════ */
export function toolRotatePages(ctx) {
  const total = ctx.pageCount() || 0;
  openModal({
    title: '旋转页面',
    bodyHTML: `
      <div class="field-row">
        <div class="field">
          <label>页码范围（留空=全部）</label>
          <input type="text" data-role="ranges" placeholder="如 1-3" />
        </div>
        <div class="field">
          <label>角度</label>
          <select data-role="angle">
            <option value="90">顺时针 90°</option>
            <option value="180">180°</option>
            <option value="270">逆时针 90°</option>
          </select>
        </div>
      </div>
      <p class="hint">旋转结果另存为新 PDF 文件。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '旋转并另存', primary: true,
        onClick: async (m) => {
          try {
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const src = ld.doc;
            const srcPath = ld.base;
            const n = src.getPageCount();
            const rangesStr = m.query('[data-role="ranges"]').value.trim();
            const targets = rangesStr ? new Set(parseRanges(rangesStr, n)) : new Set(Array.from({ length: n }, (_, i) => i));
            const angle = parseInt(m.query('[data-role="angle"]').value, 10);
            const pages = src.getPages();
            targets.forEach(i => {
              pages[i].setRotation(window.PDFLib.degrees((pages[i].getRotation().angle + angle) % 360));
            });
            await saveDoc(src, `${nameWithoutExt(srcPath)}_旋转.pdf`, '旋转页面 - 另存为');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/* ══════════ 6. 插入空白页 ══════════ */
export function toolInsertBlank(ctx) {
  const total = ctx.pageCount() || 0;
  openModal({
    title: '插入空白页',
    bodyHTML: `
      <div class="field-row">
        <div class="field">
          <label>插入位置（第几页）</label>
          <input type="number" data-role="pos" value="${total || 1}" min="1" />
        </div>
        <div class="field">
          <label>方向</label>
          <select data-role="where">
            <option value="before">在该页之前</option>
            <option value="after" selected>在该页之后</option>
          </select>
        </div>
      </div>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '插入并另存', primary: true,
        onClick: async (m) => {
          try {
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const src = ld.doc;
            const srcPath = ld.base;
            const n = src.getPageCount();
            let pos = parseInt(m.query('[data-role="pos"]').value, 10);
            pos = Math.min(Math.max(1, pos || 1), n);
            const after = m.query('[data-role="where"]').value === 'after';
            const sizeMode = m.query('[data-role="size"]').value;
            const count = Math.min(100, Math.max(1, parseInt(m.query('[data-role="count"]').value, 10) || 1));

            const refPage = src.getPage(after ? pos - 1 : Math.max(0, pos - 1));
            let w = refPage.getWidth(), h = refPage.getHeight();
            if (sizeMode === 'a4p') { w = 595.28; h = 841.89; }
            else if (sizeMode === 'a4l') { w = 841.89; h = 595.28; }

            const insertIndex = after ? pos : pos - 1;
            for (let i = 0; i < count; i++) {
              src.insertPage(insertIndex, [w, h]);
            }
            await saveDoc(src, `${nameWithoutExt(srcPath)}_插入空白页.pdf`, '插入空白页 - 另存为');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/* ══════════ 7. 文本水印 ══════════ */
export function toolWatermark(ctx) {
  const total = ctx.pageCount() || 0;
  openModal({
    title: '文本水印',
    bodyHTML: `
      <div class="field">
        <label>水印文字</label>
        <input type="text" data-role="text" placeholder="如 内部资料" />
      </div>
      <div class="field-row">
        <div class="field">
          <label>字号</label>
          <input type="number" data-role="size" value="48" min="8" max="200" />
        </div>
        <div class="field">
          <label>透明度</label>
          <input type="number" data-role="opacity" value="15" min="5" max="100" />
        </div>
      </div>
      <div class="field">
        <label>布局</label>
        <select data-role="layout">
          <option value="tile">平铺（45° 斜排）</option>
          <option value="center">页面居中</option>
        </select>
      </div>
      <div class="field">
        <label>页码范围（留空=全部页${total ? `，共 ${total} 页` : ''}）</label>
        <input type="text" data-role="ranges" placeholder="如 2-" />
      </div>
      <p class="hint">水印以图片形式嵌入（支持中文），结果另存为新 PDF。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '添加水印并另存', primary: true,
        onClick: async (m) => {
          try {
            const text = m.query('[data-role="text"]').value.trim();
            if (!text) throw new Error('水印文字不能为空');
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const src = ld.doc;
            const srcPath = ld.base;
            const n = src.getPageCount();
            const rangesStr = m.query('[data-role="ranges"]').value.trim();
            const targets = rangesStr ? new Set(parseRanges(rangesStr, n)) : new Set(Array.from({ length: n }, (_, i) => i));

            const size = parseInt(m.query('[data-role="size"]').value, 10) || 48;
            const opacity = Math.min(1, Math.max(.05, (parseInt(m.query('[data-role="opacity"]').value, 10) || 15) / 100));
            const layout = m.query('[data-role="layout"]').value;

            // 画布渲染水印文字（支持中文），嵌入为 PNG（传 dataUrl 字符串）
            const tile = await renderTextTile(text, size, layout === 'tile');
            const img = await src.embedPng(tile.dataUrl);

            const pages = src.getPages();
            targets.forEach(i => {
              const page = pages[i];
              const W = page.getWidth(), H = page.getHeight();
              if (layout === 'center') {
                const dw = Math.min(W * .9, tile.width);
                const dh = dw * (tile.height / tile.width);
                page.drawImage(img, {
                  x: (W - dw) / 2,
                  y: (H - dh) / 2,
                  width: dw,
                  height: dh,
                  opacity
                });
              } else {
                const stepX = Math.max(220, W / 2.4);
                const stepY = Math.max(180, H / 3.2);
                const dw = stepX * .95, dh = dw * (tile.height / tile.width);
                for (let y = -dh; y < H + dh; y += stepY) {
                  for (let x = -dw; x < W + dw; x += stepX) {
                    page.drawImage(img, {
                      x, y, width: dw, height: dh, opacity,
                      rotate: window.PDFLib.degrees(45)
                    });
                  }
                }
              }
            });
            await saveDoc(src, `${nameWithoutExt(srcPath)}_水印.pdf`, '文本水印 - 另存为');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/**
 * 渲染文字为 PNG（4x 超采样 → 导入 PDF 后清晰）。
 * 返回 {dataUrl, width, height}（pt 单位）。
 */
export function renderTextTile(text, fontSize, rotated, color = '#5a5f68', bold = false, ss = 4) {
  const canvas = document.createElement('canvas');
  const probe = canvas.getContext('2d');
  const font = `${bold ? '700 ' : ''}${fontSize}px "Microsoft YaHei", sans-serif`;
  probe.font = font;
  const metrics = probe.measureText(text);
  const pad = Math.ceil(fontSize * 0.25);
  const wPt = Math.ceil(metrics.width) + pad * 2;
  const hPt = Math.ceil(fontSize * 1.4) + pad * 2;
  const size = rotated ? Math.max(wPt, hPt) : wPt;
  canvas.width = Math.ceil((rotated ? size : wPt) * ss);
  canvas.height = Math.ceil((rotated ? size : hPt) * ss);
  const c2 = canvas.getContext('2d');
  c2.scale(ss, ss);
  c2.clearRect(0, 0, canvas.width / ss, canvas.height / ss);
  c2.font = font;
  c2.fillStyle = color;
  c2.textBaseline = 'middle';
  if (rotated) {
    c2.translate(canvas.width / ss / 2, canvas.height / ss / 2);
    c2.rotate(-Math.PI / 4);
    c2.textAlign = 'center';
    c2.fillText(text, 0, 0);
  } else {
    c2.fillText(text, pad, canvas.height / ss / 2);
  }
  return {
    dataUrl: canvas.toDataURL('image/png'),
    width: canvas.width / ss,
    height: canvas.height / ss
  };
}

/* ══════════ 8. 添加页码 ══════════ */
export function toolPageNumbers(ctx) {
  const total = ctx.pageCount() || 0;
  openModal({
    title: '添加页码',
    bodyHTML: `
      <div class="field-row">
        <div class="field">
          <label>位置</label>
          <select data-role="pos">
            <option value="bc">底部居中</option>
            <option value="br">底部右侧</option>
            <option value="bl">底部左侧</option>
            <option value="tc">顶部居中</option>
          </select>
        </div>
        <div class="field">
          <label>起始编号</label>
          <input type="number" data-role="start" value="1" min="0" />
        </div>
      </div>
      <div class="field">
        <label>格式</label>
        <select data-role="format">
          <option value="n">1</option>
          <option value="dn">第 1 页</option>
          <option value="slash">1 / N</option>
        </select>
      </div>
      <p class="hint">页码以图片形式绘制（支持中文），结果另存为新 PDF。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '添加页码并另存', primary: true,
        onClick: async (m) => {
          try {
            console.error('[pagenum] onClick start');
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const src = ld.doc;
            const srcPath = ld.base;
            const pages = src.getPages();
            const start = parseInt(m.query('[data-role="start"]').value, 10) || 1;
            const pos = m.query('[data-role="pos"]').value;
            const format = m.query('[data-role="format"]').value;
            const N = pages.length;

            // 每种页码文本用 canvas 渲染（支持中文），缓存同文本的 PNG
            const imgCache = new Map();
            const getImg = async (text) => {
              if (!imgCache.has(text)) {
                const tile = await renderTextTile(text, 10, false);
                const img = await src.embedPng(tile.dataUrl);
                imgCache.set(text, { img, w: tile.width, h: tile.height });
              }
              return imgCache.get(text);
            };

            for (let i = 0; i < N; i++) {
              const page = pages[i];
              const num = start + i;
              let text;
              if (format === 'dn') text = `第 ${num} 页`;
              else if (format === 'slash') text = `${num} / ${N + start - 1}`;
              else text = String(num);
              const { img, w, h } = await getImg(text);
              const { width: W, height: H } = page.getSize();
              // 文本tile上下留白约 25%，按视觉宽度居中
              const dw = w, dh = h;
              let x = (W - dw) / 2, y = 16;
              if (pos === 'br') { x = W - dw - 24; }
              else if (pos === 'bl') { x = 24; }
              else if (pos === 'tc') { y = H - dh - 16; }
              page.drawImage(img, { x, y, width: dw, height: dh, opacity: .9 });
            }
            console.error('[pagenum] before save');
            await saveDoc(src, `${nameWithoutExt(srcPath)}_页码.pdf`, '添加页码 - 另存为');
            console.error('[pagenum] done');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/* ══════════ 9. 图片转 PDF / 新建 PDF（从图片） ══════════ */
/** @returns {{doc, bytes}} 或 null */
export async function buildPdfFromImages() {
  const paths = await window.onyx.openMultiDialog([
    { name: '图片', extensions: ['png', 'jpg', 'jpeg'] }
  ]);
  if (!paths) return null;
  const out = await window.PDFLib.PDFDocument.create();
  for (const p of paths) {
    const res = await window.onyx.readPdf(p);
    if (!res.ok) throw new Error(`读取失败：${p}`);
    const u8 = new Uint8Array(res.data);
    const ext = p.toLowerCase().endsWith('.png') ? 'png' : 'jpg';
    const img = ext === 'png' ? await out.embedPng(u8) : await out.embedJpg(u8);
    const page = out.addPage([img.width, img.height]);
    page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
  }
  return out;
}

export async function toolImagesToPdf() {
  try {
    const doc = await buildPdfFromImages();
    if (!doc) return;
    await saveDoc(doc, '图片转换.pdf', '图片转 PDF - 另存为');
  } catch (err) {
    toast(err.message, true);
  }
}

/* ══════════ 10. 清理优化 ══════════ */
export function toolOptimize(ctx) {
  openModal({
    title: 'PDF 清理优化',
    bodyHTML: `
      <div class="field">
        <label style="display:flex;align-items:center;gap:6px;cursor:pointer">
          <input type="checkbox" data-role="meta" checked /> 移除文档元数据（作者/软件信息）
        </label>
      </div>
      <p class="hint">重新构建文档结构：清理未引用对象、重建交叉引用表，可减小文件体积并修复部分损坏的 PDF。</p>
    `,
    buttons: [
      { label: '取消', onClick: (m) => m.close() },
      {
        label: '清理并另存', primary: true,
        onClick: async (m) => {
          try {
            const ld = await loadSourceDoc(ctx);
            if (!ld) return;
            const srcPath = ld.base;
            const src = ld.doc;
            if (m.query('[data-role="meta"]').checked) {
              src.setTitle('');
              src.setAuthor('');
              src.setSubject('');
              src.setKeywords([]);
              src.setProducer('OnyxPDF');
              src.setCreator('OnyxPDF');
            }
            await saveDoc(src, `${nameWithoutExt(srcPath)}_优化.pdf`, '清理优化 - 另存为');
            m.close();
          } catch (err) {
            toast(err.message, true);
          }
        }
      }
    ]
  });
}

/* ══════════ 11. 新建空白 PDF ══════════ */
export async function toolNewBlank(pageCount = 1) {
  const out = await window.PDFLib.PDFDocument.create();
  for (let i = 0; i < pageCount; i++) out.addPage([595.28, 841.89]);
  const bytes = await out.save();
  return bytes;
}

/* ══════════ 12. 插入图片页 ══════════ */
/** 将所选图片作为整页插入当前文档副本，先确认位置、方向、数量和页面尺寸。 */
export function toolInsertImagePages(ctx, afterPage = null) {
  const total = ctx.pageCount?.() || 0;
  const defaultPage = Math.min(Math.max(1, afterPage || total || 1), total || 1);
  return new Promise(resolve => {
    openModal({
      title: '插入图片页',
      bodyHTML: `
        <div class="field-row">
          <div class="field">
            <label>插入位置（第几页）</label>
            <input type="number" data-role="pos" value="${defaultPage}" min="1" />
          </div>
          <div class="field">
            <label>方向</label>
            <select data-role="where">
              <option value="before">在该页之前</option>
              <option value="after" selected>在该页之后</option>
            </select>
          </div>
        </div>
        <div class="field-row">
          <div class="field">
            <label>页面尺寸</label>
            <select data-role="size">
              <option value="image" selected>与图片相同</option>
              <option value="same">与相邻页相同</option>
              <option value="a4p">A4 纵向</option>
              <option value="a4l">A4 横向</option>
            </select>
          </div>
          <div class="field">
            <label>插入页数</label>
            <input type="number" data-role="count" value="1" min="1" max="100" />
          </div>
        </div>
        <p class="hint">按所选图片顺序循环生成；选择“与图片相同”时每页保留原始宽高。</p>
      `,
      buttons: [
        { label: '取消', onClick: (modal) => { modal.close(); resolve(null); } },
        {
          label: '选择图片并插入', primary: true,
          onClick: async (modal) => {
            try {
              const paths = await window.onyx.openMultiDialog([
                { name: '图片', extensions: ['png', 'jpg', 'jpeg'] }
              ]);
              if (!paths?.length) return;
              const source = await loadSourceDoc(ctx);
              if (!source) { modal.close(); resolve(null); return; }
              const src = source.doc;
              const pageCount = src.getPageCount();
              let position = parseInt(modal.query('[data-role="pos"]').value, 10);
              position = Math.min(Math.max(1, position || 1), pageCount);
              const after = modal.query('[data-role="where"]').value === 'after';
              const sizeMode = modal.query('[data-role="size"]').value;
              const count = Math.min(100, Math.max(1, parseInt(modal.query('[data-role="count"]').value, 10) || 1));
              const refPage = src.getPage(position - 1);
              const adjacent = { width: refPage.getWidth(), height: refPage.getHeight() };
              const fixedSize = sizeMode === 'a4p' ? { width: 595.28, height: 841.89 }
                : sizeMode === 'a4l' ? { width: 841.89, height: 595.28 }
                : sizeMode === 'same' ? adjacent
                : null;
              const insertAt = after ? position : position - 1;

              for (let index = 0; index < count; index++) {
                const imagePath = paths[index % paths.length];
                const result = await window.onyx.readPdf(imagePath);
                if (!result.ok) throw new Error(`读取失败：${imagePath}`);
                const bytes = new Uint8Array(result.data);
                const isPng = imagePath.toLowerCase().endsWith('.png');
                const embedded = isPng ? await src.embedPng(bytes) : await src.embedJpg(bytes);
                const size = fixedSize || { width: embedded.width, height: embedded.height };
                const page = src.insertPage(insertAt + index, [size.width, size.height]);
                const scale = Math.min(size.width / embedded.width, size.height / embedded.height, 1);
                const width = embedded.width * scale;
                const height = embedded.height * scale;
                page.drawImage(embedded, {
                  x: (size.width - width) / 2,
                  y: (size.height - height) / 2,
                  width,
                  height
                });
              }
              resolve({ bytes: await src.save(), suffix: '_图片页.pdf' });
              modal.close();
            } catch (error) {
              toast(error.message, true);
            }
          }
        }
      ],
      onClose: () => resolve(null)
    });
  });
}

/* ══════════ 13. 页面顺序调整 ══════════ */
/** 按新顺序（0-based 页索引数组）重排文档，返回新 PDF 字节 */
export async function reorderPagesBytes(originalBytes, order) {
  const src = await window.PDFLib.PDFDocument.load(originalBytes);
  const out = await window.PDFLib.PDFDocument.create();
  const pages = await out.copyPages(src, order);
  pages.forEach(pg => out.addPage(pg));
  return out.save();
}

/** 多行文字渲染（textedit 显示与导出共用）：返回 {dataUrl,width,height} */
export function renderTextEditTile(a) {
  const SS = 3; // V12.1：3x 超采样，缩放显示清晰
  const lines = String(a.text).split('\n');
  const canvas = document.createElement('canvas');
  const probe = canvas.getContext('2d');
  const font = (a.bold ? '700 ' : '') + a.size + 'px ' + (a.font || '"Microsoft YaHei", sans-serif');
  probe.font = font;
  // V12.3.3：逐字符按 PDF 真实宽度（charWidthMap, pt）排布——与原文字宽度一致
  const measure = (ch) => (a.charWidthMap && a.charWidthMap[ch]) || probe.measureText(ch).width;
  let w = 0;
  const advances = [];
  for (const l of lines) {
    let lw = 0;
    const adv = [];
    for (const ch of l) { const c = measure(ch); adv.push(c); lw += c; }
    advances.push(adv);
    w = Math.max(w, lw);
  }
  const lh = a.size * 1.4;
  const pad = Math.ceil(a.size * 0.15);
  const targetW = Math.max(Math.ceil(Number(a.w) || 0), Math.ceil(w));
  const wPt = targetW + pad * 2, hPt = Math.ceil(lh * lines.length) + pad * 2;
  canvas.width = wPt * SS;
  canvas.height = hPt * SS;
  const c2 = canvas.getContext('2d');
  c2.scale(SS, SS);
  c2.font = font;
  c2.fillStyle = a.color || '#1f2329';
  c2.textBaseline = 'middle';
  lines.forEach((l, i) => {
    let x = pad;
    const adv = advances[i] || [];
    const lineWidth = adv.reduce((sum, advWidth) => sum + advWidth, 0);
    const alignWidth = Math.max(targetW, w);
    if (a.align === 'center') x += (alignWidth - lineWidth) / 2;
    else if (a.align === 'right') x += alignWidth - lineWidth;
    [...l].forEach((ch, j) => { c2.fillText(ch, x, pad + lh * (i + 0.5)); x += adv[j] || 0; });
  });
  return { dataUrl: canvas.toDataURL('image/png'), width: wPt, height: hPt };
}
