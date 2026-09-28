/**
 * OnyxPDF - 通用工具
 */
export const clamp = (v, min, max) => Math.min(max, Math.max(min, v));

export function debounce(fn, ms) {
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

export function fileKey(name, size) {
  return `${name}::${size}`;
}

let toastTimer = null;
export function toast(msg, isErr = false) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.toggle('err', isErr);
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add('hidden'), 2400);
}

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[c]);
}

export function formatBytes(n) {
  if (!Number.isFinite(n)) return '';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

export function timeAgo(ts) {
  const d = Date.now() - ts;
  if (d < 60e3) return '刚刚';
  if (d < 3600e3) return `${Math.floor(d / 60e3)} 分钟前`;
  if (d < 86400e3) return `${Math.floor(d / 3600e3)} 小时前`;
  return new Date(ts).toLocaleDateString('zh-CN');
}

/** 批注颜色映射（CSS 类后缀 / RGB 0-1，供 pdf-lib 导出使用） */
export const ANNOT_COLORS = {
  default: { cls: '', rgb: [1, 0.882, 0.302] },   // 黄
  green:   { cls: 'c-green', rgb: [0.49, 1, 0.62] },
  blue:    { cls: 'c-blue',  rgb: [0.494, 0.784, 1] },
  pink:    { cls: 'c-pink',  rgb: [1, 0.478, 0.82] }
};
