/**
 * OnyxPDF - 菜单栏（文件 / 工具 / 帮助）
 * 子菜单为 flyout 模式：父项 hover/click 时在右侧弹出，样式与一级菜单一致。
 */
import { escapeHtml } from './utils.js';

export const I = {
  open: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M3 4.5A1.5 1.5 0 0 1 4.5 3h3.6l2 2h5.4A1.5 1.5 0 0 1 17 6.5v8A1.5 1.5 0 0 1 15.5 16h-11A1.5 1.5 0 0 1 3 14.5z" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  newdoc: '<svg viewBox="0 0 20 20" class="me-icon"><rect x="4" y="2.5" width="12" height="15" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/><line x1="10" y1="7" x2="10" y2="13" stroke="currentColor" stroke-width="1.4"/><line x1="7" y1="10" x2="13" y2="10" stroke="currentColor" stroke-width="1.4"/></svg>',
  clock: '<svg viewBox="0 0 20 20" class="me-icon"><circle cx="10" cy="10" r="7" fill="none" stroke="currentColor" stroke-width="1.4"/><path d="M10 6v4.2l3 1.8" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  save: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M4 3h9.5L17 6.5V16a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z" fill="none" stroke="currentColor" stroke-width="1.4"/><rect x="6" y="11" width="8" height="6" fill="none" stroke="currentColor" stroke-width="1.4"/><rect x="6.5" y="3.5" width="6" height="4" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>',
  exportd: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M10 3v9M6.5 8.5 10 12l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4"/><path d="M4 14v2.2a.8.8 0 0 0 .8.8h10.4a.8.8 0 0 0 .8-.8V14" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  print: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M6 7V3h8v4M6 13H3.8A.8.8 0 0 1 3 12.2V8.8A1.8 1.8 0 0 1 4.8 7h10.4A1.8 1.8 0 0 1 17 8.8v3.4a.8.8 0 0 1-.8.8H14" fill="none" stroke="currentColor" stroke-width="1.4"/><rect x="6" y="11.5" width="8" height="5.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  info: '<svg viewBox="0 0 20 20" class="me-icon"><circle cx="10" cy="10" r="7.2" fill="none" stroke="currentColor" stroke-width="1.4"/><circle cx="10" cy="6.6" r=".9" fill="currentColor"/><line x1="10" y1="9" x2="10" y2="14" stroke="currentColor" stroke-width="1.4"/></svg>',
  quit: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M12 3H5.5A1.5 1.5 0 0 0 4 4.5v11A1.5 1.5 0 0 0 5.5 17H12" fill="none" stroke="currentColor" stroke-width="1.4"/><path d="M9 10h8M14 7l3 3-3 3" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  merge: '<svg viewBox="0 0 20 20" class="me-icon"><rect x="2.5" y="4" width="7" height="9" fill="none" stroke="currentColor" stroke-width="1.3"/><rect x="10.5" y="7" width="7" height="9" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>',
  split: '<svg viewBox="0 0 20 20" class="me-icon"><rect x="2.5" y="4.5" width="6.5" height="11" fill="none" stroke="currentColor" stroke-width="1.3"/><rect x="11" y="4.5" width="6.5" height="11" fill="none" stroke="currentColor" stroke-width="1.3" stroke-dasharray="2.5 1.8"/></svg>',
  extract: '<svg viewBox="0 0 20 20" class="me-icon"><rect x="3" y="3" width="10" height="13" fill="none" stroke="currentColor" stroke-width="1.3"/><rect x="9" y="9" width="8" height="8" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>',
  trash: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M4 6h12M8 6V4.5a.5.5 0 0 1 .5-.5h3a.5.5 0 0 1 .5.5V6M6 6l.8 9.2a1 1 0 0 0 1 .8h4.4a1 1 0 0 0 1-.8L14 6" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>',
  rotate: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M5 9a6 6 0 1 1 1.8 4.3" fill="none" stroke="currentColor" stroke-width="1.4"/><path d="M5 13.6V9h4.6" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  blank: '<svg viewBox="0 0 20 20" class="me-icon"><rect x="4" y="2.5" width="12" height="15" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  water: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M10 3.5c3 3.6 5 6.1 5 8.5a5 5 0 0 1-10 0c0-2.4 2-4.9 5-8.5z" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  hash: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M7.5 3 6 17M14 3l-1.5 14M3.5 7.5h14M2.5 12.5h14" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>',
  image: '<svg viewBox="0 0 20 20" class="me-icon"><rect x="3" y="3.5" width="14" height="13" fill="none" stroke="currentColor" stroke-width="1.4"/><circle cx="7.2" cy="8" r="1.5" fill="currentColor"/><path d="M3.5 14.5 8 10l3 3 2.5-2.5 3 3" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  broom: '<svg viewBox="0 0 20 20" class="me-icon"><path d="M13 3l4 4-8.5 8.5H4.5V11.5z" fill="none" stroke="currentColor" stroke-width="1.4"/><path d="M4 17.5h12" stroke="currentColor" stroke-width="1.4"/></svg>'
};

export function initMenus(ctx) {
  const layer = document.getElementById('menu-layer');
  let openMenu = null;          // 当前打开的顶级菜单名
  let subOpen = null;           // 当前展开子菜单的一级/二级条目元素

  function closeMenus() {
    layer.innerHTML = '';
    document.querySelectorAll('.menu-item.open').forEach(b => b.classList.remove('open'));
    openMenu = null;
    subOpen = null;
  }

  function entry(e) {
    if (e.sep) return '<div class="menu-sep"></div>';
    const dis = e.enabled && !e.enabled() ? 'disabled' : '';
    const key = e.key ? `<span class="me-key">${e.key}</span>` : '';
    const arrow = e.children ? '<span class="me-arrow">▸</span>' : '';
    return `<button class="menu-entry ${e.danger ? 'danger' : ''}" ${dis} data-id="${e.id || ''}" ${e.children ? 'data-has-children="1"' : ''}>
      ${e.icon || ''}<span class="me-label">${escapeHtml(e.label)}</span>${key}${arrow}</button>`;
  }

  /** 在条目右侧弹出 flyout 子菜单；不关闭一级菜单 */
  function openFlyout(entryBtn, items) {
    // 关闭已存在的子菜单（保留一级）
    layer.querySelectorAll('.menu-dropdown.sub').forEach(d => d.remove());
    if (subOpen) { subOpen.classList.remove('open'); subOpen = null; }

    const dd = document.createElement('div');
    dd.className = 'menu-dropdown sub';
    dd.innerHTML = items.map(entry).join('');

    const r = entryBtn.getBoundingClientRect();
    const W = 240;
    let left = r.right + 4;
    if (left + W > window.innerWidth - 4) left = Math.max(4, r.left - W - 4); // 右侧放不下则左侧
    const top = Math.max(4, Math.min(r.top - 6, window.innerHeight - 300));
    dd.style.left = `${left}px`;
    dd.style.top = `${top}px`;

    entryBtn.classList.add('open');
    subOpen = entryBtn;
    layer.appendChild(dd);

    bindEntries(dd, items);
  }

  /** 为一个 dropdown 容器的条目绑定点击行为 */
  function bindEntries(container, items) {
    container.querySelectorAll('.menu-entry:not(.disabled)').forEach(btn => {
      const item = items.find(x => !x.sep && (x.id || '') === btn.dataset.id);
      if (!item) return;
      if (item.children) {
        // hover 或 click 均展开子菜单
        btn.addEventListener('mouseenter', () => openFlyout(btn, item.children));
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          openFlyout(btn, item.children);
        });
      } else {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          closeMenus();
          if (item.action) item.action();
        });
      }
    });
  }

  function showMenu(name, anchorBtn) {
    closeMenus();
    const items = ctx.menuDefs[name]();
    const dd = document.createElement('div');
    dd.className = 'menu-dropdown';
    dd.innerHTML = items.map(entry).join('');
    const r = anchorBtn.getBoundingClientRect();
    dd.style.left = `${Math.min(r.left, window.innerWidth - 236)}px`;
    layer.appendChild(dd);
    anchorBtn.classList.add('open');
    openMenu = name;

    bindEntries(dd, items);
  }

  document.querySelectorAll('.menu-item').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (openMenu === btn.dataset.menu) closeMenus();
      else showMenu(btn.dataset.menu, btn);
    });
    btn.addEventListener('mouseenter', () => {
      if (openMenu && openMenu !== btn.dataset.menu) showMenu(btn.dataset.menu, btn);
    });
  });

  document.addEventListener('mousedown', (e) => {
    if (!e.target.closest('.menu-dropdown') && !e.target.closest('.menu-item')) closeMenus();
  });
}
