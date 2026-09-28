/**
 * OnyxPDF - 通用模态对话框
 */
export function openModal({ title, bodyHTML, buttons = [], width = 380, onClose = null }) {
  const mask = document.getElementById('modal-mask');
  mask.innerHTML = `
    <div class="modal" style="width:${width}px">
      <div class="modal-head">
        <span>${title}</span>
        <button class="modal-close" data-role="close">✕</button>
      </div>
      <div class="modal-body" data-role="body"></div>
      <div class="modal-foot" data-role="foot"></div>
    </div>`;
  const modalEl = mask.querySelector('.modal');
  const bodyEl = mask.querySelector('[data-role="body"]');
  const footEl = mask.querySelector('[data-role="foot"]');
  bodyEl.innerHTML = bodyHTML;

  const api = {
    el: modalEl,
    close() {
      mask.classList.add('hidden');
      mask.innerHTML = '';
      if (onClose) onClose();
    },
    query: (sel) => bodyEl.querySelector(sel)
  };

  mask.querySelector('[data-role="close"]').addEventListener('click', () => api.close());
  mask.addEventListener('mousedown', (e) => {
    if (e.target === mask) api.close();
  });

  for (const b of buttons) {
    const btn = document.createElement('button');
    btn.className = 'btn' + (b.primary ? ' primary' : '');
    btn.textContent = b.label;
    btn.addEventListener('click', () => b.onClick(api));
    footEl.appendChild(btn);
  }

  mask.classList.remove('hidden');
  return api;
}
