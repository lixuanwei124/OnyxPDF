'use strict';
/**
 * OnyxPDF - 预加载脚本（contextBridge 安全桥）
 * 仅暴露白名单 API，渲染进程不持有 Node 能力。
 */
const { contextBridge, ipcRenderer, webUtils } = require('electron');

const listeners = new Map();

contextBridge.exposeInMainWorld('onyx', {
  platform: () => ipcRenderer.invoke('onyx:platform'),
  version: () => ipcRenderer.invoke('onyx:version'),
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file); } catch (_) { return file.path || ''; }
  },
  openDialog: () => ipcRenderer.invoke('onyx:open-dialog'),
  openMultiDialog: (filters) => ipcRenderer.invoke('onyx:open-multi-dialog', filters),
  openDirDialog: () => ipcRenderer.invoke('onyx:open-dir-dialog'),
  writePath: (payload) => ipcRenderer.invoke('onyx:write-path', payload),
  appVersion: () => ipcRenderer.invoke('onyx:app-version'),
  testExportList: () => ipcRenderer.invoke('onyx:test-export-list'),
  readPdf: (p) => ipcRenderer.invoke('onyx:read-pdf', p),
  decryptPdf: (bytes) => ipcRenderer.invoke('onyx:decrypt-pdf', bytes),
  readFontFile: (path) => ipcRenderer.invoke('onyx:read-font-file', path),
  recentsGet: () => ipcRenderer.invoke('onyx:recents-get'),
  recentsAdd: (entry) => ipcRenderer.invoke('onyx:recents-add', entry),
  recentsClear: () => ipcRenderer.invoke('onyx:recents-clear'),
  exportPdf: (payload) => ipcRenderer.invoke('onyx:export-pdf', payload),
  print: () => ipcRenderer.invoke('onyx:print'),
  winMinimize: () => ipcRenderer.invoke('onyx:win-minimize'),
  winMaximize: () => ipcRenderer.invoke('onyx:win-maximize'),
  winClose: () => ipcRenderer.invoke('onyx:win-close'),
  isMaximized: () => ipcRenderer.invoke('onyx:is-maximized'),
  onOpenPath: (cb) => {
    const fn = (_e, p) => cb(p);
    listeners.set(cb, fn);
    ipcRenderer.on('onyx:open-path', fn);
  },
  onWinState: (cb) => {
    const fn = (_e, v) => cb(v);
    listeners.set(cb, fn);
    ipcRenderer.on('onyx:win-state', fn);
  }
});
