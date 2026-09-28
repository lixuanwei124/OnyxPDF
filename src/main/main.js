'use strict';
/**
 * OnyxPDF - 主进程
 * 职责：窗口管理、文件对话框、最近文件、打印、批注 PDF 导出、IPC 安全桥
 */
const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow = null;
const isMac = process.platform === 'darwin';

// PDF 阅读不依赖 GPU 合成；禁用硬件加速以提升兼容性（虚拟机/远程桌面环境尤其必要）
app.disableHardwareAcceleration();
if (process.env.ONYX_TEST) {
  // 虚拟机/无 GPU 环境下 Chromium GPU 进程可能反复崩溃并终止应用；
  // in-process-gpu 将 GPU 管线并入浏览器进程，彻底规避独立 GPU 进程崩溃
  app.commandLine.appendSwitch('disable-gpu');
  app.commandLine.appendSwitch('in-process-gpu');
}

/* 本地字体访问（编辑内容：系统全部字体） */
app.whenReady().then(() => {
  try {
    const { session } = require('electron');
    session.defaultSession.setPermissionRequestHandler((_wc, permission, cb) => cb(true));
    session.defaultSession.setPermissionCheckHandler(() => true);
  } catch (_) { /* ignore */ }
});

/* ---------------- 单实例锁 ---------------- */
if (process.env.ONYX_TEST) {
  process.on('uncaughtException', (e) => console.error('[main] uncaught:', e));
  process.on('unhandledRejection', (e) => console.error('[main] unhandled:', e));
}
// V12.3.3：测试可注入独立 userData（隔离 localStorage，不与正在运行的实例竞争）
if (process.env.ONYX_USER_DATA) app.setPath('userData', process.env.ONYX_USER_DATA);
const gotLock = app.requestSingleInstanceLock();
if (process.env.ONYX_TEST) console.log('[main] lock=', gotLock, 'pid=', process.pid);
if (gotLock) {
  app.on('second-instance', (_e, argv) => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
      const pdfArg = argv.find(a => /\.pdf$/i.test(a));
      if (pdfArg) mainWindow.webContents.send('onyx:open-path', pdfArg);
    }
  });
} else if (process.env.ONYX_TEST) {
  // 测试环境（受管沙箱/包装器可能误判锁）：降级多实例，供自动化探针并行
  console.warn('[main] single-instance lock unavailable (ONYX_TEST), multi-instance fallback');
} else {
  // V12.3.3：正式环境拿不到锁 = 已有实例在跑，转交后退出，杜绝多实例堆积
  app.quit();
}

/* ---------------- 窗口 ---------------- */
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 960,
    minHeight: 600,
    frame: false,
    backgroundColor: '#eef0f3',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

  mainWindow.once('ready-to-show', () => mainWindow.show());

  // 外部链接走系统浏览器，禁止在应用内打开
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.on('maximize', () => mainWindow.webContents.send('onyx:win-state', true));
  mainWindow.on('unmaximize', () => mainWindow.webContents.send('onyx:win-state', false));
}

/* ---------------- 工具 ---------------- */
function recentsFile() {
  return path.join(app.getPath('userData'), 'recents.json');
}
function readRecents() {
  try {
    return JSON.parse(fs.readFileSync(recentsFile(), 'utf8'));
  } catch (_) {
    return [];
  }
}
function writeRecents(list) {
  try {
    fs.writeFileSync(recentsFile(), JSON.stringify(list.slice(0, 10)), 'utf8');
  } catch (_) { /* 忽略写入失败 */ }
}

/* ---------------- IPC ---------------- */
ipcMain.handle('onyx:open-dialog', async () => {
  const r = await dialog.showOpenDialog(mainWindow, {
    title: '打开 PDF 文件',
    properties: ['openFile'],
    filters: [{ name: 'PDF 文档', extensions: ['pdf'] }]
  });
  if (r.canceled || !r.filePaths.length) return null;
  return r.filePaths[0];
});

ipcMain.handle('onyx:open-multi-dialog', async (_e, filters) => {
  const r = await dialog.showOpenDialog(mainWindow, {
    title: '选择文件',
    properties: ['openFile', 'multiSelections'],
    filters: filters || [{ name: 'PDF 文档', extensions: ['pdf'] }]
  });
  if (r.canceled || !r.filePaths.length) return null;
  return r.filePaths;
});

ipcMain.handle('onyx:open-dir-dialog', async () => {
  const r = await dialog.showOpenDialog(mainWindow, {
    title: '选择输出文件夹',
    properties: ['openDirectory', 'createDirectory']
  });
  if (r.canceled || !r.filePaths.length) return null;
  return r.filePaths[0];
});

ipcMain.handle('onyx:write-path', async (_e, payload) => {
  try {
    fs.writeFileSync(payload.path, Buffer.from(payload.bytes));
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
});

ipcMain.handle('onyx:app-version', () => app.getVersion());

ipcMain.handle('onyx:test-export-list', () => {
  try {
    const dir = path.join(app.getPath('userData'), 'test-export');
    return fs.existsSync(dir) ? fs.readdirSync(dir) : [];
  } catch (_) { return []; }
});

// ONYX_TEST：自动落盘代替保存对话框（测试断言用）
ipcMain.handle('onyx:test-save', (_e, payload) => {
  try {
    const dir = path.join(app.getPath('userData'), 'test-export');
    fs.mkdirSync(dir, { recursive: true });
    const fp = path.join(dir, payload.suggestedName || ('export-' + Date.now()));
    fs.writeFileSync(fp, Buffer.from(payload.bytes));
    return { ok: true, path: fp };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
});

ipcMain.handle('onyx:read-pdf', async (_e, filePath) => {
  try {
    const buf = fs.readFileSync(filePath);
    return { ok: true, data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
});

ipcMain.handle('onyx:read-font-file', async (_e, filePath) => {
  try {
    const normalized = String(filePath || '').toLowerCase().replace(/\//g, '\\');
    if (!normalized.startsWith('c:\\windows\\fonts\\')) throw new Error('字体路径不在白名单内');
    if (!/\.(ttf|otf|ttc)$/.test(normalized)) throw new Error('不支持的字体格式');
    const buffer = await fs.promises.readFile(normalized);
    return { ok: true, bytes: buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
});

ipcMain.handle('onyx:recents-get', () => readRecents());

ipcMain.handle('onyx:recents-add', (_e, entry) => {
  const list = readRecents().filter(x => x.path !== entry.path);
  list.unshift(entry);
  writeRecents(list);
  return readRecents();
});

ipcMain.handle('onyx:recents-clear', () => { writeRecents([]); return []; });

ipcMain.handle('onyx:export-pdf', async (_e, payload) => {
  // ONYX_TEST：自动落盘代替保存对话框（测试断言用）
  if (process.env.ONYX_TEST) {
    try {
      const dir = path.join(app.getPath('userData'), 'test-export');
      fs.mkdirSync(dir, { recursive: true });
      const fp = path.join(dir, payload.suggestedName || ('export-' + Date.now()));
      fs.writeFileSync(fp, Buffer.from(payload.bytes));
      return { ok: true, path: fp };
    } catch (err) {
      return { ok: false, error: String(err && err.message || err) };
    }
  }
  const r = await dialog.showSaveDialog(mainWindow, {
    title: payload.title || '另存为',
    defaultPath: payload.suggestedName || 'annotated.pdf',
    filters: payload.filters || [{ name: 'PDF 文档', extensions: ['pdf'] }]
  });
  if (r.canceled || !r.filePath) return { ok: false, canceled: true };
  try {
    fs.writeFileSync(r.filePath, Buffer.from(payload.bytes));
    shell.showItemInFolder(r.filePath);
    return { ok: true, path: r.filePath };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
});

ipcMain.handle('onyx:print', () => {
  return new Promise((resolve) => {
    if (!mainWindow) { resolve({ success: false, reason: 'no-window' }); return; }
    mainWindow.webContents.print({ silent: false, printBackground: true }, (success, reason) => {
      resolve({ success, reason });
    });
  });
});

ipcMain.handle('onyx:win-minimize', () => mainWindow && mainWindow.minimize());
ipcMain.handle('onyx:win-maximize', () => {
  if (!mainWindow) return;
  mainWindow.isMaximized() ? mainWindow.unmaximize() : mainWindow.maximize();
});
ipcMain.handle('onyx:win-close', () => mainWindow && mainWindow.close());
ipcMain.handle('onyx:is-maximized', () => mainWindow ? mainWindow.isMaximized() : false);
ipcMain.handle('onyx:platform', () => process.platform);
ipcMain.handle('onyx:version', () => app.getVersion());

/* ---------------- 生命周期 ---------------- */
app.whenReady().then(() => {
  if (process.env.ONYX_TEST) console.log('[main] ready');
  Menu.setApplicationMenu(null);
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (!isMac) app.quit();
});

