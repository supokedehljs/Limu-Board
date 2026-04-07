const { app, BrowserWindow, ipcMain, shell, clipboard, nativeImage, dialog, Menu, MenuItem, globalShortcut, Tray, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const fsSync = require('fs');

let mainWindow;
let tray = null;
let dataPath;
let activeLibraryPath = null;
let globalHotkey = 'Alt+W';
let hotkeyRegistered = false;

function getDataPath() {
  if (!dataPath) {
    dataPath = path.join(app.getAppPath(), '.whiteboard-data');
  }
  return dataPath;
}

function getDefaultLibraryPath() {
  return path.join(getDataPath(), 'library');
}

function getLibrariesConfigPath() {
  const configPath = path.join(getDefaultLibraryPath(), 'libraries.json');
  return configPath;
}

async function getLibrariesConfig() {
  const configPath = getLibrariesConfigPath();
  try {
    const data = await fs.readFile(configPath, 'utf-8');
    return JSON.parse(data);
  } catch {
    return { libraries: [], activeId: null };
  }
}

async function saveLibrariesConfig(config) {
  const configPath = getLibrariesConfigPath();
  await fs.mkdir(path.dirname(configPath), { recursive: true });
  await fs.writeFile(configPath, JSON.stringify(config, null, 2));
}

function getSettingsPath() {
  return path.join(getDataPath(), 'settings.json');
}

async function loadSettings() {
  const settingsPath = getSettingsPath();
  try {
    const data = await fs.readFile(settingsPath, 'utf-8');
    return JSON.parse(data);
  } catch {
    return { globalHotkey: 'Alt+W', localHotkeys: {} };
  }
}

async function saveSettings(settings) {
  const settingsPath = getSettingsPath();
  await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2));
}

function unregisterGlobalHotkey() {
  if (hotkeyRegistered) {
    globalShortcut.unregister(globalHotkey);
    hotkeyRegistered = false;
  }
}

function registerGlobalHotkey(accelerator) {
  unregisterGlobalHotkey();
  if (!accelerator) return;
  globalHotkey = accelerator;
  const success = globalShortcut.register(accelerator, () => {
    if (mainWindow) {
      if (mainWindow.isVisible()) {
        mainWindow.hide();
      } else {
        mainWindow.show();
        mainWindow.focus();
      }
    }
  });
  if (success) {
    hotkeyRegistered = true;
  }
}

async function getActiveLibraryPath() {
  const config = await getLibrariesConfig();
  if (config.activeId && config.libraries.length > 0) {
    const lib = config.libraries.find(l => l.id === config.activeId);
    if (lib) {
      activeLibraryPath = lib.path;
      return activeLibraryPath;
    }
  }
  const defaultPath = getDefaultLibraryPath();
  try {
    await fs.mkdir(defaultPath, { recursive: true });
  } catch {}
  activeLibraryPath = defaultPath;
  return activeLibraryPath;
}

async function getStatePath() {
  return path.join(await getActiveLibraryPath(), 'state.json');
}

async function getAssetsPath() {
  return path.join(await getActiveLibraryPath(), 'assets');
}

async function ensureDataDir() {
  const dir = getDataPath();
  await fs.mkdir(dir, { recursive: true });
  const filesDir = path.join(dir, 'files');
  await fs.mkdir(filesDir, { recursive: true });
  const libPath = await getActiveLibraryPath();
  const assetsDir = path.join(libPath, 'assets');
  await fs.mkdir(assetsDir, { recursive: true });
  return dir;
}

function generateAssetId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 8);
}

function resizeImageToMax(buffer, maxDim) {
  try {
    const img = nativeImage.createFromBuffer(Buffer.from(buffer));
    if (img.isEmpty()) return null;
    const size = img.getSize();
    if (size.width <= maxDim && size.height <= maxDim) return img;
    const scale = maxDim / Math.max(size.width, size.height);
    const newWidth = Math.round(size.width * scale);
    const newHeight = Math.round(size.height * scale);
    return img.resize({ width: newWidth, height: newHeight });
  } catch {
    return null;
  }
}

function createTray() {
  const icon = nativeImage.createEmpty();
  const iconSize = 16;
  const canvas = Buffer.alloc(iconSize * iconSize * 4);
  for (let i = 0; i < iconSize * iconSize; i++) {
    canvas[i * 4] = 74;
    canvas[i * 4 + 1] = 144;
    canvas[i * 4 + 2] = 217;
    canvas[i * 4 + 3] = 255;
  }
  const trayIcon = nativeImage.createFromBuffer(canvas, { width: iconSize, height: iconSize });
  tray = new Tray(trayIcon);
  tray.setToolTip('白板 - 双击打开');
  tray.on('double-click', () => {
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();
    }
  });
  const contextMenu = Menu.buildFromTemplate([
    { label: '显示白板', click: () => { if (mainWindow) { mainWindow.show(); mainWindow.focus(); } } },
    { label: '退出', click: () => { app.quit(); } }
  ]);
  tray.setContextMenu(contextMenu);
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    frame: false,
    backgroundColor: '#1e1e1e',
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.on('close', (e) => {
    e.preventDefault();
    mainWindow.hide();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(async () => {
  createTray();
  createWindow();
  const settings = await loadSettings();
  if (settings.globalHotkey) {
    registerGlobalHotkey(settings.globalHotkey);
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('will-quit', () => {
  globalShortcut.unregisterAll();
  if (tray) {
    tray.destroy();
    tray = null;
  }
});

ipcMain.handle('window-minimize', () => {
  mainWindow.minimize();
});

ipcMain.handle('window-maximize', () => {
  if (mainWindow.isMaximized()) {
    mainWindow.unmaximize();
  } else {
    mainWindow.maximize();
  }
});

ipcMain.handle('window-close', () => {
  mainWindow.close();
});

ipcMain.handle('window-is-maximized', () => {
  return mainWindow.isMaximized();
});

ipcMain.handle('window-unmaximize', () => {
  mainWindow.unmaximize();
});

ipcMain.handle('window-show', () => {
  if (!mainWindow.isVisible()) {
    mainWindow.show();
  }
  mainWindow.focus();
});

ipcMain.handle('window-hide', () => {
  mainWindow.hide();
});

ipcMain.handle('get-settings', async () => {
  return loadSettings();
});

ipcMain.handle('save-settings', async (event, settings) => {
  await saveSettings(settings);
  if (settings.globalHotkey) {
    registerGlobalHotkey(settings.globalHotkey);
  } else {
    unregisterGlobalHotkey();
  }
  return true;
});

ipcMain.handle('get-data-path', () => {
  return getDataPath();
});

ipcMain.handle('ensure-data-dir', async () => {
  return ensureDataDir();
});

ipcMain.handle('add-asset', async (event, { buffer, filename }) => {
  try {
    await ensureDataDir();
    const assetId = generateAssetId();
    const activeLibraryPath = await getActiveLibraryPath();
    const assetsPath = path.join(activeLibraryPath, 'assets');
    const assetDir = path.join(assetsPath, assetId);
    await fs.mkdir(assetDir, { recursive: true });

    const ext = path.extname(filename).toLowerCase();
    const savedName = `original${ext}`;
    const filePath = path.join(assetDir, savedName);
    await fs.writeFile(filePath, Buffer.from(buffer));

    let width = 0, height = 0;
    const imageExts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'];
    const isImage = imageExts.includes(ext);
    if (isImage) {
      try {
        const img = nativeImage.createFromBuffer(Buffer.from(buffer));
        const size = img.getSize();
        width = size.width;
        height = size.height;
      } catch (e) {}
    }

    const metadata = {
      id: assetId,
      assetId: assetId,
      originalName: filename,
      cardName: '',
      type: isImage ? 'image' : 'file',
      size: Buffer.from(buffer).length,
      width,
      height,
      x: 0,
      y: 0,
      timerTotal: 0,
      timerRunning: false,
      timerStartedAt: null,
      cardTags: [],
      cardAnnotation: '',
      cardUrl: '',
      addedTime: Date.now()
    };
    await fs.writeFile(path.join(assetDir, 'metadata.json'), JSON.stringify(metadata, null, 2));

    return { assetId, savedName, originalName: filename, type: metadata.type, width, height };
  } catch (err) {
    console.error('add-asset error:', err);
    throw err;
  }
});

ipcMain.handle('get-asset', async (event, assetId) => {
  try {
    const assetsPath = await getAssetsPath();
    const assetDir = path.join(assetsPath, assetId);
    try {
      await fs.access(assetDir);
    } catch {
      return null;
    }

    const metadataPath = path.join(assetDir, 'metadata.json');
    let metadata = null;
    try {
      const data = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(data);
    } catch {}

    const entries = await fs.readdir(assetDir);
    const originalFile = entries.find(f => f.startsWith('original'));
    if (!originalFile) return null;

    const buffer = await fs.readFile(path.join(assetDir, originalFile));
    const ext = path.extname(originalFile).toLowerCase();
    const mimeMap = {
      '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
      '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';
    const dataUrl = `data:${mime};base64,${buffer.toString('base64')}`;

    return { dataUrl, metadata, originalName: originalFile };
  } catch (err) {
    console.error('get-asset error:', err);
    return null;
  }
});

ipcMain.handle('get-asset-thumbnail', async (event, assetId) => {
  try {
    const assetsPath = await getAssetsPath();
    const assetDir = path.join(assetsPath, assetId);
    try {
      await fs.access(assetDir);
    } catch {
      return null;
    }

    const thumbnailPath = path.join(assetDir, 'thumbnail.png');
    try {
      await fs.access(thumbnailPath);
      const buffer = await fs.readFile(thumbnailPath);
      return `data:image/png;base64,${buffer.toString('base64')}`;
    } catch {}

    const entries = await fs.readdir(assetDir);
    const originalFile = entries.find(f => f.startsWith('original'));
    if (!originalFile) return null;

    const ext = path.extname(originalFile).toLowerCase();
    const imageExts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'];
    if (!imageExts.includes(ext)) return null;

    const buffer = await fs.readFile(path.join(assetDir, originalFile));
    const resized = resizeImageToMax(buffer, 512);
    if (resized) {
      const resizedBuffer = resized.toPNG();
      const mime = `data:image/png;base64,${resizedBuffer.toString('base64')}`;
      return mime;
    }

    const mimeMap = {
      '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
      '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';
    return `data:${mime};base64,${buffer.toString('base64')}`;
  } catch (err) {
    console.error('get-asset-thumbnail error:', err);
    return null;
  }
});

ipcMain.handle('save-file', async (event, { buffer, filename }) => {
  try {
    await ensureDataDir();
    const filesDir = path.join(getDataPath(), 'files');
    const ext = path.extname(filename);
    const id = Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
    const savedName = `item_${id}${ext}`;
    const filePath = path.join(filesDir, savedName);
    await fs.writeFile(filePath, Buffer.from(buffer));
    return { savedName, originalName: filename };
  } catch (err) {
    console.error('save-file error:', err);
    throw err;
  }
});

ipcMain.handle('read-state', async () => {
  try {
    const assetsPath = path.join(await getActiveLibraryPath(), 'assets');
    await fs.mkdir(assetsPath, { recursive: true });
    
    const entries = await fs.readdir(assetsPath);
    const items = [];
    
    for (const entry of entries) {
      const metadataPath = path.join(assetsPath, entry, 'metadata.json');
      try {
        const metadataData = await fs.readFile(metadataPath, 'utf-8');
        const metadata = JSON.parse(metadataData);
        if (metadata.type === 'tag') {
          await fs.rm(path.join(assetsPath, entry), { recursive: true, force: true });
          continue;
        }
        items.push(metadata);
      } catch {}
    }
    
    return { version: 1, zoom: 1, panX: 0, panY: 0, items };
  } catch {
    return { version: 1, zoom: 1, panX: 0, panY: 0, items: [] };
  }
});

ipcMain.handle('write-state', async (event, state) => {
  try {
    const assetsPath = path.join(await getActiveLibraryPath(), 'assets');
    await fs.mkdir(assetsPath, { recursive: true });
    
    if (state.items && Array.isArray(state.items)) {
      for (const item of state.items) {
        const metadataPath = path.join(assetsPath, item.assetId || item.id, 'metadata.json');
        try {
          await fs.access(path.dirname(metadataPath));
        } catch {
          await fs.mkdir(path.dirname(metadataPath), { recursive: true });
        }
        await fs.writeFile(metadataPath, JSON.stringify(item, null, 2));
      }
    }
  } catch (err) {
    console.error('write-state error:', err);
    throw err;
  }
});

ipcMain.handle('save-item-state', async (event, item) => {
  try {
    const assetsPath = path.join(await getActiveLibraryPath(), 'assets');
    const metadataPath = path.join(assetsPath, item.assetId || item.id, 'metadata.json');
    await fs.mkdir(path.dirname(metadataPath), { recursive: true });
    await fs.writeFile(metadataPath, JSON.stringify(item, null, 2));
  } catch (err) {
    console.error('save-item-state error:', err);
  }
});

ipcMain.handle('delete-item-state', async (event, assetId) => {
  try {
    const libPath = await getActiveLibraryPath();
    const assetsPath = path.join(libPath, 'assets');
    const assetDir = path.join(assetsPath, assetId);
    
    try {
      await fs.access(assetDir);
      await fs.rm(assetDir, { recursive: true, force: true });
    } catch {}
  } catch (err) {
    console.error('delete-item-state error:', err);
  }
});

ipcMain.handle('open-file', async (event, assetId) => {
    try {
      const activeLibraryPath = await getActiveLibraryPath();
      const assetsPath = path.join(activeLibraryPath, 'assets');
      const assetDir = path.join(assetsPath, assetId);
      try {
        await fs.access(assetDir);
      } catch {
        return null;
      }
    const entries = await fs.readdir(assetDir);
    const originalFile = entries.find(f => f.startsWith('original'));
    if (originalFile) {
      shell.openPath(path.join(assetDir, originalFile));
    }
  } catch (err) {
    console.error('open-file error:', err);
  }
});

ipcMain.handle('get-file-icon', async (event, filename) => {
  const ext = path.extname(filename).toLowerCase();
  const iconMap = {
    '.pdf': 'pdf',
    '.doc': 'word', '.docx': 'word',
    '.xls': 'excel', '.xlsx': 'excel',
    '.ppt': 'ppt', '.pptx': 'ppt',
    '.txt': 'text',
    '.zip': 'archive', '.rar': 'archive', '.7z': 'archive',
    '.mp3': 'audio', '.wav': 'audio', '.flac': 'audio',
    '.mp4': 'video', '.avi': 'video', '.mkv': 'video',
    '.obj': '3d', '.fbx': '3d', '.stl': '3d',
    '.png': 'image', '.jpg': 'image', '.jpeg': 'image', '.gif': 'image', '.bmp': 'image', '.webp': 'image'
  };
  return iconMap[ext] || 'file';
});

ipcMain.handle('read-clipboard-image', async () => {
  try {
    const image = clipboard.readImage();
    if (image.isEmpty()) {
      return null;
    }
    const buffer = image.toPNG();
    return { buffer: buffer.toString('base64'), filename: `clipboard_${Date.now()}.png` };
  } catch (err) {
    console.error('read-clipboard-image error:', err);
    return null;
  }
});

ipcMain.handle('read-clipboard-files', async () => {
  try {
    const hasFiles = clipboard.has('FileNameW');
    if (!hasFiles) return null;
    const buf = clipboard.readBuffer('FileNameW');
    if (!buf || buf.length < 20) return null;
    const pFiles = buf.readUInt32LE(0);
    if (pFiles >= buf.length) return null;
    const paths = [];
    let i = pFiles;
    while (i < buf.length - 1) {
      const code = buf.readUInt16LE(i);
      if (code === 0) {
        if (paths.length > 0) break;
        i += 2;
        continue;
      }
      let currentPath = '';
      while (i < buf.length - 1) {
        const c = buf.readUInt16LE(i);
        i += 2;
        if (c === 0) break;
        currentPath += String.fromCharCode(c);
      }
      if (currentPath.length > 0) paths.push(currentPath);
    }
    if (paths.length === 0) return null;
    const results = [];
    for (const filePath of paths) {
      try {
        const stat = await fs.stat(filePath);
        if (stat.isFile()) {
          const buffer = await fs.readFile(filePath);
          const filename = path.basename(filePath);
          results.push({ buffer: buffer.toString('base64'), filename });
        }
      } catch {}
    }
    return results.length > 0 ? results : null;
  } catch (e) {
    console.error('read-clipboard-files error:', e);
    return null;
  }
});

ipcMain.handle('select-library-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择素材库文件夹',
    properties: ['openDirectory']
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('create-library', async (event) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '新建素材库',
    defaultPath: '新素材库',
    buttonLabel: '创建',
    properties: ['showOverwriteConfirmation']
  });
  if (result.canceled || !result.filePath) return null;
  const libPath = result.filePath;
  try {
    await fs.access(libPath);
    return null;
  } catch {}
  await fs.mkdir(libPath, { recursive: true });
  const assetsDir = path.join(libPath, 'assets');
  await fs.mkdir(assetsDir, { recursive: true });
  return libPath;
});

ipcMain.handle('add-library', async (event, libPath) => {
  const config = await getLibrariesConfig();
  const existing = config.libraries.find(l => l.path === libPath);
  if (existing) return { id: existing.id, name: existing.name };
  const id = Date.now().toString(36) + Math.random().toString(36).substr(2, 6);
  const name = path.basename(libPath);
  config.libraries.push({ id, name, path: libPath });
  config.activeId = id;
  await saveLibrariesConfig(config);
  activeLibraryPath = libPath;
  await ensureDataDir();
  return { id, name };
});

ipcMain.handle('get-libraries', async () => {
  const config = await getLibrariesConfig();
  return { libraries: config.libraries, activeId: config.activeId };
});

ipcMain.handle('set-active-library', async (event, libId) => {
  const config = await getLibrariesConfig();
  const lib = config.libraries.find(l => l.id === libId);
  if (!lib) return false;
  config.activeId = libId;
  await saveLibrariesConfig(config);
  activeLibraryPath = lib.path;
  await ensureDataDir();
  mainWindow.webContents.send('library-changed', libId);
  return true;
});

ipcMain.handle('remove-library', async (event, libId) => {
  const config = await getLibrariesConfig();
  config.libraries = config.libraries.filter(l => l.id !== libId);
  if (config.activeId === libId) {
    config.activeId = config.libraries.length > 0 ? config.libraries[0].id : null;
    activeLibraryPath = config.activeId ? config.libraries.find(l => l.id === config.activeId).path : null;
  }
  await saveLibrariesConfig(config);
  return true;
});

ipcMain.handle('show-card-context-menu', async (event, { id, assetId, originalName, itemType }) => {
  const menu = new Menu();
  menu.append(new MenuItem({
    label: '添加时间',
    click: () => {
      event.sender.send('timer-action', { type: 'add', itemId: assetId || id });
    }
  }));
  menu.append(new MenuItem({
    label: '设置时间',
    click: () => {
      event.sender.send('timer-action', { type: 'set', itemId: assetId || id });
    }
  }));
  menu.append(new MenuItem({ type: 'separator' }));
  menu.append(new MenuItem({
    label: '创建分组',
    click: () => {
      event.sender.send('create-group', { itemId: assetId || id });
    }
  }));
  menu.append(new MenuItem({
    label: '取消分组',
    click: () => {
      event.sender.send('remove-from-group', { itemId: assetId || id });
    }
  }));
  menu.append(new MenuItem({ type: 'separator' }));
  menu.append(new MenuItem({
      label: '在文件浏览器中查看',
      click: async () => {
        const assetsPath = await getAssetsPath();
        const assetDir = path.join(assetsPath, assetId);
        try {
          await fs.access(assetDir);
          shell.showItemInFolder(path.join(assetDir, 'original' + path.extname(originalName)));
        } catch {}
      }
    }));
    menu.append(new MenuItem({
      label: '打开文件',
      click: async () => {
        const assetsPath = await getAssetsPath();
        const assetDir = path.join(assetsPath, assetId);
        try {
          await fs.access(assetDir);
          const entries = await fs.readdir(assetDir);
          const originalFile = entries.find(f => f.startsWith('original'));
          if (originalFile) {
            shell.openPath(path.join(assetDir, originalFile));
          }
        } catch {}
      }
    }));
  menu.append(new MenuItem({ type: 'separator' }));
  menu.append(new MenuItem({
    label: '复制文件名',
    click: () => {
      clipboard.writeText(originalName);
    }
  }));
  menu.popup({ window: BrowserWindow.fromWebContents(event.sender) });
});

ipcMain.handle('set-card-thumbnail', async (event, { assetId, buffer }) => {
    try {
      const activeLibraryPath = await getActiveLibraryPath();
      const assetsPath = path.join(activeLibraryPath, 'assets');
      const assetDir = path.join(assetsPath, assetId);
      try {
        await fs.access(assetDir);
      } catch {
        return null;
      }
    const thumbnailPath = path.join(assetDir, 'thumbnail.png');
    const buf = Buffer.from(buffer);
    await fs.writeFile(thumbnailPath, buf);
    return true;
  } catch (e) {
    console.error('set-card-thumbnail error:', e);
    return false;
  }
});

ipcMain.handle('read-file-as-base64', async (event, filePath) => {
  try {
    const buffer = await fs.readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const mimeMap = {
      '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
      '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';
    return `data:${mime};base64,${buffer.toString('base64')}`;
  } catch (err) {
    console.error('read-file-as-base64 error:', err);
    return null;
  }
});
