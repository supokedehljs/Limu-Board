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
    dataPath = path.join(app.getPath('userData'), 'limu-board-data');
  }
  return dataPath;
}

async function hasActiveLibrary() {
  const libPath = await getActiveLibraryPath();
  return libPath !== null;
}

async function getLibrariesConfig() {
  const dataPath = getDataPath();
  const configPath = path.join(dataPath, 'libraries.json');
  try {
    const data = await fs.readFile(configPath, 'utf-8');
    return JSON.parse(data);
  } catch {
    return { libraries: [], activeId: null };
  }
}

async function saveLibrariesConfig(config) {
  const dataPath = getDataPath();
  await fs.mkdir(dataPath, { recursive: true });
  const configPath = path.join(dataPath, 'libraries.json');
  await atomicWriteFile(configPath, JSON.stringify(config, null, 2));
}

async function loadSettings() {
  const libPath = await getActiveLibraryPath();
  if (!libPath) {
    return { globalHotkey: 'Alt+W', localHotkeys: {} };
  }
  const settingsPath = path.join(libPath, 'settings.json');
  try {
    const data = await fs.readFile(settingsPath, 'utf-8');
    return JSON.parse(data);
  } catch {
    return { globalHotkey: 'Alt+W', localHotkeys: {} };
  }
}

async function saveSettings(settings) {
  const libPath = await getActiveLibraryPath();
  if (!libPath) return;
  const settingsPath = path.join(libPath, 'settings.json');
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
  return null;
}

async function getStatePath() {
  const libPath = await getActiveLibraryPath();
  if (!libPath) return null;
  return path.join(libPath, 'state.json');
}

async function getAssetsPath() {
  const libPath = await getActiveLibraryPath();
  if (!libPath) return null;
  return path.join(libPath, 'assets');
}

async function ensureDataDir() {
  const libPath = await getActiveLibraryPath();
  if (!libPath) return null;
  
  const assetsDir = path.join(libPath, 'assets');
  await fs.mkdir(assetsDir, { recursive: true });
  return libPath;
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

ipcMain.handle('has-active-library', async () => {
  return await hasActiveLibrary();
});

ipcMain.handle('add-asset', async (event, { buffer, filename, cardId }) => {
  try {
    const activeLibraryPath = await getActiveLibraryPath();
    if (!activeLibraryPath) {
      throw new Error('No active library');
    }
    
    await ensureDataDir();
    const assetId = generateAssetId();
    const assetsPath = path.join(activeLibraryPath, 'assets');
    
    let assetDir;
    if (cardId) {
      assetDir = path.join(assetsPath, cardId);
    } else {
      assetDir = path.join(assetsPath, assetId);
    }
    await fs.mkdir(assetDir, { recursive: true });

    const ext = path.extname(filename).toLowerCase();
    let savedName = `original${ext}`;
    let finalPath = path.join(assetDir, savedName);
    
    let counter = 1;
    while (await fs.access(finalPath).then(() => true).catch(() => false)) {
      savedName = `original_${counter}${ext}`;
      finalPath = path.join(assetDir, savedName);
      counter++;
    }
    await fs.writeFile(finalPath, Buffer.from(buffer));

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
    let assetDir = path.join(assetsPath, assetId);
    let found = false;
    
    try {
      await fs.access(assetDir);
      found = true;
    } catch {
      const entries = await fs.readdir(assetsPath);
      for (const entry of entries) {
        const attachmentPath = path.join(assetsPath, entry, assetId);
        try {
          await fs.access(attachmentPath);
          assetDir = attachmentPath;
          found = true;
          break;
        } catch {}
      }
    }
    
    if (!found) return null;

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

ipcMain.handle('get-asset-thumbnail', async (event, assetId, cardId) => {
  try {
    const assetsPath = await getAssetsPath();
    let assetDir = null;
    let found = false;
    let originalFile = null;
    
    // First, if we have a cardId, look in the card's folder for the attachment
    if (cardId) {
      const cardDir = path.join(assetsPath, cardId);
      try {
        await fs.access(cardDir);
        const attachmentPath = path.join(cardDir, assetId);
        try {
          await fs.access(attachmentPath);
          assetDir = cardDir;
          found = true;
        } catch {}
      } catch {}
    }
    
    // If not found in card folder, try treating assetId as a folder
    if (!found) {
      assetDir = path.join(assetsPath, assetId);
      try {
        await fs.access(assetDir);
        found = true;
      } catch {
        // If not a folder, search through all asset folders
        const entries = await fs.readdir(assetsPath);
        for (const entry of entries) {
          const attachmentPath = path.join(assetsPath, entry, assetId);
          try {
            await fs.access(attachmentPath);
            assetDir = path.join(assetsPath, entry);
            found = true;
            break;
          } catch {}
        }
      }
    }
    
    if (!found) return null;

    // Look for image files in the directory
    const entries = await fs.readdir(assetDir);
    originalFile = entries.find(f => {
      const ext = path.extname(f).toLowerCase();
      return ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'].includes(ext);
    }) || entries.find(f => f.startsWith('original'));
    
    if (!originalFile) return null;

    const thumbnailPath = path.join(assetDir, 'thumbnail.png');
    try {
      await fs.access(thumbnailPath);
      const buffer = await fs.readFile(thumbnailPath);
      return `data:image/png;base64,${buffer.toString('base64')}`;
    } catch {}

    if (!originalFile) return null;

    const ext = path.extname(originalFile).toLowerCase();
    const imageExts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'];
    if (!imageExts.includes(ext)) return null;

    const buffer = await fs.readFile(path.join(assetDir, originalFile));
    const resized = resizeImageToMax(buffer, 512);
    if (resized) {
      const resizedBuffer = resized.toPNG();
      const dataUrl = `data:image/png;base64,${resizedBuffer.toString('base64')}`;
      return dataUrl;
    }

    const mimeMap = {
      '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
      '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';
    return `data:${mime};base64,${buffer.toString('base64')}`;
  } catch (err) {
    return null;
  }
});

ipcMain.handle('open-folder', async (event, assetId) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, assetId);
    try {
      await fs.access(cardDir);
      require('electron').shell.openPath(cardDir);
      return { success: true };
    } catch {
      const entries = await fs.readdir(assetsPath);
      for (const entry of entries) {
        const attachmentPath = path.join(assetsPath, entry, assetId);
        try {
          await fs.access(attachmentPath);
          const cardFolder = path.dirname(attachmentPath);
          require('electron').shell.openPath(cardFolder);
          return { success: true };
        } catch {}
      }
      return { success: false, error: 'Folder not found' };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('scan-attachments', async (event, cardId) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, cardId);
    
    try {
      await fs.access(cardDir);
    } catch {
      return { success: false, error: 'Card folder not found' };
    }
    
    const files = await fs.readdir(cardDir);
    const attachments = [];
    let thumbnailAssetId = null;
    
    for (const file of files) {
      if (file === 'metadata.json') continue;
      const filePath = path.join(cardDir, file);
      const stat = await fs.stat(filePath);
      if (!stat.isFile()) continue;
      
      attachments.push(file);
    }
    
    const exts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'];
    for (const f of attachments) {
      const ext = path.extname(f).toLowerCase();
      if (exts.includes(ext)) {
        thumbnailAssetId = f;
        break;
      }
    }
    
    return { success: true, attachments, thumbnailAssetId };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('rename-attachment', async (event, { cardId, oldName, newName }) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, cardId);
    const oldPath = path.join(cardDir, oldName);
    const newPath = path.join(cardDir, newName);
    
    await fs.access(oldPath);
    await fs.rename(oldPath, newPath);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-attachment-file', async (event, { cardId, filename }) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, cardId);
    const filePath = path.join(cardDir, filename);
    
    await fs.access(filePath);
    await fs.unlink(filePath);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-file', async (event, { buffer, filename }) => {
  try {
    const libPath = await getActiveLibraryPath();
    if (!libPath) throw new Error('No active library');
    const filesDir = path.join(libPath, 'files');
    await fs.mkdir(filesDir, { recursive: true });
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

let _cachedLibraryPath = null;

async function atomicWriteFile(filePath, data) {
  const tmp = filePath + '.tmp';
  await fs.writeFile(tmp, data, 'utf-8');
  await fs.rename(tmp, filePath);
}

async function getActiveLibraryPath() {
  if (_cachedLibraryPath !== null) return _cachedLibraryPath;
  const config = await getLibrariesConfig();
  if (config.activeId && config.libraries.length > 0) {
    const lib = config.libraries.find(l => l.id === config.activeId);
    _cachedLibraryPath = lib ? lib.path : null;
  } else {
    _cachedLibraryPath = null;
  }
  return _cachedLibraryPath;
}

ipcMain.handle('read-state', async () => {
  try {
    const libPath = await getActiveLibraryPath();
    if (!libPath) return { version: 1, zoom: 1, panX: 0, panY: 0, items: [] };

    const assetsPath = path.join(libPath, 'assets');
    await fs.mkdir(assetsPath, { recursive: true });

    let catalog = { cards: [], trash: [] };
    const catalogPath = path.join(libPath, 'card-catalog.json');
    try {
      const data = await fs.readFile(catalogPath, 'utf-8');
      catalog = JSON.parse(data);
    } catch {
      catalog = { cards: [], trash: [] };
    }
    if (!catalog.trash) catalog.trash = [];
    const trashSet = new Set(catalog.trash);

    const entries = await fs.readdir(assetsPath);
    const items = [];

    for (const entry of entries) {
      if (trashSet.has(entry)) continue;
      const metadataPath = path.join(assetsPath, entry, 'metadata.json');
      try {
        const metadataData = await fs.readFile(metadataPath, 'utf-8');
        const metadata = JSON.parse(metadataData);
        if (metadata.type === 'tag') continue;
        items.push(metadata);
      } catch {}
    }

    const currentIds = new Set(items.map(i => i.assetId));
    const missing = [...currentIds].filter(id => !catalog.cards.includes(id) && !trashSet.has(id));
    if (missing.length > 0) {
      catalog.cards = [...catalog.cards, ...missing];
      await atomicWriteFile(catalogPath, JSON.stringify(catalog, null, 2));
    }
    
    return { version: 1, zoom: 1, panX: 0, panY: 0, items };
  } catch {
    return { version: 1, zoom: 1, panX: 0, panY: 0, items: [] };
  }
});

ipcMain.handle('write-state', async (event, state) => {
  try {
    const libPath = await getActiveLibraryPath();
    if (!libPath) return;
    
    const catalog = await getCatalog(libPath);
    const trashSet = new Set(catalog.trash || []);
    const assetsPath = path.join(libPath, 'assets');
    await fs.mkdir(assetsPath, { recursive: true });
    
    if (state.items && Array.isArray(state.items)) {
      for (const item of state.items) {
        const id = item.assetId || item.id;
        if (trashSet.has(id)) continue;
        const metadataPath = path.join(assetsPath, id, 'metadata.json');
        await fs.mkdir(path.dirname(metadataPath), { recursive: true });
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
    const libPath = await getActiveLibraryPath();
    if (!libPath) return;
    const assetsPath = path.join(libPath, 'assets');
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
    if (!libPath) return;
    const assetsPath = path.join(libPath, 'assets');
    const trashPath = path.join(libPath, '.trash');
    let assetDir = path.join(assetsPath, assetId);
    let movedToTrash = false;
    
    try {
      await fs.access(assetDir);
      await fs.mkdir(trashPath, { recursive: true });
      
      let trashItemPath = path.join(trashPath, assetId);
      let counter = 1;
      while (await fs.access(trashItemPath).then(() => true).catch(() => false)) {
        trashItemPath = path.join(trashPath, `${assetId}_${counter}`);
        counter++;
      }
      await fs.rename(assetDir, trashItemPath);
      movedToTrash = true;
    } catch {
      const entries = await fs.readdir(assetsPath);
      for (const entry of entries) {
        const attachmentPath = path.join(assetsPath, entry, assetId);
        try {
          await fs.access(attachmentPath);
          assetDir = attachmentPath;
          await fs.mkdir(trashPath, { recursive: true });
          let trashItemPath = path.join(trashPath, assetId);
          let counter = 1;
          while (await fs.access(trashItemPath).then(() => true).catch(() => false)) {
            trashItemPath = path.join(trashPath, `${assetId}_${counter}`);
            counter++;
          }
          await fs.rename(assetDir, trashItemPath);
          movedToTrash = true;
          break;
        } catch {}
      }
    }
    
    await removeFromCatalog(libPath, assetId);
    
    let catalog = { cards: [], trash: [] };
    const catalogPath = path.join(libPath, 'card-catalog.json');
    try {
      const catalogData = await fs.readFile(catalogPath, 'utf-8');
      catalog = JSON.parse(catalogData);
    } catch {}
    if (!catalog.trash) catalog.trash = [];
    if (movedToTrash && !catalog.trash.includes(assetId)) {
      catalog.trash.push(assetId);
    }
    catalog.cards = catalog.cards.filter(id => id !== assetId);
    await atomicWriteFile(catalogPath, JSON.stringify(catalog, null, 2));
  } catch (err) {
    console.error('delete-item-state error:', err);
  }
});

async function getCatalog(libPath) {
  const catalogPath = path.join(libPath, 'card-catalog.json');
  try {
    const data = await fs.readFile(catalogPath, 'utf-8');
    return JSON.parse(data);
  } catch {
    return { cards: [] };
  }
}

async function saveCatalog(libPath, catalog) {
  const catalogPath = path.join(libPath, 'card-catalog.json');
  await atomicWriteFile(catalogPath, JSON.stringify(catalog, null, 2));
}

async function addToCatalog(libPath, assetId) {
  const catalog = await getCatalog(libPath);
  if (!catalog.cards.includes(assetId)) {
    catalog.cards.push(assetId);
    await saveCatalog(libPath, catalog);
  }
}

async function removeFromCatalog(libPath, assetId) {
  const catalog = await getCatalog(libPath);
  catalog.cards = catalog.cards.filter(id => id !== assetId);
  await saveCatalog(libPath, catalog);
}

ipcMain.handle('open-file', async (event, assetId) => {
    try {
      const activeLibraryPath = await getActiveLibraryPath();
      if (!activeLibraryPath) return;
      
      const assetsPath = path.join(activeLibraryPath, 'assets');
      let assetDir = path.join(assetsPath, assetId);
      let found = false;
      
      try {
        await fs.access(assetDir);
        found = true;
      } catch {
        const entries = await fs.readdir(assetsPath);
        for (const entry of entries) {
          const attachmentPath = path.join(assetsPath, entry, assetId);
          try {
            await fs.access(attachmentPath);
            assetDir = attachmentPath;
            found = true;
            break;
          } catch {}
        }
      }
      
      if (!found) return null;
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
    const formats = clipboard.availableFormats();
    if (!formats.includes('FileNameW') && !formats.includes('FileContents')) {
      return null;
    }
    const buf = clipboard.readBuffer('FileNameW');
    if (!buf || buf.length < 4) return null;
    
    const pFiles = buf.readUInt32LE(0);
    if (pFiles === 0 || pFiles >= buf.length) {
      const files = [];
      let offset = 0;
      while (offset < buf.length - 1) {
        const char = buf[offset];
        if (char === 0) {
          if (files.length > 0) break;
          offset++;
          continue;
        }
        let str = '';
        while (offset < buf.length) {
          const c = buf[offset];
          if (c === 0) break;
          str += String.fromCharCode(c);
          offset++;
        }
        if (str.length > 0 && (str.includes(':') || str.includes('\\'))) {
          files.push(str);
        }
        offset++;
      }
      if (files.length > 0) {
        const results = [];
        for (const filePath of files) {
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
      }
      return null;
    }
    
    const paths = [];
    let i = pFiles;
    while (i < buf.length - 1) {
      if (buf[i] === 0 && buf[i + 1] === 0) {
        break;
      }
      let currentPath = '';
      while (i < buf.length - 1) {
        const low = buf[i];
        const high = buf[i + 1];
        if (low === 0 && high === 0) break;
        i += 2;
        currentPath += String.fromCharCode(low + (high << 8));
      }
      if (currentPath.length > 0 && (currentPath.includes(':') || currentPath.includes('\\'))) {
        paths.push(currentPath);
      }
      i += 2;
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
  _cachedLibraryPath = null;
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
  _cachedLibraryPath = null;
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
    label: '选择所有同标签',
    click: () => {
      event.sender.send('select-same-tags', { itemId: assetId || id });
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
      let assetDir = path.join(assetsPath, assetId);
      let found = false;
      
      try {
        await fs.access(assetDir);
        found = true;
      } catch {
        const entries = await fs.readdir(assetsPath);
        for (const entry of entries) {
          const attachmentPath = path.join(assetsPath, entry, assetId);
          try {
            await fs.access(attachmentPath);
            assetDir = attachmentPath;
            found = true;
            break;
          } catch {}
        }
      }
      
      if (!found) return null;
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

ipcMain.handle('add-timer', async (event, { itemId, hours, minutes }) => {
  try {
    const assetsPath = await getAssetsPath();
    const itemDir = path.join(assetsPath, itemId);
    const metadataPath = path.join(itemDir, 'metadata.json');
    const metadata = JSON.parse(await fs.readFile(metadataPath, 'utf-8'));
    metadata.timerTotal = (metadata.timerTotal || 0) + (hours * 3600 + minutes * 60) * 1000;
    await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
    return true;
  } catch (err) {
    console.error('add-timer error:', err);
    return false;
  }
});

ipcMain.handle('set-timer', async (event, { itemId, hours, minutes }) => {
  try {
    const assetsPath = await getAssetsPath();
    const itemDir = path.join(assetsPath, itemId);
    const metadataPath = path.join(itemDir, 'metadata.json');
    const metadata = JSON.parse(await fs.readFile(metadataPath, 'utf-8'));
    metadata.timerTotal = (hours * 3600 + minutes * 60) * 1000;
    await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
    return true;
  } catch (err) {
    console.error('set-timer error:', err);
    return false;
  }
});

ipcMain.handle('save-tag-meta', async (event, meta) => {
  try {
    const libPath = await getActiveLibraryPath();
    if (!libPath) return false;
    const metaPath = path.join(libPath, 'tag-meta.json');
    await fs.writeFile(metaPath, JSON.stringify(meta, null, 2));
    return true;
  } catch (err) {
    console.error('save-tag-meta error:', err);
    return false;
  }
});

ipcMain.handle('load-tag-meta', async () => {
  try {
    const libPath = await getActiveLibraryPath();
    if (!libPath) return null;
    const metaPath = path.join(libPath, 'tag-meta.json');
    const data = await fs.readFile(metaPath, 'utf-8');
    return JSON.parse(data);
  } catch (err) {
    return null;
  }
});

ipcMain.handle('create-empty-card', async (event, { cardName }) => {
  try {
    const libPath = await getActiveLibraryPath();
    if (!libPath) return { success: false, error: 'No library' };
    
    const assetsPath = path.join(libPath, 'assets');
    const assetId = generateAssetId();
    const assetDir = path.join(assetsPath, assetId);
    await fs.mkdir(assetDir, { recursive: true });

    const newItem = {
      id: assetId,
      assetId: assetId,
      originalName: '',
      cardName: cardName || '',
      type: 'card',
      x: 0,
      y: 0,
      width: 200,
      height: 150,
      timerTotal: 0,
      timerRunning: false,
      timerStartedAt: null,
      cardTags: [],
      cardAnnotation: '',
      cardUrl: '',
      addedTime: Date.now(),
      attachments: [],
      thumbnailAssetId: null
    };

    await fs.writeFile(path.join(assetDir, 'metadata.json'), JSON.stringify(newItem, null, 2));
    await addToCatalog(libPath, assetId);
    return { assetId, item: newItem };
  } catch (err) {
    console.error('create-empty-card error:', err);
    throw err;
  }
});

ipcMain.handle('attach-asset-to-card', async (event, { cardId, assetId }) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, cardId);
    const metadataPath = path.join(cardDir, 'metadata.json');
    
    let metadata;
    try {
      const data = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(data);
    } catch {
      return { success: false, error: 'Card not found' };
    }

    if (!metadata.attachments) metadata.attachments = [];
    if (!metadata.attachments.includes(assetId)) {
      metadata.attachments.push(assetId);
    }
    if (!metadata.thumbnailAssetId && assetId) {
      metadata.thumbnailAssetId = assetId;
    }

    await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
    return { success: true, metadata };
  } catch (err) {
    console.error('attach-asset-to-card error:', err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle('detach-asset-from-card', async (event, { cardId, assetId }) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, cardId);
    const metadataPath = path.join(cardDir, 'metadata.json');
    
    let metadata;
    try {
      const data = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(data);
    } catch {
      return { success: false, error: 'Card not found' };
    }

    if (!metadata.attachments) metadata.attachments = [];
    metadata.attachments = metadata.attachments.filter(a => a !== assetId);
    
    if (metadata.thumbnailAssetId === assetId) {
      metadata.thumbnailAssetId = metadata.attachments.length > 0 ? metadata.attachments[0] : null;
    }

    await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
    return { success: true, metadata };
  } catch (err) {
    console.error('detach-asset-from-card error:', err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle('set-card-thumbnail-asset', async (event, { cardId, thumbnailAssetId }) => {
  try {
    const assetsPath = await getAssetsPath();
    const cardDir = path.join(assetsPath, cardId);
    const metadataPath = path.join(cardDir, 'metadata.json');
    
    let metadata;
    try {
      const data = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(data);
    } catch {
      return { success: false, error: 'Card not found' };
    }

    metadata.thumbnailAssetId = thumbnailAssetId;
    await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
    return { success: true, metadata };
  } catch (err) {
    console.error('set-card-thumbnail-asset error:', err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-asset', async (event, assetId) => {
  try {
    const assetsPath = await getAssetsPath();
    const assetDir = path.join(assetsPath, assetId);
    
    try {
      await fs.access(assetDir);
    } catch {
      return { success: false, error: 'Asset not found' };
    }

    await fs.rm(assetDir, { recursive: true, force: true });
    return { success: true };
  } catch (err) {
    console.error('delete-asset error:', err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle('get-attachment-thumbnails', async (event, assetIds) => {
  try {
    const results = {};
    for (const assetId of assetIds) {
      const thumbUrl = await getAssetThumbnail(assetId);
      results[assetId] = thumbUrl;
    }
    return results;
  } catch (err) {
    console.error('get-attachment-thumbnails error:', err);
    return {};
  }
});

async function getAssetThumbnail(assetId) {
  try {
    const assetsPath = await getAssetsPath();
    let assetDir = path.join(assetsPath, assetId);
    let found = false;
    let originalFile = null;
    
    try {
      await fs.access(assetDir);
      found = true;
      const entries = await fs.readdir(assetDir);
      originalFile = entries.find(f => f.startsWith('original')) || entries.find(f => {
        const ext = path.extname(f).toLowerCase();
        return ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'].includes(ext);
      });
    } catch {
      const entries = await fs.readdir(assetsPath);
      for (const entry of entries) {
        const attachmentPath = path.join(assetsPath, entry, assetId);
        try {
          await fs.access(attachmentPath);
          assetDir = attachmentPath;
          found = true;
          const entryFiles = await fs.readdir(assetDir);
          originalFile = entryFiles.find(f => f.startsWith('original')) || entryFiles.find(f => {
            const ext = path.extname(f).toLowerCase();
            return ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'].includes(ext);
          });
          break;
        } catch {}
      }
    }
    
    if (!found) return null;

    const thumbnailPath = path.join(assetDir, 'thumbnail.png');
    try {
      await fs.access(thumbnailPath);
      const buffer = await fs.readFile(thumbnailPath);
      return `data:image/png;base64,${buffer.toString('base64')}`;
    } catch {}

    if (!originalFile) return null;

    const ext = path.extname(originalFile).toLowerCase();
    const imageExts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'];
    if (!imageExts.includes(ext)) return null;

    const buffer = await fs.readFile(path.join(assetDir, originalFile));
    const resized = resizeImageToMax(buffer, 512);
    if (resized) {
      const resizedBuffer = resized.toPNG();
      const dataUrl = `data:image/png;base64,${resizedBuffer.toString('base64')}`;
      return dataUrl;
    }

    const mimeMap = {
      '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
      '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';
    return `data:${mime};base64,${buffer.toString('base64')}`;
  } catch (err) {
    return null;
  }
}
