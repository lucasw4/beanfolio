const { app, BrowserWindow, ipcMain, shell } = require('electron');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const DESKTOP_SERVER_PORT = 5180;
const DESKTOP_SERVER_HOST = '127.0.0.1';
const APP_DISPLAY_NAME = 'Beanfolio';

const BASE_WIDTH = 500;
const BASE_HEIGHT = 510;
const SIDEBAR_WIDTH = 220;

const DEV_SERVER_URL = process.env.VITE_DEV_SERVER_URL;
const isDev = Boolean(DEV_SERVER_URL);

let staticServer = null;

app.setName(APP_DISPLAY_NAME);

function createMainWindow(startUrl) {
  const window = new BrowserWindow({
    width: BASE_WIDTH,
    height: BASE_HEIGHT,
    minWidth: BASE_WIDTH,
    minHeight: 400,
    title: 'Beanfolio',
    titleBarStyle: 'hiddenInset',
    transparent: true,
    backgroundColor: '#00000000',
    vibrancy: 'under-window',
    visualEffectState: 'active',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      sandbox: true,
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  window.webContents.setWindowOpenHandler(({ url }) => {
    if (isGoogleAuthUrl(url)) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          width: 540,
          height: 700,
          minWidth: 440,
          minHeight: 560,
          autoHideMenuBar: true,
          title: 'Google Sign-In',
          webPreferences: {
            sandbox: true,
            nodeIntegration: false,
            contextIsolation: true,
          },
        },
      };
    }

    if (isExternalHttpUrl(url)) {
      shell.openExternal(url);
      return { action: 'deny' };
    }

    return { action: 'allow' };
  });

  window.webContents.on('will-navigate', (event, url) => {
    if (url === startUrl || url.startsWith(`${startUrl}/`)) {
      return;
    }

    if (isExternalHttpUrl(url)) {
      event.preventDefault();
      shell.openExternal(url);
    }
  });

  window.loadURL(startUrl);

  ipcMain.on('set-sidebar-open', (event, isOpen) => {
    const win = BrowserWindow.fromWebContents(event.sender);
    if (!win) return;
    const [, currentHeight] = win.getSize();
    const targetWidth = isOpen ? BASE_WIDTH + SIDEBAR_WIDTH : BASE_WIDTH;
    win.setSize(targetWidth, currentHeight, true);
  });
}

async function startDesktopStaticServer() {
  const distDir = path.join(__dirname, '..', 'dist');

  if (!fs.existsSync(path.join(distDir, 'index.html'))) {
    throw new Error('Missing dist/index.html. Run "npm run build" before launching desktop mode.');
  }

  staticServer = http.createServer((req, res) => {
    const requestPath = sanitizeRequestPath(req.url ?? '/');
    const requestedFilePath = path.join(distDir, requestPath);
    const safePath = path.normalize(requestedFilePath);

    if (!safePath.startsWith(path.normalize(distDir))) {
      res.writeHead(403);
      res.end('Forbidden');
      return;
    }

    const fileToServe = resolveFilePath(distDir, safePath);
    if (!fileToServe) {
      res.writeHead(404);
      res.end('Not Found');
      return;
    }

    const extension = path.extname(fileToServe);
    const contentType = getContentType(extension);

    fs.readFile(fileToServe, (error, fileBuffer) => {
      if (error) {
        res.writeHead(500);
        res.end('Internal Server Error');
        return;
      }

      res.writeHead(200, {
        'Content-Type': contentType,
      });
      res.end(fileBuffer);
    });
  });

  await new Promise((resolve, reject) => {
    staticServer.once('error', reject);
    staticServer.listen(DESKTOP_SERVER_PORT, DESKTOP_SERVER_HOST, () => resolve());
  });
}

function resolveFilePath(distDir, requestedPath) {
  if (fs.existsSync(requestedPath) && fs.statSync(requestedPath).isFile()) {
    return requestedPath;
  }

  return path.join(distDir, 'index.html');
}

function sanitizeRequestPath(rawUrl) {
  const pathWithoutQuery = rawUrl.split('?')[0] || '/';
  const decodedPath = decodeURIComponent(pathWithoutQuery);
  const normalizedPath = path.normalize(decodedPath);

  if (normalizedPath === path.sep) {
    return 'index.html';
  }

  return normalizedPath.replace(/^[/\\]+/, '');
}

function getContentType(extension) {
  const contentTypes = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.svg': 'image/svg+xml',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.webp': 'image/webp',
    '.ico': 'image/x-icon',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
  };

  return contentTypes[extension] || 'application/octet-stream';
}

function isExternalHttpUrl(url) {
  return /^https?:\/\//i.test(url);
}

function isGoogleAuthUrl(url) {
  if (!isExternalHttpUrl(url)) {
    return false;
  }

  try {
    const parsed = new URL(url);
    const host = parsed.hostname.toLowerCase();
    return (
      host === 'accounts.google.com'
      || host.endsWith('.accounts.google.com')
      || host === 'oauth2.googleapis.com'
      || host.endsWith('.oauth2.googleapis.com')
    );
  } catch {
    return false;
  }
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', () => {
  if (staticServer) {
    staticServer.close();
    staticServer = null;
  }
});

app.whenReady().then(async () => {
  if (!isDev) {
    await startDesktopStaticServer();
  }

  const appUrl = isDev
    ? DEV_SERVER_URL
    : `http://${DESKTOP_SERVER_HOST}:${DESKTOP_SERVER_PORT}`;

  createMainWindow(appUrl);

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow(appUrl);
    }
  });
});
