const { app, BrowserWindow, Tray, nativeImage, screen, Menu, ipcMain, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const https = require('https');
const http = require('http');
const os = require('os');

const TA_URL = 'https://demo1-dev.dmoeutta.dev.tungstencloud.com/forms/sene/SENE-Launchpad.form';

const config = (() => {
  const defaults = require('./ta.config.js');
  try { return { ...defaults, ...require('./ta.config.local.js') }; }
  catch { return defaults; }
})();

function postJson(url, payload, token) {
  return new Promise((resolve) => {
    let body;
    try { body = JSON.stringify(payload); }
    catch (e) { return resolve({ ok: false, message: 'Failed to serialize payload.' }); }

    let parsed;
    try { parsed = new URL(url); }
    catch (e) { return resolve({ ok: false, message: `Invalid URL: ${url}` }); }

    const isHttps = parsed.protocol === 'https:';
    const client = isHttps ? https : http;
    const headers = { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) };
    if (token) headers['Authorization'] = `Bearer ${token}`;

    const req = client.request({
      hostname: parsed.hostname,
      port: parsed.port || (isHttps ? 443 : 80),
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers,
    }, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const json = JSON.parse(data);
            resolve({ ok: true, jobId: json.jobId, workflowUrl: json.workflowUrl, status: json.status });
          } catch {
            resolve({ ok: true });
          }
        } else {
          let message = `HTTP ${res.statusCode}`;
          try {
            const json = JSON.parse(data);
            if (json.message || json.error) message += `: ${json.message || json.error}`;
          } catch { /* ignore */ }
          resolve({ ok: false, httpStatus: res.statusCode, message });
        }
      });
    });

    req.on('error', (err) => resolve({ ok: false, message: `Network error: ${err.message}` }));
    req.write(body);
    req.end();
  });
}

app.setName('TA Launchpad');

// Prevent the app from showing in the taskbar
app.dock && app.dock.hide();

let tray = null;
let win = null;

const WINDOW_WIDTH = 480;
const WINDOW_HEIGHT = 720;
const MARGIN = 12;

function getWindowPosition() {
  const trayBounds = tray.getBounds();
  const display = screen.getDisplayNearestPoint({ x: trayBounds.x, y: trayBounds.y });
  const workArea = display.workArea;

  // Position window on the right side of the screen, near the taskbar
  const x = workArea.x + workArea.width - WINDOW_WIDTH - MARGIN;
  const y = workArea.y + workArea.height - WINDOW_HEIGHT - MARGIN;

  return { x, y };
}

function createWindow() {
  win = new BrowserWindow({
    width: WINDOW_WIDTH,
    height: WINDOW_HEIGHT,
    show: false,
    frame: true,
    resizable: true,
    skipTaskbar: true,
    alwaysOnTop: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      webviewTag: true,
    },
    title: 'TA Launchpad',
  });

  win.setAlwaysOnTop(true, 'floating');
  Menu.setApplicationMenu(null);
  win.loadFile('index.html');

  win.on('closed', () => {
    win = null;
  });
}

function toggleWindow() {
  if (!win || win.isDestroyed()) {
    createWindow();
    const { x, y } = getWindowPosition();
    win.setPosition(x, y, false);
    win.setAlwaysOnTop(true, 'floating');
    win.show();
    win.focus();
    return;
  }

  if (win.isVisible()) {
    win.hide();
  } else {
    const { x, y } = getWindowPosition();
    win.setPosition(x, y, false);
    win.setAlwaysOnTop(true, 'floating');
    win.show();
    win.focus();
  }
}

ipcMain.handle('get-form-url', () => TA_URL);

ipcMain.handle('detect-word', () => {
  return new Promise((resolve) => {
    const script = `
try {
  $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
  try {
    $doc = $word.ActiveDocument
    $name = $doc.Name
    $fullName = $doc.FullName
    if ($fullName -eq $name) {
      Write-Output '{"status":"unsaved"}'
    } else {
      Write-Output ([PSCustomObject]@{status='ok';name=$name;fullName=$fullName} | ConvertTo-Json -Compress)
    }
  } catch {
    Write-Output '{"status":"no_document"}'
  }
} catch {
  Write-Output '{"status":"not_running"}'
}
`;
    const ps = spawn('powershell', ['-NonInteractive', '-NoProfile', '-Command', script]);
    let output = '';
    ps.stdout.on('data', d => { output += d.toString(); });
    ps.on('close', () => {
      try { resolve(JSON.parse(output.trim())); }
      catch { resolve({ status: 'error' }); }
    });
  });
});

// Calls our server-side wrapper endpoint, which is expected to invoke:
//   JobService.CreateJob(sessionId, processIdentity, jobInitialization)
// Electron never calls the TA SDK directly.
ipcMain.handle('submit-word-document', async (_event, documentInfo) => {
  if (!config.TA_API_BASE_URL) {
    return { ok: false, message: 'API not configured. Set TA_API_BASE_URL in ta.config.local.js.' };
  }
  const payload = {
    documentName: documentInfo.name,
    documentPath: documentInfo.fullName,
    submittedBy: os.userInfo().username,
    source: 'TotalAgility Desktop Assistant',
    submittedAt: new Date().toISOString(),
    processName: 'Word Document Review',
  };
  const url = config.TA_API_BASE_URL + config.TA_START_WORKFLOW_ENDPOINT;
  return postJson(url, payload, config.TA_API_TOKEN || null);
});

ipcMain.handle('open-external', (_event, url) => shell.openExternal(url));

app.whenReady().then(() => {
  // Keep the app running without a dock/taskbar presence
  app.setAppUserModelId('com.ta.launchpad');

  const iconPath = path.join(__dirname, 'icon.png');
  const trayIcon = nativeImage.createFromPath(iconPath);
  tray = new Tray(trayIcon.resize({ width: 16, height: 16 }));

  tray.setToolTip('TA Launchpad');
  tray.on('click', toggleWindow);
  tray.on('double-click', toggleWindow);
});

app.on('window-all-closed', (e) => {
  // Don't quit when all windows are closed — stay in tray
  e.preventDefault();
});
