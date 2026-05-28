const { app, BrowserWindow, Tray, nativeImage, screen, Menu, ipcMain, shell, webContents } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const https = require('https');
const http = require('http');
const os = require('os');
const fs = require('fs');

const TA_URL = 'https://demo1-dev.dmoeutta.dev.tungstencloud.com/forms/tacf_da/ta_cf_debugForm.form';

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

// ─── TotalAgility SDK client ─────────────────────────────────────────────────

function sdkEndpoint(service, method) {
  const base = (config.TA_SDK_BASE_URL || `${config.TA_API_BASE_URL}/TotalAgility/Services/Sdk`).replace(/\/$/, '');
  return `${base}/${service}.svc/json/${method}`;
}

// Low-level TA SDK POST — unwraps .d, throws on HTTP error.
function taPost(url, body) {
  return new Promise((resolve, reject) => {
    let bodyStr;
    try { bodyStr = JSON.stringify(body); }
    catch (e) { return reject(new Error('Failed to serialize request.')); }

    let parsed;
    try { parsed = new URL(url); }
    catch (e) { return reject(new Error(`Invalid SDK URL: ${url}`)); }

    const isHttps = parsed.protocol === 'https:';
    const client = isHttps ? https : http;

    const req = client.request({
      hostname: parsed.hostname,
      port: parsed.port || (isHttps ? 443 : 80),
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(bodyStr),
      },
    }, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const json = JSON.parse(data);
            resolve(json.d !== undefined ? json.d : json);
          } catch { resolve(data); }
        } else {
          let msg = `HTTP ${res.statusCode}`;
          try {
            const json = JSON.parse(data);
            const detail = json.ExceptionMessage || json.Message || json.error || json.message;
            if (detail) msg += `: ${detail}`;
          } catch { if (data) msg += `: ${data.substring(0, 200)}`; }
          reject(new Error(msg));
        }
      });
    });

    req.on('error', err => reject(new Error(`Network error: ${err.message}`)));
    req.write(bodyStr);
    req.end();
  });
}

let cachedSessionId = null;
let currentUrl = loadAppSettings().lastUrl ?? TA_URL;
let actionBarVisible = loadAppSettings().actionBarVisible ?? false;
let minimizeTo     = loadAppSettings().minimizeTo ?? 'bubble';
let showInTaskbar  = loadAppSettings().showInTaskbar ?? false;
let bubbleWin      = null;
let fadeTimer      = null;

// Session is always sourced from the webview (the user is already authenticated there).
// If the session is expired, the webview will automatically show the Microsoft login prompt.
// After signing in, the user retries submission and the fresh SESSION_ID is picked up.
function ensureSession() {
  if (!cachedSessionId) throw new Error('SESSION_EXPIRED');
  return cachedSessionId;
}

async function getProcessId(sessionId, processName) {
  const processes = await taPost(sdkEndpoint('ProcessService', 'GetProcessesSummary'), {
    sessionId,
    processesSummaryFilter: { AccessType: 1, UseProcessType: true, ProcessType: 0 },
  });
  const match = processes.find(p => p.Name === processName);
  if (!match) throw new Error(`Process "${processName}" not found in TotalAgility.`);
  return match.Id;
}

// Calls JobService.CreateJobWithDocuments — the TA SDK equivalent of:
//   JobService.CreateJob(sessionId, processIdentity, jobInitialization)
// FolderFields is always included (may be empty) as required by the TA SDK schema.
async function createJobWithDocument(sessionId, processId, fileBase64, mimeType) {
  return taPost(sdkEndpoint('JobService', 'CreateJobWithDocuments'), {
    sessionId,
    processIdentity: { Id: processId },
    jobWithDocsInitialization: {
      RuntimeDocumentCollection: [{ Base64Data: fileBase64, MimeType: mimeType }],
      InputVariables: [],
      FolderFields: [],
    },
  });
}

// ─────────────────────────────────────────────────────────────────────────────

// ─── Saved URL storage ───────────────────────────────────────────────────────

function urlsFile() {
  return path.join(app.getPath('userData'), 'saved-urls.json');
}

function loadUrls() {
  try { return JSON.parse(fs.readFileSync(urlsFile(), 'utf8')); }
  catch { return [{ name: 'Debug Form', url: TA_URL }]; }
}

function saveUrls(urls) {
  fs.writeFileSync(urlsFile(), JSON.stringify(urls, null, 2), 'utf8');
}

// ─── App settings storage ────────────────────────────────────────────────────

function settingsFile() {
  return path.join(app.getPath('userData'), 'settings.json');
}

function loadAppSettings() {
  try { return JSON.parse(fs.readFileSync(settingsFile(), 'utf8')); }
  catch { return {}; }
}

function saveAppSettings(patch) {
  const current = loadAppSettings();
  fs.writeFileSync(settingsFile(), JSON.stringify({ ...current, ...patch }, null, 2), 'utf8');
}

// ─────────────────────────────────────────────────────────────────────────────

app.setName('TA Desktop Agent');

// Prevent the app from showing in the taskbar
app.dock && app.dock.hide();

let tray = null;
let win = null;

const WINDOW_WIDTH = 480;
const WINDOW_WIDTH_EXPANDED = 960;
const WINDOW_HEIGHT = 720;
const MARGIN = 12;
let windowExpanded = false;

function getWindowPosition() {
  const trayBounds = tray.getBounds();
  const display = screen.getDisplayNearestPoint({ x: trayBounds.x, y: trayBounds.y });
  const workArea = display.workArea;

  // Position window on the right side of the screen, near the taskbar
  const x = workArea.x + workArea.width - WINDOW_WIDTH - MARGIN;
  const y = workArea.y + workArea.height - WINDOW_HEIGHT - MARGIN;

  return { x, y };
}

function getBubblePosition() {
  const s = loadAppSettings();
  if (s.bubbleX != null && s.bubbleY != null) return { x: s.bubbleX, y: s.bubbleY };
  const wa = screen.getPrimaryDisplay().workArea;
  return { x: wa.x + wa.width - 72, y: wa.y + wa.height - 72 };
}

function createBubbleWindow() {
  const { x, y } = getBubblePosition();
  bubbleWin = new BrowserWindow({
    width: 72, height: 72, x, y,
    frame: false, transparent: true, backgroundColor: '#00000000', resizable: false,
    skipTaskbar: true, alwaysOnTop: true,
    show: false,
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      nodeIntegration: false, contextIsolation: true,
      preload: path.join(__dirname, 'bubble-preload.js'),
    },
  });
  bubbleWin.loadFile('bubble.html');
  bubbleWin.setAlwaysOnTop(true, 'floating');
  bubbleWin.on('moved', () => {
    if (!bubbleWin || bubbleWin.isDestroyed()) return;
    const [bx, by] = bubbleWin.getPosition();
    saveAppSettings({ bubbleX: bx, bubbleY: by });
  });
  bubbleWin.on('closed', () => { bubbleWin = null; });
}

function showBubble() {
  if (minimizeTo !== 'bubble') return;
  if (!bubbleWin || bubbleWin.isDestroyed()) createBubbleWindow();
  bubbleWin.setIgnoreMouseEvents(false);
  if (!bubbleWin.isVisible()) bubbleWin.show();
  bubbleWin.setOpacity(1);
}

function hideBubble() {
  if (!bubbleWin || bubbleWin.isDestroyed()) return;
  bubbleWin.setOpacity(0);
  bubbleWin.setIgnoreMouseEvents(true);
}

function hideMain() {
  if (!win || win.isDestroyed() || !win.isVisible()) { showBubble(); return; }
  if (fadeTimer) { clearInterval(fadeTimer); fadeTimer = null; }
  let opacity = win.getOpacity();
  fadeTimer = setInterval(() => {
    opacity = Math.max(0, opacity - 0.12);
    if (win && !win.isDestroyed()) win.setOpacity(opacity);
    if (opacity <= 0) {
      clearInterval(fadeTimer); fadeTimer = null;
      if (win && !win.isDestroyed()) win.hide();
      showBubble();
    }
  }, 16);
}

function restoreMain() {
  hideBubble();
  const { x, y } = getWindowPosition();
  if (!win || win.isDestroyed()) createWindow(x, y);

  if (fadeTimer) { clearInterval(fadeTimer); fadeTimer = null; }

  win.setOpacity(0);
  win.show();
  if (win.isMinimized()) win.restore();
  win.setPosition(x, y, false);
  win.setAlwaysOnTop(true, 'floating');
  win.setIcon(appIcon());
  win.focus();

  let opacity = 0;
  fadeTimer = setInterval(() => {
    opacity = Math.min(1, opacity + 0.12);
    if (win && !win.isDestroyed()) win.setOpacity(opacity);
    if (opacity >= 1) { clearInterval(fadeTimer); fadeTimer = null; }
  }, 16);
}

function createWindow(startX, startY) {
  const pos = (startX != null && startY != null) ? { x: startX, y: startY } : {};
  win = new BrowserWindow({
    width: WINDOW_WIDTH,
    height: WINDOW_HEIGHT,
    ...pos,
    show: false,
    frame: false,
    resizable: false,
    skipTaskbar: !showInTaskbar,
    alwaysOnTop: true,
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      webviewTag: true,
    },
    title: 'TA Desktop Agent',
  });

  win.setAlwaysOnTop(true, 'floating');
  Menu.setApplicationMenu(null);
  win.loadFile('index.html');

  win.webContents.on('did-finish-load', () => {
    win.webContents.send('set-action-bar', actionBarVisible);
  });

  win.webContents.on('will-navigate', (e) => e.preventDefault());
  win.webContents.on('will-redirect', (e) => e.preventDefault());

  win.on('close', (e) => {
    e.preventDefault();
    hideMain();
  });

  win.on('minimize', () => {
    hideMain();
  });

}

const appIcon = () => nativeImage.createFromPath(path.join(__dirname, 'icon.ico'));

function toggleWindow() {
  if (win && !win.isDestroyed() && win.isVisible()) {
    hideMain();
  } else {
    restoreMain();
  }
}

ipcMain.handle('get-form-url', () => currentUrl);

ipcMain.handle('get-saved-urls', () => loadUrls());

ipcMain.handle('save-url', (_e, entry) => {
  const urls = loadUrls();
  urls.push(entry);
  saveUrls(urls);
  buildTrayMenu();
});

ipcMain.handle('delete-url', (_e, index) => {
  const urls = loadUrls();
  urls.splice(index, 1);
  saveUrls(urls);
  buildTrayMenu();
});

ipcMain.handle('update-url', (_e, index, entry) => {
  const urls = loadUrls();
  urls[index] = entry;
  saveUrls(urls);
  buildTrayMenu();
});

ipcMain.handle('navigate-to', (_e, url) => {
  currentUrl = url;
  saveAppSettings({ lastUrl: url });
  if (win && !win.isDestroyed()) win.webContents.send('navigate-to', url);
});

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

ipcMain.handle('submit-word-document', async (_event, documentInfo) => {
  if (!config.TA_API_BASE_URL) {
    return { ok: false, message: 'TA_API_BASE_URL not configured in ta.config.local.js.' };
  }

  // Use SESSION_ID and SDK URL extracted from the already-authenticated webview
  if (documentInfo.sessionId) cachedSessionId = documentInfo.sessionId;
  if (documentInfo.sdkUrl) config.TA_SDK_BASE_URL = documentInfo.sdkUrl;

  let fileBase64;
  try {
    fileBase64 = fs.readFileSync(documentInfo.fullName).toString('base64');
  } catch (e) {
    return { ok: false, message: `Could not read document: ${e.message}` };
  }

  const ext = path.extname(documentInfo.fullName).toLowerCase();
  const mimeType = ext === '.docx'
    ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    : ext === '.doc' ? 'application/msword' : 'application/octet-stream';

  async function attempt() {
    const sessionId = await ensureSession();
    const processId = await getProcessId(sessionId, config.TA_PROCESS_NAME);
    const job = await createJobWithDocument(sessionId, processId, fileBase64, mimeType);
    const jobId = job.Id;
    const workflowUrl = `${config.TA_API_BASE_URL}/forms/sene/SENE-ManageWorkflow.form?IN_JobID=${jobId}&`;
    return { ok: true, jobId, workflowUrl };
  }

  try {
    return await attempt();
  } catch (e) {
    const isSessionError = e.message === 'SESSION_EXPIRED' ||
      e.message.includes('401') ||
      e.message.toLowerCase().includes('invalid session') ||
      e.message.toLowerCase().includes('session id') ||
      e.message.toLowerCase().includes('timed out');

    if (isSessionError) {
      cachedSessionId = null;
      return {
        ok: false,
        message: 'Your session has expired. Sign in using the prompt in the app, then try again.',
      };
    }
    return { ok: false, message: e.message };
  }
});

ipcMain.handle('toggle-expand', () => {
  if (!win || win.isDestroyed()) return;

  const [startX, startY] = win.getPosition();
  const startWidth = win.getSize()[0];

  let targetX, targetWidth;
  if (windowExpanded) {
    targetWidth = WINDOW_WIDTH;
    targetX = startX + (WINDOW_WIDTH_EXPANDED - WINDOW_WIDTH);
    windowExpanded = false;
  } else {
    targetWidth = WINDOW_WIDTH_EXPANDED;
    targetX = Math.max(0, startX - (WINDOW_WIDTH_EXPANDED - WINDOW_WIDTH));
    windowExpanded = true;
  }

  const DURATION = 260;
  const TICK = 16;
  const steps = Math.ceil(DURATION / TICK);
  let step = 0;

  const timer = setInterval(() => {
    step++;
    const t = step / steps;
    const ease = t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
    const x = Math.round(startX + (targetX - startX) * ease);
    const width = Math.round(startWidth + (targetWidth - startWidth) * ease);
    if (win && !win.isDestroyed()) win.setBounds({ x, y: startY, width, height: WINDOW_HEIGHT });
    if (step >= steps) {
      clearInterval(timer);
      if (win && !win.isDestroyed()) win.setBounds({ x: targetX, y: startY, width: targetWidth, height: WINDOW_HEIGHT });
    }
  }, TICK);

  return { expanded: windowExpanded };
});

ipcMain.handle('restore-main', () => restoreMain());
ipcMain.handle('minimize-window', () => hideMain());
ipcMain.handle('close-window',    () => hideMain());

ipcMain.handle('open-external', (_event, url) => shell.openExternal(url));

ipcMain.handle('read-file', (_event, filePath) => {
  return fs.readFileSync(filePath).toString('base64');
});

ipcMain.handle('import-outlook-email', () => {
  return new Promise((resolve) => {
    const script = `
try {
  $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  Write-Output '{"status":"not_running"}'; exit
}
try {
  $sel = $ol.ActiveExplorer().Selection
  if ($sel.Count -eq 0) { Write-Output '{"status":"none"}'; exit }
  $item = $sel.Item(1)
  $safe = $item.Subject -replace '[\\\\/:*?"<>|]', '_'
  $tmp  = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $safe + '.msg')
  $item.SaveAs($tmp, 3)
  Write-Output (ConvertTo-Json @{ status='ok'; path=$tmp; subject=$item.Subject } -Compress)
} catch {
  Write-Output ('{"status":"error","message":"' + ($_.Exception.Message -replace '"',"'") + '"}')
}
`;
    const ps = spawn('powershell', ['-NonInteractive', '-NoProfile', '-Command', script]);
    let out = '';
    ps.stdout.on('data', d => { out += d.toString(); });
    ps.on('close', () => {
      let result;
      try { result = JSON.parse(out.trim()); }
      catch { result = { status: 'error', message: 'Could not parse PowerShell response.' }; }

      if (result.status === 'ok') {
        const taHostname = new URL(TA_URL).hostname;
        const webviewWC = webContents.getAllWebContents().find(wc =>
          wc.id !== win?.webContents.id &&
          wc.getURL().includes(taHostname)
        );
        if (webviewWC) {
          webviewWC.send('outlook-email-ready', { path: result.path, subject: result.subject });
        }
      }

      resolve(result);
    });
  });
});

ipcMain.handle('open-powerpdf', (_event, name = 'Power PDF Business') => {
  const dir = 'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Tungsten Power PDF Business';
  const shortcut = path.join(dir, `${name}.lnk`);
  if (!fs.existsSync(shortcut)) return { ok: false, message: `Shortcut not found: ${name}` };
  shell.openPath(shortcut);
  return { ok: true };
});


let settingsWin = null;

function openSettingsWindow() {
  if (settingsWin && !settingsWin.isDestroyed()) { settingsWin.focus(); return; }
  settingsWin = new BrowserWindow({
    width: 460,
    height: 580,
    resizable: false,
    title: 'TA Desktop Agent — Manage URLs',
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });
  settingsWin.loadFile('settings.html');
  Menu.setApplicationMenu(null);
  settingsWin.on('closed', () => { settingsWin = null; });
}

function buildTrayMenu() {
  const urls = loadUrls();
  const urlItems = urls.map((entry, i) => ({
    label: entry.name,
    click: () => {
      currentUrl = entry.url;
      saveAppSettings({ lastUrl: entry.url });
      if (win && !win.isDestroyed()) win.webContents.send('navigate-to', entry.url);
      if (!win || win.isDestroyed() || !win.isVisible()) toggleWindow();
    },
  }));

  const menu = Menu.buildFromTemplate([
    ...urlItems,
    ...(urlItems.length ? [{ type: 'separator' }] : []),
    { label: 'Manage URLs…', click: openSettingsWindow },
    { label: actionBarVisible ? 'Hide Toolbar' : 'Show Toolbar', click: () => {
        actionBarVisible = !actionBarVisible;
        saveAppSettings({ actionBarVisible });
        if (win && !win.isDestroyed()) win.webContents.send('set-action-bar', actionBarVisible);
        buildTrayMenu();
    }},
    { label: 'Show in taskbar', type: 'checkbox', checked: showInTaskbar, click: () => {
        showInTaskbar = !showInTaskbar;
        saveAppSettings({ showInTaskbar });
        if (win && !win.isDestroyed()) win.setSkipTaskbar(!showInTaskbar);
        buildTrayMenu();
    }},
    { label: 'Minimise to', submenu: [
        { label: 'Floating icon', type: 'radio', checked: minimizeTo === 'bubble', click: () => {
            minimizeTo = 'bubble';
            saveAppSettings({ minimizeTo });
            buildTrayMenu();
        }},
        { label: 'System tray', type: 'radio', checked: minimizeTo === 'tray', click: () => {
            minimizeTo = 'tray';
            hideBubble();
            saveAppSettings({ minimizeTo });
            buildTrayMenu();
        }},
    ]},
    { type: 'separator' },
    { label: 'Reload',         click: () => { if (win && !win.isDestroyed()) win.webContents.reload(); } },
    { label: 'Open in Browser', click: () => shell.openExternal(currentUrl) },
    { label: 'Dev Tools',      click: () => { if (win && !win.isDestroyed()) win.webContents.openDevTools({ mode: 'detach' }); } },
    { type: 'separator' },
    { label: 'Exit', click: () => app.exit(0) },
  ]);

  if (tray) tray.setContextMenu(menu);
}

app.on('web-contents-created', (_e, wc) => {
  wc.on('will-attach-webview', (_ev, webPreferences) => {
    webPreferences.preload = path.join(__dirname, 'webview-preload.js');
  });
});

app.whenReady().then(() => {
  const iconPath = path.join(__dirname, 'icon.ico');
  const trayIcon = nativeImage.createFromPath(iconPath);
  tray = new Tray(trayIcon);

  tray.setToolTip('TA Desktop Agent');
  tray.on('click', toggleWindow);
  tray.on('double-click', toggleWindow);

  createBubbleWindow();
  buildTrayMenu();
  toggleWindow();
});

app.on('window-all-closed', (e) => {
  // Don't quit when all windows are closed — stay in tray
  e.preventDefault();
});
