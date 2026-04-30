const { app, BrowserWindow, Tray, nativeImage, screen, Menu, ipcMain, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const https = require('https');
const http = require('http');
const os = require('os');
const fs = require('fs');

const TA_URL = 'https://demo1-dev.dmoeutta.dev.tungstencloud.com/forms/TADesktop/TADesktopLauncher.form';

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

app.setName('TA Launchpad');

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

function createWindow() {
  win = new BrowserWindow({
    width: WINDOW_WIDTH,
    height: WINDOW_HEIGHT,
    show: false,
    frame: true,
    resizable: false,
    skipTaskbar: false,
    alwaysOnTop: true,
    icon: path.join(__dirname, 'icon.ico'),
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

const appIcon = () => nativeImage.createFromPath(path.join(__dirname, 'icon.ico'));

function toggleWindow() {
  if (!win || win.isDestroyed()) {
    createWindow();
    const { x, y } = getWindowPosition();
    win.setPosition(x, y, false);
    win.setAlwaysOnTop(true, 'floating');
    win.show();
    win.setIcon(appIcon());
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
    win.setIcon(appIcon());
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
  const [x, y] = win.getPosition();
  if (windowExpanded) {
    const newX = x + (WINDOW_WIDTH_EXPANDED - WINDOW_WIDTH);
    win.setBounds({ x: newX, y, width: WINDOW_WIDTH, height: WINDOW_HEIGHT });
    windowExpanded = false;
  } else {
    const newX = Math.max(0, x - (WINDOW_WIDTH_EXPANDED - WINDOW_WIDTH));
    win.setBounds({ x: newX, y, width: WINDOW_WIDTH_EXPANDED, height: WINDOW_HEIGHT });
    windowExpanded = true;
  }
  return { expanded: windowExpanded };
});

ipcMain.handle('open-external', (_event, url) => shell.openExternal(url));

ipcMain.handle('open-powerpdf', (_event, name = 'Power PDF Business') => {
  const dir = 'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Tungsten Power PDF Business';
  const shortcut = path.join(dir, `${name}.lnk`);
  if (!fs.existsSync(shortcut)) return { ok: false, message: `Shortcut not found: ${name}` };
  shell.openPath(shortcut);
  return { ok: true };
});


app.whenReady().then(() => {
  const iconPath = path.join(__dirname, 'icon.ico');
  const trayIcon = nativeImage.createFromPath(iconPath);
  tray = new Tray(trayIcon);

  tray.setToolTip('TA Launchpad');
  tray.on('click', toggleWindow);
  tray.on('double-click', toggleWindow);

  toggleWindow();
});

app.on('window-all-closed', (e) => {
  // Don't quit when all windows are closed — stay in tray
  e.preventDefault();
});
