const { app, BrowserWindow, Tray, nativeImage, screen, Menu } = require('electron');
const path = require('path');

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
    },
    title: 'TA Launchpad',
  });

  win.setAlwaysOnTop(true, 'floating');
  Menu.setApplicationMenu(null);
  win.loadURL('https://demo1-dev.dmoeutta.dev.tungstencloud.com/forms/sene/TADesktop.form');

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
