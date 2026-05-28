const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  importOutlookEmail: () => ipcRenderer.invoke('import-outlook-email'),
  readFile:           (filePath) => ipcRenderer.invoke('read-file', filePath),
});

ipcRenderer.on('outlook-email-ready', (_e, detail) => {
  const iframe = document.querySelector('iframe');
  if (iframe && iframe.contentWindow) {
    iframe.contentWindow.postMessage({ type: 'outlook-email-imported', path: detail.path, subject: detail.subject }, '*');
  }
});
