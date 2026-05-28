const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  importOutlookEmail: () => ipcRenderer.invoke('import-outlook-email'),
  readFile:           (filePath) => ipcRenderer.invoke('read-file', filePath),
});

ipcRenderer.on('outlook-email-ready', (_e, detail) => {
  document.dispatchEvent(new CustomEvent('outlook-email-imported', { detail }));
});
