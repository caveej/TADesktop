const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  restoreMain: () => ipcRenderer.invoke('restore-main'),
});
