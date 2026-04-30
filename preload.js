const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  detectWord: () => ipcRenderer.invoke('detect-word'),
  getFormUrl: () => ipcRenderer.invoke('get-form-url'),
  submitWordDocument: (documentInfo) => ipcRenderer.invoke('submit-word-document', documentInfo),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  openPowerPDF: (name) => ipcRenderer.invoke('open-powerpdf', name),
  toggleExpand: () => ipcRenderer.invoke('toggle-expand'),
});
