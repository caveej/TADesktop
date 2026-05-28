const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  detectWord: () => ipcRenderer.invoke('detect-word'),
  getFormUrl: () => ipcRenderer.invoke('get-form-url'),
  submitWordDocument: (documentInfo) => ipcRenderer.invoke('submit-word-document', documentInfo),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  openPowerPDF: (name) => ipcRenderer.invoke('open-powerpdf', name),
  toggleExpand:  () => ipcRenderer.invoke('toggle-expand'),
  getSavedUrls:  () => ipcRenderer.invoke('get-saved-urls'),
  saveUrl:       (entry) => ipcRenderer.invoke('save-url', entry),
  deleteUrl:     (index) => ipcRenderer.invoke('delete-url', index),
  updateUrl:     (index, entry) => ipcRenderer.invoke('update-url', index, entry),
  navigateTo:    (url) => ipcRenderer.invoke('navigate-to', url),
  onNavigate:    (cb) => ipcRenderer.on('navigate-to', (_e, url) => cb(url)),
});
