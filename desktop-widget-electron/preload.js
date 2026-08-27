const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('statusApp', {
  getTabs: () => ipcRenderer.invoke('get-tabs'),
  getColumns: (tab) => ipcRenderer.invoke('get-columns', tab),
  submitEntry: (tab, values) => ipcRenderer.invoke('submit-entry', { tab, values }),
  sendReportNow: () => ipcRenderer.invoke('send-report-now'),
  getYourName: () => ipcRenderer.invoke('get-your-name'),
  hideWindow: () => ipcRenderer.invoke('hide-window'),
  openSheet: () => ipcRenderer.invoke('open-sheet'),
  fabClicked: () => ipcRenderer.invoke('fab-clicked'),
  showFabMenu: () => ipcRenderer.invoke('show-fab-menu'),
  saveYourName: (name) => ipcRenderer.invoke('save-your-name', name),
  onOpened: (callback) => ipcRenderer.on('opened', callback),
});
