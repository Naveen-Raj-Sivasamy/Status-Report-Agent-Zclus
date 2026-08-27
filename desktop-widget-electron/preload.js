const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('statusApp', {
  getTabs: () => ipcRenderer.invoke('get-tabs'),
  getColumns: (tab) => ipcRenderer.invoke('get-columns', tab),
  submitEntry: (tab, values) => ipcRenderer.invoke('submit-entry', { tab, values }),
  sendReportNow: (range) => ipcRenderer.invoke('send-report-now', range),
  downloadReport: (range) => ipcRenderer.invoke('download-report', range),
  getNextNumber: (tab, column) => ipcRenderer.invoke('get-next-number', { tab, column }),
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  checkLatestVersion: () => ipcRenderer.invoke('check-latest-version'),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  getYourName: () => ipcRenderer.invoke('get-your-name'),
  hideWindow: () => ipcRenderer.invoke('hide-window'),
  openSheet: () => ipcRenderer.invoke('open-sheet'),
  fabClicked: () => ipcRenderer.invoke('fab-clicked'),
  showFabMenu: () => ipcRenderer.invoke('show-fab-menu'),
  saveYourName: (name) => ipcRenderer.invoke('save-your-name', name),
  onOpened: (callback) => ipcRenderer.on('opened', callback),
  resizeWindow: (height) => ipcRenderer.send('resize-window', height),
});
