const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('PNDA_DESKTOP', {
  isDesktop: true,
  defaultSession: {
    user: 'compte_kasaic@pnda.cd',
    province: 'Kasaï Central',
    role: 'province'
  },
  saveFile: (options) => ipcRenderer.invoke('dialog:save-file', options),
  exportExcelWorkbook: (payload) => ipcRenderer.invoke('excel:export-workbook', payload)
});