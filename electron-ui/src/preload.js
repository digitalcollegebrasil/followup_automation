// preload.js
import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('api', {
  selectXlsx: () => ipcRenderer.invoke('select-xlsx'),
  getUserDataPath: () => ipcRenderer.invoke('get-user-data-path'),
  saveOutputs: (xlsxBase64, configJSON) => ipcRenderer.invoke('save-outputs', { xlsxBase64, configJSON }),
  runEditar: () => ipcRenderer.invoke('run-editar'),

  // NOVO: logs do processo editar.js
  onEditarLog: (cb) => ipcRenderer.on('editar-log', (_e, line) => cb(line)),
  onEditarExit: (cb) => ipcRenderer.on('editar-exit', (_e, msg) => cb(msg)),
});
