const { contextBridge, ipcRenderer } = require('electron');

// 安全地暴露API给渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 获取应用版本
  getVersion: () => ipcRenderer.invoke('get-version'),

  // 选择文件
  selectFile: () => ipcRenderer.invoke('select-file'),

  // 获取平台信息
  platform: process.platform
});