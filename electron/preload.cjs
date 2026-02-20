const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('beanfolioDesktop', {
  isDesktop: true,
  platform: process.platform,
  setSidebarOpen: (isOpen) => ipcRenderer.send('set-sidebar-open', isOpen),
});
