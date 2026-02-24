const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('beanfolioDesktop', {
  isDesktop: true,
  platform: process.platform,
  setSidebarOpen: (isOpen) => ipcRenderer.send('set-sidebar-open', isOpen),
  setAlwaysOnTop: (isPinned) => ipcRenderer.send('set-always-on-top', isPinned),
});
