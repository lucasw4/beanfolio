/// <reference types="vite/client" />

interface BeanfolioDesktopBridge {
  isDesktop: boolean;
  platform: string;
  setSidebarOpen: (isOpen: boolean) => void;
  setAlwaysOnTop: (isPinned: boolean) => void;
}

interface Window {
  beanfolioDesktop?: BeanfolioDesktopBridge;
}
