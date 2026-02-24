/// <reference types="vite/client" />

interface BeanfolioDesktopBridge {
  isDesktop: boolean;
  platform: string;
  setSidebarOpen: (isOpen: boolean) => void;
}

interface Window {
  beanfolioDesktop?: BeanfolioDesktopBridge;
}
