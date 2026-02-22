const GOOGLE_IDENTITY_SCRIPT = 'https://accounts.google.com/gsi/client';

declare global {
  interface Window {
    google?: {
      accounts: {
        oauth2: {
          initTokenClient: (config: {
            client_id: string;
            scope: string;
            ux_mode?: 'popup' | 'redirect';
            callback: (response: TokenResponse) => void;
            error_callback?: (error: { type: string }) => void;
          }) => TokenClient;
          revoke: (token: string, done: () => void) => void;
        };
      };
    };
  }
}

interface TokenClient {
  requestAccessToken: (options?: { prompt?: string }) => void;
}

interface TokenResponse {
  access_token?: string;
  error?: string;
  error_description?: string;
}

let loadPromise: Promise<void> | null = null;

export function loadGoogleIdentityScript(): Promise<void> {
  if (window.google?.accounts?.oauth2) {
    return Promise.resolve();
  }

  if (loadPromise) {
    return loadPromise;
  }

  loadPromise = new Promise((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>(
      `script[src="${GOOGLE_IDENTITY_SCRIPT}"]`,
    );

    if (existing) {
      if (window.google?.accounts?.oauth2) {
        resolve();
        return;
      }

      const timeout = window.setTimeout(() => {
        reject(new Error('Timed out while waiting for Google script.'));
      }, 10000);

      existing.addEventListener(
        'load',
        () => {
          window.clearTimeout(timeout);
          resolve();
        },
        { once: true },
      );
      existing.addEventListener(
        'error',
        () => {
          window.clearTimeout(timeout);
          reject(new Error('Failed to load Google script.'));
        },
        {
          once: true,
        },
      );
      return;
    }

    const script = document.createElement('script');
    script.src = GOOGLE_IDENTITY_SCRIPT;
    script.async = true;
    script.defer = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('Failed to load Google script.'));
    document.head.appendChild(script);
  });

  return loadPromise;
}

export async function requestGoogleAccessToken(params: {
  clientId: string;
  scopes: string[];
  prompt?: 'consent' | '';
}): Promise<string> {
  await loadGoogleIdentityScript();

  if (!window.google?.accounts?.oauth2) {
    throw new Error('Google Identity Services is unavailable.');
  }

  return new Promise((resolve, reject) => {
    const tokenClient = window.google!.accounts.oauth2.initTokenClient({
      client_id: params.clientId,
      scope: params.scopes.join(' '),
      ux_mode: 'popup',
      callback: (response) => {
        if (response.error || !response.access_token) {
          reject(new Error(response.error_description ?? response.error ?? 'OAuth token request failed.'));
          return;
        }

        resolve(response.access_token);
      },
      error_callback: (error) => {
        if (error.type === 'popup_closed') {
          reject(new Error('OAuth sign-in window was closed before completion.'));
          return;
        }

        if (error.type === 'popup_failed_to_open') {
          reject(new Error('OAuth sign-in popup failed to open.'));
          return;
        }

        reject(new Error(`OAuth request failed (${error.type}).`));
      },
    });

    tokenClient.requestAccessToken({ prompt: params.prompt ?? 'consent' });
  });
}

export function revokeGoogleToken(accessToken: string): Promise<void> {
  if (!window.google?.accounts?.oauth2) {
    return Promise.resolve();
  }

  return new Promise((resolve) => {
    window.google!.accounts.oauth2.revoke(accessToken, () => resolve());
  });
}
