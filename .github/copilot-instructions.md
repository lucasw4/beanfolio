# Beanfolio – Copilot Instructions

## Commands

```bash
npm run dev          # Electron + Vite dev (hot-reload, menu bar shows "Electron")
npm run dev:web      # Web-only dev server at http://localhost:5173
npm run build        # tsc -b && vite build → dist/
npm run desktop      # Build + launch packaged Beanfolio.app (menu bar shows "Beanfolio")
npm test             # Run all tests once (vitest run)
npm run test:watch   # Vitest in watch mode
```

Run a single test file:
```bash
npx vitest run test/grid.test.ts
```

## Architecture

The app is a macOS desktop app (Electron) wrapping a React + Vite SPA. There is no backend — all persistence goes through the Google Sheets API using short-lived OAuth tokens.

**Two rendering contexts:**
- **Dev:** Electron loads the Vite dev server at `http://localhost:5173` (`VITE_DEV_SERVER_URL` env var).
- **Prod:** Electron spawns its own static HTTP server on `127.0.0.1:5180` serving `dist/`.

**Key layers:**
- `src/App.tsx` — the entire UI lives here: grid, toolbar, formula helper, Google auth flow, save workflow. It is intentionally a single large component.
- `src/lib/` — pure utility modules (no React):
  - `grid.ts` — grid creation and value normalization helpers.
  - `cellFormat.ts` — cell style store (module-level `Map`, not React state) + color palette.
  - `date.ts` — label/timestamp helpers.
- `src/services/` — Google API wrappers:
  - `googleAuth.ts` — loads the Google Identity Services script and manages the implicit OAuth token flow (popup, no server redirect).
  - `googleLedger.ts` — Drive file discovery and Sheets `batchUpdate` append logic.
- `src/presets.ts` — Blank / Journal Entry / T-Account grid templates.
- `src/types.ts` — shared constants (`GRID_ROWS = 120`, `GRID_COLUMNS = 12`) and core interfaces.
- `electron/main.cjs` — Electron main process (CommonJS). Manages the window, static server, and IPC.
- `electron/preload.cjs` — Electron preload (sandbox + context isolation enabled).
- `test/` — Vitest tests, `environment: 'node'`, importing directly from `src/`.

## Key Conventions

**Cell style storage:** `setCellStyle` / `getCellStyle` in `cellFormat.ts` use a module-level `Map`, not React state. Clearing styles on preset load requires calling `clearAllStyles()` explicitly.

**Grid dimensions are constants:** Always import `GRID_ROWS` and `GRID_COLUMNS` from `src/types.ts`; never hard-code `120`/`12`.

**Google auth is token-only, never persisted:** `requestGoogleAccessToken` returns a short-lived access token. There is no refresh token, no local storage, no session. Re-auth is required each session.

**Google Sheets save format:** `appendLedgerBlock` always prepends 4 blank gap rows (`BLANK_GAP_ROWS`), then an optional label row, then the trimmed data block (trailing empty rows stripped). Formulas are stored as cell notes, not re-evaluated in Sheets.

**Electron files are CommonJS (`.cjs`):** `electron/main.cjs` and `electron/preload.cjs` use `require`/`module.exports`. The rest of the project uses ESM (`"type": "module"`).

**Window sizing via IPC:** The sidebar open/close state is sent from the renderer via `ipcMain.on('set-sidebar-open', ...)`, which resizes the `BrowserWindow` to `BASE_WIDTH + SIDEBAR_WIDTH` (720px) or `BASE_WIDTH` (500px).

**Required env var:** `VITE_GOOGLE_CLIENT_ID` must be set in `.env` (copy `.env.example`). The Google Cloud OAuth app must allow `http://localhost:5173` and `http://127.0.0.1:5180` as authorized JavaScript origins.
