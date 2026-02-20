import type { CellSnapshot, GridCellValue, LedgerTarget, SaveRequest } from '../types';

const DRIVE_API_BASE = 'https://www.googleapis.com/drive/v3';
const SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

const LEDGER_FILE_NAME = 'Beanfolio';
const ARCHIVE_SHEET_NAME = 'Archive';
const BLANK_GAP_ROWS = 4;

interface DriveFile {
  id: string;
  name: string;
  modifiedTime: string;
}

interface DriveFilesListResponse {
  files: DriveFile[];
}

interface SpreadsheetSheet {
  properties: {
    sheetId: number;
    title: string;
  };
}

interface SpreadsheetMetadata {
  spreadsheetId: string;
  properties: {
    title: string;
  };
  sheets: SpreadsheetSheet[];
}

interface GoogleApiError {
  error?: {
    message?: string;
  };
}

interface ApiRequestInit extends Omit<RequestInit, 'headers'> {
  headers?: Record<string, string>;
}

export async function findOrCreateMiniLedger(accessToken: string): Promise<LedgerTarget> {
  const existingFile = await findMostRecentlyModifiedLedger(accessToken);

  if (!existingFile) {
    return createMiniLedger(accessToken);
  }

  const metadata = await getSpreadsheetMetadata(accessToken, existingFile.id);
  const archiveSheet = await ensureArchiveSheet(accessToken, metadata);

  return {
    spreadsheetId: metadata.spreadsheetId,
    spreadsheetTitle: metadata.properties.title,
    sheetId: archiveSheet.sheetId,
    sheetTitle: archiveSheet.title,
  };
}

export async function appendLedgerBlock(
  accessToken: string,
  ledger: LedgerTarget,
  request: SaveRequest,
): Promise<void> {
  const rows = buildAppendRows(request.label ?? '', request.cells, request.columnCount);

  if (rows.length === 0) {
    return;
  }

  await googleApiFetch(`${SHEETS_API_BASE}/${ledger.spreadsheetId}:batchUpdate`, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          appendCells: {
            sheetId: ledger.sheetId,
            rows,
            fields: 'userEnteredValue,note',
          },
        },
      ],
    }),
  });
}

async function findMostRecentlyModifiedLedger(accessToken: string): Promise<DriveFile | null> {
  const query = [
    `name = '${LEDGER_FILE_NAME.replace(/'/g, "\\'")}'`,
    `mimeType = 'application/vnd.google-apps.spreadsheet'`,
    'trashed = false',
  ].join(' and ');

  const params = new URLSearchParams({
    q: query,
    orderBy: 'modifiedTime desc',
    fields: 'files(id,name,modifiedTime)',
    pageSize: '10',
  });

  const response = await googleApiFetch<DriveFilesListResponse>(
    `${DRIVE_API_BASE}/files?${params.toString()}`,
    accessToken,
  );

  return response.files?.[0] ?? null;
}

async function getSpreadsheetMetadata(
  accessToken: string,
  spreadsheetId: string,
): Promise<SpreadsheetMetadata> {
  const params = new URLSearchParams({
    fields: 'spreadsheetId,properties.title,sheets.properties(sheetId,title)',
  });

  return googleApiFetch<SpreadsheetMetadata>(
    `${SHEETS_API_BASE}/${spreadsheetId}?${params.toString()}`,
    accessToken,
  );
}

async function ensureArchiveSheet(
  accessToken: string,
  metadata: SpreadsheetMetadata,
): Promise<{ sheetId: number; title: string }> {
  const existing = metadata.sheets.find((sheet) => sheet.properties.title === ARCHIVE_SHEET_NAME);

  if (existing) {
    return {
      sheetId: existing.properties.sheetId,
      title: existing.properties.title,
    };
  }

  const response = await googleApiFetch<{
    replies: Array<{
      addSheet?: {
        properties: {
          sheetId: number;
          title: string;
        };
      };
    }>;
  }>(`${SHEETS_API_BASE}/${metadata.spreadsheetId}:batchUpdate`, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          addSheet: {
            properties: {
              title: ARCHIVE_SHEET_NAME,
            },
          },
        },
      ],
    }),
  });

  const properties = response.replies[0]?.addSheet?.properties;

  if (!properties) {
    throw new Error('Unable to create Archive sheet.');
  }

  return {
    sheetId: properties.sheetId,
    title: properties.title,
  };
}

async function createMiniLedger(accessToken: string): Promise<LedgerTarget> {
  const created = await googleApiFetch<SpreadsheetMetadata>(SHEETS_API_BASE, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      properties: {
        title: LEDGER_FILE_NAME,
      },
      sheets: [
        {
          properties: {
            title: ARCHIVE_SHEET_NAME,
          },
        },
      ],
    }),
  });

  const archiveSheet = created.sheets.find((sheet) => sheet.properties.title === ARCHIVE_SHEET_NAME);

  if (!archiveSheet) {
    throw new Error('Created spreadsheet is missing Archive sheet.');
  }

  return {
    spreadsheetId: created.spreadsheetId,
    spreadsheetTitle: created.properties.title,
    sheetId: archiveSheet.properties.sheetId,
    sheetTitle: archiveSheet.properties.title,
  };
}

function buildAppendRows(
  label: string,
  cells: CellSnapshot[][],
  columnCount: number,
): Array<{ values: Array<Record<string, unknown>> }> {
  const rows: Array<{ values: Array<Record<string, unknown>> }> = [];
  const safeColumnCount = Math.max(1, columnCount);

  for (let i = 0; i < BLANK_GAP_ROWS; i += 1) {
    rows.push(buildBlankRow(safeColumnCount));
  }

  const trimmedLabel = label.trim();
  if (trimmedLabel.length > 0) {
    const labelRow = buildBlankRow(safeColumnCount);
    labelRow.values[0] = {
      userEnteredValue: {
        stringValue: trimmedLabel,
      },
    };
    rows.push(labelRow);
  }

  cells.forEach((row) => {
    rows.push({
      values: row.map((cell) => toGoogleCellData(cell)),
    });
  });

  return rows;
}

function buildBlankRow(columnCount: number): { values: Array<Record<string, unknown>> } {
  return {
    values: Array.from({ length: columnCount }, () => ({})),
  };
}

function toGoogleCellData(cell: CellSnapshot): Record<string, unknown> {
  const data: Record<string, unknown> = {};
  const userEnteredValue = toUserEnteredValue(cell.displayValue);

  if (userEnteredValue) {
    data.userEnteredValue = userEnteredValue;
  }

  if (cell.formula) {
    data.note = `Formula: ${cell.formula}`;
  }

  return data;
}

function toUserEnteredValue(value: GridCellValue): Record<string, unknown> | null {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      return null;
    }

    return {
      numberValue: value,
    };
  }

  if (typeof value === 'boolean') {
    return {
      boolValue: value,
    };
  }

  const trimmed = String(value).trim();
  if (!trimmed) {
    return null;
  }

  return {
    stringValue: String(value),
  };
}

async function googleApiFetch<T>(url: string, accessToken: string, init: ApiRequestInit = {}): Promise<T> {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...init.headers,
    },
  });

  if (!response.ok) {
    const errorPayload = (await response.json().catch(() => null)) as GoogleApiError | null;
    const message = errorPayload?.error?.message ?? `Google API request failed (${response.status}).`;
    throw new Error(message);
  }

  return (await response.json()) as T;
}
