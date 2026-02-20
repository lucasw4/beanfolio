import { afterEach, describe, expect, it, vi } from 'vitest';
import { appendLedgerBlock, findOrCreateMiniLedger } from '../src/services/googleLedger';
import type { LedgerTarget } from '../src/types';

interface FetchCall {
  url: string;
  init?: RequestInit;
}

function toResponse(payload: unknown, ok = true, status = 200): Response {
  return {
    ok,
    status,
    json: async () => payload,
  } as Response;
}

function normalizeUrl(input: Parameters<typeof fetch>[0]): string {
  if (typeof input === 'string') {
    return input;
  }

  if (input instanceof URL) {
    return input.toString();
  }

  return input.url;
}

function installFetchMock(queue: Response[]) {
  const calls: FetchCall[] = [];
  const mock = vi.fn(
    async (input: Parameters<typeof fetch>[0], init?: Parameters<typeof fetch>[1]) => {
      const next = queue.shift();
      if (!next) {
        throw new Error('No mocked response available for fetch call.');
      }

      calls.push({ url: normalizeUrl(input), init });
      return next;
    },
  );

  vi.stubGlobal('fetch', mock);

  return { calls, mock };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('appendLedgerBlock', () => {
  const ledger: LedgerTarget = {
    spreadsheetId: 'spreadsheet-1',
    spreadsheetTitle: 'Beanfolio',
    sheetId: 9,
    sheetTitle: 'Archive',
  };

  it('serializes gap rows, label row, and data block for append', async () => {
    const { calls } = installFetchMock([toResponse({})]);

    await appendLedgerBlock('token-123', ledger, {
      label: 'Journal Entry - Feb 18, 2026',
      columnCount: 3,
      cells: [
        [
          { displayValue: ' Cash ' },
          { displayValue: 100, formula: '=A1*2' },
          { displayValue: null },
        ],
        [
          { displayValue: true },
          { displayValue: false },
          { displayValue: 42 },
        ],
      ],
    });

    expect(calls).toHaveLength(1);
    expect(calls[0].url).toContain('/spreadsheets/spreadsheet-1:batchUpdate');
    expect(calls[0].init?.method).toBe('POST');
    expect(calls[0].init?.headers).toMatchObject({
      Authorization: 'Bearer token-123',
      'Content-Type': 'application/json',
    });

    const body = JSON.parse(String(calls[0].init?.body));
    const rows = body.requests[0].appendCells.rows as Array<{ values: Array<Record<string, unknown>> }>;

    expect(rows).toHaveLength(7);
    expect(rows.slice(0, 4).every((row) => row.values.length === 3)).toBe(true);
    expect(rows.slice(0, 4).every((row) => row.values.every((cell) => Object.keys(cell).length === 0))).toBe(true);

    expect(rows[4].values[0]).toEqual({
      userEnteredValue: { stringValue: 'Journal Entry - Feb 18, 2026' },
    });
    expect(rows[5].values[0]).toEqual({
      userEnteredValue: { stringValue: ' Cash ' },
    });
    expect(rows[5].values[1]).toEqual({
      userEnteredValue: { numberValue: 100 },
      note: 'Formula: =A1*2',
    });
    expect(rows[5].values[2]).toEqual({});
    expect(rows[6].values[0]).toEqual({
      userEnteredValue: { boolValue: true },
    });
    expect(rows[6].values[1]).toEqual({
      userEnteredValue: { boolValue: false },
    });
    expect(rows[6].values[2]).toEqual({
      userEnteredValue: { numberValue: 42 },
    });
  });

  it('omits label row when label is blank and enforces at least one column', async () => {
    const { calls } = installFetchMock([toResponse({})]);

    await appendLedgerBlock('token-123', ledger, {
      label: '   ',
      columnCount: 0,
      cells: [],
    });

    const body = JSON.parse(String(calls[0].init?.body));
    const rows = body.requests[0].appendCells.rows as Array<{ values: Array<Record<string, unknown>> }>;

    expect(rows).toHaveLength(4);
    expect(rows.every((row) => row.values.length === 1)).toBe(true);
  });

  it('surfaces Google API error messages', async () => {
    installFetchMock([toResponse({ error: { message: 'Permission denied.' } }, false, 403)]);

    await expect(
      appendLedgerBlock('token-123', ledger, {
        label: '',
        columnCount: 3,
        cells: [],
      }),
    ).rejects.toThrow('Permission denied.');
  });
});

describe('findOrCreateMiniLedger', () => {
  it('reuses existing spreadsheet and archive sheet', async () => {
    const { calls } = installFetchMock([
      toResponse({
        files: [{ id: 'spreadsheet-1', name: 'Beanfolio', modifiedTime: '2026-02-19T00:00:00Z' }],
      }),
      toResponse({
        spreadsheetId: 'spreadsheet-1',
        properties: { title: 'Beanfolio' },
        sheets: [{ properties: { sheetId: 9, title: 'Archive' } }],
      }),
    ]);

    const target = await findOrCreateMiniLedger('token-123');

    expect(target).toEqual({
      spreadsheetId: 'spreadsheet-1',
      spreadsheetTitle: 'Beanfolio',
      sheetId: 9,
      sheetTitle: 'Archive',
    });

    expect(calls).toHaveLength(2);
    expect(calls[0].url).toContain('drive/v3/files?');
    expect(calls[1].url).toContain('/spreadsheets/spreadsheet-1?');
  });

  it('adds archive sheet when missing', async () => {
    const { calls } = installFetchMock([
      toResponse({
        files: [{ id: 'spreadsheet-1', name: 'Beanfolio', modifiedTime: '2026-02-19T00:00:00Z' }],
      }),
      toResponse({
        spreadsheetId: 'spreadsheet-1',
        properties: { title: 'Beanfolio' },
        sheets: [{ properties: { sheetId: 1, title: 'Sheet1' } }],
      }),
      toResponse({
        replies: [{ addSheet: { properties: { sheetId: 9, title: 'Archive' } } }],
      }),
    ]);

    const target = await findOrCreateMiniLedger('token-123');
    expect(target.sheetTitle).toBe('Archive');
    expect(target.sheetId).toBe(9);

    expect(calls).toHaveLength(3);
    expect(calls[2].url).toContain('/spreadsheets/spreadsheet-1:batchUpdate');
  });

  it('creates spreadsheet when none exists', async () => {
    const { calls } = installFetchMock([
      toResponse({ files: [] }),
      toResponse({
        spreadsheetId: 'spreadsheet-created',
        properties: { title: 'Beanfolio' },
        sheets: [{ properties: { sheetId: 7, title: 'Archive' } }],
      }),
    ]);

    const target = await findOrCreateMiniLedger('token-123');

    expect(target).toEqual({
      spreadsheetId: 'spreadsheet-created',
      spreadsheetTitle: 'Beanfolio',
      sheetId: 7,
      sheetTitle: 'Archive',
    });

    expect(calls).toHaveLength(2);
    expect(calls[1].url).toBe('https://sheets.googleapis.com/v4/spreadsheets');
    expect(calls[1].init?.method).toBe('POST');
  });
});
