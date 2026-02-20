import { describe, expect, it } from 'vitest';
import { HyperFormula } from 'hyperformula';
import {
  buildPresetGrid,
  getPresetColumnHeaders,
  getPresetColumnWidths,
  getPresetMergeCells,
} from '../src/presets';
import { GRID_COLUMNS } from '../src/types';

function sum(values: number[]): number {
  return values.reduce((total, value) => total + value, 0);
}

describe('presets', () => {
  it('uses A-series column labels for all presets', () => {
    const expected = Array.from({ length: GRID_COLUMNS }, (_, index) => String.fromCharCode(65 + index));
    expect(getPresetColumnHeaders('blank')).toEqual(expected);
    expect(getPresetColumnHeaders('journal_entry')).toEqual(expected);
    expect(getPresetColumnHeaders('t_account')).toEqual(expected);
  });

  it('keeps total width equal for journal and t-account presets', () => {
    const journalWidth = sum(getPresetColumnWidths('journal_entry'));
    const tAccountWidth = sum(getPresetColumnWidths('t_account'));

    expect(journalWidth).toBe(tAccountWidth);
  });

  it('uses wider equal column sizing for blank preset', () => {
    const blankWidths = getPresetColumnWidths('blank');
    const tAccountWidths = getPresetColumnWidths('t_account');

    expect(new Set(blankWidths).size).toBe(1);
    expect(sum(blankWidths)).toBeGreaterThan(sum(tAccountWidths));
  });

  it('builds blank preset with no seeded data', () => {
    const blank = buildPresetGrid('blank');
    expect(blank.length).toBeGreaterThan(0);
    expect(blank.flat().every((value) => value === null)).toBe(true);
  });

  it('builds journal entry preset with row-1 headers and index seed', () => {
    const journal = buildPresetGrid('journal_entry');
    expect(journal[0].slice(0, 4)).toEqual(['#', 'Account Name', 'Debit', 'Credit']);
    expect(journal[1][0]).toBe(1);
  });

  it('builds t-account preset with merged titles and formulas', () => {
    const tAccount = buildPresetGrid('t_account');
    const merges = getPresetMergeCells('t_account');

    expect(merges).toEqual([
      { row: 2, col: 1, rowspan: 1, colspan: 2 },
      { row: 13, col: 1, rowspan: 1, colspan: 2 },
    ]);

    expect(typeof tAccount[11][1]).toBe('string');
    expect(String(tAccount[11][1])).toContain('COUNTA');
    expect(typeof tAccount[22][2]).toBe('string');
    expect(String(tAccount[22][2])).toContain('ABS');
  });

  it('evaluates t-account totals as debit positive and credit negative', () => {
    const source = buildPresetGrid('t_account');
    expect(source[0]).toHaveLength(GRID_COLUMNS);

    const hf = HyperFormula.buildFromArray(source, { licenseKey: 'gpl-v3' });

    expect(hf.getCellValue({ sheet: 0, row: 11, col: 1 })).toBe('');
    expect(hf.getCellValue({ sheet: 0, row: 11, col: 2 })).toBe('');

    hf.setCellContents({ sheet: 0, row: 4, col: 1 }, [[100]]);
    hf.setCellContents({ sheet: 0, row: 5, col: 2 }, [[30]]);

    expect(hf.getCellValue({ sheet: 0, row: 11, col: 1 })).toBe(70);
    expect(hf.getCellValue({ sheet: 0, row: 11, col: 2 })).toBe('');

    hf.setCellContents({ sheet: 0, row: 15, col: 1 }, [[40]]);
    hf.setCellContents({ sheet: 0, row: 16, col: 2 }, [[110]]);

    expect(hf.getCellValue({ sheet: 0, row: 22, col: 1 })).toBe('');
    expect(hf.getCellValue({ sheet: 0, row: 22, col: 2 })).toBe(70);
  });
});
