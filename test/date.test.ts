import { describe, expect, it } from 'vitest';
import { buildDefaultLabel, formatShortLocalDateTime } from '../src/lib/date';

describe('date label helpers', () => {
  it('formats local date/time into a non-empty string', () => {
    const formatted = formatShortLocalDateTime(new Date('2026-02-18T18:30:00.000Z'));
    expect(typeof formatted).toBe('string');
    expect(formatted.length).toBeGreaterThan(0);
  });

  it('builds preset-specific default label prefixes', () => {
    expect(buildDefaultLabel('blank')).toMatch(/^Blank Sheet - /);
    expect(buildDefaultLabel('journal_entry')).toMatch(/^Journal Entry - /);
    expect(buildDefaultLabel('t_account')).toMatch(/^T-Account - /);
  });
});
