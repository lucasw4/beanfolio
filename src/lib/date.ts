import type { PresetType } from '../types';

export function formatShortLocalDateTime(date = new Date()): string {
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(date);
}

export function buildDefaultLabel(preset: PresetType = 'blank'): string {
  const prefix =
    preset === 'blank'
      ? 'Blank Sheet'
      : preset === 'journal_entry'
      ? 'Journal Entry'
      : preset === 't_account'
        ? 'T-Account'
        : 'Ledger Entry';

  return `${prefix} - ${formatShortLocalDateTime()}`;
}
