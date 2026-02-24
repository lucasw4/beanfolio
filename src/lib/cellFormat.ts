export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fillColor?: string;
  textColor?: string;
  commaOff?: boolean;
  percent?: boolean;
  borderTop?: number;
  borderRight?: number;
  borderBottom?: number;
  borderLeft?: number;
  borderTopStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
  borderRightStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
  borderBottomStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
  borderLeftStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
}

type CellKey = string;

function key(row: number, col: number): CellKey {
  return `${row},${col}`;
}

const store = new Map<CellKey, CellStyle>();

export function getCellStyle(row: number, col: number): CellStyle | undefined {
  return store.get(key(row, col));
}

export function setCellStyle(row: number, col: number, patch: Partial<CellStyle>): void {
  const k = key(row, col);
  const existing = store.get(k) ?? {};
  const merged = { ...existing, ...patch };

  const isEmpty = !merged.bold && !merged.italic && !merged.underline
    && !merged.fillColor && !merged.textColor && !merged.commaOff && !merged.percent
    && !merged.borderTop && !merged.borderRight && !merged.borderBottom && !merged.borderLeft
    && !merged.borderTopStyle && !merged.borderRightStyle && !merged.borderBottomStyle && !merged.borderLeftStyle;

  if (isEmpty) {
    store.delete(k);
  } else {
    store.set(k, merged);
  }
}

export function clearAllStyles(): void {
  store.clear();
}

export const PRESET_BASE_COLORS = [
  '#ef4444', '#f97316', '#eab308', '#22c55e',
  '#06b6d4', '#3b82f6', '#8b5cf6', '#ec4899',
  '#ffffff', '#a1a1aa', '#52525b', '#18181b',
];

export function formatNumberWithCommas(value: number, decimals?: number): string {
  const opts: Intl.NumberFormatOptions = {
    useGrouping: true,
    maximumFractionDigits: 20,
  };
  if (decimals !== undefined) {
    opts.minimumFractionDigits = decimals;
    opts.maximumFractionDigits = decimals;
  }
  return new Intl.NumberFormat('en-US', opts).format(value);
}
