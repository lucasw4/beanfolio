import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties, type MutableRefObject } from 'react';
import type Handsontable from 'handsontable';
import { HotTable } from '@handsontable/react';
import { HyperFormula } from 'hyperformula';
import './App.css';
import { buildDefaultLabel } from './lib/date';
import { hasNonEmptyCells, isFormula, normalizeDisplayValue } from './lib/grid';
import {
  getCellStyle,
  setCellStyle,
  clearAllStyles,
  PRESET_BASE_COLORS,
  formatNumberWithCommas,
  type CellStyle,
} from './lib/cellFormat';
import {
  buildPresetGrid,
  getPresetColumnCount,
  getPresetColumnHeaders,
  getPresetMergeCells,
  getPresetColumnWidths,
  getPresetHiddenColumns,
  isTAccountDividerCell,
  isTAccountLabelCell,
  isTAccountSumBorderCell,
  isTAccountTotalCell,
  isTAccountTitleCell,
  isTAccountTopBorderCell,
} from './presets';
import { loadGoogleIdentityScript, requestGoogleAccessToken, revokeGoogleToken } from './services/googleAuth';
import { appendLedgerBlock, findOrCreateMiniLedger } from './services/googleLedger';
import { GRID_COLUMNS, GRID_ROWS, type CellSnapshot, type LedgerTarget, type PresetType } from './types';

const GOOGLE_SCOPES = [
  'openid',
  'profile',
  'email',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.metadata.readonly',
] as const;

const ROW_HEADER_WIDTH = 34;
const DROPDOWN_TRANSITION_MS = 220;
const POP_OUT_APP_WIDTH = 500;
const POP_OUT_APP_HEIGHT = 510;
const SIDEBAR_WIDTH = 220;
const POP_OUT_CHROME_WIDTH = 40;
const POP_OUT_CHROME_HEIGHT = 110;

const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID;

interface FormulaDoc {
  name: string;
  params: string[];
  description: string;
}

interface FormulaSuggestState {
  mode: 'suggest';
  style: CSSProperties;
  query: string;
  suggestions: FormulaDoc[];
}

interface FormulaFunctionState {
  mode: 'function';
  style: CSSProperties;
  exact: FormulaDoc;
  activeParamIndex: number;
}

type FormulaPopupState = FormulaSuggestState | FormulaFunctionState;

interface FormulaEditSession {
  editRow: number;
  editCol: number;
  cursorRow: number;
  cursorCol: number;
  pendingReferenceRange: { start: number; end: number } | null;
  lastReferenceInput: 'keyboard' | 'mouse' | null;
}

type ExportFormat = 'csv' | 'tsv' | 'xlsx' | 'ods';

type OpenEditorContext = {
  editor: Handsontable.editors.BaseEditor & { TEXTAREA?: HTMLTextAreaElement };
  input: string;
  caretPosition: number;
  row: number;
  col: number;
};

const A1_REFERENCE_TOKEN_RE = /^\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$/;
const A1_REFERENCE_PART_RE = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/;
const PALETTE_LIGHTNESS_MIN = -3;
const PALETTE_LIGHTNESS_MAX = 3;

const REFERENCE_COLORS = [
  { bg: 'rgba(37, 99, 235, 0.12)', border: '#2563eb', text: '#2563eb' },
  { bg: 'rgba(220, 38, 38, 0.12)', border: '#dc2626', text: '#dc2626' },
  { bg: 'rgba(124, 58, 237, 0.12)', border: '#7c3aed', text: '#7c3aed' },
  { bg: 'rgba(5, 150, 105, 0.12)', border: '#059669', text: '#059669' },
  { bg: 'rgba(217, 119, 6, 0.12)', border: '#d97706', text: '#d97706' },
  { bg: 'rgba(219, 39, 119, 0.12)', border: '#db2777', text: '#db2777' },
] as const;

interface FormulaReferenceHighlight {
  ref: string;
  row: number;
  col: number;
  colorIndex: number;
}

const FORMULA_DOCS: FormulaDoc[] = [
  {
    name: 'ABS',
    params: ['number'],
    description: 'Absolute value of a number.',
  },
  {
    name: 'AVERAGE',
    params: ['number1', '[number2]', '...'],
    description: 'Arithmetic mean of numbers and ranges.',
  },
  {
    name: 'AVERAGEIFS',
    params: ['average_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    description: 'Average of values that satisfy one or more criteria.',
  },
  {
    name: 'COUNTIFS',
    params: ['criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    description: 'Counts cells that satisfy one or more criteria.',
  },
  {
    name: 'DB',
    params: ['cost', 'salvage', 'life', 'period', '[month]'],
    description: 'Fixed-declining balance depreciation for a period.',
  },
  {
    name: 'DDB',
    params: ['cost', 'salvage', 'life', 'period', '[factor]'],
    description: 'Double-declining (or custom factor) depreciation.',
  },
  {
    name: 'EDATE',
    params: ['start_date', 'months'],
    description: 'Date shifted by a number of months.',
  },
  {
    name: 'EFFECT',
    params: ['nominal_rate', 'npery'],
    description: 'Effective annual interest rate.',
  },
  {
    name: 'EOMONTH',
    params: ['start_date', 'months'],
    description: 'Last day of month shifted by a number of months.',
  },
  {
    name: 'PV',
    params: ['rate', 'nper', 'pmt', '[fv]', '[type]'],
    description: 'Present value of a series of future payments.',
  },
  {
    name: 'FV',
    params: ['rate', 'nper', 'pmt', '[pv]', '[type]'],
    description: 'Future value of an investment or loan.',
  },
  {
    name: 'PMT',
    params: ['rate', 'nper', 'pv', '[fv]', '[type]'],
    description: 'Periodic payment amount for a loan or annuity.',
  },
  {
    name: 'NPV',
    params: ['rate', 'value1', '[value2]', '...'],
    description: 'Net present value of a discounted cash flow series.',
  },
  {
    name: 'NPER',
    params: ['rate', 'pmt', 'pv', '[fv]', '[type]'],
    description: 'Number of periods needed for an investment or loan.',
  },
  {
    name: 'IRR',
    params: ['values', '[guess]'],
    description: 'Internal rate of return for periodic cash flows.',
  },
  {
    name: 'INDEX',
    params: ['array', 'row_num', '[column_num]'],
    description: 'Returns the value at a row and column within a range.',
  },
  {
    name: 'IPMT',
    params: ['rate', 'per', 'nper', 'pv', '[fv]', '[type]'],
    description: 'Interest portion of a payment for a given period.',
  },
  {
    name: 'MATCH',
    params: ['lookup_value', 'lookup_array', '[match_type]'],
    description: 'Position of a value in a range.',
  },
  {
    name: 'MIRR',
    params: ['values', 'finance_rate', 'reinvest_rate'],
    description: 'Modified internal rate of return.',
  },
  {
    name: 'NETWORKDAYS',
    params: ['start_date', 'end_date', '[holidays]'],
    description: 'Number of working days between two dates.',
  },
  {
    name: 'NOMINAL',
    params: ['effect_rate', 'npery'],
    description: 'Nominal annual interest rate.',
  },
  {
    name: 'SUM',
    params: ['number1', '[number2]', '...'],
    description: 'Adds numbers and ranges.',
  },
  {
    name: 'SUMIFS',
    params: ['sum_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    description: 'Sums values that satisfy one or more criteria.',
  },
  {
    name: 'IF',
    params: ['logical_test', 'value_if_true', 'value_if_false'],
    description: 'Conditional branching expression.',
  },
  {
    name: 'PPMT',
    params: ['rate', 'per', 'nper', 'pv', '[fv]', '[type]'],
    description: 'Principal portion of a payment for a given period.',
  },
  {
    name: 'RATE',
    params: ['nper', 'pmt', 'pv', '[fv]', '[type]', '[guess]'],
    description: 'Interest rate per period for an annuity.',
  },
  {
    name: 'ROUND',
    params: ['number', 'num_digits'],
    description: 'Rounds a number to a specified digit count.',
  },
  {
    name: 'SLN',
    params: ['cost', 'salvage', 'life'],
    description: 'Straight-line depreciation.',
  },
  {
    name: 'SYD',
    params: ['cost', 'salvage', 'life', 'period'],
    description: "Sum-of-years'-digits depreciation.",
  },
  {
    name: 'XNPV',
    params: ['rate', 'values', 'dates'],
    description: 'Net present value for irregular cash-flow dates.',
  },
  {
    name: 'YEARFRAC',
    params: ['start_date', 'end_date', '[basis]'],
    description: 'Fraction of a year between two dates.',
  },
];

const FORMULA_DOC_BY_NAME = new Map(FORMULA_DOCS.map((doc) => [doc.name, doc]));
const PRESET_OPTIONS: Array<{ value: PresetType; label: string }> = [
  { value: 'blank', label: 'Blank' },
  { value: 'journal_entry', label: 'Journal Entry' },
  { value: 't_account', label: 'T-Account' },
];

function App() {
  const hotRef = useRef<any>(null);
  const gridStageRef = useRef<HTMLElement | null>(null);
  const formulaEditSessionRef = useRef<FormulaEditSession | null>(null);
  const activeReferencesRef = useRef<FormulaReferenceHighlight[]>([]);
  const formulaOverlayRef = useRef<HTMLDivElement | null>(null);
  const [gridData, setGridData] = useState(() => buildPresetGrid('blank'));
  const [gridEpoch, setGridEpoch] = useState(0);
  const [selectedPreset, setSelectedPreset] = useState<PresetType>('blank');
  const [label, setLabel] = useState(() => buildDefaultLabel('blank'));
  const [formulaPopup, setFormulaPopup] = useState<FormulaPopupState | null>(null);

  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [ledgerTarget, setLedgerTarget] = useState<LedgerTarget | null>(null);
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [isSigningIn, setIsSigningIn] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [statusMessage, setStatusMessage] = useState('Sign in to connect your Beanfolio archive.');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const [defaultDecimals, setDefaultDecimals] = useState<number | null>(null);
  const [defaultCommas, setDefaultCommas] = useState(true);

  const [colorPickerTarget, setColorPickerTarget] = useState<'fill' | 'text' | null>(null);
  const [fillPaletteLightness, setFillPaletteLightness] = useState(1);
  const [textPaletteLightness, setTextPaletteLightness] = useState(-1);
  const [roundPopupOpen, setRoundPopupOpen] = useState(false);
  const [presetMenuOpen, setPresetMenuOpen] = useState(false);
  const [presetMenuVisible, setPresetMenuVisible] = useState(false);
  const [saveMenuOpen, setSaveMenuOpen] = useState(false);
  const [saveMenuVisible, setSaveMenuVisible] = useState(false);
  const [popOutBlockedMessage, setPopOutBlockedMessage] = useState<string | null>(null);
  const colorPickerRef = useRef<HTMLDivElement>(null);
  const roundPopupRef = useRef<HTMLDivElement>(null);
  const presetMenuRef = useRef<HTMLDivElement>(null);
  const saveMenuRef = useRef<HTMLDivElement>(null);
  const savedSelectionRef = useRef<Array<[number, number, number, number]>>([]);
  const isDesktopApp = Boolean(window.beanfolioDesktop?.isDesktop);
  const isWebPopOutMode = useMemo(() => {
    if (isDesktopApp) {
      return false;
    }

    const search = new URLSearchParams(window.location.search);
    return search.get('mode') === 'app';
  }, [isDesktopApp]);
  const isWebLanding = !isDesktopApp && !isWebPopOutMode;

  const closePresetMenu = useCallback(() => {
    setPresetMenuOpen(false);
  }, []);

  const openPresetMenu = useCallback(() => {
    setPresetMenuVisible(true);
    setPresetMenuOpen(true);
  }, []);

  const togglePresetMenu = useCallback(() => {
    if (presetMenuOpen) {
      closePresetMenu();
      return;
    }

    if (!presetMenuVisible) {
      openPresetMenu();
      return;
    }

    setPresetMenuOpen(true);
  }, [closePresetMenu, openPresetMenu, presetMenuOpen, presetMenuVisible]);

  const closeSaveMenu = useCallback(() => {
    setSaveMenuOpen(false);
  }, []);

  const openSaveMenu = useCallback(() => {
    setSaveMenuVisible(true);
    setSaveMenuOpen(true);
  }, []);

  const toggleSaveMenu = useCallback(() => {
    if (saveMenuOpen) {
      closeSaveMenu();
      return;
    }

    if (!saveMenuVisible) {
      openSaveMenu();
      return;
    }

    setSaveMenuOpen(true);
  }, [closeSaveMenu, openSaveMenu, saveMenuOpen, saveMenuVisible]);

  const activeColumnCount = useMemo(() => getPresetColumnCount(selectedPreset), [selectedPreset]);
  const columnHeaders = useMemo(() => getPresetColumnHeaders(selectedPreset), [selectedPreset]);
  const hiddenColumns = useMemo(() => getPresetHiddenColumns(selectedPreset), [selectedPreset]);
  const columnWidths = useMemo(() => getPresetColumnWidths(selectedPreset), [selectedPreset]);
  const mergeCells = useMemo(() => getPresetMergeCells(selectedPreset), [selectedPreset]);

  useEffect(() => {
    loadGoogleIdentityScript()
      .then(() => setIsAuthReady(true))
      .catch((error: unknown) => {
        setErrorMessage(getErrorMessage(error));
        setStatusMessage('Google sign-in failed to initialize.');
      });
  }, []);

  useEffect(() => {
    if (!colorPickerTarget && !roundPopupOpen && !presetMenuVisible && !saveMenuVisible) return;

    const handleClick = (e: MouseEvent) => {
      if (colorPickerTarget && colorPickerRef.current && !colorPickerRef.current.contains(e.target as Node)) {
        setColorPickerTarget(null);
      }
      if (roundPopupOpen && roundPopupRef.current && !roundPopupRef.current.contains(e.target as Node)) {
        setRoundPopupOpen(false);
      }
      if (presetMenuVisible && presetMenuRef.current && !presetMenuRef.current.contains(e.target as Node)) {
        closePresetMenu();
      }
      if (saveMenuVisible && saveMenuRef.current && !saveMenuRef.current.contains(e.target as Node)) {
        closeSaveMenu();
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [closePresetMenu, closeSaveMenu, colorPickerTarget, presetMenuVisible, roundPopupOpen, saveMenuVisible]);

  useEffect(() => {
    if (presetMenuOpen) {
      return;
    }

    const timer = window.setTimeout(() => setPresetMenuVisible(false), DROPDOWN_TRANSITION_MS);
    return () => window.clearTimeout(timer);
  }, [presetMenuOpen]);

  useEffect(() => {
    if (saveMenuOpen) {
      return;
    }

    const timer = window.setTimeout(() => setSaveMenuVisible(false), DROPDOWN_TRANSITION_MS);
    return () => window.clearTimeout(timer);
  }, [saveMenuOpen]);

  useEffect(() => {
    const html = document.documentElement;
    const body = document.body;

    if (isWebLanding) {
      html.classList.add('web-landing-mode');
      body.classList.add('web-landing-mode');
      html.classList.remove('web-app-mode');
      body.classList.remove('web-app-mode');
    } else if (isWebPopOutMode) {
      html.classList.add('web-app-mode');
      body.classList.add('web-app-mode');
      html.classList.remove('web-landing-mode');
      body.classList.remove('web-landing-mode');
    } else {
      html.classList.remove('web-landing-mode');
      body.classList.remove('web-landing-mode');
      html.classList.remove('web-app-mode');
      body.classList.remove('web-app-mode');
    }

    return () => {
      html.classList.remove('web-landing-mode');
      body.classList.remove('web-landing-mode');
      html.classList.remove('web-app-mode');
      body.classList.remove('web-app-mode');
    };
  }, [isWebLanding, isWebPopOutMode]);

  const getSelectedRanges = useCallback((): Array<[number, number, number, number]> => {
    return savedSelectionRef.current;
  }, []);

  const forEachSelectedCell = useCallback((fn: (row: number, col: number) => void) => {
    for (const [r1, c1, r2, c2] of getSelectedRanges()) {
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          fn(r, c);
        }
      }
    }
  }, [getSelectedRanges]);

  const restoreSelection = useCallback(() => {
    const hot = getHotInstance(hotRef);
    if (!hot) return;
    const saved = savedSelectionRef.current;
    if (saved.length > 0) {
      hot.selectCells(saved);
    }
    hot.listen();
  }, []);

  const rerenderGrid = useCallback(() => {
    const hot = getHotInstance(hotRef);
    if (!hot) return;
    hot.render();
    restoreSelection();
  }, [restoreSelection]);

  const toggleStyleProp = useCallback((prop: keyof CellStyle) => {
    const ranges = getSelectedRanges();
    if (ranges.length === 0) return;
    let allSet = true;
    for (const [r1, c1, r2, c2] of ranges) {
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (!getCellStyle(r, c)?.[prop]) { allSet = false; break; }
        }
        if (!allSet) break;
      }
      if (!allSet) break;
    }
    forEachSelectedCell((r, c) => setCellStyle(r, c, { [prop]: !allSet ? true : undefined }));
    rerenderGrid();
  }, [getSelectedRanges, forEachSelectedCell, rerenderGrid]);

  const applyColor = useCallback((target: 'fill' | 'text', color: string) => {
    const prop = target === 'fill' ? 'fillColor' : 'textColor';
    forEachSelectedCell((r, c) => setCellStyle(r, c, { [prop]: color }));
    setColorPickerTarget(null);
    rerenderGrid();
  }, [forEachSelectedCell, rerenderGrid]);

  const setPaletteLightness = useCallback((target: 'fill' | 'text', nextValue: number) => {
    const clamped = clamp(nextValue, PALETTE_LIGHTNESS_MIN, PALETTE_LIGHTNESS_MAX);
    if (target === 'fill') {
      setFillPaletteLightness(clamped);
      return;
    }

    setTextPaletteLightness(clamped);
  }, []);

  const resetPaletteLightness = useCallback((target: 'fill' | 'text') => {
    if (target === 'fill') {
      setFillPaletteLightness(0);
      return;
    }

    setTextPaletteLightness(0);
  }, []);

  const toggleComma = useCallback(() => {
    const ranges = getSelectedRanges();
    if (ranges.length === 0) return;
    let allOff = true;
    for (const [r1, c1, r2, c2] of ranges) {
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (!getCellStyle(r, c)?.commaOff) { allOff = false; break; }
        }
        if (!allOff) break;
      }
      if (!allOff) break;
    }
    forEachSelectedCell((r, c) => setCellStyle(r, c, { commaOff: allOff ? undefined : true }));
    rerenderGrid();
  }, [getSelectedRanges, forEachSelectedCell, rerenderGrid]);

  const togglePercent = useCallback(() => {
    const ranges = getSelectedRanges();
    if (ranges.length === 0) return;
    let allPercent = true;
    for (const [r1, c1, r2, c2] of ranges) {
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (!getCellStyle(r, c)?.percent) { allPercent = false; break; }
        }
        if (!allPercent) break;
      }
      if (!allPercent) break;
    }
    forEachSelectedCell((r, c) => setCellStyle(r, c, { percent: allPercent ? undefined : true }));
    rerenderGrid();
  }, [getSelectedRanges, forEachSelectedCell, rerenderGrid]);

  const applyRound = useCallback((decimals: number) => {
    const hot = getHotInstance(hotRef);
    if (!hot) return;
    const ranges = getSelectedRanges();
    if (ranges.length === 0) return;
    const changes: Array<[number, number, string | number]> = [];
    for (const [r1, c1, r2, c2] of ranges) {
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const display = hot.getDataAtCell(r, c);
          const num = typeof display === 'number' ? display : parseFloat(String(display));
          if (!Number.isFinite(num)) continue;
          const src = hot.getSourceDataAtCell(r, c);
          if (isFormula(src)) {
            changes.push([r, c, wrapFormulaInRound(src, decimals)]);
          } else {
            const factor = Math.pow(10, decimals);
            changes.push([r, c, Math.round(num * factor) / factor]);
          }
        }
      }
    }
    if (changes.length > 0) {
      hot.batch(() => {
        for (const [r, c, v] of changes) {
          hot.setDataAtCell(r, c, v);
        }
      });
    }
    setRoundPopupOpen(false);
    restoreSelection();
  }, [getSelectedRanges, restoreSelection]);

  const handleSignIn = async () => {
    setErrorMessage(null);

    if (!GOOGLE_CLIENT_ID) {
      setErrorMessage('Missing VITE_GOOGLE_CLIENT_ID in environment configuration.');
      return;
    }

    setIsSigningIn(true);

    try {
      const token = await requestGoogleAccessToken({
        clientId: GOOGLE_CLIENT_ID,
        scopes: [...GOOGLE_SCOPES],
        prompt: accessToken ? '' : 'consent',
      });

      const ledger = await findOrCreateMiniLedger(token);
      setAccessToken(token);
      setLedgerTarget(ledger);
      setStatusMessage(`Connected to ${ledger.spreadsheetTitle} / ${ledger.sheetTitle}.`);
    } catch (error: unknown) {
      setErrorMessage(getErrorMessage(error));
      setStatusMessage('Unable to complete sign-in.');
    } finally {
      setIsSigningIn(false);
    }
  };

  const handleSignOut = async () => {
    if (accessToken) {
      await revokeGoogleToken(accessToken);
    }

    setAccessToken(null);
    setLedgerTarget(null);
    setStatusMessage('Signed out.');
  };

  const handlePresetSelect = (nextPreset: PresetType) => {
    const hot = getHotInstance(hotRef);
    const hasCurrentData = hot
      ? sourceHasNonEmptyCells(hot.getSourceDataArray() as unknown[][])
      : hasNonEmptyCells(gridData);

    if (hasCurrentData) {
      const confirmed = window.confirm('Replace current grid content with this preset?');
      if (!confirmed) {
        return;
      }
    }

    clearAllStyles();
    setSelectedPreset(nextPreset);
    setGridData(buildPresetGrid(nextPreset));
    setGridEpoch((value) => value + 1);
    setLabel(buildDefaultLabel(nextPreset));
  };

  const handleSave = async () => {
    setErrorMessage(null);

    if (!accessToken || !ledgerTarget) {
      setErrorMessage('Sign in first to save into Google Sheets.');
      return;
    }

    const hot = getHotInstance(hotRef);
    if (!hot) {
      setErrorMessage('Grid is not ready yet.');
      return;
    }

    setIsSaving(true);

    try {
      const snapshot = trimTrailingEmptyRows(buildSnapshot(hot, activeColumnCount));

      await appendLedgerBlock(accessToken, ledgerTarget, {
        label,
        cells: snapshot,
        columnCount: activeColumnCount,
      });

      setStatusMessage(`Saved block at ${new Intl.DateTimeFormat(undefined, {
        month: 'short',
        day: 'numeric',
        year: 'numeric',
        hour: 'numeric',
        minute: '2-digit',
      }).format(new Date())}.`);
    } catch (error: unknown) {
      setErrorMessage(getErrorMessage(error));
      setStatusMessage('Save failed.');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDownload = (format: ExportFormat) => {
    setErrorMessage(null);
    closeSaveMenu();

    const hot = getHotInstance(hotRef);
    if (!hot) {
      setErrorMessage('Grid is not ready yet.');
      return;
    }

    try {
      const snapshot = trimTrailingEmptyRows(buildSnapshot(hot, activeColumnCount));
      const exportFile = buildExportFile(snapshot, format, label);
      downloadFile(exportFile);
      setStatusMessage(`Downloaded ${exportFile.fileName}.`);
    } catch (error: unknown) {
      setErrorMessage(getErrorMessage(error));
      setStatusMessage('Download failed.');
    }
  };

  const handleSaveMainClick = () => {
    closeSaveMenu();
    if (!accessToken || !ledgerTarget) {
      void handleSignIn();
      return;
    }

    void handleSave();
  };

  const refreshFormulaPopup = () => {
    const hot = getHotInstance(hotRef);
    if (!hot) {
      setFormulaPopup(null);
      formulaEditSessionRef.current = null;
      updateFormulaReferences(hotRef, activeReferencesRef, null);
      updateFormulaOverlay(formulaOverlayRef, null, null, []);
      return;
    }

    const editorContext = getOpenEditorContext(hot);

    if (!editorContext) {
      setFormulaPopup(null);
      formulaEditSessionRef.current = null;
      updateFormulaReferences(hotRef, activeReferencesRef, null);
      updateFormulaOverlay(formulaOverlayRef, null, null, []);
      return;
    }

    if (!isFormulaInput(editorContext.input)) {
      updateFormulaReferences(hotRef, activeReferencesRef, null);
      updateFormulaOverlay(formulaOverlayRef, editorContext.editor.TEXTAREA, null, []);
      setFormulaPopup(buildFormulaPopupState(editorContext.input, editorContext.caretPosition, gridStageRef.current));
      return;
    }

    syncFormulaSession(formulaEditSessionRef, editorContext);
    updateFormulaReferences(hotRef, activeReferencesRef, editorContext.input);
    updateFormulaOverlay(formulaOverlayRef, editorContext.editor.TEXTAREA, editorContext.input, activeReferencesRef.current);

    setFormulaPopup(buildFormulaPopupState(editorContext.input, editorContext.caretPosition, gridStageRef.current));
  };

  const toggleAbsoluteReference = () => {
    const hot = getHotInstance(hotRef);
    if (!hot) {
      return;
    }

    const editorContext = getOpenEditorContext(hot);
    if (!editorContext || !isFormulaInput(editorContext.input)) {
      return;
    }

    const session = syncFormulaSession(formulaEditSessionRef, editorContext);
    const currentInput = String(editorContext.editor.getValue() ?? '');
    const referenceRange = detectReferenceTokenRange(currentInput, editorContext.caretPosition);

    if (referenceRange) {
      const currentReference = currentInput.slice(referenceRange.start, referenceRange.end);
      const nextReference = cycleAbsoluteReferenceToken(currentReference);

      if (nextReference) {
        const nextInput =
          `${currentInput.slice(0, referenceRange.start)}${nextReference}${currentInput.slice(referenceRange.end)}`;
        const nextEnd = referenceRange.start + nextReference.length;

        editorContext.editor.setValue(nextInput);
        if (editorContext.editor.TEXTAREA) {
          editorContext.editor.TEXTAREA.setSelectionRange(nextEnd, nextEnd);
        }

        session.pendingReferenceRange = {
          start: referenceRange.start,
          end: nextEnd,
        };
      }
    } else {
      upsertFormulaReference(editorContext, session, toAbsoluteA1Reference(session.cursorRow, session.cursorCol));
    }

    editorContext.editor.focus();
    hot.listen();
    window.requestAnimationFrame(refreshFormulaPopup);
  };

  const handleAfterBeginEditing = () => {
    window.requestAnimationFrame(refreshFormulaPopup);
  };

  const handleAfterSelectionEnd = useCallback(() => {
    const hot = getHotInstance(hotRef);
    if (!hot) return;
    const sel = hot.getSelected();
    if (sel && sel.length > 0) {
      savedSelectionRef.current = sel.map(([r1, c1, r2, c2]) => [
        Math.min(r1, r2), Math.min(c1, c2),
        Math.max(r1, r2), Math.max(c1, c2),
      ] as [number, number, number, number]);
    }
    window.requestAnimationFrame(refreshFormulaPopup);
  }, []);

  const handleBeforeKeyDown = useCallback<NonNullable<Handsontable.GridSettings['beforeKeyDown']>>((event) => {
    const hot = getHotInstance(hotRef);
    if (!hot) {
      return;
    }

    const editorContext = getOpenEditorContext(hot);
    if (!editorContext || !isFormulaInput(editorContext.input)) {
      window.requestAnimationFrame(refreshFormulaPopup);
      return;
    }

    const session = syncFormulaSession(formulaEditSessionRef, editorContext);

    if (isArrowKey(event.key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      (event as any).isImmediatePropagationEnabled = false;

      const [nextRow, nextCol] = moveReferenceCursor(
        session.cursorRow,
        session.cursorCol,
        event.key,
        activeColumnCount - 1,
      );

      session.cursorRow = nextRow;
      session.cursorCol = nextCol;
      session.lastReferenceInput = 'keyboard';

      upsertFormulaReference(editorContext, session, toA1Reference(nextRow, nextCol));
      updateFormulaReferences(hotRef, activeReferencesRef, String(editorContext.editor.getValue() ?? ''));
      window.requestAnimationFrame(refreshFormulaPopup);
      return;
    }

    if (isFormulaOperator(event.key) && session.lastReferenceInput === 'keyboard') {
      session.cursorRow = session.editRow;
      session.cursorCol = session.editCol;
    }

    if (resetsReferenceRange(event.key)) {
      session.pendingReferenceRange = null;
    }

    window.requestAnimationFrame(refreshFormulaPopup);
  }, [activeColumnCount]);

  const handleBeforeOnCellMouseDown = useCallback<NonNullable<Handsontable.GridSettings['beforeOnCellMouseDown']>>((
    event,
    coords,
    _td,
    controller,
  ) => {
    const hot = getHotInstance(hotRef);
    if (!hot) {
      return;
    }

    const editorContext = getOpenEditorContext(hot);
    if (!editorContext || !isFormulaInput(editorContext.input)) {
      return;
    }

    if (coords.row < 0 || coords.col < 0 || coords.col >= activeColumnCount) {
      return;
    }

    const session = syncFormulaSession(formulaEditSessionRef, editorContext);
    if (coords.row === session.editRow && coords.col === session.editCol) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    controller.row = false;
    controller.column = false;
    controller.cell = false;

    session.cursorRow = coords.row;
    session.cursorCol = coords.col;
    session.lastReferenceInput = 'mouse';

    upsertFormulaReference(editorContext, session, toA1Reference(coords.row, coords.col));
    updateFormulaReferences(hotRef, activeReferencesRef, String(editorContext.editor.getValue() ?? ''));
    editorContext.editor.focus();
    hot.listen();
    window.requestAnimationFrame(refreshFormulaPopup);
  }, [activeColumnCount]);

  const handleBeforeRenderer = useCallback<NonNullable<Handsontable.GridSettings['beforeRenderer']>>((
    td,
    row,
    column,
  ) => {
    td.classList.remove(
      'formula-ref-cell',
      'journal-credit-indent',
      'journal-index-cell',
      'journal-header-cell',
      't-account-top-border',
      't-account-divider',
      't-account-title',
      't-account-label',
      't-account-sum-border',
      't-account-total-cell',
    );

    if (selectedPreset === 'journal_entry') {
      if (column === 0) {
        td.classList.add('journal-index-cell');
      }

      if (row === 0 && column <= 3) {
        td.classList.add('journal-header-cell');
      }

      if (column === 1) {
        const hot = getHotInstance(hotRef);
        const creditValue = hot?.getSourceDataAtCell(row, 3);
        if (row >= 1 && hasMeaningfulValue(creditValue)) {
          td.classList.add('journal-credit-indent');
        }
      }
    }

    if (selectedPreset === 't_account') {
      if (isTAccountTopBorderCell(row, column)) {
        td.classList.add('t-account-top-border');
      }
      if (isTAccountDividerCell(row, column)) {
        td.classList.add('t-account-divider');
      }
      if (isTAccountTitleCell(row, column)) {
        td.classList.add('t-account-title');
      }
      if (isTAccountLabelCell(row, column)) {
        td.classList.add('t-account-label');
      }
      if (isTAccountSumBorderCell(row, column)) {
        td.classList.add('t-account-sum-border');
      }
      if (isTAccountTotalCell(row, column)) {
        td.classList.add('t-account-total-cell');
      }
    }

    const refs = activeReferencesRef.current;
    for (let i = 0; i < REFERENCE_COLORS.length; i++) {
      td.classList.remove(`formula-ref-cell-${i}`);
    }
    for (const highlight of refs) {
      if (row === highlight.row && column === highlight.col) {
        td.classList.add(`formula-ref-cell-${highlight.colorIndex}`);
        break;
      }
    }

    const style = getCellStyle(row, column);
    td.style.fontWeight = style?.bold ? '700' : '';
    td.style.fontStyle = style?.italic ? 'italic' : '';
    td.style.textDecoration = style?.underline ? 'underline' : '';
    td.style.backgroundColor = style?.fillColor ?? '';
    td.style.color = style?.textColor ?? '';
  }, [selectedPreset]);

  const handleAfterRenderer = useCallback((
    td: HTMLTableCellElement,
    row: number,
    col: number,
    _prop: string | number,
    value: unknown,
  ) => {
    if (typeof value !== 'number' || !Number.isFinite(value)) return;
    const style = getCellStyle(row, col);
    if (style?.percent) {
      td.textContent = (value * 100).toFixed(2) + '%';
      return;
    }
    const useCommas = style?.commaOff ? false : defaultCommas;
    if (useCommas) {
      td.textContent = formatNumberWithCommas(value, defaultDecimals ?? undefined);
    } else if (defaultDecimals !== null) {
      td.textContent = value.toFixed(defaultDecimals);
    }
  }, [defaultDecimals, defaultCommas]);

  const handleAfterDeselect = () => {
    const hot = getHotInstance(hotRef);
    const editorContext = hot ? getOpenEditorContext(hot) : null;
    if (editorContext && isFormulaInput(editorContext.input)) {
      return;
    }

    setFormulaPopup(null);
    formulaEditSessionRef.current = null;
    updateFormulaReferences(hotRef, activeReferencesRef, null);
    updateFormulaOverlay(formulaOverlayRef, null, null, []);
  };

  const toggleSidebar = () => {
    const next = !sidebarOpen;
    setSidebarOpen(next);
    closePresetMenu();
    closeSaveMenu();
    window.beanfolioDesktop?.setSidebarOpen?.(next);

    if (isWebPopOutMode) {
      const appWidth = next ? POP_OUT_APP_WIDTH + SIDEBAR_WIDTH : POP_OUT_APP_WIDTH;
      const windowWidth = appWidth + POP_OUT_CHROME_WIDTH;
      const windowHeight = POP_OUT_APP_HEIGHT + POP_OUT_CHROME_HEIGHT;
      window.resizeTo(windowWidth, windowHeight);
    }
  };

  const handleOpenPopOut = useCallback(() => {
    const width = POP_OUT_APP_WIDTH + POP_OUT_CHROME_WIDTH;
    const height = POP_OUT_APP_HEIGHT + POP_OUT_CHROME_HEIGHT;
    const left = Math.max(window.screenX + Math.round((window.outerWidth - width) / 2), 0);
    const top = Math.max(window.screenY + Math.round((window.outerHeight - height) / 2), 0);
    const popOutUrl = new URL(window.location.href);

    popOutUrl.searchParams.set('mode', 'app');

    const features = [
      `width=${width}`,
      `height=${height}`,
      `left=${left}`,
      `top=${top}`,
      'noopener',
      'noreferrer',
    ].join(',');

    setPopOutBlockedMessage(null);
    const popOutWindow = window.open(popOutUrl.toString(), '_blank', features);
    if (!popOutWindow) {
      setPopOutBlockedMessage('Your browser blocked the pop-out window. Allow pop-ups and try again.');
      return;
    }

    popOutWindow.focus();
  }, []);

  const appShellClassName = isWebLanding ? 'app-shell app-shell-embedded' : 'app-shell';

  const appContent = (
    <div className={appShellClassName} data-sidebar-open={sidebarOpen || undefined}>
      <div className="drag-bar">
        <button
          className={sidebarOpen ? 'sidebar-toggle is-open' : 'sidebar-toggle'}
          type="button"
          onClick={toggleSidebar}
          aria-label={sidebarOpen ? 'Close menu' : 'Open menu'}
          aria-expanded={sidebarOpen}
          aria-controls="app-sidebar"
        >
          <span className="sidebar-toggle-icon" aria-hidden="true">
            <span className="sidebar-toggle-line sidebar-toggle-line-1" />
            <span className="sidebar-toggle-line sidebar-toggle-line-2" />
            <span className="sidebar-toggle-line sidebar-toggle-line-3" />
          </span>
        </button>
      </div>

      <div className="format-toolbar" role="toolbar" aria-label="Formatting">
        <div className="fmt-group">
          <button
            className="fmt-btn"
            type="button"
            title="Bold"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => toggleStyleProp('bold')}
          >
            <strong>B</strong>
          </button>
          <button
            className="fmt-btn"
            type="button"
            title="Italic"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => toggleStyleProp('italic')}
          >
            <em>I</em>
          </button>
          <button
            className="fmt-btn"
            type="button"
            title="Underline"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => toggleStyleProp('underline')}
          >
            <span style={{ textDecoration: 'underline' }}>U</span>
          </button>
        </div>

        <span className="fmt-divider" />

        <div className="fmt-group" ref={colorPickerRef}>
          <div className="color-picker-anchor">
            <button
              className={colorPickerTarget === 'fill' ? 'fmt-btn fmt-btn-fill is-active' : 'fmt-btn fmt-btn-fill'}
              type="button"
              title="Fill color"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setColorPickerTarget(colorPickerTarget === 'fill' ? null : 'fill')}
            >
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
                <rect x="1" y="10" width="12" height="3" rx="0.5" fill="currentColor" />
                <path d="M7 1.5L3 8.5h8L7 1.5z" fill="currentColor" />
              </svg>
            </button>
            {colorPickerTarget === 'fill' ? (
              <div className="color-picker-popup">
                <div className="color-adjust-row">
                  <input
                    className="color-adjust-slider"
                    type="range"
                    min={PALETTE_LIGHTNESS_MIN}
                    max={PALETTE_LIGHTNESS_MAX}
                    step={1}
                    value={fillPaletteLightness}
                    onMouseDown={(e) => e.stopPropagation()}
                    onChange={(e) => setPaletteLightness('fill', Number(e.target.value))}
                    aria-label="Fill palette lightness"
                  />
                  <button
                    className="color-adjust-reset"
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => resetPaletteLightness('fill')}
                    title="Reset lightness"
                    aria-label="Reset lightness"
                  >
                    <svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">
                      <path d="M9.5 6a3.5 3.5 0 1 1-1.1-2.55" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" />
                      <path d="M8.5 1.9v2.1h-2.1" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </button>
                </div>
                <div className="color-adjust-label">{formatLightnessLabel(fillPaletteLightness)}</div>
                <div className="color-grid">
                  {PRESET_BASE_COLORS
                    .map((color) => shiftHexColorLightness(color, fillPaletteLightness))
                    .map((c, index) => (
                      <button
                        key={`${c}-${index}`}
                        className="color-swatch"
                        type="button"
                        style={{ background: c }}
                        title={c}
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => applyColor('fill', c)}
                      />
                    ))}
                </div>
                <button
                  className="color-clear-btn"
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => {
                    forEachSelectedCell((r, c) => setCellStyle(r, c, { fillColor: undefined }));
                    setColorPickerTarget(null);
                    rerenderGrid();
                  }}
                >
                  Clear
                </button>
              </div>
            ) : null}
          </div>

          <div className="color-picker-anchor">
            <button
              className={colorPickerTarget === 'text' ? 'fmt-btn fmt-btn-text-color is-active' : 'fmt-btn fmt-btn-text-color'}
              type="button"
              title="Text color"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setColorPickerTarget(colorPickerTarget === 'text' ? null : 'text')}
            >
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
                <rect x="1" y="11" width="12" height="2" rx="0.5" fill="currentColor" />
                <text x="7" y="9.5" textAnchor="middle" fontSize="9" fontWeight="700" fill="currentColor">A</text>
              </svg>
            </button>
            {colorPickerTarget === 'text' ? (
              <div className="color-picker-popup">
                <div className="color-adjust-row">
                  <input
                    className="color-adjust-slider"
                    type="range"
                    min={PALETTE_LIGHTNESS_MIN}
                    max={PALETTE_LIGHTNESS_MAX}
                    step={1}
                    value={textPaletteLightness}
                    onMouseDown={(e) => e.stopPropagation()}
                    onChange={(e) => setPaletteLightness('text', Number(e.target.value))}
                    aria-label="Text palette lightness"
                  />
                  <button
                    className="color-adjust-reset"
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => resetPaletteLightness('text')}
                    title="Reset lightness"
                    aria-label="Reset lightness"
                  >
                    <svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">
                      <path d="M9.5 6a3.5 3.5 0 1 1-1.1-2.55" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" />
                      <path d="M8.5 1.9v2.1h-2.1" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </button>
                </div>
                <div className="color-adjust-label">{formatLightnessLabel(textPaletteLightness)}</div>
                <div className="color-grid">
                  {PRESET_BASE_COLORS
                    .map((color) => shiftHexColorLightness(color, textPaletteLightness))
                    .map((c, index) => (
                      <button
                        key={`${c}-${index}`}
                        className="color-swatch"
                        type="button"
                        style={{ background: c }}
                        title={c}
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => applyColor('text', c)}
                      />
                    ))}
                </div>
                <button
                  className="color-clear-btn"
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => {
                    forEachSelectedCell((r, c) => setCellStyle(r, c, { textColor: undefined }));
                    setColorPickerTarget(null);
                    rerenderGrid();
                  }}
                >
                  Clear
                </button>
              </div>
            ) : null}
          </div>
        </div>

        <span className="fmt-divider" />

        <div className="fmt-group">
          <button
            className="fmt-btn"
            type="button"
            title="Toggle absolute reference ($A$1) while editing a formula"
            onMouseDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
              toggleAbsoluteReference();
            }}
          >
            <span className="fmt-label">$</span>
          </button>
          <button
            className="fmt-btn"
            type="button"
            title="Percent format"
            onMouseDown={(e) => e.preventDefault()}
            onClick={togglePercent}
          >
            <span className="fmt-label">%</span>
          </button>
          <button
            className="fmt-btn"
            type="button"
            title="Round to 2dp (shift=0dp)"
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
              applyRound(e.shiftKey ? 0 : 2);
            }}
          >
            <span className="fmt-label">.0</span>
          </button>
          <div className="fmt-caret-wrap" ref={roundPopupRef}>
            <button
              className="fmt-btn fmt-btn-caret"
              type="button"
              title="Pick decimal places"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setRoundPopupOpen(!roundPopupOpen)}
            >
              <svg width="8" height="8" viewBox="0 0 8 8" fill="none" aria-hidden="true">
                <path d="M1.5 3L4 5.5L6.5 3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
            {roundPopupOpen && (
              <div className="round-popup">
                {[0, 1, 2, 3, 4, 5, 6].map((d) => (
                  <button
                    key={d}
                    className="round-option"
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => applyRound(d)}
                  >
                    {d} dp
                  </button>
                ))}
              </div>
            )}
          </div>
          <button
            className="fmt-btn"
            type="button"
            title="Toggle comma separators"
            onMouseDown={(e) => e.preventDefault()}
            onClick={toggleComma}
          >
            <span className="fmt-label">,</span>
          </button>
        </div>
      </div>

      <div className="app-body">
        <aside id="app-sidebar" className={sidebarOpen ? 'sidebar is-open' : 'sidebar'} aria-hidden={!sidebarOpen}>
            <div className="sidebar-controls">
              <div className="sidebar-top-row">
                <div className="sidebar-field">
                  <span className="label">Preset</span>
                  <div className="sidebar-preset" ref={presetMenuRef}>
                    <button
                      id="preset-select"
                      className={presetMenuOpen ? 'input sidebar-input sidebar-preset-trigger is-open' : 'input sidebar-input sidebar-preset-trigger'}
                      type="button"
                      aria-haspopup="menu"
                      aria-expanded={presetMenuOpen}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={togglePresetMenu}
                    >
                      <span>{getPresetOptionLabel(selectedPreset)}</span>
                      <svg width="10" height="6" viewBox="0 0 10 6" fill="none" aria-hidden="true">
                        <path d="M1 1L5 5L9 1" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </button>

                    {presetMenuVisible ? (
                      <div
                        className={presetMenuOpen ? 'sidebar-save-menu sidebar-preset-menu is-open' : 'sidebar-save-menu sidebar-preset-menu is-closing'}
                        role="menu"
                        aria-label="Preset options"
                      >
                        {PRESET_OPTIONS.map((option) => (
                          <button
                            key={option.value}
                            className={option.value === selectedPreset ? 'sidebar-save-option sidebar-preset-option is-selected' : 'sidebar-save-option sidebar-preset-option'}
                            type="button"
                            role="menuitemradio"
                            aria-checked={option.value === selectedPreset}
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => {
                              closePresetMenu();
                              handlePresetSelect(option.value);
                            }}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    ) : null}
                  </div>
                </div>
              </div>

              <div className="sidebar-field">
                <span className="label">Label</span>
                <input
                  id="save-label"
                  className="input sidebar-input"
                  type="text"
                  value={label}
                  onChange={(event) => setLabel(event.target.value)}
                  placeholder="Optional save label"
                />
              </div>

              <div className="sidebar-row-pair">
                <div className="sidebar-field">
                  <span className="label">Decimals</span>
                  <select
                    className="input sidebar-input"
                    value={defaultDecimals === null ? 'auto' : String(defaultDecimals)}
                    onChange={(e) => {
                      const v = e.target.value;
                      setDefaultDecimals(v === 'auto' ? null : Number(v));
                      getHotInstance(hotRef)?.render();
                    }}
                  >
                    <option value="auto">Auto</option>
                    <option value="0">0</option>
                    <option value="1">1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="5">5</option>
                    <option value="6">6</option>
                  </select>
                </div>
                <div className="sidebar-field">
                  <span className="label">Format</span>
                  <select
                    className="input sidebar-input"
                    value={defaultCommas ? 'comma' : 'plain'}
                    onChange={(e) => {
                      setDefaultCommas(e.target.value === 'comma');
                      getHotInstance(hotRef)?.render();
                    }}
                  >
                    <option value="comma">1,000</option>
                    <option value="plain">1000</option>
                  </select>
                </div>
              </div>

              <div className="sidebar-save" ref={saveMenuRef}>
                <div className={saveMenuOpen ? 'sidebar-save-split is-open' : 'sidebar-save-split'}>
                  <button
                    className="sidebar-save-main"
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={handleSaveMainClick}
                    disabled={isSaving || isSigningIn || ((!accessToken || !ledgerTarget) && !isAuthReady)}
                  >
                    <span className="sidebar-save-main-content">
                      <svg className="sidebar-save-google-icon" width="14" height="14" viewBox="0 0 48 48" aria-hidden="true">
                        <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                        <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                        <path fill="#FBBC05" d="M10.53 28.59a14.5 14.5 0 0 1 0-9.18l-7.98-6.19a24.0 24.0 0 0 0 0 21.56l7.98-6.19z"/>
                        <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
                      </svg>
                      <span>{isSigningIn ? 'Connecting...' : !accessToken || !ledgerTarget ? 'Connect to Sync' : isSaving ? 'Saving...' : 'Save & Download'}</span>
                    </span>
                  </button>
                  <button
                    className="sidebar-save-caret"
                    type="button"
                    title="Download options"
                    aria-label="Open download options"
                    aria-haspopup="menu"
                    aria-expanded={saveMenuOpen}
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={toggleSaveMenu}
                  >
                    <svg width="10" height="10" viewBox="0 0 10 10" fill="none" aria-hidden="true">
                      <path d="M2 4L5 7L8 4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </button>
                </div>

                {saveMenuVisible ? (
                  <div
                    className={saveMenuOpen ? 'sidebar-save-menu is-open' : 'sidebar-save-menu is-closing'}
                    role="menu"
                    aria-label="Download as"
                  >
                    <button
                      className="sidebar-save-option"
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleDownload('csv')}
                    >
                      Download CSV (.csv)
                    </button>
                    <button
                      className="sidebar-save-option"
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleDownload('tsv')}
                    >
                      Download TSV (.tsv)
                    </button>
                    <button
                      className="sidebar-save-option"
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleDownload('xlsx')}
                    >
                      Download Excel (.xlsx)
                    </button>
                    <button
                      className="sidebar-save-option"
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleDownload('ods')}
                    >
                      Download OpenDocument (.ods)
                    </button>
                    {accessToken ? (
                      <button
                        className="sidebar-save-option"
                        type="button"
                        role="menuitem"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => {
                          closeSaveMenu();
                          void handleSignOut();
                        }}
                      >
                        Disconnect Google
                      </button>
                    ) : null}
                  </div>
                ) : null}
              </div>
            </div>

            <footer className="sidebar-footer">
              <span>{statusMessage}</span>
              {errorMessage ? <span className="error">{errorMessage}</span> : null}
            </footer>
          </aside>

        <main ref={gridStageRef} className="grid-stage" aria-live="polite">
          {formulaPopup ? (
            <aside className="formula-popover" aria-live="polite" style={formulaPopup.style}>
              <div className="formula-popover-header">Formula Helper</div>
              {formulaPopup.mode === 'function' ? (
                <div>
                  <div className="formula-signature">
                    <span className="formula-fn-name">{formulaPopup.exact.name}</span>
                    <span>(</span>
                    {formulaPopup.exact.params.map((param, index) => (
                      <span key={`${formulaPopup.exact.name}-${param}-${index}`}>
                        {index > 0 ? <span className="formula-separator">, </span> : null}
                        <span className={index === formulaPopup.activeParamIndex ? 'formula-param active' : 'formula-param'}>
                          {param}
                        </span>
                      </span>
                    ))}
                    <span>)</span>
                  </div>
                  <div className="formula-description">{formulaPopup.exact.description}</div>
                </div>
              ) : (
                <div>
                  <div className="formula-suggestions">
                    {formulaPopup.suggestions.map((doc) => (
                      <div key={doc.name} className="formula-suggestion">
                        <div className="formula-signature">{formatFormulaSignature(doc)}</div>
                        <div className="formula-description">{doc.description}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </aside>
          ) : null}

          <div className="grid-frame">
            <HotTable
              key={`grid-${selectedPreset}-${gridEpoch}`}
              ref={hotRef}
              data={gridData}
              colHeaders={columnHeaders}
              colWidths={columnWidths}
              rowHeaders
              rowHeaderWidth={ROW_HEADER_WIDTH}
              width="100%"
              height="100%"
              outsideClickDeselects={false}
              stretchH="all"
              manualColumnResize
              contextMenu
              autoWrapRow
              autoWrapCol
              rowHeights={22}
              columnHeaderHeight={22}
              minRows={GRID_ROWS}
              minCols={GRID_COLUMNS}
              maxRows={GRID_ROWS}
              maxCols={GRID_COLUMNS}
              formulas={{
                engine: HyperFormula,
                sheetName: 'Entry',
              }}
              hiddenColumns={{
                columns: hiddenColumns,
                indicators: false,
              }}
              mergeCells={mergeCells}
              afterBeginEditing={handleAfterBeginEditing}
              beforeKeyDown={handleBeforeKeyDown}
              beforeOnCellMouseDown={handleBeforeOnCellMouseDown}
              beforeRenderer={handleBeforeRenderer}
              afterRenderer={handleAfterRenderer}
              afterSelectionEnd={handleAfterSelectionEnd}
              afterDeselect={handleAfterDeselect}
              licenseKey="non-commercial-and-evaluation"
            />
          </div>
        </main>
      </div>
    </div>
  );

  if (!isWebLanding) {
    return appContent;
  }

  return (
    <div className="web-shell">
      <header className="web-hero">
        <h1 className="web-title">Beanfolio</h1>
        <p className="web-subtitle">Your mini spreadsheet app for quick calculations</p>
      </header>

      <div className={sidebarOpen ? 'web-stage is-sidebar-open' : 'web-stage'}>
        <div className="web-controls">
          <div className="popout-arrow" aria-hidden="true">
            <span className="popout-arrow-text">Need a separate window?</span>
            <svg className="popout-arrow-icon" viewBox="0 0 120 14" fill="none" focusable="false">
              <path d="M1 7H112" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
              <path d="M105 2L112 7L105 12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
          <button className="popout-button" type="button" onClick={handleOpenPopOut}>
            Pop Out App
          </button>
        </div>

        <div className={sidebarOpen ? 'embedded-app-frame is-sidebar-open' : 'embedded-app-frame'}>
          {appContent}
        </div>

        {popOutBlockedMessage ? (
          <p className="popout-message" role="status">{popOutBlockedMessage}</p>
        ) : null}
      </div>
    </div>
  );
}

function getHotInstance(ref: MutableRefObject<any>): Handsontable | null {
  return ref.current?.hotInstance ?? null;
}

function sourceHasNonEmptyCells(source: unknown[][]): boolean {
  return source.some((row) => row.some((cell) => normalizeDisplayValue(cell) !== null));
}

function hasMeaningfulValue(value: unknown): boolean {
  return normalizeDisplayValue(value) !== null;
}

function buildFormulaPopupState(
  input: string,
  caretPosition: number,
  container: HTMLElement | null,
): FormulaPopupState | null {
  const style = computeFormulaPopupStyle(container);
  const functionCallMatch = input.match(/^\s*=\s*([A-Za-z][A-Za-z0-9._]*)\s*\(/);

  if (functionCallMatch) {
    const functionName = functionCallMatch[1].toUpperCase();
    const exact = FORMULA_DOC_BY_NAME.get(functionName);

    if (!exact) {
      return null;
    }

    const openParenIndex = input.indexOf('(');
    const activeParamIndex = computeActiveParamIndex(input, openParenIndex + 1, caretPosition);

    return {
      mode: 'function',
      style,
      exact,
      activeParamIndex: Math.max(0, Math.min(activeParamIndex, exact.params.length - 1)),
    };
  }

  const suggestMatch = input.match(/^\s*=\s*([A-Za-z][A-Za-z0-9._]*)?\s*$/);
  if (!suggestMatch) {
    return null;
  }

  const query = (suggestMatch[1] ?? '').toUpperCase();
  if (query.length === 0) {
    return null;
  }

  const suggestions = FORMULA_DOCS.filter((doc) => doc.name.startsWith(query));
  if (suggestions.length === 0) {
    return null;
  }

  return {
    mode: 'suggest',
    style,
    query,
    suggestions,
  };

}

function computeFormulaPopupStyle(container: HTMLElement | null): CSSProperties {
  if (!container) {
    return {};
  }

  const rect = container.getBoundingClientRect();

  return {
    bottom: Math.max(window.innerHeight - rect.bottom + 8, 8),
    left: rect.left + 8,
    width: Math.max(Math.min(260, rect.width - 16), 180),
  };
}

function computeActiveParamIndex(input: string, argsStartIndex: number, caretPosition: number): number {
  let depth = 0;
  let commas = 0;
  const endIndex = Math.min(Math.max(caretPosition, argsStartIndex), input.length);

  for (let i = argsStartIndex; i < endIndex; i += 1) {
    const char = input[i];

    if (char === '(') {
      depth += 1;
      continue;
    }

    if (char === ')') {
      if (depth === 0) {
        break;
      }
      depth -= 1;
      continue;
    }

    if (char === ',' && depth === 0) {
      commas += 1;
    }
  }

  return commas;
}

function isFormulaInput(input: string): boolean {
  return input.trimStart().startsWith('=');
}

function updateFormulaReferences(
  hotRef: MutableRefObject<any>,
  activeReferencesRef: MutableRefObject<FormulaReferenceHighlight[]>,
  formula: string | null,
): void {
  if (!formula) {
    if (activeReferencesRef.current.length === 0) return;
    activeReferencesRef.current = [];
    getHotInstance(hotRef)?.render();
    return;
  }

  const parsed = parseFormulaReferences(formula);
  const colorMap = new Map<string, number>();
  let nextColor = 0;

  const highlights: FormulaReferenceHighlight[] = parsed.map((p) => {
    const key = `${p.row},${p.col}`;
    if (!colorMap.has(key)) {
      colorMap.set(key, nextColor % REFERENCE_COLORS.length);
      nextColor++;
    }
    return { ref: p.ref, row: p.row, col: p.col, colorIndex: colorMap.get(key)! };
  });

  activeReferencesRef.current = highlights;
  getHotInstance(hotRef)?.render();
}

function updateFormulaOverlay(
  overlayRef: MutableRefObject<HTMLDivElement | null>,
  textarea: HTMLTextAreaElement | null | undefined,
  formula: string | null,
  highlights: FormulaReferenceHighlight[],
): void {
  if (!textarea || !formula || highlights.length === 0) {
    if (overlayRef.current) {
      overlayRef.current.style.display = 'none';
    }
    if (textarea) {
      textarea.style.removeProperty('color');
      textarea.style.removeProperty('caret-color');
      textarea.style.removeProperty('background');
      textarea.style.removeProperty('position');
      textarea.style.removeProperty('z-index');
    }
    return;
  }

  let overlay = overlayRef.current;
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.className = 'formula-overlay';
    overlayRef.current = overlay;
  }

  const holder = textarea.parentElement;
  if (holder && overlay.parentElement !== holder) {
    holder.appendChild(overlay);
  }

  // Build color map
  const parsed = parseFormulaReferences(formula);
  const colorMap = new Map<string, number>();
  let nextColor = 0;
  for (const p of parsed) {
    const key = `${p.row},${p.col}`;
    if (!colorMap.has(key)) {
      colorMap.set(key, nextColor % REFERENCE_COLORS.length);
      nextColor++;
    }
  }

  // Build highlighted HTML — color only, NO font-weight changes
  let html = '';
  let lastIndex = 0;
  for (const p of parsed) {
    if (p.start > lastIndex) {
      html += escapeHtml(formula.slice(lastIndex, p.start));
    }
    const ci = colorMap.get(`${p.row},${p.col}`) ?? 0;
    html += `<span style="color:${REFERENCE_COLORS[ci].text}">${escapeHtml(p.ref)}</span>`;
    lastIndex = p.end;
  }
  if (lastIndex < formula.length) {
    html += escapeHtml(formula.slice(lastIndex));
  }
  overlay.innerHTML = html;

  // Copy every text-affecting property from the textarea
  const cs = window.getComputedStyle(textarea);
  const borderT = parseFloat(cs.borderTopWidth) || 0;
  const borderL = parseFloat(cs.borderLeftWidth) || 0;
  const borderR = parseFloat(cs.borderRightWidth) || 0;
  const borderB = parseFloat(cs.borderBottomWidth) || 0;

  overlay.style.display = 'block';
  overlay.style.position = 'absolute';
  overlay.style.top = (textarea.offsetTop + borderT) + 'px';
  overlay.style.left = (textarea.offsetLeft + borderL) + 'px';
  overlay.style.width = (textarea.offsetWidth - borderL - borderR) + 'px';
  overlay.style.height = (textarea.offsetHeight - borderT - borderB) + 'px';
  overlay.style.font = cs.font;
  overlay.style.padding = cs.padding;
  overlay.style.border = 'none';
  overlay.style.boxSizing = 'border-box';
  overlay.style.lineHeight = cs.lineHeight;
  overlay.style.letterSpacing = cs.letterSpacing;
  overlay.style.wordSpacing = cs.wordSpacing;
  overlay.style.textIndent = cs.textIndent;
  overlay.style.textTransform = cs.textTransform;
  overlay.style.whiteSpace = cs.whiteSpace;
  overlay.style.overflow = 'hidden';
  overlay.style.pointerEvents = 'none';
  overlay.style.backgroundColor = '#fff';
  overlay.style.zIndex = '0';

  // Textarea on top: transparent text, visible caret
  textarea.style.position = 'relative';
  textarea.style.zIndex = '1';
  textarea.style.color = 'transparent';
  textarea.style.caretColor = '#000';
  textarea.style.background = 'transparent';
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function syncFormulaSession(
  ref: MutableRefObject<FormulaEditSession | null>,
  editorContext: OpenEditorContext,
): FormulaEditSession {
  const existing = ref.current;
  if (existing && existing.editRow === editorContext.row && existing.editCol === editorContext.col) {
    return existing;
  }

  const created: FormulaEditSession = {
    editRow: editorContext.row,
    editCol: editorContext.col,
    cursorRow: editorContext.row,
    cursorCol: editorContext.col,
    pendingReferenceRange: null,
    lastReferenceInput: null,
  };

  ref.current = created;
  return created;
}

function isArrowKey(key: string): key is 'ArrowUp' | 'ArrowDown' | 'ArrowLeft' | 'ArrowRight' {
  return key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight';
}

function moveReferenceCursor(
  row: number,
  col: number,
  key: 'ArrowUp' | 'ArrowDown' | 'ArrowLeft' | 'ArrowRight',
  maxVisibleCol: number,
): [number, number] {
  if (key === 'ArrowUp') {
    return [Math.max(row - 1, 0), col];
  }

  if (key === 'ArrowDown') {
    return [Math.min(row + 1, GRID_ROWS - 1), col];
  }

  if (key === 'ArrowLeft') {
    return [row, Math.max(col - 1, 0)];
  }

  return [row, Math.min(col + 1, Math.max(maxVisibleCol, 0))];
}

function resetsReferenceRange(key: string): boolean {
  return key.length === 1 || key === 'Backspace' || key === 'Delete';
}

const FORMULA_OPERATORS = new Set(['+', '-', '*', '/', '^', '(', ',']);
function isFormulaOperator(key: string): boolean {
  return FORMULA_OPERATORS.has(key);
}

function upsertFormulaReference(
  editorContext: OpenEditorContext,
  session: FormulaEditSession,
  reference: string,
): void {
  const currentInput = String(editorContext.editor.getValue() ?? '');
  const pending = session.pendingReferenceRange;

  let start: number;
  let end: number;
  let nextInput: string;

  if (pending && pending.start >= 0 && pending.end >= pending.start && pending.end <= currentInput.length) {
    start = pending.start;
    end = pending.end;
    nextInput = `${currentInput.slice(0, start)}${reference}${currentInput.slice(end)}`;
  } else {
    const detected = detectReferenceTokenRange(currentInput, editorContext.caretPosition);
    if (detected) {
      start = detected.start;
      end = detected.end;
      nextInput = `${currentInput.slice(0, start)}${reference}${currentInput.slice(end)}`;
    } else {
      start = editorContext.caretPosition;
      end = editorContext.caretPosition;
      nextInput = `${currentInput.slice(0, start)}${reference}${currentInput.slice(end)}`;
    }
  }

  const nextEnd = start + reference.length;
  editorContext.editor.setValue(nextInput);
  if (editorContext.editor.TEXTAREA) {
    editorContext.editor.TEXTAREA.setSelectionRange(nextEnd, nextEnd);
  }

  session.pendingReferenceRange = {
    start,
    end: nextEnd,
  };
}

function detectReferenceTokenRange(
  input: string,
  caretPosition: number,
): { start: number; end: number } | null {
  const isRefChar = (char: string): boolean => /[A-Za-z0-9:$]/.test(char);

  let seed = -1;
  if (caretPosition > 0 && isRefChar(input[caretPosition - 1] ?? '')) {
    seed = caretPosition - 1;
  } else if (caretPosition < input.length && isRefChar(input[caretPosition] ?? '')) {
    seed = caretPosition;
  }

  if (seed < 0) {
    return null;
  }

  let start = seed;
  let end = seed + 1;

  while (start > 0 && isRefChar(input[start - 1] ?? '')) {
    start -= 1;
  }
  while (end < input.length && isRefChar(input[end] ?? '')) {
    end += 1;
  }

  const candidate = input.slice(start, end);
  if (!A1_REFERENCE_TOKEN_RE.test(candidate)) {
    return null;
  }

  return {
    start,
    end,
  };
}

type AbsoluteReferenceMode = 'relative' | 'absolute' | 'row' | 'column';

interface ParsedA1ReferencePart {
  colAbsolute: boolean;
  column: string;
  rowAbsolute: boolean;
  row: string;
}

function cycleAbsoluteReferenceToken(token: string): string | null {
  if (!A1_REFERENCE_TOKEN_RE.test(token)) {
    return null;
  }

  const references = token.split(':');
  const parsedParts: ParsedA1ReferencePart[] = [];

  for (const reference of references) {
    const parsed = parseA1ReferencePart(reference);
    if (!parsed) {
      return null;
    }
    parsedParts.push(parsed);
  }

  const nextMode = nextAbsoluteReferenceMode(getAbsoluteReferenceMode(parsedParts[0]));
  return parsedParts.map((part) => formatA1ReferencePart(part, nextMode)).join(':');
}

function parseA1ReferencePart(reference: string): ParsedA1ReferencePart | null {
  const match = reference.match(A1_REFERENCE_PART_RE);
  if (!match) {
    return null;
  }

  const [, colPrefix, column, rowPrefix, row] = match;
  return {
    colAbsolute: colPrefix === '$',
    column: column.toUpperCase(),
    rowAbsolute: rowPrefix === '$',
    row,
  };
}

function getAbsoluteReferenceMode(reference: ParsedA1ReferencePart): AbsoluteReferenceMode {
  if (reference.colAbsolute && reference.rowAbsolute) {
    return 'absolute';
  }

  if (!reference.colAbsolute && reference.rowAbsolute) {
    return 'row';
  }

  if (reference.colAbsolute && !reference.rowAbsolute) {
    return 'column';
  }

  return 'relative';
}

function nextAbsoluteReferenceMode(mode: AbsoluteReferenceMode): AbsoluteReferenceMode {
  if (mode === 'relative') {
    return 'absolute';
  }

  if (mode === 'absolute') {
    return 'row';
  }

  if (mode === 'row') {
    return 'column';
  }

  return 'relative';
}

function formatA1ReferencePart(reference: ParsedA1ReferencePart, mode: AbsoluteReferenceMode): string {
  const columnPrefix = mode === 'absolute' || mode === 'column' ? '$' : '';
  const rowPrefix = mode === 'absolute' || mode === 'row' ? '$' : '';
  return `${columnPrefix}${reference.column}${rowPrefix}${reference.row}`;
}

function toA1Reference(row: number, col: number): string {
  return `${columnToLetters(col)}${row + 1}`;
}

function toAbsoluteA1Reference(row: number, col: number): string {
  return `$${columnToLetters(col)}$${row + 1}`;
}

function lettersToColumnIndex(letters: string): number {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
}

const A1_INLINE_RE = /\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?/g;

interface ParsedFormulaRef {
  ref: string;
  start: number;
  end: number;
  row: number;
  col: number;
}

function parseFormulaReferences(input: string): ParsedFormulaRef[] {
  A1_INLINE_RE.lastIndex = 0;
  const results: ParsedFormulaRef[] = [];
  let match: RegExpExecArray | null;

  while ((match = A1_INLINE_RE.exec(input)) !== null) {
    const ref = match[0];
    const mainPart = ref.includes(':') ? ref.split(':')[0] : ref;
    const parsed = parseA1ReferencePart(mainPart);
    if (!parsed) continue;
    const row = parseInt(parsed.row, 10) - 1;
    const col = lettersToColumnIndex(parsed.column);
    if (row < 0 || col < 0) continue;
    results.push({ ref, start: match.index, end: match.index + ref.length, row, col });
  }

  return results;
}

function shiftHexColorLightness(color: string, steps: number): string {
  const normalized = normalizeHexColor(color);
  if (!normalized) {
    return color;
  }

  const [r, g, b] = normalized;
  const [h, s, l] = rgbToHsl(r, g, b);
  const nextLightness = clamp(l + steps * 0.08, 0, 1);
  const [nextR, nextG, nextB] = hslToRgb(h, s, nextLightness);
  return rgbToHex(nextR, nextG, nextB);
}

function normalizeHexColor(color: string): [number, number, number] | null {
  const source = color.trim().replace(/^#/, '');
  const sixDigit =
    source.length === 3
      ? source.split('').map((char) => `${char}${char}`).join('')
      : source;

  if (sixDigit.length !== 6 || /[^0-9a-f]/i.test(sixDigit)) {
    return null;
  }

  const r = Number.parseInt(sixDigit.slice(0, 2), 16);
  const g = Number.parseInt(sixDigit.slice(2, 4), 16);
  const b = Number.parseInt(sixDigit.slice(4, 6), 16);
  return [r, g, b];
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${toHexByte(r)}${toHexByte(g)}${toHexByte(b)}`;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const nr = r / 255;
  const ng = g / 255;
  const nb = b / 255;

  const max = Math.max(nr, ng, nb);
  const min = Math.min(nr, ng, nb);
  const delta = max - min;

  let hue = 0;
  const lightness = (max + min) / 2;

  if (delta === 0) {
    return [0, 0, lightness];
  }

  const saturation = delta / (1 - Math.abs(2 * lightness - 1));

  if (max === nr) {
    hue = ((ng - nb) / delta) % 6;
  } else if (max === ng) {
    hue = (nb - nr) / delta + 2;
  } else {
    hue = (nr - ng) / delta + 4;
  }

  hue = (hue * 60 + 360) % 360;
  return [hue, saturation, lightness];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const value = Math.round(l * 255);
    return [value, value, value];
  }

  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const segment = h / 60;
  const x = chroma * (1 - Math.abs((segment % 2) - 1));

  let r1 = 0;
  let g1 = 0;
  let b1 = 0;

  if (segment >= 0 && segment < 1) {
    r1 = chroma;
    g1 = x;
  } else if (segment >= 1 && segment < 2) {
    r1 = x;
    g1 = chroma;
  } else if (segment >= 2 && segment < 3) {
    g1 = chroma;
    b1 = x;
  } else if (segment >= 3 && segment < 4) {
    g1 = x;
    b1 = chroma;
  } else if (segment >= 4 && segment < 5) {
    r1 = x;
    b1 = chroma;
  } else {
    r1 = chroma;
    b1 = x;
  }

  const m = l - chroma / 2;
  return [
    Math.round((r1 + m) * 255),
    Math.round((g1 + m) * 255),
    Math.round((b1 + m) * 255),
  ];
}

function toHexByte(value: number): string {
  return clamp(Math.round(value), 0, 255).toString(16).padStart(2, '0');
}

function formatLightnessLabel(value: number): string {
  if (value === 0) {
    return 'Lightness 0';
  }

  return value > 0 ? `Lightness +${value}` : `Lightness ${value}`;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function columnToLetters(index: number): string {
  let n = index + 1;
  let letters = '';

  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }

  return letters;
}

function getOpenEditorContext(
  hot: Handsontable,
): OpenEditorContext | null {
  const editor = hot.getActiveEditor() as (Handsontable.editors.BaseEditor & { TEXTAREA?: HTMLTextAreaElement }) | null;

  if (!editor || !editor.isOpened()) {
    return null;
  }

  const textarea = editor.TEXTAREA;
  const input = textarea ? textarea.value : String(editor.getValue() ?? '');
  const caretPosition =
    textarea && typeof textarea.selectionStart === 'number'
      ? textarea.selectionStart
      : input.length;

  return {
    editor,
    input,
    caretPosition,
    row: editor.row,
    col: editor.col,
  };
}

const ROUND_WRAP_RE = /^(\s*=\s*)ROUND\s*\(([\s\S]+),\s*\d+\s*\)\s*$/i;

function wrapFormulaInRound(formula: string, decimals: number): string {
  const match = formula.match(ROUND_WRAP_RE);
  if (match) {
    return `${match[1]}ROUND(${match[2].trim()},${decimals})`;
  }
  const inner = formula.replace(/^\s*=\s*/, '');
  return `=ROUND(${inner},${decimals})`;
}

function formatFormulaSignature(doc: FormulaDoc): string {
  return `${doc.name}(${doc.params.join(', ')})`;
}

function getPresetOptionLabel(preset: PresetType): string {
  const option = PRESET_OPTIONS.find((item) => item.value === preset);
  return option?.label ?? 'Preset';
}

function buildSnapshot(hot: Handsontable, columnCount: number): CellSnapshot[][] {
  return Array.from({ length: GRID_ROWS }, (_, rowIndex) =>
    Array.from({ length: columnCount }, (_, columnIndex) => {
      const sourceValue = hot.getSourceDataAtCell(rowIndex, columnIndex);
      const displayValue = hot.getDataAtCell(rowIndex, columnIndex);

      return {
        displayValue: normalizeDisplayValue(displayValue),
        formula: isFormula(sourceValue) ? sourceValue : undefined,
      };
    }),
  );
}

function trimTrailingEmptyRows(cells: CellSnapshot[][]): CellSnapshot[][] {
  for (let rowIndex = cells.length - 1; rowIndex >= 0; rowIndex -= 1) {
    const hasContent = cells[rowIndex].some((cell) => cell.displayValue !== null || Boolean(cell.formula));
    if (hasContent) {
      return cells.slice(0, rowIndex + 1);
    }
  }

  return [];
}

interface ExportFile {
  fileName: string;
  mimeType: string;
  payload: Array<string | ArrayBuffer>;
}

function buildExportFile(cells: CellSnapshot[][], format: ExportFormat, label: string): ExportFile {
  const preparedLabel = sanitizeFileName(label);
  const baseName = preparedLabel.length > 0 ? preparedLabel : `beanfolio-${new Date().toISOString().replace(/[:]/g, '-').slice(0, 19)}`;

  if (format === 'csv') {
    return {
      fileName: `${baseName}.csv`,
      mimeType: 'text/csv;charset=utf-8',
      payload: [`\uFEFF${serializeDelimited(cells, ',')}`],
    };
  }

  if (format === 'tsv') {
    return {
      fileName: `${baseName}.tsv`,
      mimeType: 'text/tab-separated-values;charset=utf-8',
      payload: [`\uFEFF${serializeDelimited(cells, '\t')}`],
    };
  }

  if (format === 'xlsx') {
    return {
      fileName: `${baseName}.xlsx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      payload: [toArrayBuffer(buildXlsxWorkbookArchive(cells))],
    };
  }

  return {
    fileName: `${baseName}.ods`,
    mimeType: 'application/vnd.oasis.opendocument.spreadsheet',
    payload: [toArrayBuffer(buildOdsWorkbookArchive(cells))],
  };
}

function sanitizeFileName(input: string): string {
  return input
    .trim()
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/[.\s]+$/g, '')
    .slice(0, 80);
}

function serializeDelimited(cells: CellSnapshot[][], delimiter: string): string {
  return cells
    .map((row) => row.map((cell) => encodeDelimitedCell(resolveSnapshotValue(cell), delimiter)).join(delimiter))
    .join('\r\n');
}

function resolveSnapshotValue(cell: CellSnapshot): string | number | boolean | null {
  if (cell.formula) {
    return cell.formula;
  }

  return cell.displayValue;
}

function encodeDelimitedCell(value: string | number | boolean | null, delimiter: string): string {
  if (value === null) {
    return '';
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  const needsQuoting =
    value.includes('"')
    || value.includes('\n')
    || value.includes('\r')
    || value.includes(delimiter)
    || value.trim() !== value;

  if (!needsQuoting) {
    return value;
  }

  return `"${value.replace(/"/g, '""')}"`;
}

function buildXlsxWorkbookArchive(cells: CellSnapshot[][]): Uint8Array {
  const entries: ZipEntry[] = [
    { name: '[Content_Types].xml', data: encodeUtf8(buildXlsxContentTypesXml()) },
    { name: '_rels/.rels', data: encodeUtf8(buildXlsxRootRelationshipsXml()) },
    { name: 'xl/workbook.xml', data: encodeUtf8(buildXlsxWorkbookXml()) },
    { name: 'xl/_rels/workbook.xml.rels', data: encodeUtf8(buildXlsxWorkbookRelationshipsXml()) },
    { name: 'xl/styles.xml', data: encodeUtf8(buildXlsxStylesXml()) },
    { name: 'xl/worksheets/sheet1.xml', data: encodeUtf8(buildXlsxSheetXml(cells)) },
  ];

  return createZipArchive(entries);
}

function buildXlsxContentTypesXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '</Types>',
  ].join('');
}

function buildXlsxRootRelationshipsXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>',
  ].join('');
}

function buildXlsxWorkbookXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>',
    '</workbook>',
  ].join('');
}

function buildXlsxWorkbookRelationshipsXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '</Relationships>',
  ].join('');
}

function buildXlsxStylesXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>',
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    '</styleSheet>',
  ].join('');
}

function buildXlsxSheetXml(cells: CellSnapshot[][]): string {
  const rows = cells
    .map((row, rowIndex) => buildXlsxRowXml(row, rowIndex))
    .join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<sheetData>${rows}</sheetData>`,
    '</worksheet>',
  ].join('');
}

function buildXlsxRowXml(row: CellSnapshot[], rowIndex: number): string {
  const cellsXml = row
    .map((cell, columnIndex) => buildXlsxCellXml(cell, rowIndex, columnIndex))
    .filter((item) => item.length > 0)
    .join('');

  if (cellsXml.length === 0) {
    return `<row r="${rowIndex + 1}"/>`;
  }

  return `<row r="${rowIndex + 1}">${cellsXml}</row>`;
}

function buildXlsxCellXml(cell: CellSnapshot, rowIndex: number, columnIndex: number): string {
  const cellRef = `${columnToLetters(columnIndex)}${rowIndex + 1}`;

  if (cell.formula) {
    const formula = cell.formula.trim().replace(/^=/, '');
    if (formula.length === 0) {
      return '';
    }

    return `<c r="${cellRef}"><f>${escapeXml(formula)}</f></c>`;
  }

  const value = cell.displayValue;
  if (value === null) {
    return '';
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      return '';
    }

    return `<c r="${cellRef}"><v>${value}</v></c>`;
  }

  if (typeof value === 'boolean') {
    return `<c r="${cellRef}" t="b"><v>${value ? 1 : 0}</v></c>`;
  }

  const preserveSpace = shouldPreserveXmlSpace(value);
  const spaceAttr = preserveSpace ? ' xml:space="preserve"' : '';
  return `<c r="${cellRef}" t="inlineStr"><is><t${spaceAttr}>${escapeXml(value)}</t></is></c>`;
}

function buildOdsWorkbookArchive(cells: CellSnapshot[][]): Uint8Array {
  const entries: ZipEntry[] = [
    { name: 'mimetype', data: encodeUtf8('application/vnd.oasis.opendocument.spreadsheet') },
    { name: 'content.xml', data: encodeUtf8(buildOdsContentXml(cells)) },
    { name: 'styles.xml', data: encodeUtf8(buildOdsStylesXml()) },
    { name: 'meta.xml', data: encodeUtf8(buildOdsMetaXml()) },
    { name: 'settings.xml', data: encodeUtf8(buildOdsSettingsXml()) },
    { name: 'META-INF/manifest.xml', data: encodeUtf8(buildOdsManifestXml()) },
  ];

  return createZipArchive(entries);
}

function buildOdsContentXml(cells: CellSnapshot[][]): string {
  const rows = cells.map((row) => buildOdsRowXml(row)).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<office:document-content',
    ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
    ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
    ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
    ' office:version="1.2">',
    '<office:body><office:spreadsheet><table:table table:name="Sheet1">',
    rows,
    '</table:table></office:spreadsheet></office:body>',
    '</office:document-content>',
  ].join('');
}

function buildOdsRowXml(row: CellSnapshot[]): string {
  const rowCells = row.map((cell) => buildOdsCellXml(cell)).join('');
  return `<table:table-row>${rowCells}</table:table-row>`;
}

function buildOdsCellXml(cell: CellSnapshot): string {
  const value = resolveSnapshotValue(cell);

  if (value === null) {
    return '<table:table-cell/>';
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      return '<table:table-cell/>';
    }

    return `<table:table-cell office:value-type="float" office:value="${value}"><text:p>${escapeXml(String(value))}</text:p></table:table-cell>`;
  }

  if (typeof value === 'boolean') {
    const normalized = value ? 'true' : 'false';
    return `<table:table-cell office:value-type="boolean" office:boolean-value="${normalized}"><text:p>${value ? 'TRUE' : 'FALSE'}</text:p></table:table-cell>`;
  }

  return `<table:table-cell office:value-type="string"><text:p>${escapeXml(value)}</text:p></table:table-cell>`;
}

function buildOdsStylesXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">',
    '<office:styles/>',
    '</office:document-styles>',
  ].join('');
}

function buildOdsMetaXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">',
    '<office:meta/>',
    '</office:document-meta>',
  ].join('');
}

function buildOdsSettingsXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">',
    '<office:settings/>',
    '</office:document-settings>',
  ].join('');
}

function buildOdsManifestXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">',
    '<manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/"/>',
    '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>',
    '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>',
    '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>',
    '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>',
    '</manifest:manifest>',
  ].join('');
}

interface ZipEntry {
  name: string;
  data: Uint8Array;
}

const UTF8_ENCODER = new TextEncoder();

function encodeUtf8(value: string): Uint8Array {
  return UTF8_ENCODER.encode(value);
}

function createZipArchive(entries: ZipEntry[]): Uint8Array {
  const localChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];
  const centralDirectoryEntries: Array<{
    nameBytes: Uint8Array;
    crc32: number;
    size: number;
    localOffset: number;
  }> = [];

  const timestamp = new Date();
  const dosTime = toDosTime(timestamp);
  const dosDate = toDosDate(timestamp);
  let offset = 0;

  entries.forEach((entry) => {
    const nameBytes = encodeUtf8(entry.name);
    const size = entry.data.length;
    const crc = crc32(entry.data);
    const localHeader = buildZipLocalFileHeader(nameBytes.length, crc, size, dosTime, dosDate);

    localChunks.push(localHeader, nameBytes, entry.data);
    centralDirectoryEntries.push({
      nameBytes,
      crc32: crc,
      size,
      localOffset: offset,
    });

    offset += localHeader.length + nameBytes.length + size;
  });

  let centralDirectorySize = 0;
  centralDirectoryEntries.forEach((entry) => {
    const centralHeader = buildZipCentralDirectoryHeader(
      entry.nameBytes.length,
      entry.crc32,
      entry.size,
      dosTime,
      dosDate,
      entry.localOffset,
    );

    centralChunks.push(centralHeader, entry.nameBytes);
    centralDirectorySize += centralHeader.length + entry.nameBytes.length;
  });

  const endOfCentralDirectory = buildZipEndOfCentralDirectoryRecord(
    entries.length,
    centralDirectorySize,
    offset,
  );

  return concatUint8Arrays([...localChunks, ...centralChunks, endOfCentralDirectory]);
}

function buildZipLocalFileHeader(
  fileNameLength: number,
  crc: number,
  size: number,
  dosTime: number,
  dosDate: number,
): Uint8Array {
  const buffer = new Uint8Array(30);
  const view = new DataView(buffer.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, dosTime, true);
  view.setUint16(12, dosDate, true);
  view.setUint32(14, crc >>> 0, true);
  view.setUint32(18, size, true);
  view.setUint32(22, size, true);
  view.setUint16(26, fileNameLength, true);
  view.setUint16(28, 0, true);
  return buffer;
}

function buildZipCentralDirectoryHeader(
  fileNameLength: number,
  crc: number,
  size: number,
  dosTime: number,
  dosDate: number,
  localOffset: number,
): Uint8Array {
  const buffer = new Uint8Array(46);
  const view = new DataView(buffer.buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, dosTime, true);
  view.setUint16(14, dosDate, true);
  view.setUint32(16, crc >>> 0, true);
  view.setUint32(20, size, true);
  view.setUint32(24, size, true);
  view.setUint16(28, fileNameLength, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, localOffset, true);
  return buffer;
}

function buildZipEndOfCentralDirectoryRecord(
  entryCount: number,
  centralDirectorySize: number,
  centralDirectoryOffset: number,
): Uint8Array {
  const buffer = new Uint8Array(22);
  const view = new DataView(buffer.buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, entryCount, true);
  view.setUint16(10, entryCount, true);
  view.setUint32(12, centralDirectorySize, true);
  view.setUint32(16, centralDirectoryOffset, true);
  view.setUint16(20, 0, true);
  return buffer;
}

function concatUint8Arrays(chunks: Uint8Array[]): Uint8Array {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(totalLength);
  let offset = 0;

  chunks.forEach((chunk) => {
    merged.set(chunk, offset);
    offset += chunk.length;
  });

  return merged;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}

function toDosTime(date: Date): number {
  const seconds = Math.floor(date.getSeconds() / 2);
  return ((date.getHours() & 0x1f) << 11) | ((date.getMinutes() & 0x3f) << 5) | (seconds & 0x1f);
}

function toDosDate(date: Date): number {
  const year = Math.max(date.getFullYear(), 1980);
  return (((year - 1980) & 0x7f) << 9) | (((date.getMonth() + 1) & 0x0f) << 5) | (date.getDate() & 0x1f);
}

let _crc32Table: Uint32Array | null = null;

function getCrc32Table(): Uint32Array {
  return (_crc32Table ??= buildCrc32Table());
}

function buildCrc32Table(): Uint32Array {
  const table = new Uint32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
    }
    table[index] = value >>> 0;
  }

  return table;
}

function crc32(data: Uint8Array): number {
  const table = getCrc32Table();
  let checksum = 0xffffffff;

  for (let index = 0; index < data.length; index += 1) {
    const tableIndex = (checksum ^ data[index]) & 0xff;
    checksum = (checksum >>> 8) ^ table[tableIndex];
  }

  return (checksum ^ 0xffffffff) >>> 0;
}

function shouldPreserveXmlSpace(value: string): boolean {
  return value.trim() !== value || /[\n\r\t]| {2,}/.test(value);
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function downloadFile(file: ExportFile): void {
  const blob = new Blob(file.payload, { type: file.mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = file.fileName;
  anchor.style.display = 'none';
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1500);
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return 'Unexpected error.';
}

export default App;
