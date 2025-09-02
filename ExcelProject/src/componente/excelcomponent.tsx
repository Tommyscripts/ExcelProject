import React, { useState, useRef, useEffect, useCallback } from 'react';
import { TooltipCooldown, Toolbar } from './Auxiliares/TooltipCooldown';
import Modal from './Auxiliares/Modal';
// Iconos antiguos eliminados junto con el menú viejo
import * as XLSX from 'xlsx';
import Cell from './Auxiliares/Cell';
import { useVirtualizer } from '@tanstack/react-virtual';

const INITIAL_COLS = 25;
const INITIAL_ROWS = 25;

function getColName(idx: number) {
  return String.fromCharCode(65 + idx);
}

const createInitialData = (rows: number, cols: number) =>
  Array.from({ length: rows }, () => Array(cols).fill(''));

const ExcelComponent: React.FC = () => {
  // Estado para modo compacto/expandido de funciones
  const [compactMode, setCompactMode] = useState(true);
  // Default = light (false). We ignore any previous localStorage value on load to ensure default light.
  const _initDark = false;
  const [darkMode, setDarkMode] = useState<boolean>(_initDark);

  // Cambia la clase del body para modo claro/oscuro
  useEffect(() => {
    const root = document.documentElement;
    if (darkMode) {
      root.classList.add('dark');
      try { window.localStorage?.setItem('darkMode', 'true'); } catch (e) { }
    } else {
      root.classList.remove('dark');
      try { window.localStorage?.setItem('darkMode', 'false'); } catch (e) { }
    }
  }, [darkMode]);

  // Ensure body inline background/color reflect the current mode so default is light.
  useEffect(() => {
    try {
      if (darkMode) {
        document.body.style.background = '#0f1720'; // gray-900
        document.body.style.color = '#f8fafc';
      } else {
        document.body.style.background = '#ffffff';
        document.body.style.color = '#111827';
      }
    } catch (e) { /* ignore */ }
  }, [darkMode]);
  // Clase común para botones (alineación/altura consistente)
  const BTN = 'inline-flex items-center justify-center h-9 px-3 rounded text-sm leading-none align-middle font-medium';
  const [data, setData] = useState<string[][]>(createInitialData(INITIAL_ROWS, INITIAL_COLS));
  const [selected, setSelected] = useState<{ row: number; col: number } | null>({ row: 0, col: 0 });
  // Para selección múltiple
  const [selectionStart, setSelectionStart] = useState<{ row: number; col: number } | null>(null);
  const [selectionEnd, setSelectionEnd] = useState<{ row: number; col: number } | null>(null);
  // Para resize de columnas
  const [colWidths, setColWidths] = useState<number[]>(Array(INITIAL_COLS).fill(96));
  const resizingCol = useRef<number | null>(null);
  const startX = useRef<number>(0);
  const startWidth = useRef<number>(0);
  // Flag para saber si el usuario mantiene el mouse pulsado (drag selection)
  const isMouseDown = useRef<boolean>(false);
  // Estado reactivo para cuando el usuario está arrastrando (permite re-render para la animación)
  const [isDragging, setIsDragging] = useState<boolean>(false);
  // Indica si hubo movimiento (drag) desde el último mousedown, para evitar que el click posterior lo interprete como un nuevo click
  const hadDragSinceMouseDown = useRef<boolean>(false);
  // Track Shift key state globally to require Shift for multi-selection
  const shiftPressed = useRef<boolean>(false);

  // Header drag state: permite arrastrar sobre los encabezados para seleccionar columnas
  const headerDragging = useRef<boolean>(false);
  const headerDragStart = useRef<number | null>(null);
  // Track Ctrl/Cmd key for additive selection
  const ctrlPressed = useRef<boolean>(false);

  // Selecciones adicionales no contiguas (para Ctrl+click)
  const [extraSelections, setExtraSelections] = useState<Array<{ r1: number; c1: number; r2: number; c2: number }>>([]);

  React.useEffect(() => {
    const kd = (e: KeyboardEvent) => { if (e.key === 'Shift') shiftPressed.current = true; if (e.key === 'Control' || e.key === 'Meta') ctrlPressed.current = true; };
    const ku = (e: KeyboardEvent) => { if (e.key === 'Shift') shiftPressed.current = false; if (e.key === 'Control' || e.key === 'Meta') ctrlPressed.current = false; };
    window.addEventListener('keydown', kd);
    window.addEventListener('keyup', ku);
    return () => { window.removeEventListener('keydown', kd); window.removeEventListener('keyup', ku); };
  }, []);
  // Merges: regiones combinadas
  const [merges, setMerges] = useState<Array<{ r: number; c: number; rows: number; cols: number }>>([]);
  const [showCombineConfirm, setShowCombineConfirm] = useState(false);
  const [pendingCombineBox, setPendingCombineBox] = useState<{ r1: number; r2: number; c1: number; c2: number } | null>(null);

  // Helper: devuelve el merge que cubre (row,col) si existe
  const getCoveringMerge = useCallback((row: number, col: number) => {
    return merges.find(m => row >= m.r && row < m.r + m.rows && col >= m.c && col < m.c + m.cols) ?? null;
  }, [merges]);

  // --- Historial para Undo/Redo ---
  const HISTORY_LIMIT = 50;
  const [history, setHistory] = useState<any[]>([]);
  const [future, setFuture] = useState<any[]>([]);
  const pushHistory = useCallback((snapshot?: any) => {
    const snap = snapshot ?? { data: JSON.parse(JSON.stringify(data)), merges: JSON.parse(JSON.stringify(merges)), colWidths: [...colWidths], selected, selectionStart, selectionEnd };
    setHistory(h => {
      const next = [...h, snap];
      if (next.length > HISTORY_LIMIT) next.shift();
      return next;
    });
    setFuture([]);
  }, [data, merges, colWidths, selected, selectionStart, selectionEnd]);

  const undo = () => {
    setHistory(h => {
      if (h.length === 0) return h;
      const last = h[h.length - 1];
      setFuture(f => [{ data: JSON.parse(JSON.stringify(data)), merges: JSON.parse(JSON.stringify(merges)), colWidths: [...colWidths], selected, selectionStart, selectionEnd }, ...f]);
      if (last.data) setData(last.data);
      if (last.merges) setMerges(last.merges);
      if (last.colWidths) setColWidths(last.colWidths);
      setSelected(last.selected ?? null);
      setSelectionStart(last.selectionStart ?? null);
      setSelectionEnd(last.selectionEnd ?? null);
      return h.slice(0, -1);
    });
  };

  const redo = () => {
    setFuture(f => {
      if (f.length === 0) return f;
      const [first, ...rest] = f;
      setHistory(h => [...h, { data: JSON.parse(JSON.stringify(data)), merges: JSON.parse(JSON.stringify(merges)), colWidths: [...colWidths], selected, selectionStart, selectionEnd }]);
      if (first.data) setData(first.data);
      if (first.merges) setMerges(first.merges);
      if (first.colWidths) setColWidths(first.colWidths);
      setSelected(first.selected ?? null);
      setSelectionStart(first.selectionStart ?? null);
      setSelectionEnd(first.selectionEnd ?? null);
      return rest;
    });
  };

  // --- Find & Replace ---
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [matchRegex, setMatchRegex] = useState(false);
  const [searchMatches, setSearchMatches] = useState<Array<{ row: number; col: number; text: string }>>([]);
  const [currentMatchIndex, setCurrentMatchIndex] = useState<number>(0);

  const findMatches = useCallback((scopeAll = true) => {
    const q = findText;
    if (!q) { setSearchMatches([]); setCurrentMatchIndex(0); return []; }
    let re: RegExp | null = null;
    if (matchRegex) {
      try { re = new RegExp(q, caseSensitive ? '' : 'i'); } catch (e) { re = null; }
    }
    const matches: Array<{ row: number; col: number; text: string }> = [];
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < (data[0] || []).length; c++) {
        const cell = String(data[r][c] ?? '');
        if (!cell) continue;
        if (matchRegex && re) {
          if (re.test(cell)) matches.push({ row: r, col: c, text: cell });
        } else {
          const a = caseSensitive ? cell : cell.toLowerCase();
          const b = caseSensitive ? q : q.toLowerCase();
          if (a.indexOf(b) !== -1) matches.push({ row: r, col: c, text: cell });
        }
      }
    }
    setSearchMatches(matches);
    setCurrentMatchIndex(0);
    return matches;
  }, [findText, caseSensitive, matchRegex, data]);

  const goToMatch = (idx: number) => {
    if (!searchMatches.length) return;
    const m = searchMatches[(idx + searchMatches.length) % searchMatches.length];
    setSelected({ row: m.row, col: m.col });
    setSelectionStart({ row: m.row, col: m.col });
    setSelectionEnd({ row: m.row, col: m.col });
    setCurrentMatchIndex((idx + searchMatches.length) % searchMatches.length);
  };

  const replaceCurrent = () => {
    if (!searchMatches.length) return;
    const m = searchMatches[currentMatchIndex];
    pushHistory();
    setData(prev => {
      const next = prev.map(r => [...r]);
      const original = String(next[m.row][m.col] ?? '');
      let replaced = original;
      if (matchRegex) {
        try { replaced = original.replace(new RegExp(findText, caseSensitive ? '' : 'i'), replaceText); } catch (e) { replaced = original; }
      } else {
        if (caseSensitive) replaced = original.split(findText).join(replaceText);
        else replaced = original.split(new RegExp(findText, 'gi')).join(replaceText);
      }
      next[m.row][m.col] = replaced;
      return next;
    });
    // refresh matches
    setTimeout(() => findMatches(), 0);
  };

  const replaceAll = () => {
    const matches = findMatches();
    if (!matches.length) return;
    pushHistory();
    setData(prev => {
      const next = prev.map(r => [...r]);
      for (const m of matches) {
        const original = String(next[m.row][m.col] ?? '');
        let replaced = original;
        if (matchRegex) {
          try { replaced = original.replace(new RegExp(findText, caseSensitive ? '' : 'i'), replaceText); } catch (e) { replaced = original; }
        } else {
          if (caseSensitive) replaced = original.split(findText).join(replaceText);
          else replaced = original.split(new RegExp(findText, 'gi')).join(replaceText);
        }
        next[m.row][m.col] = replaced;
      }
      return next;
    });
    setTimeout(() => findMatches(), 0);
  };

  // --- Freeze panes (filas/columnas) ---
  const [freezeRows, setFreezeRows] = useState<number>(0);
  const [freezeCols, setFreezeCols] = useState<number>(0);



  const getColLefts = useCallback(() => {
    const lefts: number[] = [];
    let acc = 40; // gutter (row header) ancho aproximado
    for (let i = 0; i < colWidths.length; i++) {
      lefts[i] = acc;
      acc += (colWidths[i] ?? 96);
    }
    return lefts;
  }, [colWidths]);

  // --- Sort Range ---
  const sortSelectionByColumn = () => {
    // Sorting is handled via the Sort modal (showSortModal) which calls sortRangeByColumn
    return;
  };

  // --- Conditional Formatting (simple) ---
  type CFRule = { id: string; type: 'gt' | 'lt' | 'eq' | 'contains'; value: string; bg?: string; color?: string; scope?: { r1: number, c1: number, r2: number, c2: number } | null };
  const [conditionalFormats, setConditionalFormats] = useState<CFRule[]>([]);
  // --- Local storage persistence ---
  const STORAGE_KEY = 'meloexcel_v1';
  const loadFromLocal = useCallback(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (parsed.data) setData(parsed.data);
      if (parsed.merges) setMerges(parsed.merges);
      if (parsed.colWidths) setColWidths(parsed.colWidths);
      if (typeof parsed.darkMode === 'boolean') { setDarkMode(parsed.darkMode); }
      if (parsed.conditionalFormats) setConditionalFormats(parsed.conditionalFormats);
      if (typeof parsed.freezeRows === 'number') setFreezeRows(parsed.freezeRows);
      if (typeof parsed.freezeCols === 'number') setFreezeCols(parsed.freezeCols);
    } catch (e) {
      console.warn('Failed to load local save:', e);
    }
  }, []);

  const saveToLocal = useCallback(() => {
    try {
      const payload = {
        data,
        merges,
        colWidths,
        darkMode,
        conditionalFormats,
        freezeRows,
        freezeCols,
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
      console.warn('Failed to save to localStorage:', e);
    }
  }, [data, merges, colWidths, darkMode, conditionalFormats, freezeRows, freezeCols]);

  const clearLocal = useCallback(() => {
    try { window.localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
  }, []);

  // load once on mount
  useEffect(() => { loadFromLocal(); }, [loadFromLocal]);

  // save whenever relevant state changes (but skip the initial run to avoid overwriting loaded data)
  const skipFirstSave = React.useRef(true);
  useEffect(() => {
    if (skipFirstSave.current) { skipFirstSave.current = false; return; }
    saveToLocal();
  }, [saveToLocal]);
  // Replaced by the CF modal UI (showCFModal)

  // --- Mejoras UX: modales para Sort y Conditional Format + toast ---
  const [showSortModal, setShowSortModal] = useState(false);
  const [sortColInput, setSortColInput] = useState(getColName(selectionStart?.col ?? 0));
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  const [showCFModal, setShowCFModal] = useState(false);
  const [cfTypeInput, setCfTypeInput] = useState<CFRule['type']>('gt');
  const [cfValueInput, setCfValueInput] = useState('0');
  const [cfColorInput, setCfColorInput] = useState('#fffbcc');
  const [cfScopeInput, setCfScopeInput] = useState<'selection' | 'sheet'>('selection');

  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const showToast = (msg: string, ms = 2500) => {
    setToastMsg(msg);
    window.setTimeout(() => setToastMsg(null), ms);
  };

  // Clipboard for copy/paste ranges
  const [copiedRange, setCopiedRange] = useState<string[][] | null>(null);
  const [pasteHighlight, setPasteHighlight] = useState<{ r1: number; c1: number; r2: number; c2: number } | null>(null);
  const [pasteFade, setPasteFade] = useState<boolean>(false);
  const serializeRange = (arr: string[][]) => arr.map(r => r.join('\t')).join('\n');
  const parseSerialized = (s: string) => s.split(/\r?\n/).map(row => row.split('\t'));

  // Copy selection to internal clipboard and try system clipboard
  const copySelectionToClipboard = async () => {
    if (!selectionStart || !selectionEnd) { showToast('No hay selección para copiar'); return; }
    const r1 = Math.min(selectionStart.row, selectionEnd.row);
    const r2 = Math.max(selectionStart.row, selectionEnd.row);
    const c1 = Math.min(selectionStart.col, selectionEnd.col);
    const c2 = Math.max(selectionStart.col, selectionEnd.col);
    const out: string[][] = [];
    for (let rr = r1; rr <= r2; rr++) {
      const row: string[] = [];
      for (let cc = c1; cc <= c2; cc++) {
        row.push(String(data[rr] && data[rr][cc] != null ? data[rr][cc] : ''));
      }
      out.push(row);
    }
    setCopiedRange(out);
    const txt = serializeRange(out);
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) await navigator.clipboard.writeText(txt);
    } catch (e) {
      // ignore
    }
    showToast('Copiado');
  };

  // Paste from internal or system clipboard at selected cell
  const pasteClipboardAtSelection = async () => {
    if (!selected) { showToast('Selecciona una celda para pegar'); return; }
    let toPaste: string[][] | null = copiedRange;
    if (!toPaste) {
      try {
        if (navigator.clipboard && navigator.clipboard.readText) {
          const txt = await navigator.clipboard.readText();
          if (txt) toPaste = parseSerialized(txt);
        }
      } catch (e) {
        // ignore
      }
    }
    if (!toPaste) { showToast('No hay datos en el portapapeles'); return; }
    const startRow = selected.row;
    const startCol = selected.col;
    pushHistory();
    setData(prev => {
      const next = prev.map(r => [...r]);
      const needCols = Math.max(0, startCol + toPaste[0].length - next[0].length);
      if (needCols > 0) {
        for (let rr = 0; rr < next.length; rr++) for (let i = 0; i < needCols; i++) next[rr].push('');
      }
      const needRows = Math.max(0, startRow + toPaste.length - next.length);
      if (needRows > 0) {
        for (let i = 0; i < needRows; i++) next.push(Array(next[0].length).fill(''));
      }
      for (let r = 0; r < toPaste.length; r++) {
        for (let c = 0; c < toPaste[r].length; c++) {
          const rr = startRow + r;
          const cc = startCol + c;
          next[rr][cc] = toPaste[r][c];
        }
      }
      return next;
    });
    // Ensure colWidths length matches new width synchronously
    setColWidths(prev => {
      const newLen = Math.max(prev.length, (toPaste[0]?.length ?? 0) + startCol, (data[0]?.length ?? 0));
      if (newLen <= prev.length) return prev;
      return [...prev, ...Array(newLen - prev.length).fill(96)];
    });
    // highlight temporal con fade
    setPasteHighlight({ r1: startRow, c1: startCol, r2: startRow + toPaste.length - 1, c2: startCol + toPaste[0].length - 1 });
    setPasteFade(false);
    // trigger fade shortly before clearing
    window.setTimeout(() => setPasteFade(true), 1100);
    window.setTimeout(() => { setPasteHighlight(null); setPasteFade(false); }, 1500);
    showToast('Pegado');
  };

  // Paste transposed (rows<->cols)
  const pasteTransposedAtSelection = async () => {
    if (!selected) { showToast('Selecciona una celda para pegar'); return; }
    let toPaste: string[][] | null = copiedRange;
    if (!toPaste) {
      try {
        if (navigator.clipboard && navigator.clipboard.readText) {
          const txt = await navigator.clipboard.readText();
          if (txt) toPaste = parseSerialized(txt);
        }
      } catch (e) {
        // ignore
      }
    }
    if (!toPaste) { showToast('No hay datos en el portapapeles'); return; }
    const startRow = selected.row;
    const startCol = selected.col;
    pushHistory();
    setData(prev => {
      const next = prev.map(r => [...r]);
      const transRows = toPaste[0].length;
      const transCols = toPaste.length;
      const needCols = Math.max(0, startCol + transCols - next[0].length);
      if (needCols > 0) {
        for (let i = 0; i < needCols; i++) for (let rr = 0; rr < next.length; rr++) next[rr].push('');
      }
      const needRows = Math.max(0, startRow + transRows - next.length);
      if (needRows > 0) {
        for (let i = 0; i < needRows; i++) next.push(Array(next[0].length).fill(''));
      }
      for (let r = 0; r < transRows; r++) {
        for (let c = 0; c < transCols; c++) {
          const rr = startRow + r;
          const cc = startCol + c;
          next[rr][cc] = toPaste[c][r];
        }
      }
      return next;
    });
    // highlight temporal para transposed con fade
    setPasteHighlight({ r1: startRow, c1: startCol, r2: startRow + (toPaste[0].length) - 1, c2: startCol + (toPaste.length) - 1 });
    setPasteFade(false);
    window.setTimeout(() => setPasteFade(true), 1100);
    window.setTimeout(() => { setPasteHighlight(null); setPasteFade(false); }, 1500);
    showToast('Pegado (transpuesto)');
  };

  // Global keyboard handlers for Ctrl+C / Ctrl+V
  React.useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const meta = isMac ? e.metaKey : e.ctrlKey;
      if (!meta) return;
      if (e.key.toLowerCase() === 'c') {
        e.preventDefault();
        copySelectionToClipboard();
      }
      if (e.key.toLowerCase() === 'v') {
        e.preventDefault();
        if (e.shiftKey) pasteTransposedAtSelection();
        else pasteClipboardAtSelection();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [selectionStart, selectionEnd, selected, copiedRange, data]);

  // Escucha global para borrar contenido con la tecla Supr/Delete
  React.useEffect(() => {
    const onDel = (e: KeyboardEvent) => {
      // Solo si no hay meta/ctrl (no interferir con atajos como Ctrl+Delete)
      if (e.ctrlKey || e.metaKey) return;
      if (e.key === 'Delete' || e.key === 'Del') {
        if (!selected) return;
        e.preventDefault();
        pushHistory();
        setData(prev => {
          const next = prev.map(r => [...r]);
          if (next[selected.row] && next[selected.row][selected.col] != null) next[selected.row][selected.col] = '';
          return next;
        });
        // Si la celda formaba parte de un merge, dejamos el merge pero limpiamos el contenido del bloque.
        // Alternativa: separar merge; por ahora solo limpiamos la celda activa.
        showToast('Contenido de la celda eliminado');
      }
    };
    window.addEventListener('keydown', onDel);
    return () => window.removeEventListener('keydown', onDel);
  }, [selected, pushHistory]);

  // Aplica orden usando parámetros (reutiliza la lógica existente)
  const sortRangeByColumn = (colIdx: number, dir: 'asc' | 'desc') => {
    if (!selectionStart || !selectionEnd) { showToast('No hay selección para ordenar'); return; }
    const r1 = Math.min(selectionStart.row, selectionEnd.row);
    const r2 = Math.max(selectionStart.row, selectionEnd.row);
    const c1 = Math.min(selectionStart.col, selectionEnd.col);
    const c2 = Math.max(selectionStart.col, selectionEnd.col);
    if (colIdx < c1 || colIdx > c2) { showToast('La columna debe estar dentro de la selección'); return; }
    pushHistory();
    setData(prev => {
      const next = prev.map(r => [...r]);
      const slice = next.slice(r1, r2 + 1);
      slice.sort((a, b) => {
        const va = a[colIdx] ?? '';
        const vb = b[colIdx] ?? '';
        const na = Number(String(va).replace(/,/g, ''));
        const nb = Number(String(vb).replace(/,/g, ''));
        if (!Number.isNaN(na) && !Number.isNaN(nb)) return dir === 'asc' ? na - nb : nb - na;
        const sa = String(va);
        const sb = String(vb);
        return dir === 'asc' ? sa.localeCompare(sb) : sb.localeCompare(sa);
      });
      for (let i = r1; i <= r2; i++) next[i] = slice[i - r1];
      return next;
    });
    setSelectionStart(null);
    setSelectionEnd(null);
    showToast('Rango ordenado');
  };

  const addConditionalFormat = (type: CFRule['type'], value: string, bg: string, scopeChoice: 'selection' | 'sheet' = 'selection') => {
    const id = String(Date.now());
    let scope = null as null | { r1: number; c1: number; r2: number; c2: number };
    if (scopeChoice === 'selection' && selectionStart && selectionEnd) {
      const r1 = Math.min(selectionStart.row, selectionEnd.row);
      const r2 = Math.max(selectionStart.row, selectionEnd.row);
      const c1 = Math.min(selectionStart.col, selectionEnd.col);
      const c2 = Math.max(selectionStart.col, selectionEnd.col);
      scope = { r1, c1, r2, c2 };
    }
    setConditionalFormats(prev => [...prev, { id, type, value, bg, color: undefined, scope }]);
    showToast('Formato condicional añadido');
  };

  // Cell component moved to ./Cell.tsx (memoized there). Use that instead of inline definition.

  const addRow = () => {
    setData(prev => [...prev, Array(prev[0].length).fill('')]);
  };

  const addCol = () => {
    setData(prev => prev.map(row => [...row, '']));
    setColWidths(prev => [...prev, 96]);
  };

  // Selección de fila y columna
  const [selectedRow, setSelectedRow] = useState<number>(0);
  const [selectedCol, setSelectedCol] = useState<number>(0);

  const deleteRow = () => {
    if (data.length > 1 && selectedRow >= 0 && selectedRow < data.length) {
      setData(prev => prev.filter((_, idx) => idx !== selectedRow));
      // ajustar merges: eliminar merges que intersecten la fila eliminada, y desplazar merges posteriores
      setMerges(prev => prev
        .map(m => {
          if (m.r > selectedRow) return { ...m, r: m.r - 1 };
          return m;
        })
        .filter(m => !(m.r + m.rows - 1 < 0 || m.rows <= 0))
        .filter(m => !(selectedRow >= m.r && selectedRow < m.r + m.rows))
      );
      setSelectedRow(selectedRow > 0 ? selectedRow - 1 : 0);
    }
  };

  const deleteCol = () => {
    if (data[0].length > 1 && selectedCol >= 0 && selectedCol < data[0].length) {
      setData(prev => prev.map(row => row.filter((_, idx) => idx !== selectedCol)));
      setColWidths(prev => prev.filter((_, idx) => idx !== selectedCol));
      // ajustar merges: eliminar merges que intersecten la columna eliminada, y desplazar merges posteriores
      setMerges(prev => prev
        .map(m => {
          if (m.c > selectedCol) return { ...m, c: m.c - 1 };
          return m;
        })
        .filter(m => !(selectedCol >= m.c && selectedCol < m.c + m.cols))
      );
      setSelectedCol(selectedCol > 0 ? selectedCol - 1 : 0);
    }
  };

  const handleChange = useCallback((row: number, col: number, value: string) => {
    setData(prev => {
      if (!prev[row]) return prev;
      const next = prev.map(r => [...r]);
      next[row][col] = value;
      return next;
    });
  }, [setData]);

  const commitCell = useCallback((row: number, col: number, value: string) => {
    handleChange(row, col, value);
  }, [handleChange]);

  const handleCellClick = useCallback((row: number, col: number) => {
    // Si justo hubo un drag, ignorar el click (es el click generado al soltar)
    if (hadDragSinceMouseDown.current) {
      hadDragSinceMouseDown.current = false;
      return;
    }
    // Ctrl+click: toggle la celda en extraSelections
    if (ctrlPressed.current) {
      const exists = extraSelections.find(s => s.r1 === row && s.r2 === row && s.c1 === col && s.c2 === col);
      if (exists) setExtraSelections(prev => prev.filter(s => !(s.r1 === row && s.r2 === row && s.c1 === col && s.c2 === col)));
      else setExtraSelections(prev => [...prev, { r1: row, c1: col, r2: row, c2: col }]);
      return;
    }
    // Para click normal, limpiar selecciones aditivas previas
    setExtraSelections([]);
    // Si el Shift está presionado, expandir o crear selección usando la celda actualmente seleccionada como origen
    if (shiftPressed.current) {
      // Si hay una celda activa (selected), usarla como anchor y expandir hasta la celda clicada
      if (selected) {
        setSelectionStart({ row: selected.row, col: selected.col });
        setSelectionEnd({ row, col });
        // mantener `selected` como anchor (no lo movemos)
        return;
      }
      // Si no hay celda activa, tratar Shift+click como click normal: establecer anchor en la celda clicada
      setSelected({ row, col });
      setSelectionStart(null);
      setSelectionEnd(null);
      return;
    }
    // Sin Shift: cancelar cualquier selección múltiple y seleccionar solo la celda clickeada
    if (selectionStart && selectionEnd) {
      setSelectionStart(null);
      setSelectionEnd(null);
      setSelected({ row, col });
      return;
    }
    setSelected({ row, col });
  }, [extraSelections, selectionStart, selectionEnd, selected]);


  const handleCellMouseDown = useCallback((row: number, col: number) => {
    // Iniciar posible selección por mousedown.
    // Si Shift está presionado y hay una celda previamente seleccionada, usarla como ancla.
    // Si no, anclar la selección a la celda clicada para evitar rangos accidentales.
    // reset drag flag at start of a new mousedown
    hadDragSinceMouseDown.current = false;
    isMouseDown.current = true;
    // no activamos isDragging hasta que el usuario mueva el mouse
    setIsDragging(false);
    if (shiftPressed.current && selected) {
      setSelectionStart({ row: selected.row, col: selected.col });
    } else {
      setSelectionStart({ row, col });
      setSelected({ row, col });
    }
    setSelectionEnd({ row, col });
  }, [selected]);

  const handleCellMouseEnter = useCallback((row: number, col: number) => {
    // Solo actualizar la selección por hover si estamos arrastrando con el mouse (mousedown)
    if (!isMouseDown.current) return;
    // Si la celda pertenece a un merge, hacer snap al bloque completo
    const covering = getCoveringMerge(row, col);
    if (covering) {
      setSelectionEnd({ row: covering.r + covering.rows - 1, col: covering.c + covering.cols - 1 });
      setSelectionStart({ row: covering.r, col: covering.c });
    } else if (selectionStart) {
      setSelectionEnd({ row, col });
    }
    // marcar drag solo cuando hay movimiento real (señal simple: selectionEnd distinta de selectionStart)
    hadDragSinceMouseDown.current = true;
  }, [selectionStart]);

  const handleCellMouseUp = useCallback(() => {
    // No limpiamos selectionStart aquí para que la selección persista después de soltar.
    isMouseDown.current = false;
    // reset drag flag when mouse is released
    hadDragSinceMouseDown.current = false;
    setIsDragging(false);
  }, []);

  const onCellMouseDown = useCallback((r: number, c: number) => { handleCellMouseDown(r, c); }, [handleCellMouseDown]);
  const onCellMouseEnter = useCallback((r: number, c: number) => { handleCellMouseEnter(r, c); }, [handleCellMouseEnter]);
  const onCellMouseUp = useCallback(() => { handleCellMouseUp(); }, [handleCellMouseUp]);
  const onCellClick = useCallback((r: number, c: number) => { handleCellClick(r, c); }, [handleCellClick]);

  // Listener global para asegurar que isDragging se desactive si el usuario suelta fuera de la tabla
  React.useEffect(() => {
    const onUp = () => {
      if (isMouseDown.current) {
        isMouseDown.current = false;
      }
      setIsDragging(false);
      hadDragSinceMouseDown.current = false;
    };
    window.addEventListener('mouseup', onUp);
    return () => window.removeEventListener('mouseup', onUp);
  }, []);

  const isCellSelected = (row: number, col: number) => {
    if (!selectionStart || !selectionEnd) return false;
    const minRow = Math.min(selectionStart.row, selectionEnd.row);
    const maxRow = Math.max(selectionStart.row, selectionEnd.row);
    const minCol = Math.min(selectionStart.col, selectionEnd.col);
    const maxCol = Math.max(selectionStart.col, selectionEnd.col);
    return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
  };

  // Comprueba si una celda está dentro de cualquier selección (principal o extra)
  const isCellInSelections = (row: number, col: number) => {
    if (isCellSelected(row, col)) return true;
    for (const s of extraSelections) {
      if (row >= s.r1 && row <= s.r2 && col >= s.c1 && col <= s.c2) return true;
    }
    return false;
  };

  const combineCells = () => {
    // Determinar la región a combinar.
    // Si hay extraSelections priorizarlas (usuario hizo Ctrl+click). Requerimos que formen
    // un rectángulo contiguo para evitar borrar celdas no seleccionadas.
    let r1: number, r2: number, c1: number, c2: number;
    if (extraSelections && extraSelections.length) {
      r1 = Math.min(...extraSelections.map(s => s.r1));
      r2 = Math.max(...extraSelections.map(s => s.r2));
      c1 = Math.min(...extraSelections.map(s => s.c1));
      c2 = Math.max(...extraSelections.map(s => s.c2));
      // comprobar contigüidad: todos los cells dentro del bbox deben estar presentes en extraSelections
      const expected = (r2 - r1 + 1) * (c2 - c1 + 1);
      const present = new Set<string>();
      for (const s of extraSelections) {
        for (let rr = s.r1; rr <= s.r2; rr++) {
          for (let cc = s.c1; cc <= s.c2; cc++) {
            present.add(`${rr},${cc}`);
          }
        }
      }
      if (present.size !== expected) {
        // propose to combine the bounding box but ask confirmation
        setPendingCombineBox({ r1, r2, c1, c2 });
        setShowCombineConfirm(true);
        return;
      }
    } else if (selectionStart && selectionEnd) {
      r1 = Math.min(selectionStart.row, selectionEnd.row);
      r2 = Math.max(selectionStart.row, selectionEnd.row);
      c1 = Math.min(selectionStart.col, selectionEnd.col);
      c2 = Math.max(selectionStart.col, selectionEnd.col);
    } else {
      return;
    }
    // Expand selection to fully include any existing merges that partially intersect
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of merges) {
        const mr1 = m.r;
        const mr2 = m.r + m.rows - 1;
        const mc1 = m.c;
        const mc2 = m.c + m.cols - 1;
        // if merge intersects selection (not completely outside)
        const intersects = !(mr2 < r1 || mr1 > r2 || mc2 < c1 || mc1 > c2);
        if (intersects) {
          const nr1 = Math.min(r1, mr1);
          const nr2 = Math.max(r2, mr2);
          const nc1 = Math.min(c1, mc1);
          const nc2 = Math.max(c2, mc2);
          if (nr1 !== r1 || nr2 !== r2 || nc1 !== c1 || nc2 !== c2) {
            r1 = nr1; r2 = nr2; c1 = nc1; c2 = nc2;
            changed = true;
          }
        }
      }
    }
    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;
    if (rows === 1 && cols === 1) return; // nada que combinar
    // push history for undo
    pushHistory();

    // Concatenar contenidos en la celda superior-izquierda y vaciar las demás
    setData(prev => {
      const next = prev.map(r => [...r]);
      const parts: string[] = [];
      for (let rr = r1; rr <= r2; rr++) {
        for (let cc = c1; cc <= c2; cc++) {
          const v = String(next[rr] && next[rr][cc] != null ? next[rr][cc] : '');
          if (v !== '') parts.push(v);
          if (!(rr === r1 && cc === c1)) next[rr][cc] = '';
        }
      }
      next[r1][c1] = parts.join(' ');
      return next;
    });

    // Reemplazar merges que intersecten la región por el nuevo merge de forma atómica
    setMerges(prev => {
      const filtered = prev.filter(m => (m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2));
      return [...filtered, { r: r1, c: c1, rows, cols }];
    });
    // seleccionar la celda combinada y limpiar selecciones previas
    setSelected({ row: r1, col: c1 });
    setSelectionStart(null);
    setSelectionEnd(null);
    setExtraSelections([]);
    showToast('Celdas combinadas');
  };

  const performCombineRegion = (r1: number, r2: number, c1: number, c2: number) => {
    // push history for undo
    pushHistory();
    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;
    setData(prev => {
      const next = prev.map(r => [...r]);
      const parts: string[] = [];
      for (let rr = r1; rr <= r2; rr++) {
        for (let cc = c1; cc <= c2; cc++) {
          const v = String(next[rr] && next[rr][cc] != null ? next[rr][cc] : '');
          if (v !== '') parts.push(v);
          if (!(rr === r1 && cc === c1)) next[rr][cc] = '';
        }
      }
      next[r1][c1] = parts.join(' ');
      return next;
    });
    setMerges(prev => {
      const filtered = prev.filter(m => (m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2));
      return [...filtered, { r: r1, c: c1, rows, cols }];
    });
    setSelected({ row: r1, col: c1 });
    setSelectionStart(null);
    setSelectionEnd(null);
    setExtraSelections([]);
    showToast('Celdas combinadas');
  };

  const separateCells = () => {
    // Si hay selección principal, separar merges que intersecten
    if (selectionStart && selectionEnd) {
      const r1 = Math.min(selectionStart.row, selectionEnd.row);
      const r2 = Math.max(selectionStart.row, selectionEnd.row);
      const c1 = Math.min(selectionStart.col, selectionEnd.col);
      const c2 = Math.max(selectionStart.col, selectionEnd.col);
      setMerges(prev => prev.filter(m => (m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2)));
      setSelectionStart(null);
      setSelectionEnd(null);
      setExtraSelections([]);
      showToast('Celdas separadas');
      return;
    }
    // Si hay extraSelections, separar merges que intersecten cualquiera de ellas
    if (extraSelections && extraSelections.length) {
      setMerges(prev => prev.filter(m => {
        for (const s of extraSelections) {
          const r1 = s.r1, r2 = s.r2, c1 = s.c1, c2 = s.c2;
          // si intersecta, excluir (i.e., eliminar merge)
          if (!((m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2))) {
            return false;
          }
        }
        return true;
      }));
      setExtraSelections([]);
      showToast('Celdas separadas');
      return;
    }
    // si no hay selección múltiple, separar el merge que contenga la celda seleccionada
    if (selected) {
      const { row, col } = selected;
      setMerges(prev => prev.filter(m => !(row >= m.r && row < m.r + m.rows && col >= m.c && col < m.c + m.cols)));
      showToast('Celdas separadas');
    }
  };

  // Lógica para resize de columnas
  React.useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (resizingCol.current !== null) {
        const delta = e.clientX - startX.current;
        setColWidths(widths => {
          const newWidths = [...widths];
          newWidths[resizingCol.current!] = Math.max(40, startWidth.current + delta);
          return newWidths;
        });
      }
    };
    const handleMouseUp = () => {
      resizingCol.current = null;
      document.body.style.cursor = '';
    };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  // Import / Export Excel
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const tableRef = useRef<HTMLTableElement | null>(null);
  const parentRef = useRef<HTMLDivElement | null>(null);

  const rowVirtualizer = useVirtualizer({
    count: data.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => 36,
    overscan: 10,
  });

  // Mientras se arrastra, hacemos seguimiento del elemento bajo el cursor y actualizamos selectionEnd por celda
  React.useEffect(() => {
    // Usar coordenadas relativas a la tabla y `colWidths` para calcular fila/col
    const onMove = (e: MouseEvent) => {
      if (!isMouseDown.current || !shiftPressed.current) return;
      const table = tableRef.current;
      if (!table) return;
      const rect = table.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      if (x < 0 || y < 0 || x > rect.width || y > rect.height) return;

      // calcular ancho del primer TH (gutter de filas)
      const firstTh = table.querySelector('thead th');
      const gutter = firstTh ? firstTh.getBoundingClientRect().width : 40;
      // altura del header (thead)
      const thead = table.querySelector('thead');
      const headerHeight = thead ? thead.getBoundingClientRect().height : 36;

      // calcular columna usando colWidths
      let cx = x - gutter;
      if (cx < 0) return;
      let acc = 0;
      let col = 0;
      for (let i = 0; i < colWidths.length; i++) {
        acc += (colWidths[i] ?? 96);
        if (cx <= acc) { col = i; break; }
        // si llegamos al final sin romper, col será última
        if (i === colWidths.length - 1) col = i;
      }

      // calcular fila usando altura fija (36px por fila)
      const relY = y - headerHeight; // 0 en la primera fila
      const rowHeight = 36;
      let row = Math.floor(relY / rowHeight);
      if (row < 0) row = 0;
      if (row >= data.length) row = data.length - 1;

      // si la celda pertenece a un merge, snap a su bounding box
      const covering = merges.find(m => row >= m.r && row < m.r + m.rows && col >= m.c && col < m.c + m.cols);
      if (covering) {
        setSelectionStart({ row: covering.r, col: covering.c });
        setSelectionEnd({ row: covering.r + covering.rows - 1, col: covering.c + covering.cols - 1 });
      } else {
        setSelectionEnd({ row, col });
      }
      // marcar drag
      hadDragSinceMouseDown.current = true;
    };
    window.addEventListener('mousemove', onMove);
    return () => window.removeEventListener('mousemove', onMove);
  }, []);

  // Manejo específico para arrastrar sobre encabezados (selección por columnas)
  React.useEffect(() => {
    const onMoveHeader = (e: MouseEvent) => {
      if (!headerDragging.current) return;
      const el = document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null;
      if (!el) return;
      const th = el.closest('th[data-header-col]') as HTMLElement | null;
      if (!th) return;
      const col = Number(th.getAttribute('data-header-col'));
      if (Number.isNaN(col)) return;
      const start = headerDragStart.current ?? col;
      const c1 = Math.min(start, col);
      const c2 = Math.max(start, col);
      // seleccionar columnas completas entre c1 y c2
      setSelectionStart({ row: 0, col: c1 });
      setSelectionEnd({ row: data.length - 1, col: c2 });
      setSelected({ row: 0, col: c1 });
    };
    const onUpHeader = (e: MouseEvent) => {
      if (headerDragging.current) {
        headerDragging.current = false;
        headerDragStart.current = null;
        hadDragSinceMouseDown.current = false;
      }
    };
    window.addEventListener('mousemove', onMoveHeader);
    window.addEventListener('mouseup', onUpHeader);
    return () => {
      window.removeEventListener('mousemove', onMoveHeader);
      window.removeEventListener('mouseup', onUpHeader);
    };
  }, [data.length]);

  // Barra de fórmulas: texto editable y resultado calculado (suma rápida y funciones)
  const [formulaText, setFormulaText] = useState<string>('');
  const [computedSum, setComputedSum] = useState<number | null>(null);

  // Helpers para referencias de celda/rango
  const colNameToIndex = (name: string) => {
    let col = 0;
    const s = name.toUpperCase();
    for (let i = 0; i < s.length; i++) {
      col = col * 26 + (s.charCodeAt(i) - 64);
    }
    return col - 1; // A -> 0
  };

  const parseCellRef = (ref: string) => {
    const m = ref.trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!m) return null;
    const col = colNameToIndex(m[1]);
    const row = parseInt(m[2], 10) - 1;
    return { row, col };
  };

  const collectValuesInRange = (r1: number, c1: number, r2: number, c2: number) => {
    const vals: number[] = [];
    for (let rr = Math.max(0, r1); rr <= Math.min(data.length - 1, r2); rr++) {
      for (let cc = Math.max(0, c1); cc <= Math.min((data[0] || []).length - 1, c2); cc++) {
        const v = data[rr] && data[rr][cc] != null ? String(data[rr][cc]).trim() : '';
        if (v === '') continue;
        const n = Number(v.toString().replace(/,/g, ''));
        if (!Number.isNaN(n)) vals.push(n);
      }
    }
    return vals;
  };

  const parseRangeArg = (arg: string) => {
    const parts = arg.split(':').map(p => p.trim());
    if (parts.length === 1) {
      const ref = parseCellRef(parts[0]);
      if (!ref) return null;
      return { r1: ref.row, c1: ref.col, r2: ref.row, c2: ref.col };
    }
    const a = parseCellRef(parts[0]);
    const b = parseCellRef(parts[1]);
    if (!a || !b) return null;
    return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
  };

  const computeFormulaResult = (formula: string | null) => {
    // si formula es null, y hay selección, devolver suma/avg/count automático (preview)
    if ((!formula || formula.trim() === '') && selectionStart && selectionEnd) {
      const r1 = Math.min(selectionStart.row, selectionEnd.row);
      const r2 = Math.max(selectionStart.row, selectionEnd.row);
      const c1 = Math.min(selectionStart.col, selectionEnd.col);
      const c2 = Math.max(selectionStart.col, selectionEnd.col);
      const vals = collectValuesInRange(r1, c1, r2, c2);
      if (!vals.length) return null;
      const sum = vals.reduce((a, b) => a + b, 0);
      return { value: sum, func: 'SUM' } as any;
    }

    if (!formula) return null;
    const f = formula.trim();
    if (!f.startsWith('=')) {
      // si no empieza con =, intentar parsear como número simple
      const n = Number(f.replace(/,/g, ''));
      return Number.isNaN(n) ? null : { value: n, func: 'VAL' };
    }
    // parsear =FUNC(arg)
    const m = f.match(/^=([A-Z]+)\(([^)]*)\)$/i);
    if (!m) return null;
    const func = m[1].toUpperCase();
    const arg = m[2].trim();
    // soporte múltiples argumentos separados por ","
    const argParts = arg.split(',').map(s => s.trim()).filter(Boolean);
    let allVals: number[] = [];
    for (const ap of argParts) {
      const range = parseRangeArg(ap);
      if (!range) continue;
      const vals = collectValuesInRange(range.r1, range.c1, range.r2, range.c2);
      allVals = allVals.concat(vals);
    }
    if (!allVals.length) return { value: 0, func: func };
    if (func === 'SUM') {
      return { value: allVals.reduce((a, b) => a + b, 0), func };
    }
    if (func === 'AVERAGE' || func === 'AVG') {
      return { value: allVals.reduce((a, b) => a + b, 0) / allVals.length, func: 'AVERAGE' };
    }
    if (func === 'COUNT') {
      return { value: allVals.length, func };
    }
    // CONCAT: concatena los valores como texto
    if (func === 'CONCAT') {
      // Para CONCAT, obtener los valores como texto
      let allTexts: string[] = [];
      for (const ap of argParts) {
        const range = parseRangeArg(ap);
        if (!range) continue;
        for (let r = range.r1; r <= range.r2; r++) {
          for (let c = range.c1; c <= range.c2; c++) {
            const val = data[r]?.[c];
            if (val !== undefined && val !== null) allTexts.push(String(val));
          }
        }
      }
      return { value: allTexts.join(''), func };
    }

      // ABS: valor absoluto de un número
      if (func === 'ABS') {
        if (argParts.length === 1) {
          const n = Number(argParts[0]);
          return { value: Math.abs(n), func };
        }
        return { value: 0, func };
      }

      // TRIM: elimina espacios extra de un texto
      if (func === 'TRIM') {
        if (argParts.length === 1) {
          return { value: String(argParts[0]).trim(), func };
        }
        return { value: '', func };
      }

      // NOT: invierte el valor lógico
      if (func === 'NOT') {
        if (argParts.length === 1) {
          const v = argParts[0].toLowerCase();
          return { value: !(v === 'true' || v === '1'), func };
        }
        return { value: false, func };
      }

      // AND: todas las condiciones deben ser verdaderas
      if (func === 'AND') {
        const res = argParts.every(p => p.toLowerCase() === 'true' || p === '1');
        return { value: res, func };
      }

      // OR: alguna condición debe ser verdadera
      if (func === 'OR') {
        const res = argParts.some(p => p.toLowerCase() === 'true' || p === '1');
        return { value: res, func };
      }

      // SUMIF: suma valores que cumplen una condición simple (ejemplo: >5)
      if (func === 'SUMIF') {
        // SUMIF(rango, condicion)
        if (argParts.length === 2) {
          const range = parseRangeArg(argParts[0]);
          if (!range) return { value: 0, func };
          const cond = argParts[1];
          let sum = 0;
          for (let r = range.r1; r <= range.r2; r++) {
            for (let c = range.c1; c <= range.c2; c++) {
              const val = Number(data[r]?.[c]);
              if (eval(`val${cond}`)) sum += val;
            }
          }
          return { value: sum, func };
        }
        return { value: 0, func };
      }

      // COUNTIF: cuenta valores que cumplen una condición simple (ejemplo: >5)
      if (func === 'COUNTIF') {
        // COUNTIF(rango, condicion)
        if (argParts.length === 2) {
          const range = parseRangeArg(argParts[0]);
          if (!range) return { value: 0, func };
          const cond = argParts[1];
          let count = 0;
          for (let r = range.r1; r <= range.r2; r++) {
            for (let c = range.c1; c <= range.c2; c++) {
              const val = Number(data[r]?.[c]);
              if (eval(`val${cond}`)) count++;
            }
          }
          return { value: count, func };
        }
        return { value: 0, func };
      }

      // VLOOKUP: busca un valor en una columna y devuelve el relacionado de otra columna
      if (func === 'VLOOKUP') {
        // VLOOKUP(valor, rango_busqueda, col_retorno)
        if (argParts.length >= 3) {
          const searchVal = argParts[0];
          const range = parseRangeArg(argParts[1]);
          const retColIdx = Number(argParts[2]);
          if (!range || isNaN(retColIdx)) return { value: '', func };
          for (let r = range.r1; r <= range.r2; r++) {
            if (String(data[r]?.[range.c1]) === searchVal) {
              return { value: data[r]?.[range.c1 + retColIdx] ?? '', func };
            }
          }
          return { value: '', func };
        }
        return { value: '', func };
      }

    // DATE: construye una fecha a partir de argumentos (año, mes, día)
    if (func === 'DATE') {
      // Espera argumentos: año, mes, día
      if (argParts.length === 3) {
        const year = Number(argParts[0]);
        const month = Number(argParts[1]) - 1; // JS: 0-based
        const day = Number(argParts[2]);
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) {
          return { value: d.toISOString().slice(0, 10), func };
        }
      }
      return { value: '', func };
    }

    // TODAY: devuelve la fecha actual
    if (func === 'TODAY') {
      const d = new Date();
      return { value: d.toISOString().slice(0, 10), func };
    }
    return null;
  };

  // actualizar preview cuando cambia la selección, datos o formulaText
  React.useEffect(() => {
    const res = computeFormulaResult(formulaText || null);
    if (res && typeof res.value === 'number') {
      setComputedSum(res.value);
      // si no hay formulaText, mostrar el número en la barra (comodidad)
      // Nota: no sobreescribimos `formulaText` automáticamente con el preview, evita escribir valores inesperados al pulsar V
    } else {
      setComputedSum(null);
      // no tocar formulaText
    }
  }, [selectionStart, selectionEnd, data, formulaText]);

  const applyFormulaToSelection = () => {
    // si formulaText empieza con '=', evaluar y aplicar valor a celda activa (o a la celda superior-izquierda de selección)
    const res = computeFormulaResult(formulaText || null);
    if (!res) return;
    // si hay selección principal, aplicar en la celda superior-izquierda
    if (selectionStart && selectionEnd) {
      const targetRow = Math.min(selectionStart.row, selectionEnd.row);
      const targetCol = Math.min(selectionStart.col, selectionEnd.col);
      // buscar la siguiente celda vacía a la derecha de targetCol
      const startCol = targetCol + 1;
      let placedCol = -1;
      setData(prev => {
        const next = prev.map(r => [...r]);
        let resultCol = startCol;
        // buscar dentro de las columnas existentes
        while (resultCol < next[0].length && next[targetRow] && String(next[targetRow][resultCol] ?? '').trim() !== '') {
          resultCol++;
        }
        // si necesitamos ampliar columnas
        if (resultCol >= next[0].length) {
          const need = resultCol - next[0].length + 1;
          for (let i = 0; i < need; i++) {
            for (let rr = 0; rr < next.length; rr++) next[rr].push('');
          }
          // ajustar colWidths fuera del setData
        }
        next[targetRow][resultCol] = `Resultado = ${res.value}`;
        placedCol = resultCol;
        return next;
      });
      // si se añadió/extendieron columnas, ajustar colWidths hasta el tamaño necesario
      setColWidths(prev => {
        const cur = prev.slice();
        const needed = Math.max(0, (data[0]?.length ?? 0) - cur.length);
        if (needed > 0) return [...cur, ...Array(needed).fill(96)];
        return cur;
      });
      if (placedCol >= 0) setSelected({ row: targetRow, col: placedCol });
      setSelectionStart(null);
      setSelectionEnd(null);
      setFormulaText(String(res.value));
      return;
    }

    // no aplicamos a extraSelections automáticamente para evitar efectos inesperados

    // fallback: aplicar en la celda seleccionada
    const targetRow = selected?.row ?? 0;
    const targetCol = selected?.col ?? 0;
    // escribir resultado en la siguiente celda vacía a la derecha
    const startCol = targetCol + 1;
    let placedCol = -1;
    setData(prev => {
      const next = prev.map(r => [...r]);
      let resultCol = startCol;
      while (resultCol < next[0].length && next[targetRow] && String(next[targetRow][resultCol] ?? '').trim() !== '') resultCol++;
      if (resultCol >= next[0].length) {
        const need = resultCol - next[0].length + 1;
        for (let i = 0; i < need; i++) {
          for (let rr = 0; rr < next.length; rr++) next[rr].push('');
        }
      }
      next[targetRow][resultCol] = `Resultado = ${res.value}`;
      placedCol = resultCol;
      return next;
    });
    // ajustar colWidths si se añadieron columnas
    setColWidths(prev => {
      const cur = prev.slice();
      const needed = Math.max(0, (data[0]?.length ?? 0) - cur.length);
      if (needed > 0) return [...cur, ...Array(needed).fill(96)];
      return cur;
    });
    if (placedCol >= 0) setSelected({ row: targetRow, col: placedCol });
    setFormulaText(String(res.value));
  };

  const clearFormulaAndSelection = () => {
    setFormulaText('');
    setComputedSum(null);
    setSelectionStart(null);
    setSelectionEnd(null);
    setExtraSelections([]);
    setSelected(null);
  };

  // Quick action: build formula from current selection and apply (SUM/AVERAGE/COUNT)
  const applyQuickFunc = (fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'CONCAT') => {
    if (!selectionStart || !selectionEnd) return;
    const r1 = Math.min(selectionStart.row, selectionEnd.row);
    const r2 = Math.max(selectionStart.row, selectionEnd.row);
    const c1 = Math.min(selectionStart.col, selectionEnd.col);
    const c2 = Math.max(selectionStart.col, selectionEnd.col);
    // construir referencia A1:B2
    const colIdxToName = (idx: number) => {
      let s = '';
      let n = idx + 1;
      while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    };
    const rangeStr = `${colIdxToName(c1)}${r1 + 1}:${colIdxToName(c2)}${r2 + 1}`;
    const formula = `=${fn}(${rangeStr})`;
    setFormulaText(formula);
    // aplicar inmediatamente
    const res = computeFormulaResult(formula);
    if (res) {
      setData(prev => {
        const next = prev.map(r => [...r]);
        next[r1][c1] = String(res.value);
        return next;
      });
      setSelected({ row: r1, col: c1 });
      setSelectionStart(null);
      setSelectionEnd(null);
      setComputedSum(res.value as number);
    }
  };

  const exportToExcel = () => {
    // convertir data a hoja
    const ws = XLSX.utils.aoa_to_sheet(data.map(r => [...r]));
    // añadir merges al worksheet
    if (merges.length) {
      ws['!merges'] = merges.map(m => ({ s: { r: m.r, c: m.c }, e: { r: m.r + m.rows - 1, c: m.c + m.cols - 1 } }));
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'export.xlsx');
  };

  // Limpia la selección actual o toda la hoja si no hay selección
  const clearSelectedOrAll = () => {
    console.debug('clearSelectedOrAll invoked', { selectionStart, selectionEnd, extraSelections, selected });
    pushHistory();
    setData(prev => {
      const next = prev.map(r => [...r]);
      // selección principal
      if (selectionStart && selectionEnd) {
        const r1 = Math.min(selectionStart.row, selectionEnd.row);
        const r2 = Math.max(selectionStart.row, selectionEnd.row);
        const c1 = Math.min(selectionStart.col, selectionEnd.col);
        const c2 = Math.max(selectionStart.col, selectionEnd.col);
        for (let rr = r1; rr <= r2; rr++) {
          for (let cc = c1; cc <= c2; cc++) {
            if (next[rr] && next[rr][cc] != null) next[rr][cc] = '';
          }
        }
        // limpiar selección visual
        // eliminar merges que intersecten la región limpiada
        setMerges(prev => prev.filter(m => (m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2)));
        setSelectionStart(null);
        setSelectionEnd(null);
        setExtraSelections([]);
        setSelected(null);
        showToast('Rango limpiado y merges separados si existían');
        return next;
      }
      // selecciones adicionales
      if (extraSelections && extraSelections.length) {
        for (const s of extraSelections) {
          for (let rr = s.r1; rr <= s.r2; rr++) {
            for (let cc = s.c1; cc <= s.c2; cc++) {
              if (next[rr] && next[rr][cc] != null) next[rr][cc] = '';
            }
          }
        }
        setExtraSelections([]);
        // eliminar merges que intersecten cualquiera de las selecciones adicionales
        setMerges(prev => prev.filter(m => {
          for (const s of extraSelections) {
            const r1 = s.r1, r2 = s.r2, c1 = s.c1, c2 = s.c2;
            if (!((m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2))) {
              return false; // intersecta -> quitar
            }
          }
          return true;
        }));
        setSelected(null);
        showToast('Selecciones limpiadas y merges separados si existían');
        return next;
      }
      // celda individual seleccionada
      if (selected) {
        const { row, col } = selected;
        if (next[row] && next[row][col] != null) next[row][col] = '';
        setSelected(null);
        // eliminar cualquier merge que contenga esta celda
        setMerges(prev => prev.filter(m => !(row >= m.r && row < m.r + m.rows && col >= m.c && col < m.c + m.cols)));
        showToast('Celda limpiada y merge separado si existía');
        return next;
      }
      // sin selección: confirmar y limpiar toda la hoja
      if (!window.confirm('No hay selección. ¿Limpiar toda la hoja?')) return prev;
      for (let rr = 0; rr < next.length; rr++) {
        for (let cc = 0; cc < next[rr].length; cc++) next[rr][cc] = '';
      }
      // borrar todos los merges
      setMerges([]);
      setSelected(null);
      setSelectionStart(null);
      setSelectionEnd(null);
      setExtraSelections([]);
      showToast('Hoja completamente limpiada');
      return next;
    });
  };

  // aplicar estilos condicionales antes de render
  const computeCellStyles = (r: number, c: number) => {
    let style: React.CSSProperties = {};
    let cls = '';
    // search matches -> use tailwind ring
    if (searchMatches.find(m => m.row === r && m.col === c)) {
      cls += ' ring-2 ring-orange-400/70';
    }
    // conditional formats
    for (const cf of conditionalFormats) {
      const scope = cf.scope;
      const inScope = !scope || (r >= cf.scope!.r1 && r <= cf.scope!.r2 && c >= cf.scope!.c1 && c <= cf.scope!.c2);
      if (!inScope) continue;
      const cellVal = String(data[r][c] ?? '');
      if (cf.type === 'gt') { if (!Number.isNaN(Number(cellVal)) && Number(cellVal) > Number(cf.value)) { style.background = cf.bg; } }
      if (cf.type === 'lt') { if (!Number.isNaN(Number(cellVal)) && Number(cellVal) < Number(cf.value)) { style.background = cf.bg; } }
      if (cf.type === 'eq') { if (cellVal === cf.value) style.background = cf.bg; }
      if (cf.type === 'contains') { if (cellVal.indexOf(cf.value) !== -1) style.background = cf.bg; }
    }
    // highlight temporal después de pegar
    if (pasteHighlight) {
      if (r >= pasteHighlight.r1 && r <= pasteHighlight.r2 && c >= pasteHighlight.c1 && c <= pasteHighlight.c2) {
        // use tailwind ring + translucent bg for highlight and a fade class
        cls += ' ring-4 ring-emerald-400/80';
        if (!style.background) style.background = 'rgba(0,200,120,0.12)';
        if (pasteFade) cls += ' opacity-0 transition-opacity duration-700';
        else cls += ' transition-opacity duration-700 opacity-100';
      }
    }
    return { style, cls } as any;
  };

  const onFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const arrayBuffer = ev.target?.result as ArrayBuffer | null;
      if (!arrayBuffer) return;
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1 });
      // leer merges si existen
      const sheetMerges: any[] = worksheet['!merges'] || [];
      const importedMerges = sheetMerges.map(m => ({ r: m.s.r, c: m.s.c, rows: m.e.r - m.s.r + 1, cols: m.e.c - m.s.c + 1 }));
      // normalizar filas/columnas
      const maxCols = Math.max(...json.map((r: any) => r.length));
      const normalized = json.map((r: any) => {
        const arr = Array.from({ length: maxCols }, (_, i) => (r[i] !== undefined ? String(r[i]) : ''));
        return arr;
      });
      setData(normalized.length ? normalized : createInitialData(INITIAL_ROWS, INITIAL_COLS));
      setMerges(importedMerges);
      // ajustar colWidths
      setColWidths(prev => {
        const needed = normalized[0]?.length || INITIAL_COLS;
        return Array.from({ length: needed }, (_, i) => prev[i] ?? 96);
      });
    };
    reader.readAsArrayBuffer(file);
    // limpiar
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Estado para mostrar/ocultar el menú COUNT/COUNTIF
  const [showCountMenu, setShowCountMenu] = useState(false);
  // Estado para mostrar/ocultar el menú SUM/SUMIF
  const [showSumMenu, setShowSumMenu] = useState(false);
  // Estilos globales para modo oscuro/claro
  // We rely on Tailwind `dark:` variants for styling. No injected CSS needed.

  return (
    <div className={"min-h-screen p-6 bg-gray-50 dark:bg-gray-900 dark:text-gray-100"}>
      <Toolbar
        BTN={BTN}
        data={data}
        selectedRow={selectedRow}
        setSelectedRow={setSelectedRow}
        addRow={addRow}
        deleteRow={deleteRow}
        selectedCol={selectedCol}
        setSelectedCol={setSelectedCol}
        addCol={addCol}
        deleteCol={deleteCol}
        fileInputRef={fileInputRef}
        onFileInputChange={onFileInputChange}
        exportToExcel={exportToExcel}
        clearSelectedOrAll={clearSelectedOrAll}
        darkMode={darkMode}
        setDarkMode={setDarkMode}
        combineCells={combineCells}
        separateCells={separateCells}
        showToast={showToast}
        copySelectionToClipboard={copySelectionToClipboard}
        pasteClipboardAtSelection={pasteClipboardAtSelection}
        pasteTransposedAtSelection={pasteTransposedAtSelection}
        undo={undo}
        redo={redo}
        history={history}
        future={future}
        findText={findText}
        setFindText={setFindText}
        replaceText={replaceText}
        setReplaceText={setReplaceText}
        findMatches={findMatches}
        replaceCurrent={replaceCurrent}
        replaceAll={replaceAll}
        freezeRows={freezeRows}
        setFreezeRows={setFreezeRows}
        freezeCols={freezeCols}
        setFreezeCols={setFreezeCols}
        setShowSortModal={setShowSortModal}
        getColName={getColName}
        selectionStart={selectionStart}
        showCFModal={showCFModal}
        setShowCFModal={setShowCFModal}
        setCfTypeInput={setCfTypeInput}
        setCfValueInput={setCfValueInput}
        setCfColorInput={setCfColorInput}
        cfScopeInput={cfScopeInput}
        setCfScopeInput={setCfScopeInput}
        compactMode={compactMode}
        setCompactMode={setCompactMode}
        showSumMenu={showSumMenu}
        setShowSumMenu={setShowSumMenu}
        applyQuickFunc={applyQuickFunc}
        showCountMenu={showCountMenu}
        setShowCountMenu={setShowCountMenu}
        setFormulaText={setFormulaText}
        applyFormulaToSelection={applyFormulaToSelection}
        clearFormulaAndSelection={clearFormulaAndSelection}
        computedSum={computedSum}
      />

  {/* Menú viejo eliminado: mantenemos solo el Toolbar superior y la tabla */}

      <div className="overflow-auto border rounded shadow-sm bg-white dark:bg-gray-800">
        <table ref={tableRef} className="min-w-[1200px] w-full table-fixed" style={{ borderCollapse: 'separate', borderSpacing: 0 }}>
          <thead className="sticky top-0 bg-gray-100 dark:bg-gray-700">
            <tr>
              <th className="bg-gray-100 text-gray-700 border-r border-gray-200 w-10 dark:bg-gray-700 dark:text-gray-200 dark:border-gray-600">&nbsp;</th>
              {Array.from({ length: data[0].length }).map((_, colIdx) => (
                <th key={colIdx} data-header-col={colIdx} className="bg-gray-100 text-gray-700 text-center font-medium border-b border-gray-200 relative dark:bg-gray-700 dark:text-gray-200" style={{ width: colWidths[colIdx], height: 36 }}
                  onMouseDown={e => {
                    // iniciar drag de selección por columna (sin tocar filas)
                    e.preventDefault();
                    e.stopPropagation();
                    headerDragging.current = true;
                    headerDragStart.current = colIdx;
                    // iniciar selección: si Shift está presionado, expandir desde selectionStart o selected; si no, iniciar nueva selección en toda la columna
                    if (shiftPressed.current) {
                      if (selectionStart && selectionEnd) {
                        // mantener selecciónStart y solo cambiar selectionEnd columna
                        setSelectionEnd(prev => prev ? { row: prev.row, col: colIdx } : { row: 0, col: colIdx });
                      } else {
                        // si no hay selectionStart, anclar la selección a la columna clicada (no a `selected` que pueda pertenecer a otra columna)
                        setSelectionStart({ row: 0, col: colIdx });
                        setSelectionEnd({ row: data.length - 1, col: colIdx });
                      }
                    } else {
                      // seleccionar solo la columna completa (todas las filas)
                      setSelectionStart({ row: 0, col: colIdx });
                      setSelectionEnd({ row: data.length - 1, col: colIdx });
                      setSelected({ row: 0, col: colIdx });
                    }
                  }}
                  onClick={e => {
                    e.stopPropagation();
                    if (hadDragSinceMouseDown.current) { hadDragSinceMouseDown.current = false; return; }
                    // Ctrl+click toggleea la columna en extraSelections
                    if (ctrlPressed.current) {
                      const exists = extraSelections.find(s => s.c1 === colIdx && s.c2 === colIdx);
                      if (exists) {
                        setExtraSelections(prev => prev.filter(s => !(s.c1 === colIdx && s.c2 === colIdx)));
                      } else {
                        setExtraSelections(prev => [...prev, { r1: 0, c1: colIdx, r2: data.length - 1, c2: colIdx }]);
                      }
                      return;
                    }
                    // limpiar selecciones aditivas previas y establecer nueva selección
                    setExtraSelections([]);
                    setSelectionStart({ row: 0, col: colIdx });
                    setSelectionEnd({ row: data.length - 1, col: colIdx });
                    setSelected({ row: 0, col: colIdx });
                  }}
                >
                  <div className="flex items-center justify-between px-2">
                    <span className="select-none">{getColName(colIdx)}</span>
                    <div
                      className="absolute top-0 bottom-0 cursor-col-resize z-40"
                      style={{ right: -6, width: 12 }}
                      onMouseDown={e => {
                        e.stopPropagation();
                        e.preventDefault();
                        resizingCol.current = colIdx;
                        startX.current = e.clientX;
                        startWidth.current = colWidths[colIdx] ?? 96;
                        document.body.style.cursor = 'col-resize';
                      }}
                      onDoubleClick={() => {
                        // resetear ancho a valor por defecto
                        setColWidths(prev => {
                          const next = [...prev];
                          next[colIdx] = 96;
                          return next;
                        });
                      }}
                    />
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rowIdx) => (
              <tr key={rowIdx} className="align-top">
                <th className="bg-gray-50 text-gray-600 border-r border-gray-100 w-10 text-center font-medium dark:bg-gray-700 dark:text-gray-200" style={{ height: 36 }}>{rowIdx + 1}</th>
                {row.map((cell, colIdx) => {
                  // comprobar si esta celda está cubierta por un merge existente (y si no es la cabecera del merge)
                  const covering = merges.find(m => rowIdx >= m.r && rowIdx < m.r + m.rows && colIdx >= m.c && colIdx < m.c + m.cols);
                  if (covering) {
                    // si no es la celda inicial del merge, omitimos el render (queda cubierta)
                    if (covering.r !== rowIdx || covering.c !== colIdx) return null;
                    // si es la celda inicial, renderizamos con rowSpan/colSpan
                    const isMultiSelected = isCellInSelections(rowIdx, colIdx) || isCellInSelections(covering.r + covering.rows - 1, covering.c + covering.cols - 1);
                    // calcular ancho como suma de anchos de las columnas cubiertas
                    const spanWidth = colWidths.slice(colIdx, colIdx + covering.cols).reduce((a, b) => a + (b ?? 96), 0);
                    const cs = computeCellStyles(rowIdx, colIdx);
                    return (
                      <Cell
                        key={colIdx}
                        value={cell}
                        row={rowIdx}
                        col={colIdx}
                        rowSpan={covering.rows}
                        colSpan={covering.cols}
                        width={spanWidth}
                        height={36 * covering.rows}
                        isSelected={!!(selected && selected.row === rowIdx && selected.col === colIdx)}
                        isMultiSelected={isMultiSelected}
                        isDragging={isDragging}
                        onCommit={commitCell}
                        onMouseDown={onCellMouseDown}
                        onMouseEnter={onCellMouseEnter}
                        onMouseUp={onCellMouseUp}
                        onClick={onCellClick}
                        extraStyle={cs.style}
                        extraClass={cs.cls}
                      />
                    );
                  }

                  const isSelected = selected && selected.row === rowIdx && selected.col === colIdx;
                  const isMultiSelected = isCellInSelections(rowIdx, colIdx);
                  const cs = computeCellStyles(rowIdx, colIdx);
                  return (
                    <Cell
                      key={colIdx}
                      value={cell}
                      row={rowIdx}
                      col={colIdx}
                      width={colWidths[colIdx]}
                      height={36}
                      isSelected={!!isSelected}
                      isMultiSelected={isMultiSelected}
                      isDragging={isDragging}
                      onCommit={commitCell}
                      onMouseDown={onCellMouseDown}
                      onMouseEnter={onCellMouseEnter}
                      onMouseUp={onCellMouseUp}
                      onClick={onCellClick}
                      extraStyle={cs.style}
                      extraClass={cs.cls}
                    />
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Modales y Toasts */}
      {showSortModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40">
          <div className="bg-white p-4 rounded shadow-lg w-80 dark:bg-gray-800 dark:text-gray-100">
            <h3 className="font-bold mb-2">Sort Range</h3>
            <div className="mb-2">
              <label className="block text-xs">Columna (ej: A)</label>
              <input value={sortColInput} onChange={e => setSortColInput(e.target.value)} className="w-full px-2 h-9 border rounded bg-white text-gray-800 dark:bg-gray-700 dark:text-gray-100 dark:border-gray-600" />
            </div>
            <div className="mb-3">
              <label className="block text-xs">Dirección</label>
              <div className="relative group">
                <select value={sortDirection} onChange={e => setSortDirection(e.target.value as any)} className="w-full pr-8 px-2 h-9 border rounded bg-white text-gray-800 dark:bg-gray-700 dark:text-gray-100 dark:border-gray-600 appearance-none">
                  <option value="asc">Ascendente</option>
                  <option value="desc">Descendente</option>
                </select>
                <span className="pointer-events-none absolute right-2 top-1/2 transform -translate-y-1/2 text-gray-500 dark:text-gray-300 transition-transform duration-200 group-focus-within:rotate-180">
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" className="w-4 h-4">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                  </svg>
                </span>
              </div>
            </div>
            <div className="flex justify-end gap-2">
              <button onClick={() => setShowSortModal(false)} className={`${BTN} border`}>Cancelar</button>
              <button onClick={() => { const ci = colNameToIndex(sortColInput); setShowSortModal(false); sortRangeByColumn(ci, sortDirection); }} className={`${BTN} bg-indigo-600 text-white`}>Aplicar</button>
            </div>
          </div>
        </div>
      )}

      {showCombineConfirm && pendingCombineBox && (
        <Modal title="Confirmar combinación" onClose={() => { setShowCombineConfirm(false); setPendingCombineBox(null); }}>
          <div className="mb-3 text-sm">La selección con Ctrl+click no es contigua. Se propone combinar el rectángulo {`(${pendingCombineBox.r1 + 1}, ${pendingCombineBox.c1 + 1}) - (${pendingCombineBox.r2 + 1}, ${pendingCombineBox.c2 + 1})`}. ¿Continuar?</div>
          <div className="flex justify-end gap-2">
            <button onClick={() => { setShowCombineConfirm(false); setPendingCombineBox(null); }} className={`${BTN} border`}>Cancelar</button>
            <button onClick={() => { if (pendingCombineBox) performCombineRegion(pendingCombineBox.r1, pendingCombineBox.r2, pendingCombineBox.c1, pendingCombineBox.c2); setShowCombineConfirm(false); setPendingCombineBox(null); }} className={`${BTN} bg-purple-600 text-white`}>Confirmar</button>
          </div>
        </Modal>
      )}

      {showCFModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40">
          <div className="bg-white p-4 rounded shadow-lg w-80 dark:bg-gray-800 dark:text-gray-100">
            <h3 className="font-bold mb-2">Add Conditional Format</h3>
            <div className="mb-2">
              <label className="block text-xs">Tipo</label>
              <div className="relative group">
                <select value={cfTypeInput} onChange={e => setCfTypeInput(e.target.value as any)} className="w-full pr-8 px-2 h-9 border rounded bg-white text-gray-800 dark:bg-gray-700 dark:text-gray-100 dark:border-gray-600 appearance-none">
                  <option value="gt">Greater than (gt)</option>
                  <option value="lt">Less than (lt)</option>
                  <option value="eq">Equals (eq)</option>
                  <option value="contains">Contains</option>
                </select>
                <span className="pointer-events-none absolute right-2 top-1/2 transform -translate-y-1/2 text-gray-500 dark:text-gray-300 transition-transform duration-200 group-focus-within:rotate-180">
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" className="w-4 h-4">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                  </svg>
                </span>
              </div>
            </div>
            <div className="mb-2">
              <label className="block text-xs">Valor</label>
              <input value={cfValueInput} onChange={e => setCfValueInput(e.target.value)} className="w-full px-2 h-9 border rounded bg-white text-gray-800 dark:bg-gray-700 dark:text-gray-100 dark:border-gray-600" />
            </div>
            <div className="mb-2">
              <label className="block text-xs">Scope</label>
              <div className="relative group">
                <select value={cfScopeInput} onChange={e => setCfScopeInput(e.target.value as any)} className="w-full pr-8 px-2 h-9 border rounded bg-white text-gray-800 dark:bg-gray-700 dark:text-gray-100 dark:border-gray-600 appearance-none">
                  <option value="selection">Selección actual</option>
                  <option value="sheet">Toda la hoja</option>
                </select>
                <span className="pointer-events-none absolute right-2 top-1/2 transform -translate-y-1/2 text-gray-500 dark:text-gray-300 transition-transform duration-200 group-focus-within:rotate-180">
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" className="w-4 h-4">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                  </svg>
                </span>
              </div>
            </div>
            <div className="mb-3">
              <label className="block text-xs">Color de fondo</label>
              <input type="color" value={cfColorInput} onChange={e => setCfColorInput(e.target.value)} className="w-full h-9 p-1 border rounded" />
            </div>
            <div className="flex justify-end gap-2">
              <button onClick={() => setShowCFModal(false)} className={`${BTN} border`}>Cancelar</button>
              <button
                onClick={() => { setShowCFModal(false); addConditionalFormat(cfTypeInput, cfValueInput, cfColorInput, cfScopeInput); }}
                disabled={cfScopeInput === 'selection' && !(selectionStart && selectionEnd)}
                title={cfScopeInput === 'selection' && !(selectionStart && selectionEnd) ? 'Selecciona un rango antes de añadir formato condicional' : 'Añadir formato condicional'}
                className={
                  (cfScopeInput === 'selection' && !(selectionStart && selectionEnd))
                    ? `${BTN} bg-pink-300 text-white opacity-60 cursor-not-allowed`
                    : `${BTN} bg-pink-600 text-white`
                }
              >Añadir</button>
            </div>
          </div>
        </div>
      )}

      {toastMsg && (
        <div className="fixed right-4 bottom-6 z-60 bg-black text-white h-9 px-3 flex items-center rounded shadow">{toastMsg}</div>
      )}

    </div>
  );
}

export default ExcelComponent;