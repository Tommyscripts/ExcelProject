import React, { useState, useRef, useEffect, useCallback } from 'react';
import TooltipCooldown from './Auxiliares/TooltipCooldown';
import { FaTrash } from 'react-icons/fa';
import * as XLSX from 'xlsx';

const INITIAL_COLS = 22; // A-V
const INITIAL_ROWS = 46;

function getColName(idx: number) {
  return String.fromCharCode(65 + idx);
}

const createInitialData = (rows: number, cols: number) =>
  Array.from({ length: rows }, () => Array(cols).fill(''));

const ExcelComponent: React.FC = () => {
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

  // --- Historial para Undo/Redo ---
  const HISTORY_LIMIT = 50;
  const [history, setHistory] = useState<any[]>([]);
  const [future, setFuture] = useState<any[]>([]);
  const pushHistory = useCallback((snapshot?: any) => {
    const snap = snapshot ?? JSON.parse(JSON.stringify({ data, merges, colWidths, selected, selectionStart, selectionEnd }));
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
      setFuture(f => [JSON.parse(JSON.stringify({ data, merges, colWidths, selected, selectionStart, selectionEnd })), ...f]);
      // restaurar
      setData(last.data ?? last);
      setMerges(last.merges ?? last.merges ?? last.merges ?? (last.merges || []));
      setColWidths(last.colWidths ?? (last.colWidths || colWidths));
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
      // push current into history
      setHistory(h => [...h, JSON.parse(JSON.stringify({ data, merges, colWidths, selected, selectionStart, selectionEnd }))]);
      setData(first.data ?? first);
      setMerges(first.merges ?? (first.merges || []));
      setColWidths(first.colWidths ?? (first.colWidths || colWidths));
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
    if (!selectionStart || !selectionEnd) return;
    const r1 = Math.min(selectionStart.row, selectionEnd.row);
    const r2 = Math.max(selectionStart.row, selectionEnd.row);
    const c1 = Math.min(selectionStart.col, selectionEnd.col);
    const c2 = Math.max(selectionStart.col, selectionEnd.col);
    const colLetter = window.prompt('Columna para ordenar (ej: A, B, ... relative a hoja):', getColName(c1));
    if (!colLetter) return;
    const colIdx = colNameToIndex(colLetter);
    if (colIdx < c1 || colIdx > c2) {
      alert('La columna debe estar dentro de la selección.');
      return;
    }
    const dir = window.prompt('Dirección: asc o desc', 'asc') || 'asc';
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
  };

  // --- Conditional Formatting (simple) ---
  type CFRule = { id: string; type: 'gt'|'lt'|'eq'|'contains'; value: string; bg?: string; color?: string; scope?: { r1:number,c1:number,r2:number,c2:number } | null };
  const [conditionalFormats, setConditionalFormats] = useState<CFRule[]>([]);
  const addConditionalFormatViaPrompt = () => {
    const type = window.prompt('Tipo de regla: gt, lt, eq, contains', 'gt') as any;
    if (!type) return;
    const value = window.prompt('Valor (para contains escribe texto)', '0') || '0';
    const bg = window.prompt('Color de fondo (css)', '#fffbcc') || '#fffbcc';
    const id = String(Date.now());
    setConditionalFormats(prev => [...prev, { id, type, value, bg, color: undefined, scope: null }]);
  };

  // Memoized cell component to reduce re-renders while typing
  // Defined inside component to capture types but memoized to avoid re-render when props unchanged
  const Cell = React.useMemo(() => React.memo(function CellInner(props: {
    value: string;
    row: number;
    col: number;
    rowSpan?: number;
    colSpan?: number;
    width?: number;
    height?: number;
    isSelected?: boolean;
    isMultiSelected?: boolean;
    isDragging?: boolean;
    onCommit: (row: number, col: number, value: string) => void;
    onMouseDown: (r: number, c: number) => void;
    onMouseEnter: (r: number, c: number) => void;
    onMouseUp: () => void;
    onClick: (r: number, c: number) => void;
  }) {
    const { value, row, col, rowSpan, colSpan, width, height, isSelected, isMultiSelected, isDragging, onCommit, onMouseDown, onMouseEnter, onMouseUp, onClick } = props;
  // support additional style props passed via data-attributes
  const [extraStyle, setExtraStyle] = useState<React.CSSProperties>({});
  const [extraClass, setExtraClass] = useState<string>('');
    const [local, setLocal] = useState<string>(value ?? '');
    const isEditing = useRef(false);
    const commitTimer = useRef<number | null>(null);

    useEffect(() => {
      // update local when value prop changes, unless user is editing
      if (!isEditing.current) setLocal(value ?? '');
    }, [value]);

    const doCommit = useCallback(() => {
      if (commitTimer.current) {
        window.clearTimeout(commitTimer.current);
        commitTimer.current = null;
      }
      if (local !== value) onCommit(row, col, local);
    }, [local, value, row, col, onCommit]);

    const onChange = (e: React.ChangeEvent<HTMLInputElement>) => {
      setLocal(e.target.value);
      // debounce commit
      if (commitTimer.current) window.clearTimeout(commitTimer.current);
      // commit after 300ms of inactivity
      commitTimer.current = window.setTimeout(() => {
        commitTimer.current = null;
        if (local !== e.target.value) {
          // local was updated above but closure may reference old value; use e.target.value
          onCommit(row, col, e.target.value);
        } else {
          doCommit();
        }
      }, 300) as unknown as number;
    };

    const handleFocus = () => { isEditing.current = true; };
    const handleBlur = () => { isEditing.current = false; doCommit(); };
    const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === 'Enter') {
        (e.target as HTMLInputElement).blur();
      }
    };

  const inputRef = useRef<HTMLInputElement | null>(null);
  const [hasOverflow, setHasOverflow] = useState(false);
  // Mostrar tooltip con un cooldown para permitir manejar la celda sin que aparezca instantáneamente
  const HOVER_DELAY_MS = 1500;
  const hoverTimer = useRef<number | null>(null);
  const [showTooltip, setShowTooltip] = useState(false);

    useEffect(() => {
  if (!inputRef.current) return;
      const el = inputRef.current;
      const check = () => {
        try {
          const over = el.scrollWidth > el.clientWidth + 4; // small tolerance
          setHasOverflow(!!over);
        } catch (err) {
          setHasOverflow(false);
        }
      };
      check();
      const ro = new ResizeObserver(check);
      ro.observe(el);
      return () => ro.disconnect();
    }, [local, width]);

    // leer atributos data-extra-style/data-extra-class del td parent al montar
    useEffect(() => {
      try {
        const td = inputRef.current?.closest('td') as HTMLElement | null;
        if (!td) return;
        const s = (td.getAttribute('data-extra-style') || '');
        const c = (td.getAttribute('data-extra-class') || '');
        if (s) {
          try { setExtraStyle(JSON.parse(s)); } catch (e) { /* ignore */ }
        }
        if (c) setExtraClass(c);
      } catch (e) {
        // ignore
      }
    }, []);

    // limpiar timer si el componente se desmonta
    useEffect(() => {
      return () => {
        if (hoverTimer.current) {
          window.clearTimeout(hoverTimer.current);
          hoverTimer.current = null;
        }
      };
    }, []);

    return (
      <td rowSpan={rowSpan} colSpan={colSpan} data-cell-row={row} data-cell-col={col}
  className={`p-0 relative ${isSelected || isMultiSelected || isDragging ? 'ring-2 ring-blue-400 bg-blue-50' : 'border border-gray-200'} ${extraClass}`}
  style={{ width, height, ...(extraStyle || {}) }}
        onMouseDown={() => onMouseDown(row, col)}
        onMouseEnter={() => {
          onMouseEnter(row, col);
          // si la celda está seleccionada, mostramos el tooltip inmediatamente
          if (isSelected) {
            if (hoverTimer.current) { window.clearTimeout(hoverTimer.current); hoverTimer.current = null; }
            setShowTooltip(true);
            return;
          }
          // iniciar cooldown para mostrar tooltip
          if (hoverTimer.current) window.clearTimeout(hoverTimer.current);
          setShowTooltip(false);
          hoverTimer.current = window.setTimeout(() => {
            hoverTimer.current = null;
            setShowTooltip(true);
          }, HOVER_DELAY_MS) as unknown as number;
        }}
        onMouseLeave={() => {
          // cancelar cooldown y ocultar tooltip
          if (hoverTimer.current) {
            window.clearTimeout(hoverTimer.current);
            hoverTimer.current = null;
          }
          setShowTooltip(false);
        }}
        onMouseUp={() => {
          // al soltar, ejecutar handler y ocultar tooltip
          if (hoverTimer.current) {
            window.clearTimeout(hoverTimer.current);
            hoverTimer.current = null;
          }
          setShowTooltip(false);
          onMouseUp();
        }}
        onClick={() => {
          if (hoverTimer.current) {
            window.clearTimeout(hoverTimer.current);
            hoverTimer.current = null;
          }
          setShowTooltip(false);
          onClick(row, col);
        }}>
        <input
          ref={inputRef}
          type="text"
          value={local}
          onChange={onChange}
          onMouseEnter={() => {
            // Si la celda está seleccionada, mostramos tooltip inmediatamente.
            if (isSelected) {
              if (hoverTimer.current) { window.clearTimeout(hoverTimer.current); hoverTimer.current = null; }
              setShowTooltip(true);
              return;
            }
            // Si ya hay un timer o el tooltip ya se muestra, no reiniciamos nada.
            if (showTooltip || hoverTimer.current) return;
            hoverTimer.current = window.setTimeout(() => { hoverTimer.current = null; setShowTooltip(true); }, HOVER_DELAY_MS) as unknown as number;
          }}
          onMouseLeave={() => {
            // no cancelamos aquí para evitar flicker cuando se mueve entre td y input;
            // la salida del td es la que cancela el tooltip.
          }}
          onFocus={handleFocus}
          onBlur={handleBlur}
          onKeyDown={handleKeyDown}
          className="w-full h-full px-2 bg-transparent focus:outline-none text-sm cursor-text"
          tabIndex={0}
        />
  {(showTooltip && !isEditing.current && String(local).trim() !== '') && (
          <div className="absolute left-0 top-0 z-50 p-2 bg-white border rounded shadow-md text-sm max-w-[400px] break-words">
            <div className="text-xs text-gray-500 mb-1">Texto completo</div>
            {local}
          </div>
        )}
      </td>
    );
  }), []);

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

  const handleChange = (row: number, col: number, value: string) => {
    setData(prev => {
      if (!prev[row]) return prev;
      const next = prev.map(r => [...r]);
      next[row][col] = value;
      return next;
    });
  };

  const commitCell = useCallback((row: number, col: number, value: string) => {
    handleChange(row, col, value);
  }, [handleChange]);

  const onCellMouseDown = useCallback((r: number, c: number) => handleCellMouseDown(r, c), [/* deps none, uses stable handleCellMouseDown */]);
  const onCellMouseEnter = useCallback((r: number, c: number) => handleCellMouseEnter(r, c), []);
  const onCellMouseUp = useCallback(() => handleCellMouseUp(), []);
  const onCellClick = useCallback((r: number, c: number) => handleCellClick(r, c), []);

  const handleCellClick = (row: number, col: number) => {
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
  };

  const handleCellMouseDown = (row: number, col: number) => {
    // Iniciar posible selección por mousedown.
    // Si Shift está presionado y hay una celda previamente seleccionada, usarla como ancla.
    // Si no, anclar la selección a la celda clicada para evitar rangos accidentales.
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
  };

  const handleCellMouseEnter = (row: number, col: number) => {
    if (selectionStart) {
      setSelectionEnd({ row, col });
    }
    // Si el usuario mantiene pulsado el ratón y mueve sobre otras celdas,
    // marcamos que hubo un arrastre para que el click posterior no lo reinterprete.
    if (isMouseDown.current) {
      hadDragSinceMouseDown.current = true;
    }
  };

  const handleCellMouseUp = () => {
  // No limpiamos selectionStart aquí para que la selección persista después de soltar.
  // La selección se limpiará explícitamente en combineCells/separateCells o cuando el usuario la reemplace.
  isMouseDown.current = false;
  setIsDragging(false);
  };

  // Listener global para asegurar que isDragging se desactive si el usuario suelta fuera de la tabla
  React.useEffect(() => {
    const onUp = () => {
      if (isMouseDown.current) {
        isMouseDown.current = false;
      }
      setIsDragging(false);
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
    // Determinar la región a combinar: preferimos la selección principal
    let r1: number, r2: number, c1: number, c2: number;
    if (selectionStart && selectionEnd) {
      r1 = Math.min(selectionStart.row, selectionEnd.row);
      r2 = Math.max(selectionStart.row, selectionEnd.row);
      c1 = Math.min(selectionStart.col, selectionEnd.col);
      c2 = Math.max(selectionStart.col, selectionEnd.col);
    } else if (extraSelections && extraSelections.length) {
      // si no hay selección principal, usar bounding box de extraSelections
      r1 = Math.min(...extraSelections.map(s => s.r1));
      r2 = Math.max(...extraSelections.map(s => s.r2));
      c1 = Math.min(...extraSelections.map(s => s.c1));
      c2 = Math.max(...extraSelections.map(s => s.c2));
    } else {
      return;
    }
    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;
    if (rows === 1 && cols === 1) return; // nada que combinar
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

    // Eliminar merges que intersecten la región seleccionada
    setMerges(prev => prev.filter(m => (m.r + m.rows - 1 < r1) || (m.r > r2) || (m.c + m.cols - 1 < c1) || (m.c > c2)));

    // Añadir nuevo merge
    setMerges(prev => [...prev, { r: r1, c: c1, rows, cols }]);
    // limpiar selección principal y adicionales
    setSelectionStart(null);
    setSelectionEnd(null);
    setExtraSelections([]);
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
      return;
    }
    // si no hay selección múltiple, separar el merge que contenga la celda seleccionada
    if (selected) {
      const { row, col } = selected;
      setMerges(prev => prev.filter(m => !(row >= m.r && row < m.r + m.rows && col >= m.c && col < m.c + m.cols)));
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

      setSelectionEnd({ row, col });
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
        hadDragSinceMouseDown.current = true;
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
      // escribir resultado en la celda adyacente a la derecha como 'Resultado = <valor>'
      const resultCol = targetCol + 1;
      if (resultCol >= data[0].length) {
        // ampliar tabla con una columna adicional
        setData(prev => prev.map(r => [...r, '']));
        setColWidths(prev => [...prev, 96]);
      }
      setData(prev => {
        const next = prev.map(r => [...r]);
        next[targetRow][resultCol] = `Resultado = ${res.value}`;
        return next;
      });
      setSelected({ row: targetRow, col: resultCol });
      setSelectionStart(null);
      setSelectionEnd(null);
      setFormulaText(String(res.value));
      return;
    }

  // no aplicamos a extraSelections automáticamente para evitar efectos inesperados

    // fallback: aplicar en la celda seleccionada
    const targetRow = selected?.row ?? 0;
    const targetCol = selected?.col ?? 0;
    // escribir resultado en la celda adyacente a la derecha si es posible
    const resultCol = targetCol + 1;
    if (resultCol >= data[0].length) {
      setData(prev => prev.map(r => [...r, '']));
      setColWidths(prev => [...prev, 96]);
    }
    setData(prev => {
      const next = prev.map(r => [...r]);
      next[targetRow][resultCol] = `Resultado = ${res.value}`;
      return next;
    });
    setSelected({ row: targetRow, col: resultCol });
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
  const applyQuickFunc = (fn: 'SUM' | 'AVERAGE' | 'COUNT') => {
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
      // poner en celda superior-izquierda
      setData(prev => {
        const next = prev.map(r => [...r]);
        next[r1][c1] = String(res.value);
        return next;
      });
      setSelected({ row: r1, col: c1 });
      setSelectionStart(null);
      setSelectionEnd(null);
      setComputedSum(res.value);
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
        setSelectionStart(null);
        setSelectionEnd(null);
        setExtraSelections([]);
        setSelected(null);
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
        setSelected(null);
        return next;
      }
      // celda individual seleccionada
      if (selected) {
        const { row, col } = selected;
        if (next[row] && next[row][col] != null) next[row][col] = '';
        setSelected(null);
        return next;
      }
      // sin selección: confirmar y limpiar toda la hoja
      if (!window.confirm('No hay selección. ¿Limpiar toda la hoja?')) return prev;
      for (let rr = 0; rr < next.length; rr++) {
        for (let cc = 0; cc < next[rr].length; cc++) next[rr][cc] = '';
      }
      setSelected(null);
      setSelectionStart(null);
      setSelectionEnd(null);
      setExtraSelections([]);
      return next;
    });
  };

  // aplicar estilos condicionales antes de render
  const computeCellStyles = (r:number,c:number) => {
    let style: React.CSSProperties = {};
    let cls = '';
    // search matches
    if (searchMatches.find(m=>m.row===r && m.col===c)) {
      cls += ' search-match';
      style.outline = '2px solid rgba(255,165,0,0.7)';
    }
    // conditional formats
    for (const cf of conditionalFormats) {
      const scope = cf.scope;
      const inScope = !scope || (r>=cf.scope!.r1 && r<=cf.scope!.r2 && c>=cf.scope!.c1 && c<=cf.scope!.c2);
      if (!inScope) continue;
      const cellVal = String(data[r][c] ?? '');
      if (cf.type === 'gt') { if (!Number.isNaN(Number(cellVal)) && Number(cellVal) > Number(cf.value)) { style.background = cf.bg; } }
      if (cf.type === 'lt') { if (!Number.isNaN(Number(cellVal)) && Number(cellVal) < Number(cf.value)) { style.background = cf.bg; } }
      if (cf.type === 'eq') { if (cellVal === cf.value) style.background = cf.bg; }
      if (cf.type === 'contains') { if (cellVal.indexOf(cf.value) !== -1) style.background = cf.bg; }
    }
    return { style, cls } as any;
  };

  const onFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const dataBuf = ev.target?.result;
      const workbook = XLSX.read(dataBuf, { type: 'binary' });
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
    reader.readAsBinaryString(file);
    // limpiar
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <div className="bg-gray-50 min-h-screen p-6">
      <div className="max-w-full mx-auto">
        <div className="flex items-center justify-between mb-3 gap-3">
          <div className="flex items-center gap-2">
            <TooltipCooldown content="Agrega una nueva fila al final de la tabla" cooldown={1500}>
              <button onClick={addRow} className="bg-blue-600 text-white px-2 py-1 rounded hover:bg-blue-700 transition text-sm">+ Fila</button>
            </TooltipCooldown>
            <TooltipCooldown content="Elimina la fila seleccionada" cooldown={1500}>
              <button onClick={deleteRow} className="bg-red-600 text-white px-2 py-1 rounded hover:bg-red-700 transition text-sm flex items-center"><FaTrash className="mr-1" />Fila</button>
            </TooltipCooldown>
            <select value={selectedRow} onChange={e => setSelectedRow(Number(e.target.value))} className="px-2 py-1 rounded border text-sm">
              {data.map((_, idx) => (
                <option key={idx} value={idx}>{`Fila ${idx + 1}`}</option>
              ))}
            </select>
          </div>
          <div className="flex items-center gap-2">
            <TooltipCooldown content="Agrega una nueva columna al final de la tabla" cooldown={1500}>
              <button onClick={addCol} className="bg-green-600 text-white px-2 py-1 rounded hover:bg-green-700 transition text-sm">+ Col</button>
            </TooltipCooldown>
            <TooltipCooldown content="Elimina la columna seleccionada" cooldown={1500}>
              <button onClick={deleteCol} className="bg-red-600 text-white px-2 py-1 rounded hover:bg-red-700 transition text-sm flex items-center"><FaTrash className="mr-1" />Col</button>
            </TooltipCooldown>
            <select value={selectedCol} onChange={e => setSelectedCol(Number(e.target.value))} className="px-2 py-1 rounded border text-sm">
              {data[0].map((_, idx) => (
                <option key={idx} value={idx}>{getColName(idx)}</option>
              ))}
            </select>
          </div>
          <div className="flex items-center gap-2">
            <TooltipCooldown content="Importa datos desde un archivo Excel (.xlsx, .xls, .csv)" cooldown={1500}>
              <button onClick={() => fileInputRef.current?.click()} className="bg-yellow-500 hover:bg-yellow-600 text-black font-semibold px-3 py-1 rounded text-sm">Importar Excel</button>
            </TooltipCooldown>
            <TooltipCooldown content="Exporta la tabla actual a un archivo Excel (.xlsx)" cooldown={1500}>
              <button onClick={exportToExcel} className="bg-blue-500 hover:bg-blue-600 text-white font-semibold px-3 py-1 rounded text-sm">Exportar Excel</button>
            </TooltipCooldown>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFileInputChange} style={{ display: 'none' }} />
            <TooltipCooldown content="Limpia las celdas seleccionadas o toda la hoja si no hay selección" cooldown={1500}>
              <button
                onClick={() => clearSelectedOrAll()}
                className="bg-gray-700 hover:bg-gray-800 text-white font-semibold px-3 py-1 rounded text-sm"
              >
                Clean
              </button>
            </TooltipCooldown>
          </div>
        </div>

          {/* CONTROLES ADICIONALES: Undo/Redo, Find/Replace, Freeze, Sort, Conditional Formatting */}
          <div className="flex items-center gap-3 mb-3">
            <div className="flex items-center gap-2">
              <TooltipCooldown content="Deshace la última acción (Ctrl+Z)" cooldown={1500}>
                <button onClick={() => { undo(); }} disabled={!history.length} className="px-2 py-1 bg-gray-200 rounded">Undo</button>
              </TooltipCooldown>
              <TooltipCooldown content="Rehace la última acción deshecha (Ctrl+Y)" cooldown={1500}>
                <button onClick={() => { redo(); }} disabled={!future.length} className="px-2 py-1 bg-gray-200 rounded">Redo</button>
              </TooltipCooldown>
            </div>

            <div className="flex items-center gap-2 p-2 border rounded bg-white">
              <input placeholder="Buscar" value={findText} onChange={e=>setFindText(e.target.value)} className="px-2 py-1 border rounded text-sm" />
              <input placeholder="Reemplazar" value={replaceText} onChange={e=>setReplaceText(e.target.value)} className="px-2 py-1 border rounded text-sm" />
              <TooltipCooldown content="Busca el texto en la hoja" cooldown={1500}>
                <button onClick={() => { findMatches(); }} className="px-2 py-1 bg-blue-500 text-white rounded text-sm">Find</button>
              </TooltipCooldown>
              <TooltipCooldown content="Reemplaza la coincidencia actual por el texto indicado" cooldown={1500}>
                <button onClick={() => replaceCurrent()} className="px-2 py-1 bg-yellow-400 text-black rounded text-sm">Replace</button>
              </TooltipCooldown>
              <TooltipCooldown content="Reemplaza todas las coincidencias por el texto indicado" cooldown={1500}>
                <button onClick={() => replaceAll()} className="px-2 py-1 bg-red-400 text-white rounded text-sm">Replace All</button>
              </TooltipCooldown>
            </div>

            <div className="flex items-center gap-2">
              <label className="text-sm">Freeze R:</label>
              <input type="number" value={freezeRows} onChange={e=>setFreezeRows(Number(e.target.value)||0)} className="w-16 px-2 py-1 border rounded text-sm" />
              <label className="text-sm">C:</label>
              <input type="number" value={freezeCols} onChange={e=>setFreezeCols(Number(e.target.value)||0)} className="w-16 px-2 py-1 border rounded text-sm" />
            </div>

            <div className="flex items-center gap-2">
              <TooltipCooldown content="Ordena el rango seleccionado por la columna actual" cooldown={1500}>
                <button onClick={() => sortSelectionByColumn()} className="px-2 py-1 bg-indigo-500 text-white rounded text-sm">Sort Range</button>
              </TooltipCooldown>
            </div>

            <div className="flex items-center gap-2">
              <TooltipCooldown content="Agrega formato condicional al rango seleccionado" cooldown={1500}>
                <button onClick={() => addConditionalFormatViaPrompt()} className="px-2 py-1 bg-pink-500 text-white rounded text-sm">Add Conditional Format</button>
              </TooltipCooldown>
            </div>
          </div>

        <div className="flex gap-2 mb-4">
          <TooltipCooldown content="Combina las celdas seleccionadas en una sola" cooldown={1500}>
            <button onClick={combineCells} className="bg-purple-600 hover:bg-purple-700 text-white px-3 py-1 rounded text-sm">Combinar celdas</button>
          </TooltipCooldown>
          <TooltipCooldown content="Separa las celdas combinadas seleccionadas" cooldown={1500}>
            <button onClick={separateCells} className="bg-pink-600 hover:bg-pink-700 text-white px-3 py-1 rounded text-sm">Separar celdas</button>
          </TooltipCooldown>
        </div>

        {/* Barra de fórmulas similar a Excel */}
        <div className="mb-3 bg-white border rounded p-2 shadow-sm flex items-center gap-3">
          <div className="text-sm text-gray-600 w-16 text-center font-medium">fx</div>
          <textarea
            value={formulaText}
            onChange={e => setFormulaText(e.target.value)}
            className="flex-1 p-2 border rounded resize-none h-10 text-sm"
            placeholder="Barra de fórmulas"
          />
          <div className="flex items-center gap-2">
            <TooltipCooldown content="Aplica la fórmula escrita a la selección (V)" cooldown={1500}>
              <button onClick={applyFormulaToSelection} className="bg-green-600 hover:bg-green-700 text-white px-3 py-1 rounded font-semibold">V</button>
            </TooltipCooldown>
            <TooltipCooldown content="Limpia la fórmula y la selección (X)" cooldown={1500}>
              <button onClick={clearFormulaAndSelection} className="bg-red-600 hover:bg-red-700 text-white px-3 py-1 rounded font-semibold">X</button>
            </TooltipCooldown>
          </div>
          <div className="ml-4 text-sm text-gray-700">
            {computedSum != null ? <span className="font-semibold">Preview: {computedSum.toLocaleString()}</span> : <span className="text-gray-400">Preview: —</span>}
          </div>
          <div className="ml-4 flex items-center gap-2">
            <TooltipCooldown content="Suma los valores de la selección" cooldown={1500}>
              <button onClick={() => applyQuickFunc('SUM')} className="bg-blue-600 hover:bg-blue-700 text-white px-2 py-1 rounded text-sm">SUM</button>
            </TooltipCooldown>
            <TooltipCooldown content="Calcula el promedio de la selección" cooldown={1500}>
              <button onClick={() => applyQuickFunc('AVERAGE')} className="bg-indigo-600 hover:bg-indigo-700 text-white px-2 py-1 rounded text-sm">AVERAGE</button>
            </TooltipCooldown>
            <TooltipCooldown content="Cuenta las celdas con valor en la selección" cooldown={1500}>
              <button onClick={() => applyQuickFunc('COUNT')} className="bg-gray-600 hover:bg-gray-700 text-white px-2 py-1 rounded text-sm">COUNT</button>
            </TooltipCooldown>
          </div>
        </div>

        <div className="overflow-auto border rounded bg-white shadow-sm">
          <table ref={tableRef} className="min-w-[1200px] w-full table-fixed" style={{ borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead className="sticky top-0 bg-gray-100">
              <tr>
                <th className="bg-gray-100 text-gray-700 border-r border-gray-200 w-10">&nbsp;</th>
                {Array.from({ length: data[0].length }).map((_, colIdx) => (
                  <th key={colIdx} data-header-col={colIdx} className="bg-gray-100 text-gray-700 text-center font-medium border-b border-gray-200 relative" style={{ width: colWidths[colIdx], height: 36 }}
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
                  <th className="bg-gray-50 text-gray-600 border-r border-gray-100 w-10 text-center font-medium" style={{ height: 36 }}>{rowIdx + 1}</th>
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
              // pasar estilos extra como props en data-attributes (Cell usa internal state)
              data-extra-style={JSON.stringify(cs.style)}
              data-extra-class={cs.cls}
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
            data-extra-style={JSON.stringify(cs.style)}
            data-extra-class={cs.cls}
                      />
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default ExcelComponent;