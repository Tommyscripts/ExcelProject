import { useCallback, useEffect, useRef, useState } from 'react';

export type Range = { r1:number; c1:number; r2:number; c2:number };
export type Merge = { r:number; c:number; rows:number; cols:number };
export type CFRule = { id:string; type: 'gt'|'lt'|'eq'|'contains'; value:string; bg?:string; color?:string; scope?: Range | null };

const STORAGE_KEY = 'meloexcel_v1';

export function createInitialData(rows:number, cols:number) {
  return Array.from({ length: rows }, () => Array(cols).fill('')) as string[][];
}

export default function useSpreadsheet(initialRows = 46, initialCols = 22) {
  const INITIAL_ROWS = initialRows;
  const INITIAL_COLS = initialCols;
  const [data, setData] = useState<string[][]>(() => createInitialData(INITIAL_ROWS, INITIAL_COLS));
  const [merges, setMerges] = useState<Merge[]>([]);
  const [colWidths, setColWidths] = useState<number[]>(Array(INITIAL_COLS).fill(96));
  const [darkMode, setDarkMode] = useState<boolean>(false);

  // selection state left to consumer

  // history
  const HISTORY_LIMIT = 50;
  const [history, setHistory] = useState<any[]>([]);
  const [future, setFuture] = useState<any[]>([]);

  const pushHistory = useCallback((snapshot?: any) => {
    const snap = snapshot ?? { data: JSON.parse(JSON.stringify(data)), merges: JSON.parse(JSON.stringify(merges)), colWidths: [...colWidths] };
    setHistory(h => {
      const next = [...h, snap];
      if (next.length > HISTORY_LIMIT) next.shift();
      return next;
    });
    setFuture([]);
  }, [data, merges, colWidths]);

  const undo = useCallback(() => {
    setHistory(h => {
      if (h.length === 0) return h;
      const last = h[h.length - 1];
      setFuture(f => [{ data: JSON.parse(JSON.stringify(data)), merges: JSON.parse(JSON.stringify(merges)), colWidths: [...colWidths] }, ...f]);
      if (last.data) setData(last.data);
      if (last.merges) setMerges(last.merges);
      if (last.colWidths) setColWidths(last.colWidths);
      return h.slice(0, -1);
    });
  }, [data, merges, colWidths]);

  const redo = useCallback(() => {
    setFuture(f => {
      if (f.length === 0) return f;
      const [first, ...rest] = f;
      setHistory(h => [...h, { data: JSON.parse(JSON.stringify(data)), merges: JSON.parse(JSON.stringify(merges)), colWidths: [...colWidths] }]);
      if (first.data) setData(first.data);
      if (first.merges) setMerges(first.merges);
      if (first.colWidths) setColWidths(first.colWidths);
      return rest;
    });
  }, [data, merges, colWidths]);

  // persistence
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (parsed.data) setData(parsed.data);
      if (parsed.merges) setMerges(parsed.merges);
      if (parsed.colWidths) setColWidths(parsed.colWidths);
      if (typeof parsed.darkMode === 'boolean') setDarkMode(parsed.darkMode);
    } catch (e) {
      // ignore
    }
  }, []);

  useEffect(() => {
    const payload = { data, merges, colWidths, darkMode };
    try { window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); } catch (e) {}
  }, [data, merges, colWidths, darkMode]);

  // basic mutators
  const setCell = useCallback((row:number, col:number, value:string) => {
    setData(prev => {
      if (!prev[row]) return prev;
      const next = prev.map(r => [...r]);
      next[row][col] = value;
      return next;
    });
  }, []);

  const addRow = useCallback(() => {
    setData(prev => [...prev, Array(prev[0].length).fill('')]);
  }, []);
  const addCol = useCallback(() => {
    setData(prev => prev.map(row => [...row, '']));
    setColWidths(prev => [...prev, 96]);
  }, []);

  const deleteRow = useCallback((idx:number) => {
    setData(prev => prev.filter((_, i) => i !== idx));
    setMerges(prev => prev.map(m => m.r > idx ? { ...m, r: m.r - 1 } : m).filter(m => m.rows > 0));
  }, []);

  const deleteCol = useCallback((idx:number) => {
    setData(prev => prev.map(row => row.filter((_, i) => i !== idx)));
    setColWidths(prev => prev.filter((_, i) => i !== idx));
    setMerges(prev => prev.map(m => m.c > idx ? { ...m, c: m.c - 1 } : m).filter(m => m.cols > 0));
  }, []);

  // paste helper that expands grid if needed
  const pasteAt = useCallback((startRow:number, startCol:number, toPaste:string[][]) => {
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
      for (let r = 0; r < toPaste.length; r++) for (let c = 0; c < toPaste[r].length; c++) next[startRow + r][startCol + c] = toPaste[r][c];
      return next;
    });
    setColWidths(prev => {
      const newLen = Math.max(prev.length, (toPaste[0]?.length ?? 0) + startCol);
      if (newLen <= prev.length) return prev;
      return [...prev, ...Array(newLen - prev.length).fill(96)];
    });
  }, [pushHistory]);

  // import/export helpers
  const importSheet = useCallback((normalized:string[][], importedMerges:Merge[], neededCols:number) => {
    setData(normalized.length ? normalized : createInitialData(INITIAL_ROWS, INITIAL_COLS));
    setMerges(importedMerges);
    setColWidths(prev => Array.from({ length: neededCols }, (_, i) => prev[i] ?? 96));
  }, []);

  const exportToExcel = useCallback(() => {
    // consumer can call XLSX externally; keep helper placeholder
  }, []);

  return {
    data, setData, merges, setMerges, colWidths, setColWidths, darkMode, setDarkMode,
    history, future, pushHistory, undo, redo,
    setCell, addRow, addCol, deleteRow, deleteCol, pasteAt, importSheet, exportToExcel
  } as const;
}
