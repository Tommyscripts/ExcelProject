import React, { useEffect, useRef, useState, useCallback } from 'react';

export interface CellProps {
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
  extraStyle?: React.CSSProperties | null;
  extraClass?: string | null;
}

const Cell: React.FC<CellProps> = React.memo(function CellInner(props: CellProps) {
  const { value, row, col, rowSpan, colSpan, width, height, isSelected, isMultiSelected, isDragging, onCommit, onMouseDown, onMouseEnter, onMouseUp, onClick, extraStyle, extraClass } = props;
  const [local, setLocal] = useState<string>(value ?? '');
  const isEditing = useRef(false);
  const commitTimer = useRef<number | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [showTooltip, setShowTooltip] = useState(false);
  const HOVER_DELAY_MS = 1500;
  const hoverTimer = useRef<number | null>(null);

  useEffect(() => {
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
    if (commitTimer.current) window.clearTimeout(commitTimer.current);
    commitTimer.current = window.setTimeout(() => {
      commitTimer.current = null;
      onCommit(row, col, e.target.value);
    }, 300) as unknown as number;
  };

  const handleFocus = () => { isEditing.current = true; };
  const handleBlur = () => { isEditing.current = false; doCommit(); };
  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
  };

  useEffect(() => {
    return () => {
      if (hoverTimer.current) {
        window.clearTimeout(hoverTimer.current);
        hoverTimer.current = null;
      }
      if (commitTimer.current) {
        window.clearTimeout(commitTimer.current);
        commitTimer.current = null;
      }
    };
  }, []);

  return (
    <td rowSpan={rowSpan} colSpan={colSpan} data-cell-row={row} data-cell-col={col}
      className={`p-0 relative ${isSelected || isMultiSelected || isDragging ? 'ring-2 ring-blue-400 bg-blue-50 dark:bg-blue-900 dark:bg-opacity-50 dark:text-white' : 'border border-gray-200 dark:border-gray-700'} ${extraClass || ''}`}
      style={{ width, height, ...(extraStyle || {}) }}
      onMouseDown={() => onMouseDown(row, col)}
      onMouseEnter={() => {
        onMouseEnter(row, col);
        if (isSelected) {
          if (hoverTimer.current) { window.clearTimeout(hoverTimer.current); hoverTimer.current = null; }
          setShowTooltip(true);
          return;
        }
        if (hoverTimer.current) window.clearTimeout(hoverTimer.current);
        setShowTooltip(false);
        hoverTimer.current = window.setTimeout(() => { hoverTimer.current = null; setShowTooltip(true); }, HOVER_DELAY_MS) as unknown as number;
      }}
      onMouseLeave={() => {
        if (hoverTimer.current) { window.clearTimeout(hoverTimer.current); hoverTimer.current = null; }
        setShowTooltip(false);
      }}
      onMouseUp={() => {
        if (hoverTimer.current) { window.clearTimeout(hoverTimer.current); hoverTimer.current = null; }
        setShowTooltip(false);
        onMouseUp();
      }}
      onClick={() => {
        if (hoverTimer.current) { window.clearTimeout(hoverTimer.current); hoverTimer.current = null; }
        setShowTooltip(false);
        onClick(row, col);
      }}
    >
      <input
        ref={inputRef}
        type="text"
        value={local}
        onChange={onChange}
        onFocus={handleFocus}
        onBlur={handleBlur}
        onKeyDown={handleKeyDown}
        className="w-full h-full px-2 bg-transparent focus:outline-none text-sm cursor-text text-gray-900 dark:text-gray-100"
        tabIndex={0}
      />
      {(showTooltip && !isEditing.current && String(local).trim() !== '') && (
        <div className="absolute left-0 top-0 z-50 p-2 bg-white border rounded shadow-md text-sm max-w-[400px] break-words dark:bg-gray-800 dark:border-gray-700 dark:text-gray-100">
          <div className="text-xs text-gray-500 mb-1">Texto completo</div>
          {local}
        </div>
      )}
    </td>
  );
});

export default Cell;
