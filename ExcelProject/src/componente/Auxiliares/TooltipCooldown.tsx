
import React, { useState, useRef } from 'react';

// TooltipCooldown: solo tooltip, ahora con Tailwind
export const TooltipCooldown: React.FC<{ content: string; cooldown?: number; children: React.ReactNode }> = ({ content, cooldown = 1500, children }) => {
  const [visible, setVisible] = useState(false);
  const timeoutRef = useRef<number | null>(null);
  const handleMouseEnter = () => {
    timeoutRef.current = window.setTimeout(() => setVisible(true), cooldown);
  };
  const handleMouseLeave = () => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    setVisible(false);
  };
  return (
    <span className="relative inline-block" onMouseEnter={handleMouseEnter} onMouseLeave={handleMouseLeave}>
      {children}
      {visible && (
        <span className="absolute bottom-[120%] left-1/2 -translate-x-1/2 bg-gray-900 text-white px-3 py-1.5 rounded text-xs whitespace-nowrap z-50 shadow-lg">
          {content}
        </span>
      )}
    </span>
  );
};

// Toolbar: toda la cabecera de botones y controles
export const Toolbar: React.FC<any> = ({
  BTN,
  data,
  selectedRow,
  setSelectedRow,
  addRow,
  deleteRow,
  selectedCol,
  setSelectedCol,
  addCol,
  deleteCol,
  fileInputRef,
  onFileInputChange,
  exportToExcel,
  clearSelectedOrAll,
  darkMode,
  setDarkMode,
  combineCells,
  separateCells,
  showToast,
  copySelectionToClipboard,
  pasteClipboardAtSelection,
  pasteTransposedAtSelection,
  undo,
  redo,
  history,
  future,
  findText,
  setFindText,
  replaceText,
  setReplaceText,
  findMatches,
  replaceCurrent,
  replaceAll,
  freezeRows,
  setFreezeRows,
  freezeCols,
  setFreezeCols,
  setShowSortModal,
  getColName,
  selectionStart,
  showCFModal,
  setShowCFModal,
  setCfTypeInput,
  setCfValueInput,
  setCfColorInput,
  cfScopeInput,
  setCfScopeInput,
  compactMode,
  setCompactMode,
  showSumMenu,
  setShowSumMenu,
  applyQuickFunc,
  showCountMenu,
  setShowCountMenu,
  setFormulaText,
  applyFormulaToSelection,
  clearFormulaAndSelection,
  computedSum,
  ...rest
}) => (
  <>
    {/* Botones de cabecera y controles principales */}
    <div className="max-w-full mx-auto relative">
      <div className="flex items-center justify-between mb-3 gap-3">
        <div className="flex items-center gap-2">
          <div className="flex gap-4 mb-2 w-full justify-start">
            <TooltipCooldown content="Combina las celdas seleccionadas en una sola" cooldown={1500}>
              <button onClick={combineCells} className={`${BTN} bg-purple-600 hover:bg-purple-700 text-white`}>Combinar celdas</button>
            </TooltipCooldown>
            <TooltipCooldown content="Separa las celdas combinadas seleccionadas" cooldown={1500}>
              <button onClick={separateCells} className={`${BTN} bg-pink-600 hover:bg-pink-700 text-white`}>Separar celdas</button>
            </TooltipCooldown>
          </div>
          <TooltipCooldown content="Agrega una nueva fila al final de la tabla" cooldown={1500}>
            <button onClick={addRow} className={`${BTN} bg-blue-600 text-white hover:bg-blue-700 transition`}>+ Fila</button>
          </TooltipCooldown>
          <TooltipCooldown content="Elimina la fila seleccionada" cooldown={1500}>
            <button onClick={deleteRow} className={`${BTN} bg-red-600 text-white hover:bg-red-700 transition`}>Fila</button>
          </TooltipCooldown>
          <select value={selectedRow} onChange={e => setSelectedRow(Number(e.target.value))} className="px-2 h-9 rounded border text-sm bg-white dark:bg-gray-700 dark:text-gray-100">
            {data.map((_, idx) => (
              <option key={idx} value={idx}>{`Fila ${idx + 1}`}</option>
            ))}
          </select>
        </div>
        <div className="flex items-center gap-2">
          <TooltipCooldown content="Agrega una nueva columna al final de la tabla" cooldown={1500}>
            <button onClick={addCol} className={`${BTN} bg-green-600 text-white hover:bg-green-700 transition`}>+ Col</button>
          </TooltipCooldown>
          <TooltipCooldown content="Elimina la columna seleccionada" cooldown={1500}>
            <button onClick={deleteCol} className={`${BTN} bg-red-600 text-white hover:bg-red-700 transition`}>Col</button>
          </TooltipCooldown>
          <select value={selectedCol} onChange={e => setSelectedCol(Number(e.target.value))} className="px-2 h-9 rounded border text-sm bg-white dark:bg-gray-700 dark:text-gray-100">
            {data[0].map((_, idx) => (
              <option key={idx} value={idx}>{getColName(idx)}</option>
            ))}
          </select>
        </div>
        <div className="flex items-center gap-2">
          <TooltipCooldown content="Importa datos desde un archivo Excel (.xlsx, .xls,.csv)" cooldown={1500}>
            <button onClick={() => fileInputRef.current?.click()} className={`${BTN} bg-yellow-500 hover:bg-yellow-600 text-black font-bold shadow-md border border-yellow-700`}>Importar Excel</button>
          </TooltipCooldown>
          <TooltipCooldown content="Exporta la tabla actual a un archivo Excel (.xlsx)" cooldown={1500}>
            <button onClick={exportToExcel} className={`${BTN} bg-blue-500 hover:bg-blue-700 text-white font-bold shadow-md border border-blue-800`}>Exportar Excel</button>
          </TooltipCooldown>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFileInputChange} className="hidden" />
          <TooltipCooldown content="Limpia las celdas seleccionadas o toda la hoja si no hay selección" cooldown={1500}>
            <button onClick={() => clearSelectedOrAll()} className={`${BTN} bg-gray-800 hover:bg-gray-900 text-white font-bold shadow-md border border-gray-900`}>Clean</button>
          </TooltipCooldown>
        </div>
        <div className="absolute right-12 top-16 z-10">
          <button
            className={"w-11 h-11 rounded-full flex items-center justify-center shadow-md border-4 transition-colors " + (darkMode ? 'bg-gray-800 border-amber-400 text-amber-300' : 'bg-amber-300 border-amber-400 text-gray-800')}
            title={darkMode ? "Modo claro" : "Modo oscuro"}
            onClick={() => setDarkMode((d: boolean) => !d)}
          >
            <span className="text-xl">💡</span>
          </button>
        </div>
      </div>
    </div>
    {/* ...puedes seguir añadiendo los demás controles aquí, siguiendo el mismo patrón... */}
  </>
);
