
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
  <div className="space-y-6">
    {/* Cabecera Editar */}
    <div className="flex flex-wrap gap-3 bg-gray-800/80 p-4 rounded-xl">
      <button onClick={copySelectionToClipboard} className="flex items-center gap-2 px-6 py-3 rounded-lg bg-green-500 text-white font-semibold shadow hover:bg-green-600 transition"><span>📋</span>Copiar</button>
      <button onClick={pasteClipboardAtSelection} className="flex items-center gap-2 px-6 py-3 rounded-lg bg-green-400 text-white font-semibold shadow hover:bg-green-500 transition"><span>📥</span>Pegar</button>
      <button onClick={undo} className="flex items-center gap-2 px-6 py-3 rounded-lg bg-yellow-400 text-gray-900 font-semibold shadow hover:bg-yellow-500 transition"><span>↩️</span>Deshacer</button>
      <div className="flex items-center gap-2 px-6 py-3 rounded-lg bg-gray-700 text-white font-semibold shadow">
        <span>🔍</span>
        <input value={findText} onChange={e => setFindText(e.target.value)} placeholder="Buscar" className="bg-transparent outline-none text-white placeholder:text-gray-300" />
      </div>
      <div className="flex items-center gap-2 px-6 py-3 rounded-lg bg-gray-700 text-white font-semibold shadow">
        <span>🖊️</span>
        <input value={replaceText} onChange={e => setReplaceText(e.target.value)} placeholder="Reemplazar" className="bg-transparent outline-none text-white placeholder:text-gray-300" />
      </div>
    </div>
    {/* Secciones principales */}
    <div className="flex flex-wrap gap-6">
      {/* Filas/Columnas */}
      <div className="bg-blue-500/90 rounded-xl p-6 min-w-[220px] flex flex-col gap-3 text-white shadow-lg">
        <div className="text-lg font-bold mb-2">Filas/Columnas</div>
          <TooltipCooldown content="Añade una nueva fila debajo" cooldown={1500}>
            <button onClick={addRow} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-400 hover:bg-blue-600 font-semibold"><span>➕</span>Insertar fila</button>
          </TooltipCooldown>
          <TooltipCooldown content="Elimina la fila seleccionada" cooldown={1500}>
            <button onClick={deleteRow} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-red-500 hover:bg-red-700 font-semibold"><span>🗑️</span>Eliminar fila</button>
          </TooltipCooldown>
          <TooltipCooldown content="Añade una nueva columna a la derecha" cooldown={1500}>
            <button onClick={addCol} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-400 hover:bg-blue-600 font-semibold"><span>➕</span>Insertar columna</button>
          </TooltipCooldown>
          <TooltipCooldown content="Elimina la columna seleccionada" cooldown={1500}>
            <button onClick={deleteCol} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-red-500 hover:bg-red-700 font-semibold"><span>🗑️</span>Eliminar columna</button>
          </TooltipCooldown>
      </div>
      {/* Funciones */}
      <div className="bg-orange-500/90 rounded-xl p-6 min-w-[220px] flex flex-col gap-3 text-white shadow-lg">
        <div className="text-lg font-bold mb-2">Funciones</div>
          <TooltipCooldown content="Suma los valores seleccionados" cooldown={1500}>
            <button onClick={() => applyQuickFunc('SUM')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-orange-400 hover:bg-orange-600 font-semibold"><span>∑</span>SUM</button>
          </TooltipCooldown>
          <TooltipCooldown content="Cuenta el número de celdas seleccionadas" cooldown={1500}>
            <button onClick={() => applyQuickFunc('COUNT')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-orange-400 hover:bg-orange-600 font-semibold"><span>234</span>COUNT</button>
          </TooltipCooldown>
          <TooltipCooldown content="Calcula el promedio de los valores seleccionados" cooldown={1500}>
            <button onClick={() => applyQuickFunc('AVERAGE')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-orange-400 hover:bg-orange-600 font-semibold"><span>📊</span>AVERAGE</button>
          </TooltipCooldown>
          <TooltipCooldown content="Aplica la función ABS a la celda seleccionada" cooldown={1500}>
            <button onClick={() => setFormulaText('=ABS(A1)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-orange-400 hover:bg-orange-600 font-semibold"><span>✖️</span>ABS</button>
          </TooltipCooldown>
      </div>
      {/* Formato */}
      <div className="bg-green-500/90 rounded-xl p-6 min-w-[220px] flex flex-col gap-3 text-white shadow-lg">
        <div className="text-lg font-bold mb-2">Formato</div>
          <TooltipCooldown content="Aplica la función lógica AND" cooldown={1500}>
            <button onClick={() => setFormulaText('=AND(true,false)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-green-400 hover:bg-green-600 font-semibold"><span>&&</span>AND</button>
          </TooltipCooldown>
          <TooltipCooldown content="Aplica la función lógica OR" cooldown={1500}>
            <button onClick={() => setFormulaText('=OR(true,false)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-green-400 hover:bg-green-600 font-semibold"><span>||</span>OR</button>
          </TooltipCooldown>
          <TooltipCooldown content="Aplica la función lógica NOT" cooldown={1500}>
            <button onClick={() => setFormulaText('=NOT(true)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-green-400 hover:bg-green-600 font-semibold"><span>!</span>NOT</button>
          </TooltipCooldown>
          <TooltipCooldown content="Elimina espacios en la celda seleccionada" cooldown={1500}>
            <button onClick={() => setFormulaText('=TRIM(A1)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-green-400 hover:bg-green-600 font-semibold"><span>---</span>TRIM</button>
          </TooltipCooldown>
      </div>
    </div>
    {/* Barra de fórmulas */}
    <div className="bg-gray-800/80 rounded-xl p-4 flex items-center gap-3 mt-4">
      <span className="text-white text-lg font-bold px-2">fx</span>
      <input value={rest.formulaText} onChange={e => rest.setFormulaText(e.target.value)} placeholder="Barra de fórmulas" className="flex-1 px-3 py-2 rounded bg-gray-900 text-white placeholder:text-gray-400" />
      <button onClick={applyFormulaToSelection} className="px-4 py-2 rounded-lg bg-green-500 text-white font-bold shadow hover:bg-green-600 transition">✔️</button>
      <button onClick={clearFormulaAndSelection} className="px-4 py-2 rounded-lg bg-red-500 text-white font-bold shadow hover:bg-red-600 transition">❌</button>
    {/* Botones SUM eliminados */}
    </div>
  </div>
);
