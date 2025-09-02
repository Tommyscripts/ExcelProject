
import React, { useState, useRef } from 'react';
import { FaLightbulb, FaRegLightbulb } from 'react-icons/fa';

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
        <TooltipCooldown content="Reemplaza el texto buscado por el nuevo valor" cooldown={1500}>
          <button onClick={replaceCurrent} className="px-4 py-2 rounded-lg bg-blue-500 text-white font-bold shadow hover:bg-blue-700 transition">Reemplazar</button>
        </TooltipCooldown>
      </div>
      <div className="flex items-center gap-2 ml-auto">
        {/* Toggle claro/oscuro */}
        <TooltipCooldown content="Cambiar tema claro/oscuro" cooldown={800}>
          <button
            aria-label="Cambiar tema"
            onClick={() => setDarkMode(!darkMode)}
            className="p-2 rounded-full bg-gray-700 hover:bg-gray-600 text-white shadow"
          >
            {darkMode ? <FaLightbulb className="w-5 h-5 text-yellow-300" /> : <FaRegLightbulb className="w-5 h-5" />}
          </button>
        </TooltipCooldown>
        <TooltipCooldown content="Importa datos desde un archivo" cooldown={1500}>
          <>
            <button onClick={() => fileInputRef.current?.click()} className="px-4 py-2 rounded-lg bg-indigo-500 text-white font-bold shadow hover:bg-indigo-700 transition">Importar</button>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              ref={fileInputRef}
              style={{ display: 'none' }}
              onChange={onFileInputChange}
            />
          </>
        </TooltipCooldown>
        <TooltipCooldown content="Exporta los datos a un archivo" cooldown={1500}>
          <button onClick={exportToExcel} className="px-4 py-2 rounded-lg bg-green-700 text-white font-bold shadow hover:bg-green-900 transition">Exportar</button>
        </TooltipCooldown>
      </div>
    </div>
    {/* Secciones principales */}
    <div className="flex flex-wrap gap-6">
  {/* Filas/Columnas */}
  <div className="bg-blue-500/90 rounded-xl p-6 w-full max-w-3xl flex flex-col gap-3 text-white shadow-lg">
        <div className="text-lg font-bold">Filas/Columnas</div>
        <div className="flex flex-wrap gap-3 items-center">
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
          <TooltipCooldown content="Combina las celdas seleccionadas en una sola" cooldown={1500}>
            <button onClick={combineCells} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-purple-500 hover:bg-purple-700 font-semibold"><span>🔗</span>Combinar celdas</button>
          </TooltipCooldown>
          <TooltipCooldown content="Separa las celdas combinadas seleccionadas" cooldown={1500}>
            <button onClick={separateCells} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-purple-400 hover:bg-purple-600 font-semibold"><span>✂️</span>Separar celdas</button>
          </TooltipCooldown>
          <TooltipCooldown content="Limpia la celda seleccionada o toda la tabla si no hay selección. Incluye celdas combinadas." cooldown={1500}>
            <button onClick={clearSelectedOrAll} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-gray-300 hover:bg-gray-400 font-semibold text-gray-900"><span>🧹</span>Limpiar</button>
          </TooltipCooldown>
        </div>
      </div>
      {/* Funciones */}
  <div className="bg-orange-500/90 rounded-xl p-6 w-full max-w-3xl flex flex-col gap-3 text-white shadow-lg">
        <div className="text-lg font-bold">Funciones</div>
        <div className="flex flex-wrap gap-3 items-center">
        {/* SUM y SUMIF */}
        <div className="relative inline-block">
          <TooltipCooldown content="Opciones de suma" cooldown={1500}>
            <button onClick={() => setShowSumMenu(!showSumMenu)} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 font-semibold">
              SUM <span className="ml-1">▼</span>
            </button>
          </TooltipCooldown>
          {showSumMenu && (
            <div className="absolute left-0 top-full mt-2 w-32 bg-blue-900 rounded shadow-lg z-50 flex flex-col">
              <button onClick={() => applyQuickFunc('SUM')} className="px-4 py-2 text-blue-300 hover:bg-gray-700 text-left">SUM</button>
              <button onClick={() => applyQuickFunc('SUMIF')} className="px-4 py-2 text-blue-300 hover:bg-gray-700 text-left">SUMIF</button>
            </div>
          )}
        </div>
        {/* COUNT y COUNTIF */}
        <div className="relative inline-block">
          <TooltipCooldown content="Opciones de conteo" cooldown={1500}>
            <button onClick={() => setShowCountMenu(!showCountMenu)} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 font-semibold">
              COUNT <span className="ml-1">▼</span>
            </button>
          </TooltipCooldown>
          {showCountMenu && (
            <div className="absolute left-0 top-full mt-2 w-32 bg-blue-900 rounded shadow-lg z-50 flex flex-col">
              <button onClick={() => applyQuickFunc('COUNT')} className="px-4 py-2 text-blue-300 hover:bg-gray-700 text-left">COUNT</button>
              <button onClick={() => applyQuickFunc('COUNTIF')} className="px-4 py-2 text-blue-300 hover:bg-gray-700 text-left">COUNTIF</button>
            </div>
          )}
        </div>
        {/* AVERAGE y ABS siguen igual */}
        <TooltipCooldown content="Calcula el promedio de los valores seleccionados" cooldown={1500}>
          <button onClick={() => applyQuickFunc('AVERAGE')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-orange-400 hover:bg-orange-600 font-semibold"><span>📊</span>AVERAGE</button>
        </TooltipCooldown>
        <TooltipCooldown content="Aplica la función ABS a la celda seleccionada" cooldown={1500}>
          <button onClick={() => setFormulaText('=ABS(A1)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-orange-400 hover:bg-orange-600 font-semibold"><span>✖️</span>ABS</button>
        </TooltipCooldown>
        {/* MAX, MIN y ROUND */}
        <TooltipCooldown content="Devuelve el valor máximo de la selección" cooldown={1500}>
          <button onClick={() => applyQuickFunc('MAX')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-yellow-400 hover:bg-yellow-500 font-semibold"><span>⬆️</span>MAX</button>
        </TooltipCooldown>
        <TooltipCooldown content="Devuelve el valor mínimo de la selección" cooldown={1500}>
          <button onClick={() => applyQuickFunc('MIN')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-yellow-600 hover:bg-yellow-700 font-semibold"><span>⬇️</span>MIN</button>
        </TooltipCooldown>
        <TooltipCooldown content="Redondea los valores seleccionados al entero más cercano" cooldown={1500}>
          <button onClick={() => setFormulaText('=ROUND(A1)')} className="flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-300 hover:bg-blue-400 font-semibold text-gray-900"><span>🔵</span>ROUND</button>
        </TooltipCooldown>
        </div>
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
