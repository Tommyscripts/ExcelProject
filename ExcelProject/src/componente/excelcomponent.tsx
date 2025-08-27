import React, { useState } from 'react';

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
  const addRow = () => {
    setData(prev => [...prev, Array(prev[0].length).fill('')]);
  };

  const addCol = () => {
    setData(prev => prev.map(row => [...row, '']));
  };

  const handleChange = (row: number, col: number, value: string) => {
    const newData = data.map((r, i) =>
      i === row ? r.map((c, j) => (j === col ? value : c)) : r
    );
    setData(newData);
  };

  const handleCellClick = (row: number, col: number) => {
    setSelected({ row, col });
  };

  return (
    <div className="overflow-auto bg-[#222] min-h-screen flex flex-col items-center justify-center p-4">
      <div className="mb-2 flex gap-2">
        <button onClick={addRow} className="bg-blue-600 text-white px-3 py-1 rounded hover:bg-blue-700">Añadir fila</button>
        <button onClick={addCol} className="bg-green-600 text-white px-3 py-1 rounded hover:bg-green-700">Añadir columna</button>
      </div>
      <table className="border-collapse shadow-xl" style={{ minWidth: 1200 }}>
        <thead>
          <tr>
            <th className="bg-black text-white border border-[#444] w-10 h-8"></th>
            {Array.from({ length: data[0].length }).map((_, colIdx) => (
              <th
                key={colIdx}
                className="bg-black text-white border border-[#444] w-24 h-8 text-center font-bold"
              >
                {getColName(colIdx)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((row, rowIdx) => (
            <tr key={rowIdx}>
              <th className="bg-black text-white border border-[#444] w-10 h-8 text-center font-bold">{rowIdx + 1}</th>
              {row.map((cell, colIdx) => {
                const isSelected = selected && selected.row === rowIdx && selected.col === colIdx;
                return (
                  <td
                    key={colIdx}
                    className={`border border-[#ccc] w-24 h-8 p-0 relative ${isSelected ? 'outline outline-2 outline-blue-500 z-10' : ''}`}
                    onClick={() => handleCellClick(rowIdx, colIdx)}
                  >
                    <input
                      type="text"
                      value={cell}
                      onChange={e => handleChange(rowIdx, colIdx, e.target.value)}
                      className={`w-full h-full px-2 bg-white focus:outline-none ${isSelected ? '' : 'cursor-pointer'}`}
                      style={{ fontSize: 15, color: '#222' }}
                      onFocus={() => handleCellClick(rowIdx, colIdx)}
                      tabIndex={0}
                    />
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default ExcelComponent;