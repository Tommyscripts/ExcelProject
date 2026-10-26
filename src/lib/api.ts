// Pequeño helper para unificar la base de la API y funcionar tanto en dev con proxy
// como en preview/build (donde no hay proxy de Vite).

// Si existe VITE_API_BASE, se usa. En su defecto, si el puerto es uno típico de Vite
// (5173/5174 para dev, 4173/4174 para preview) apuntamos al backend local en 3000.
const guessApiBase = (): string => {
  try {
    const envBase = (import.meta as any)?.env?.VITE_API_BASE;
    if (typeof envBase === 'string') return envBase;
  } catch {}
  try {
    const port = (window?.location?.port || '').trim();
    if (['5173', '5174', '4173', '4174'].includes(port)) return 'http://localhost:3000';
  } catch {}
  // Por defecto, relativo (misma origin) — útil si sirves frontend y backend juntos.
  return '';
};

export const API_BASE = guessApiBase();

export const apiUrl = (path: string) => `${API_BASE}${path}`;

export const apiFetch: typeof fetch = (input: any, init?: RequestInit) => {
  const url = typeof input === 'string' ? apiUrl(input) : input;
  return fetch(url, init);
};

// Construye cabeceras para endpoints admin usando localStorage 'user'.
// Tolera formatos antiguos (cadena 'admin') convirtiéndolos a objeto.
export const buildAdminHeaders = (base?: HeadersInit): HeadersInit => {
  const headers: any = { 'Content-Type': 'application/json', ...(base || {}) };
  try {
    const raw = localStorage.getItem('user');
    if (raw) {
      let obj: any = null;
      try { obj = JSON.parse(raw); } catch { obj = { username: raw, role: raw === 'admin' ? 'admin' : 'usuario' }; }
      headers['x-user'] = JSON.stringify(obj);
    }
  } catch {}
  return headers;
};
