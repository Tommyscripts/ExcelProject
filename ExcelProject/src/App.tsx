import * as React from 'react';
import { useState, useEffect } from 'react';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import ExcelView from './view/excelview';
import LoginView from './view/loginview';
import Header from './componente/Header';
import AdminPanel from './componente/AdminPanel';

type UserObj = {
  id: number;
  username: string;
  email?: string;
  role?: string;
} | null;

const App: React.FC = () => {
  const [user, setUser] = useState<UserObj>(null);

  useEffect(() => {
    const raw = localStorage.getItem('user');
    if (raw) {
      try {
        // Si es JSON válido, parsear normalmente
        const parsed = JSON.parse(raw);
        // Normalizar: si username es 'admin' asegurar role 'admin'
        if (parsed && parsed.username === 'admin' && (!parsed.role || parsed.role !== 'admin')) {
          parsed.role = 'admin';
        }
  setUser(parsed);
  // Persistir normalización
  try { localStorage.setItem('user', JSON.stringify(parsed)); } catch {}
      } catch {
        // Valores antiguos podían guardar solo el username como string.
        // Normalizar a un objeto user básico (role por defecto 'usuario').
        const normalized = { id: 0, username: raw, role: raw === 'admin' ? 'admin' : 'usuario' };
        setUser(normalized);
  try { localStorage.setItem('user', JSON.stringify(normalized)); } catch {}
      }
    }
  }, []);

  const handleLogout = () => {
    localStorage.removeItem('user');
    setUser(null);
  };

  const [adminOpen, setAdminOpen] = useState(false);
  const [adminTab, setAdminTab] = useState<'create'|'list'|undefined>(undefined);

  // Componente que protege rutas: si no hay usuario, redirige a /login
  const ProtectedRoute: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    if (!user) return <Navigate to="/login" replace />;
    return <>{children}</>;
  };

  return (
    <BrowserRouter>
      <Routes>
        <Route path="/login" element={<LoginView onLoginSuccess={(u?: UserObj) => { if (u) setUser(u); else setUser(null); }} />} />
        <Route
          path="/*"
          element={
            <ProtectedRoute>
              <div>
                <Header user={user} onLogout={handleLogout} onOpenAdmin={(tab) => { setAdminTab(tab); setAdminOpen(true); }} />
                <ExcelView />
                {adminOpen && <AdminPanel initialTab={adminTab} onClose={() => { setAdminOpen(false); setAdminTab(undefined); }} />}
              </div>
            </ProtectedRoute>
          }
        />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
