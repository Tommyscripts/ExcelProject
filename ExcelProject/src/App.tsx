import * as React from 'react';
import { useState, useEffect } from 'react';
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
      } catch {
        // Valores antiguos podían guardar solo el username como string.
        // Normalizar a un objeto user básico (role por defecto 'usuario').
        const normalized = { id: 0, username: raw, role: raw === 'admin' ? 'admin' : 'usuario' };
        setUser(normalized);
      }
    }
  }, []);

  const handleLogout = () => {
    localStorage.removeItem('user');
    setUser(null);
  };

  const [adminOpen, setAdminOpen] = useState(false);
  const [adminTab, setAdminTab] = useState<'create'|'list'|undefined>(undefined);

  if (!user) {
    return <LoginView onLoginSuccess={(u?: UserObj) => { if (u) setUser(u); else setUser(null); }} />;
  }

  return (
    <div>
  <Header user={user} onLogout={handleLogout} onOpenAdmin={(tab) => { setAdminTab(tab); setAdminOpen(true); }} />
      <ExcelView />
  {adminOpen && <AdminPanel initialTab={adminTab} onClose={() => { setAdminOpen(false); setAdminTab(undefined); }} />}
    </div>
  );
}

export default App;
