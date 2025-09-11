import * as React from 'react';
import { useState, useEffect } from 'react';
import ExcelView from './view/excelview';
import LoginView from './view/loginview';

const App: React.FC = () => {
  const [user, setUser] = useState<string | null>(null);

  useEffect(() => {
    const u = localStorage.getItem('user');
    setUser(u);
  }, []);

  const handleLogout = () => {
    localStorage.removeItem('user');
    setUser(null);
  };

  if (!user) {
    return <LoginView onLoginSuccess={() => setUser(localStorage.getItem('user'))} />;
  }

  return (
    <div>
      <div className="p-4 bg-gray-100 flex justify-between items-center">
        <h1 className="text-xl font-semibold">Excel Project</h1>
        <div>
          <span className="mr-4">{user}</span>
          <button onClick={handleLogout} className="py-1 px-3 bg-red-500 text-white rounded">Cerrar sesión</button>
        </div>
      </div>
      <ExcelView />
    </div>
  );
}

export default App;
