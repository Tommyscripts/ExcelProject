import * as React from 'react';
import { useState, useEffect } from 'react';
import ExcelView from './view/excelview';
import LoginView from './view/loginview';
import Header from './componente/Header';

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
      <Header user={user} onLogout={handleLogout} />
      <ExcelView />
    </div>
  );
}

export default App;
