import React from 'react';

type Props = {
  user: string | null;
  onLogout: () => void;
};

const Header: React.FC<Props> = ({ user, onLogout }) => {
  return (
    <header className="p-4 flex justify-between items-center" style={{ backgroundColor: 'var(--card)', borderBottom: '1px solid rgba(0,0,0,0.04)' }}>
      <h1 className="text-lg font-semibold" style={{ color: 'var(--accent)' }}>Excel Project</h1>
      <div className="flex items-center space-x-3">
        <span className="text-sm mr-2" style={{ color: 'var(--muted)' }}>{user}</span>
        <button onClick={onLogout} className="py-1 px-3 rounded text-white" style={{ background: '#ef4444' }}>Cerrar sesión</button>
      </div>
    </header>
  );
};

export default Header;
