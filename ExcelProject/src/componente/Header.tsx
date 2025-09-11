import React from 'react';

type UserObj = {
  id: number;
  username: string;
  email?: string;
  role?: string;
} | null;

type Props = {
  user: UserObj;
  onLogout: () => void;
  onOpenAdmin?: (tab?: 'create'|'list') => void;
};

const Header: React.FC<Props> = ({ user, onLogout, onOpenAdmin }) => {
  return (
    <header className="p-4 flex justify-between items-center" style={{ backgroundColor: 'var(--card)', borderBottom: '1px solid rgba(0,0,0,0.04)' }}>
      <h1 className="text-lg font-semibold" style={{ color: 'var(--accent)' }}>Excel Project</h1>
      <div className="flex items-center space-x-3">
        <div className="text-sm mr-2" style={{ color: 'var(--muted)' }}>
          <div>{user ? user.username : ''}</div>
          <div className="text-xs" style={{ color: 'var(--muted)' }}>{user ? (user.role || 'usuario') : ''}</div>
        </div>
        {user && user.role === 'admin' && (
          <>
            <button onClick={() => onOpenAdmin && onOpenAdmin('create')} className="py-1 px-3 rounded text-white mr-2" style={{ background: '#2563eb' }}>Registro</button>
            <button onClick={() => onOpenAdmin && onOpenAdmin('list')} className="py-1 px-3 rounded text-white mr-2" style={{ background: '#10b981' }}>Usuarios</button>
          </>
        )}
        <button onClick={onLogout} className="py-1 px-3 rounded text-white" style={{ background: '#ef4444' }}>Cerrar sesión</button>
      </div>
    </header>
  );
};

export default Header;
