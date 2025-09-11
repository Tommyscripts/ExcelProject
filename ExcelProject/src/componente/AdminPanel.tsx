import React, { useEffect, useState } from 'react';

type User = { id: number; username: string; email?: string; role?: string };

const AdminPanel: React.FC<{ onClose?: () => void, initialTab?: 'create'|'list' }> = ({ onClose, initialTab }) => {
  const [users, setUsers] = useState<User[]>([]);
  const [activeTab, setActiveTab] = useState<'create'|'list'>(initialTab || 'create');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [form, setForm] = useState({ username: '', email: '', password: '', role: 'usuario' });

  const fetchUsers = async () => {
    setLoading(true);
    setError(null);
    try {
      const rawUser = localStorage.getItem('user');
      const headers: any = { 'Content-Type': 'application/json' };
      if (rawUser) headers['x-user'] = rawUser;
      const res = await fetch('/auth/users', { headers });
      const data = await res.json();
      if (!res.ok) throw new Error(data && data.error ? data.error : 'Error fetching');
      setUsers(data.users || []);
    } catch (err: any) {
      setError(err.message || String(err));
    } finally { setLoading(false); }
  };

  useEffect(() => { fetchUsers(); }, []);
  useEffect(() => { if (initialTab) setActiveTab(initialTab); }, [initialTab]);

  const submitCreate = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    try {
      const rawUser = localStorage.getItem('user');
      const headers: any = { 'Content-Type': 'application/json' };
      if (rawUser) headers['x-user'] = rawUser;
      const res = await fetch('/auth/users', { method: 'POST', headers, body: JSON.stringify(form) });
      const data = await res.json();
      if (!res.ok) throw new Error(data && data.error ? data.error : 'Error creating');
      setForm({ username: '', email: '', password: '', role: 'usuario' });
      fetchUsers();
    } catch (err: any) { setError(err.message || String(err)); }
  };

  const deleteUser = async (id: number) => {
    if (!confirm('Eliminar usuario?')) return;
    try {
      const rawUser = localStorage.getItem('user');
      const headers: any = { 'Content-Type': 'application/json' };
      if (rawUser) headers['x-user'] = rawUser;
      const res = await fetch(`/auth/users/${id}`, { method: 'DELETE', headers });
      const data = await res.json();
      if (!res.ok) throw new Error(data && data.error ? data.error : 'Error deleting');
      fetchUsers();
    } catch (err: any) { setError(err.message || String(err)); }
  };

  return (
    <div className="fixed inset-0 flex items-start justify-center p-6 z-50 backdrop-blur-sm bg-black/40 dark:bg-white/10">
      <div className="w-full max-w-3xl relative rounded-2xl shadow-2xl p-6 ring-1 ring-inset transition-transform" style={{ maxHeight: '90vh', overflowY: 'auto', backgroundColor: 'var(--card)', border: '1px solid rgba(255,255,255,0.06)' }}>
        {/* Close button top-right (adjusted for contrast and placement) */}
        <button
          aria-label="Cerrar panel"
          title="Cerrar"
          onClick={() => { if (onClose) onClose(); }}
          className="absolute top-4 right-8 w-11 h-11 rounded-full flex items-center justify-center shadow-lg hover:scale-105 transition-transform"
          style={{ backgroundColor: 'var(--card)', border: '1px solid rgba(0,0,0,0.08)', color: 'var(--muted)', zIndex: 60, transform: 'translate(6px, 6px)' }}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ color: 'var(--muted)' }}>
            <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2.25" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>

        <div className="flex items-center justify-between mb-4">
          <div>
            <h3 className="text-lg font-semibold" style={{ color: 'var(--accent)' }}>Panel de administración</h3>
            <p className="text-sm" style={{ color: 'var(--muted)' }}>Gestiona usuarios y roles</p>
          </div>
        </div>

        <div className="mb-6">
          <div className="flex items-center gap-3 mb-4">
            <button className={`px-4 py-2 rounded-full transition ${activeTab === 'create' ? 'text-white' : ''}`} onClick={() => setActiveTab('create')} style={{ background: activeTab === 'create' ? 'linear-gradient(90deg,var(--accent),var(--accent-2))' : 'transparent', border: activeTab === 'create' ? 'none' : '1px solid rgba(0,0,0,0.06)', color: activeTab === 'create' ? '#fff' : 'var(--muted)' }}>Crear</button>
            <button className={`px-4 py-2 rounded-full transition ${activeTab === 'list' ? 'text-white' : ''}`} onClick={() => setActiveTab('list')} style={{ background: activeTab === 'list' ? 'linear-gradient(90deg,#7dd3fc,#60a5fa)' : 'transparent', border: activeTab === 'list' ? 'none' : '1px solid rgba(0,0,0,0.06)', color: activeTab === 'list' ? '#fff' : 'var(--muted)' }}>Lista</button>
          </div>

          {error && <div className="text-sm mb-3" style={{ color: '#e11d48' }}>{error}</div>}

          {activeTab === 'create' && (
            <form onSubmit={submitCreate} className="grid grid-cols-1 md:grid-cols-4 gap-3">
              <input required placeholder="Usuario" value={form.username} onChange={e => setForm({ ...form, username: e.target.value })} className="p-3 rounded-lg shadow-sm border border-transparent focus:outline-none focus:ring-2 focus:ring-offset-1" style={{ backgroundColor: 'rgba(255,255,255,0.9)', color: 'var(--text)' }} />
              <input required placeholder="Email" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} className="p-3 rounded-lg shadow-sm border border-transparent focus:outline-none focus:ring-2 focus:ring-offset-1" style={{ backgroundColor: 'rgba(255,255,255,0.9)', color: 'var(--text)' }} />
              <input required placeholder="Contraseña" type="password" value={form.password} onChange={e => setForm({ ...form, password: e.target.value })} className="p-3 rounded-lg shadow-sm border border-transparent focus:outline-none focus:ring-2 focus:ring-offset-1" style={{ backgroundColor: 'rgba(255,255,255,0.9)', color: 'var(--text)' }} />
              <select value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} className="p-3 rounded-lg shadow-sm border border-transparent focus:outline-none focus:ring-2 focus:ring-offset-1" style={{ backgroundColor: 'rgba(255,255,255,0.9)', color: 'var(--text)' }}>
                <option value="usuario">usuario</option>
                <option value="admin">admin</option>
              </select>

              <div className="col-span-full text-right">
                <button type="submit" className="px-5 py-2 rounded-lg font-medium shadow hover:scale-105 transition-transform" style={{ background: 'linear-gradient(90deg,#34d399,#10b981)', color: '#fff', boxShadow: '0 6px 18px rgba(16,185,129,0.18)' }}>Crear</button>
              </div>
            </form>
          )}
        </div>
        {activeTab === 'list' && (
          <div>
            <h4 className="font-medium mb-2" style={{ color: 'var(--accent)' }}>Usuarios</h4>
            {loading ? <div style={{ color: 'var(--muted)' }}>Cargando...</div> : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm rounded-lg" style={{ borderCollapse: 'separate', borderSpacing: '0 8px' }}>
                  <thead>
                    <tr className="text-left text-xs" style={{ color: 'var(--muted)' }}><th>ID</th><th>Usuario</th><th>Email</th><th>Role</th><th>Acciones</th></tr>
                  </thead>
                  <tbody>
                    {users.map(u => (
                      <tr key={u.id} className="bg-[rgba(0,0,0,0.03)] rounded-md" style={{ marginBottom: '8px' }}>
                        <td className="py-2 pr-4">{u.id}</td>
                        <td>{u.username}</td>
                        <td>{u.email}</td>
                        <td>{u.role}</td>
                        <td>
                          <button className="px-3 py-1 rounded-md text-white" onClick={() => deleteUser(u.id)} style={{ background: 'linear-gradient(90deg,#f97316,#fb923c)' }}>Eliminar</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default AdminPanel;
