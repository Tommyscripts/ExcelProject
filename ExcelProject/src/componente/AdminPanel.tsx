import React, { useEffect, useState } from 'react';
import { apiFetch, buildAdminHeaders } from '../lib/api';

type User = { id: number; username: string; email?: string; role?: string };

const AdminPanel: React.FC<{ onClose?: () => void, initialTab?: 'create'|'list' }> = ({ onClose, initialTab }) => {
  const [users, setUsers] = useState<User[]>([]);
  const [activeTab, setActiveTab] = useState<'create'|'list'>(initialTab || 'create');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [fieldErrors, setFieldErrors] = useState<{ username?: string; email?: string; password?: string }>({});
  const [debugMsg, setDebugMsg] = useState<string | null>(null);
  const [form, setForm] = useState({ username: '', email: '', password: '', role: 'usuario' });

  const fetchUsers = async () => {
    setLoading(true);
    setError(null);
    try {
  const headers = buildAdminHeaders();
  const res = await apiFetch('/auth/users', { headers });
      const data = await res.json();
      if (!res.ok) throw new Error(data && data.error ? data.error : 'Error fetching');
      setUsers(data.users || []);
    } catch (err: any) {
      setError(err.message || String(err));
    } finally { setLoading(false); }
  };

  useEffect(() => { fetchUsers(); }, []);
  useEffect(() => { if (initialTab) setActiveTab(initialTab); }, [initialTab]);

  const submitCreate = async (e?: any) => {
  try { if (e && typeof e.preventDefault === 'function') e.preventDefault(); } catch {}
  setDebugMsg('Invocando creación...');
  console.log('[AdminPanel] submitCreate invoked', form);
    setError(null);
    setSuccess(null);
    setLoading(true);
    try {
  const headers = buildAdminHeaders();
      console.log('[AdminPanel] headers', headers);
  console.log('[AdminPanel] body', JSON.stringify(form));
  const res = await apiFetch('/auth/users', { method: 'POST', headers, body: JSON.stringify(form) });
      const data = await res.json();
      if (!res.ok) {
        // Map backend validation messages to field-specific errors when possible
        const msg = data && data.error ? String(data.error) : 'Error creating';
        const lower = msg.toLowerCase();
        const fe: any = {};
        if (lower.includes('usuario') || lower.includes('username')) fe.username = msg;
        if (lower.includes('email')) fe.email = msg;
        if (lower.includes('password') || lower.includes('contrase')) fe.password = msg;
        // generic exists message
        if (!Object.keys(fe).length && /existe|already|duplicate/.test(lower)) {
          // assume username/email conflict
          fe.username = msg;
          fe.email = msg;
        }
        setFieldErrors(fe);
        throw new Error(msg);
      }
      setForm({ username: '', email: '', password: '', role: 'usuario' });
      setFieldErrors({});
      setSuccess('Usuario creado correctamente');
      setDebugMsg('Creación completada');
      fetchUsers();
    } catch (err: any) { setError(err.message || String(err)); }
    finally { setLoading(false); setTimeout(() => setDebugMsg(null), 3500); }
  };

  const deleteUser = async (id: number) => {
    if (!confirm('Eliminar usuario?')) return;
    try {
  const headers = buildAdminHeaders();
  const res = await apiFetch(`/auth/users/${id}`, { method: 'DELETE', headers });
      const data = await res.json();
      if (!res.ok) throw new Error(data && data.error ? data.error : 'Error deleting');
      fetchUsers();
    } catch (err: any) { setError(err.message || String(err)); }
  };

  return (
    <div className="fixed inset-0 flex items-start justify-center p-6 backdrop-blur-sm bg-black/40 dark:bg-black/50 z-[9999]">
      <div className="w-full max-w-3xl relative rounded-2xl shadow-2xl p-6 ring-1 ring-inset transition-transform bg-[color:var(--card)] text-[color:var(--text)]" style={{ maxHeight: '90vh', overflowY: 'auto', border: '1px solid rgba(255,255,255,0.06)' }}>
        {/* Close button top-right (adjusted for contrast and placement) */}
          <button
          aria-label="Cerrar panel"
          title="Cerrar"
          onClick={() => { if (onClose) onClose(); }}
          className="absolute top-4 right-8 w-11 h-11 rounded-full flex items-center justify-center shadow-lg hover:scale-105 transition-transform bg-[color:var(--card)] text-[color:var(--muted)] z-[11000]"
          style={{ border: '1px solid rgba(0,0,0,0.08)', transform: 'translate(6px, 6px)' }}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ color: 'var(--muted)' }}>
            <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2.25" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>

        <div className="flex items-center justify-between mb-4">
          <div>
            <h3 className="text-lg font-semibold text-[color:var(--accent)]">Panel de administración</h3>
            <p className="text-sm text-[color:var(--muted)]">Gestiona usuarios y roles</p>
          </div>
        </div>

        <div className="mb-6">
          <div className="flex items-center gap-3 mb-4">
            <button className={`px-4 py-2 rounded-full transition ${activeTab === 'create' ? 'text-white' : ''}`} onClick={() => setActiveTab('create')} style={{ background: activeTab === 'create' ? 'linear-gradient(90deg,var(--accent),var(--accent-2))' : 'transparent', border: activeTab === 'create' ? 'none' : '1px solid rgba(0,0,0,0.06)', color: activeTab === 'create' ? '#fff' : 'var(--muted)' }}>Crear</button>
            <button className={`px-4 py-2 rounded-full transition ${activeTab === 'list' ? 'text-white' : ''}`} onClick={() => setActiveTab('list')} style={{ background: activeTab === 'list' ? 'linear-gradient(90deg,#7dd3fc,#60a5fa)' : 'transparent', border: activeTab === 'list' ? 'none' : '1px solid rgba(0,0,0,0.06)', color: activeTab === 'list' ? '#fff' : 'var(--muted)' }}>Lista</button>
          </div>

          {error && <div className="text-sm mb-3 text-[color:#e11d48]">{error}</div>}
          {success && <div className="text-sm mb-3 text-[color:var(--accent)]">{success}</div>}
          {debugMsg && <div className="text-sm mb-3 text-[color:var(--muted)]">{debugMsg}</div>}

          {activeTab === 'create' && (
            <form onSubmit={(e) => { e.preventDefault(); return false; }} className="grid grid-cols-1 md:grid-cols-4 gap-3">
              <div className="min-w-0">
                <input required placeholder="Usuario" value={form.username} onChange={e => { setForm({ ...form, username: e.target.value }); setFieldErrors({ ...fieldErrors, username: undefined }); }} className={`w-full min-w-0 h-12 p-3 rounded-lg shadow-sm border ${fieldErrors.username ? 'border-red-400' : 'border-transparent'} focus:outline-none focus:ring-2 focus:ring-offset-1`} style={{ backgroundColor: 'var(--card)', color: 'var(--text)', boxSizing: 'border-box' }} />
                {fieldErrors.username && <div className="text-xs mt-1 text-[color:#e11d48]">{fieldErrors.username}</div>}
              </div>
              <div className="min-w-0">
                <input required placeholder="Email" value={form.email} onChange={e => { setForm({ ...form, email: e.target.value }); setFieldErrors({ ...fieldErrors, email: undefined }); }} className={`w-full min-w-0 h-12 p-3 rounded-lg shadow-sm border ${fieldErrors.email ? 'border-red-400' : 'border-transparent'} focus:outline-none focus:ring-2 focus:ring-offset-1`} style={{ backgroundColor: 'var(--card)', color: 'var(--text)', boxSizing: 'border-box' }} />
                {fieldErrors.email && <div className="text-xs mt-1 text-[color:#e11d48]">{fieldErrors.email}</div>}
              </div>
              <div className="min-w-0">
                <input required placeholder="Contraseña" type="password" value={form.password} onChange={e => { setForm({ ...form, password: e.target.value }); setFieldErrors({ ...fieldErrors, password: undefined }); }} className={`w-full min-w-0 h-12 p-3 rounded-lg shadow-sm border ${fieldErrors.password ? 'border-red-400' : 'border-transparent'} focus:outline-none focus:ring-2 focus:ring-offset-1`} style={{ backgroundColor: 'var(--card)', color: 'var(--text)', boxSizing: 'border-box' }} />
                {fieldErrors.password && <div className="text-xs mt-1 text-[color:#e11d48]">{fieldErrors.password}</div>}
              </div>
              <div className="min-w-0">
                <select value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} className="w-full min-w-0 h-12 p-3 rounded-lg shadow-sm border border-transparent focus:outline-none focus:ring-2 focus:ring-offset-1" style={{ backgroundColor: 'var(--card)', color: 'var(--text)', boxSizing: 'border-box' }}>
                  <option value="usuario">usuario</option>
                  <option value="admin">admin</option>
                </select>
              </div>

              <div className="col-span-full text-right">
                <button disabled={loading} type="button" onClick={() => submitCreate()} className={`px-5 py-2 rounded-lg font-medium shadow hover:scale-105 transition-transform ${loading ? 'opacity-60 cursor-not-allowed' : ''}`} style={{ background: 'linear-gradient(90deg,#34d399,#10b981)', color: '#fff', boxShadow: '0 6px 18px rgba(16,185,129,0.18)' }}>{loading ? 'Creando...' : 'Crear'}</button>
              </div>
            </form>
          )}
        </div>
        {activeTab === 'list' && (
          <div>
            <h4 className="font-medium mb-2 text-[color:var(--accent)]">Usuarios</h4>
            {loading ? <div className="text-[color:var(--muted)]">Cargando...</div> : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm rounded-lg" style={{ borderCollapse: 'separate', borderSpacing: '0 8px' }}>
                  <thead>
                    <tr className="text-left text-xs text-[color:var(--muted)]"><th>ID</th><th>Usuario</th><th>Email</th><th>Role</th><th>Acciones</th></tr>
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
