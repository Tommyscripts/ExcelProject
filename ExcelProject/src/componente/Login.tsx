import React, { useState } from 'react';

type Props = {
  onSuccess?: () => void;
};

const Login: React.FC<Props> = ({ onSuccess }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
  const url = '/auth/login'; // usar proxy de Vite en dev
  const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password }),
      });
      let data: any = null;
      try {
        data = await res.json();
      } catch (jsonErr) {
        console.error('No JSON in response', jsonErr);
      }
      if (!res.ok) {
        setError((data && data.error) || `Error en el servidor: ${res.status}`);
      } else {
        // marcar como autenticado localmente
        localStorage.setItem('user', username);
        if (onSuccess) onSuccess();
      }
    } catch (err: any) {
      console.error('Fetch error:', err);
      // mostrar más info para debugging
      setError(err.message ? `${err.message}` : 'Error de conexión (failed to fetch)');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center" style={{ backgroundColor: 'var(--bg-pastel)' }}>
      <div className="max-w-sm w-full p-6 rounded-2xl shadow-xl glass" style={{ backgroundColor: 'var(--card)' }}>
        <div className="text-center mb-6">
          <h2 className="text-2xl font-bold" style={{ color: 'var(--accent)' }}>Iniciar sesión</h2>
          <p className="text-sm mt-1" style={{ color: 'var(--muted)' }}>Accede a tu panel</p>
        </div>
        {error && (
          <div className="mb-4 text-sm text-red-700 bg-red-50 p-3 rounded">{error}</div>
        )}
        <form onSubmit={submit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium" style={{ color: 'var(--muted)' }}>Usuario</label>
            <input
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="mt-1 block w-full px-4 py-3 border border-transparent rounded-lg shadow-sm focus:outline-none"
              placeholder="Usuario"
              required
              style={{ backgroundColor: 'rgba(255,255,255,0.6)' }}
            />
          </div>
          <div>
            <label className="block text-sm font-medium" style={{ color: 'var(--muted)' }}>Contraseña</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="mt-1 block w-full px-4 py-3 border border-transparent rounded-lg shadow-sm focus:outline-none"
              placeholder="Contraseña"
              required
              style={{ backgroundColor: 'rgba(255,255,255,0.6)' }}
            />
          </div>
          <button
            type="submit"
            className="w-full py-3 px-4 rounded-lg disabled:opacity-60"
            disabled={loading}
            style={{ background: 'linear-gradient(90deg,var(--accent),var(--accent-2))', color: '#fff' }}
          >
            {loading ? 'Entrando...' : 'Entrar'}
          </button>
        </form>
        <div className="mt-6 text-center text-xs" style={{ color: 'var(--muted)' }}>
          <span>¿Olvidaste tu contraseña? Contacta con el administrador.</span>
        </div>
      </div>
    </div>
  );
};

export default Login;
