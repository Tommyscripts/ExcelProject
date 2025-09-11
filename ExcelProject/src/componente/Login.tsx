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
    <div className="min-h-screen flex items-center justify-center bg-gray-50">
      <div className="max-w-md w-full bg-white p-8 rounded-lg shadow">
        <h2 className="text-2xl font-semibold mb-6 text-center">Iniciar sesión</h2>
        {error && (
          <div className="mb-4 text-sm text-red-600 bg-red-50 p-2 rounded">{error}</div>
        )}
        <form onSubmit={submit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700">Usuario</label>
            <input
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="mt-1 block w-full px-3 py-2 border rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              placeholder="Usuario"
              required
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700">Contraseña</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="mt-1 block w-full px-3 py-2 border rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              placeholder="Contraseña"
              required
            />
          </div>
          <button
            type="submit"
            className="w-full py-2 px-4 bg-indigo-600 text-white rounded hover:bg-indigo-700 disabled:opacity-50"
            disabled={loading}
          >
            {loading ? 'Entrando...' : 'Entrar'}
          </button>
        </form>
  <p className="mt-4 text-xs text-gray-500 text-center">&nbsp;</p>
      </div>
    </div>
  );
};

export default Login;
