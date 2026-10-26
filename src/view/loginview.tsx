import React from 'react';
import { useNavigate } from 'react-router-dom';
import Login from '../componente/Login';

type UserShape = { id: number; username: string; email?: string; role?: string };
type Props = { onLoginSuccess?: (user?: UserShape) => void };

const LoginView: React.FC<Props> = ({ onLoginSuccess }) => {
  const navigate = useNavigate();

  const handleSuccess = (user?: UserShape) => {
    if (onLoginSuccess) onLoginSuccess(user);
    // Después de login exitoso, navegar a la raíz para mostrar el contenido protegido
    navigate('/', { replace: true });
  };

  return (
    <div>
      <Login onSuccess={handleSuccess} />
    </div>
  );
};

export default LoginView;
