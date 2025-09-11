import React from 'react';
import Login from '../componente/Login';

type Props = { onLoginSuccess?: () => void };

const LoginView: React.FC<Props> = ({ onLoginSuccess }) => {
  return (
    <div>
      <Login onSuccess={onLoginSuccess} />
    </div>
  );
};

export default LoginView;
