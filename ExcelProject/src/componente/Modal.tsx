import React, { useEffect, useRef } from 'react';

interface ModalProps {
  title?: string;
  children?: React.ReactNode;
  onClose: () => void;
}

const Modal: React.FC<ModalProps> = ({ title, children, onClose }) => {
  const ref = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  useEffect(() => {
    const el = ref.current;
    const prev = document.activeElement as HTMLElement | null;
    if (el) {
      const first = el.querySelector<HTMLElement>('button, input, select, textarea');
      first?.focus();
    }
    return () => { prev?.focus?.(); };
  }, []);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40" role="dialog" aria-modal="true">
      <div ref={ref} className="bg-white p-4 rounded shadow-lg w-80 dark:bg-gray-800 dark:text-gray-100">
        {title && <h3 className="font-bold mb-2">{title}</h3>}
        <div>
          {children}
        </div>
        <div className="flex justify-end mt-3">
          <button onClick={onClose} className="px-3 py-1 border rounded">Cerrar</button>
        </div>
      </div>
    </div>
  );
};

export default Modal;
