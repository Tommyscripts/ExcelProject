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
    <div className="fixed inset-0 z-50 flex items-center justify-center" role="dialog" aria-modal="true" style={{ background: 'rgba(0,0,0,0.4)' }}>
      <div ref={ref} className="p-4 rounded shadow-lg w-80" style={{ background: 'var(--card)', color: 'var(--text)' }}>
        {title && <h3 className="font-bold mb-2">{title}</h3>}
        <div>
          {children}
        </div>
        <div className="flex justify-end mt-3">
          <button onClick={onClose} className={`inline-flex items-center h-9 px-3 rounded text-sm border`} style={{ background: 'transparent', color: 'var(--muted)', borderColor: 'rgba(0,0,0,0.06)' }}>Cerrar</button>
        </div>
      </div>
    </div>
  );
};

export default Modal;
