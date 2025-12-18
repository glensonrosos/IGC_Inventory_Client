import React, { createContext, useCallback, useContext, useState } from 'react';
import { Alert, Snackbar } from '@mui/material';

type Toast = { open: boolean; message: string; severity: 'success'|'error'|'info'|'warning' };

const ToastCtx = createContext<{ success: (m: string)=>void; error: (m: string)=>void; info: (m: string)=>void; warning: (m: string)=>void } | null>(null);

export const useToast = () => {
  const ctx = useContext(ToastCtx);
  if (!ctx) throw new Error('useToast must be used inside ToastProvider');
  return ctx;
};

export default function ToastProvider({ children }: { children: React.ReactNode }) {
  const [toast, setToast] = useState<Toast>({ open: false, message: '', severity: 'info' });

  const show = useCallback((severity: Toast['severity'], message: string) => {
    setToast({ open: true, message, severity });
  }, []);

  const api = {
    success: (m: string) => show('success', m),
    error: (m: string) => show('error', m),
    info: (m: string) => show('info', m),
    warning: (m: string) => show('warning', m)
  };

  return (
    <ToastCtx.Provider value={api}>
      {children}
      <Snackbar open={toast.open} autoHideDuration={3000} onClose={()=>setToast((t)=>({ ...t, open: false }))} anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }}>
        <Alert onClose={()=>setToast((t)=>({ ...t, open: false }))} severity={toast.severity} variant="filled" sx={{ width: '100%' }}>
          {toast.message}
        </Alert>
      </Snackbar>
    </ToastCtx.Provider>
  );
}
