import React, { createContext, useCallback, useContext, useState } from 'react';
import { AlertCircle, CheckCircle2, Info, AlertTriangle, X } from 'lucide-react';
import type { ToastItem, ToastVariant } from '../types/workspace';

interface ToastContextValue {
  pushToast: (title: string, description?: string, variant?: ToastVariant) => void;
  success: (title: string, description?: string) => void;
  error: (title: string, description?: string) => void;
  info: (title: string, description?: string) => void;
  warning: (title: string, description?: string) => void;
}

const ToastContext = createContext<ToastContextValue | null>(null);

const VARIANT_STYLE: Record<ToastVariant, string> = {
  info: 'border-blue-200 bg-white text-slate-800 dark:border-blue-800 dark:bg-slate-900 dark:text-slate-100',
  success: 'border-emerald-200 bg-white text-slate-800 dark:border-emerald-800 dark:bg-slate-900 dark:text-slate-100',
  warning: 'border-amber-200 bg-white text-slate-800 dark:border-amber-800 dark:bg-slate-900 dark:text-slate-100',
  error: 'border-red-200 bg-white text-slate-800 dark:border-red-800 dark:bg-slate-900 dark:text-slate-100',
};

const VARIANT_ICON: Record<ToastVariant, React.ReactNode> = {
  info: <Info className="h-4 w-4 text-blue-600 dark:text-blue-400 shrink-0 mt-0.5" />,
  success: <CheckCircle2 className="h-4 w-4 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5" />,
  warning: <AlertTriangle className="h-4 w-4 text-amber-600 dark:text-amber-400 shrink-0 mt-0.5" />,
  error: <AlertCircle className="h-4 w-4 text-red-600 dark:text-red-400 shrink-0 mt-0.5" />,
};

export function ToastProvider({ children }: { children: React.ReactNode }) {
  const [toasts, setToasts] = useState<ToastItem[]>([]);

  const removeToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((item) => item.id !== id));
  }, []);

  const pushToast = useCallback(
    (title: string, description?: string, variant: ToastVariant = 'info') => {
      const id = `toast_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      setToasts((prev) => [...prev.slice(-4), { id, title, description, variant }]);
      window.setTimeout(() => {
        removeToast(id);
      }, variant === 'error' ? 6000 : 4000);
    },
    [removeToast]
  );

  const success = useCallback((t: string, d?: string) => pushToast(t, d, 'success'), [pushToast]);
  const error = useCallback((t: string, d?: string) => pushToast(t, d, 'error'), [pushToast]);
  const info = useCallback((t: string, d?: string) => pushToast(t, d, 'info'), [pushToast]);
  const warning = useCallback((t: string, d?: string) => pushToast(t, d, 'warning'), [pushToast]);

  return (
    <ToastContext.Provider value={{ pushToast, success, error, info, warning }}>
      {children}
      <div
        aria-live="polite"
        className="pointer-events-none fixed bottom-4 right-4 z-[90] flex w-full max-w-sm flex-col gap-2 px-4 sm:px-0"
      >
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`pointer-events-auto flex items-start gap-2.5 rounded-lg border p-3.5 shadow-elevated transition-all ${VARIANT_STYLE[t.variant]}`}
          >
            {VARIANT_ICON[t.variant]}
            <div className="min-w-0 flex-1">
              <p className="text-sm font-medium leading-5">{t.title}</p>
              {t.description ? (
                <p className="mt-0.5 text-xs text-slate-500 dark:text-slate-400 break-words">{t.description}</p>
              ) : null}
            </div>
            <button
              type="button"
              onClick={() => removeToast(t.id)}
              className="rounded p-0.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700 dark:hover:bg-slate-800 dark:hover:text-slate-200"
              aria-label="Đóng thông báo"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}

export function useToast(): ToastContextValue {
  const ctx = useContext(ToastContext);
  if (!ctx) throw new Error('useToast must be used inside ToastProvider');
  return ctx;
}
