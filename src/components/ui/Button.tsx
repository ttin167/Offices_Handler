import React from 'react';
import { Loader2 } from 'lucide-react';
import type { DomainAccent } from '../../types/workspace';
import { ACCENT_CLASSES } from '../../config/navigation';

export interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'outline' | 'ghost' | 'danger';
  accent?: DomainAccent;
  size?: 'xs' | 'sm' | 'md' | 'lg';
  loading?: boolean;
  icon?: React.ReactNode;
}

const SIZE_CLASSES: Record<NonNullable<ButtonProps['size']>, string> = {
  xs: 'px-2 py-1 text-xs gap-1 rounded',
  sm: 'px-3 py-1.5 text-xs gap-1.5 rounded-md',
  md: 'px-4 py-2 text-sm gap-2 rounded-lg',
  lg: 'px-5 py-2.5 text-sm gap-2 rounded-lg',
};

export function Button({
  variant = 'primary',
  accent = 'studio',
  size = 'md',
  loading = false,
  icon,
  children,
  className = '',
  disabled,
  type = 'button',
  ...rest
}: ButtonProps) {
  let variantStyle = '';
  if (variant === 'primary') {
    variantStyle = `${ACCENT_CLASSES[accent].button} shadow-subtle border border-transparent`;
  } else if (variant === 'secondary') {
    variantStyle =
      'bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 dark:text-slate-100 dark:border-slate-700';
  } else if (variant === 'outline') {
    variantStyle =
      'bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 dark:bg-slate-900 dark:hover:bg-slate-800 dark:text-slate-200 dark:border-slate-700';
  } else if (variant === 'ghost') {
    variantStyle =
      'bg-transparent hover:bg-slate-100 text-slate-600 hover:text-slate-900 border border-transparent dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100';
  } else if (variant === 'danger') {
    variantStyle =
      'bg-red-600 hover:bg-red-700 text-white border border-transparent shadow-subtle';
  }

  return (
    <button
      type={type}
      disabled={disabled || loading}
      className={`inline-flex items-center justify-center font-medium transition-colors focus:outline-none focus:ring-2 focus:ring-indigo-500/40 disabled:cursor-not-allowed disabled:opacity-50 ${SIZE_CLASSES[size]} ${variantStyle} ${className}`}
      {...rest}
    >
      {loading ? <Loader2 className="h-4 w-4 animate-spin shrink-0" /> : icon ? <span className="shrink-0">{icon}</span> : null}
      {children ? <span>{children}</span> : null}
    </button>
  );
}
