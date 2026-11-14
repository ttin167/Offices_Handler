import React from 'react';
import { FileQuestion } from 'lucide-react';

export function EmptyState({
  icon,
  title,
  description,
  action,
}: {
  icon?: React.ReactNode;
  title: string;
  description?: string;
  action?: React.ReactNode;
}) {
  return (
    <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 bg-slate-50/60 px-6 py-10 text-center dark:border-slate-800 dark:bg-slate-900/40">
      <div className="mb-3 flex h-11 w-11 items-center justify-center rounded-full bg-white text-slate-400 shadow-subtle border border-slate-200 dark:bg-slate-800 dark:border-slate-700">
        {icon || <FileQuestion className="h-5 w-5" />}
      </div>
      <h4 className="text-sm font-medium text-slate-900 dark:text-slate-100">{title}</h4>
      {description ? (
        <p className="mt-1 max-w-sm text-xs text-slate-500 dark:text-slate-400">{description}</p>
      ) : null}
      {action ? <div className="mt-4">{action}</div> : null}
    </div>
  );
}
