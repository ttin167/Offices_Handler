import React from 'react';
import { Badge } from '../ui/Badge';
import type { DomainAccent } from '../../types/workspace';
import { ACCENT_CLASSES } from '../../config/navigation';
import { InlinePreviewPane, type PreviewTarget } from '../preview/InlinePreviewPane';

const ACCENT_TO_BADGE_TONE: Record<DomainAccent, 'slate' | 'blue' | 'green' | 'red' | 'orange' | 'indigo'> = {
  neutral: 'slate',
  word: 'blue',
  excel: 'green',
  pdf: 'red',
  slide: 'orange',
  studio: 'indigo',
};

export function SplitWorkspace({
  title,
  subtitle,
  badgeText,
  accent = 'studio',
  leftPanel,
  resultPanel,
  previewTarget,
  rightHeaderExtra,
}: {
  title: string;
  subtitle: string;
  badgeText?: string;
  accent?: DomainAccent;
  leftPanel: React.ReactNode;
  resultPanel?: React.ReactNode;
  previewTarget?: PreviewTarget | null;
  rightHeaderExtra?: React.ReactNode;
}) {
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-start justify-between gap-3 border-b border-slate-200 pb-3.5 dark:border-slate-800">
        <div>
          <div className="flex items-center gap-2">
            <h1 className="text-lg font-semibold tracking-tight text-slate-900 dark:text-slate-50">{title}</h1>
            {badgeText ? <Badge tone={ACCENT_TO_BADGE_TONE[accent]}>{badgeText}</Badge> : null}
          </div>
          <p className="mt-0.5 text-xs text-slate-500 dark:text-slate-400">{subtitle}</p>
        </div>
        {rightHeaderExtra ? <div className="flex items-center gap-2">{rightHeaderExtra}</div> : null}
      </div>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-12">
        <div className="space-y-4 lg:col-span-5">
          <div className={`rounded-xl border border-slate-200 bg-white p-4 shadow-card dark:border-slate-800 dark:bg-slate-900 border-t-2 ${ACCENT_CLASSES[accent].border}`}>
            {leftPanel}
          </div>
        </div>

        <div className="space-y-4 lg:col-span-7">
          {resultPanel}
          {previewTarget !== undefined ? <InlinePreviewPane target={previewTarget} /> : null}
        </div>
      </div>
    </div>
  );
}
