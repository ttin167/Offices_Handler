import React from 'react';
import { ArrowRight, CheckCircle2, Download, Printer } from 'lucide-react';
import { Button } from './Button';
import { Badge } from './Badge';
import { useWorkspace } from '../../context/WorkspaceContext';
import { suggestToolsForExtension } from '../../config/navigation';
import type { SessionFile } from '../../types/workspace';

export function ResultPreviewCard({
  title,
  summaryItems,
  outputFile,
  extraActions,
  children,
}: {
  title: string;
  summaryItems?: Array<{ label: string; value: string | number }>;
  outputFile?: SessionFile | null;
  extraActions?: React.ReactNode;
  children?: React.ReactNode;
}) {
  const { navigateToTool, openPrintModal } = useWorkspace();

  const chainedTools = outputFile
    ? suggestToolsForExtension(outputFile.ext).slice(0, 3)
    : [];

  const canPrint = outputFile
    ? ['pdf', 'docx', 'xlsx', 'xls', 'png', 'jpg'].includes(outputFile.ext.toLowerCase())
    : false;

  return (
    <div className="space-y-3 rounded-xl border border-emerald-200 bg-white p-4 shadow-card dark:border-emerald-900/60 dark:bg-slate-900">
      <div className="flex flex-wrap items-start justify-between gap-2">
        <div className="flex items-center gap-2">
          <CheckCircle2 className="h-5 w-5 shrink-0 text-emerald-600 dark:text-emerald-400" />
          <div>
            <h4 className="text-sm font-semibold text-slate-900 dark:text-slate-100">{title}</h4>
            {outputFile ? (
              <p className="text-xs text-slate-500 dark:text-slate-400">{outputFile.filename}</p>
            ) : null}
          </div>
        </div>
        {outputFile ? <Badge tone="green">.{outputFile.ext.toUpperCase()}</Badge> : null}
      </div>

      {summaryItems && summaryItems.length > 0 ? (
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-3">
          {summaryItems.map((item) => (
            <div
              key={item.label}
              className="rounded-lg border border-slate-100 bg-slate-50 px-2.5 py-2 dark:border-slate-800 dark:bg-slate-950/60"
            >
              <div className="text-[11px] text-slate-500 dark:text-slate-400">{item.label}</div>
              <div className="mt-0.5 text-xs font-semibold text-slate-900 dark:text-slate-100">{item.value}</div>
            </div>
          ))}
        </div>
      ) : null}

      {children}

      <div className="flex flex-wrap items-center gap-2 pt-1">
        {outputFile?.downloadUrl ? (
          <a
            href={outputFile.downloadUrl}
            download={outputFile.filename}
            className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-600 px-3.5 py-2 text-xs font-medium text-white shadow-subtle hover:bg-emerald-700"
          >
            <Download className="h-3.5 w-3.5" />
            <span>Tải xuống</span>
          </a>
        ) : null}

        {outputFile && canPrint ? (
          <Button
            size="sm"
            variant="outline"
            icon={<Printer className="h-3.5 w-3.5" />}
            onClick={() => openPrintModal({ filePath: outputFile.filePath, filename: outputFile.filename })}
          >
            In ngay
          </Button>
        ) : null}

        {extraActions}
      </div>

      {outputFile && chainedTools.length > 0 ? (
        <div className="flex flex-wrap items-center gap-1.5 border-t border-slate-100 pt-2.5 dark:border-slate-800">
          <span className="text-[11px] font-medium text-slate-500">Chuyển tiếp tệp này sang:</span>
          {chainedTools.map((tool) => (
            <button
              key={tool.id}
              type="button"
              onClick={() => navigateToTool(tool.id, outputFile)}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2 py-1 text-[11px] font-medium text-slate-700 hover:border-indigo-300 hover:bg-indigo-50 hover:text-indigo-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-300"
            >
              <span>{tool.label}</span>
              <ArrowRight className="h-3 w-3" />
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}
