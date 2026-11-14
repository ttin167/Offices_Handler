import React from 'react';
import {
  BookOpen,
  Briefcase,
  Calculator,
  ChevronRight,
  FileCheck2,
  FileDiff,
  FileSpreadsheet,
  FileText,
  GitBranch,
  Home,
  Layers,
  Mail,
  Presentation,
  RefreshCw,
  ScanText,
  ShieldAlert,
  Sparkles,
  Stamp,
  Wrench,
} from 'lucide-react';
import { DOMAIN_GROUPS, TOOL_CATALOG } from '../../config/navigation';
import { useWorkspace } from '../../context/WorkspaceContext';
import type { ToolId } from '../../types/workspace';

const TOOL_ICONS: Record<ToolId, React.ReactNode> = {
  home: <Home className="h-4 w-4" />,
  'word-polish': <Sparkles className="h-4 w-4" />,
  'word-merge-split': <Layers className="h-4 w-4" />,
  'word-mail-merge': <Mail className="h-4 w-4" />,
  'word-quick-patch': <FileText className="h-4 w-4" />,
  'word-diff': <FileDiff className="h-4 w-4" />,
  'pdf-pro': <Stamp className="h-4 w-4" />,
  'excel-power': <FileSpreadsheet className="h-4 w-4" />,
  'excel-reporting': <Briefcase className="h-4 w-4" />,
  'excel-tracker': <FileCheck2 className="h-4 w-4" />,
  'excel-formulas': <BookOpen className="h-4 w-4" />,
  'excel-template': <Calculator className="h-4 w-4" />,
  'convert-universal': <RefreshCw className="h-4 w-4" />,
  'ocr-vietnamese': <ScanText className="h-4 w-4" />,
  'qa-security': <ShieldAlert className="h-4 w-4" />,
  'presentation-studio': <Presentation className="h-4 w-4" />,
  'diagram-studio': <GitBranch className="h-4 w-4" />,
};

const ACTIVE_DOMAIN_COLOR: Record<string, string> = {
  neutral: 'bg-slate-900 text-white dark:bg-slate-100 dark:text-slate-900',
  word: 'bg-blue-50 text-blue-700 border-blue-200 dark:bg-blue-950/60 dark:text-blue-300 dark:border-blue-800',
  pdf: 'bg-red-50 text-red-700 border-red-200 dark:bg-red-950/60 dark:text-red-300 dark:border-red-800',
  excel: 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/60 dark:text-emerald-300 dark:border-emerald-800',
  convert: 'bg-indigo-50 text-indigo-700 border-indigo-200 dark:bg-indigo-950/60 dark:text-indigo-300 dark:border-indigo-800',
  visual: 'bg-orange-50 text-orange-700 border-orange-200 dark:bg-orange-950/60 dark:text-orange-300 dark:border-orange-800',
};

export function LeftSidebar() {
  const { activeTool, navigateToTool, sidebarCollapsed, setDiagnosticsOpen } = useWorkspace();

  return (
    <aside
      className={`shrink-0 border-r border-slate-200 bg-white transition-all duration-200 dark:border-slate-800 dark:bg-slate-900 ${
        sidebarCollapsed ? 'w-16' : 'w-64'
      } flex flex-col justify-between overflow-y-auto`}
    >
      <nav className="space-y-4 p-2.5">
        {DOMAIN_GROUPS.map((group) => {
          const items = TOOL_CATALOG.filter((t) => t.domain === group.id);
          return (
            <div key={group.id} className="space-y-1">
              {!sidebarCollapsed ? (
                <div className="px-2.5 pt-1 text-[10px] font-semibold uppercase tracking-wider text-slate-400 dark:text-slate-500">
                  {group.label}
                </div>
              ) : (
                <div className="mx-2 border-t border-slate-100 dark:border-slate-800" />
              )}

              {items.map((item) => {
                const active = activeTool === item.id;
                const activeClass =
                  group.id === 'home' && active
                    ? ACTIVE_DOMAIN_COLOR.neutral
                    : active
                    ? `${ACTIVE_DOMAIN_COLOR[group.id]} border`
                    : 'border border-transparent text-slate-600 hover:bg-slate-100/80 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800/70 dark:hover:text-slate-100';

                return (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => navigateToTool(item.id)}
                    title={sidebarCollapsed ? item.label : undefined}
                    className={`flex w-full items-center gap-2.5 rounded-lg px-2.5 py-2 text-left text-xs font-medium transition-colors ${activeClass} ${
                      sidebarCollapsed ? 'justify-center px-0' : ''
                    }`}
                  >
                    <span className="shrink-0">{TOOL_ICONS[item.id]}</span>
                    {!sidebarCollapsed ? (
                      <>
                        <span className="flex-1 truncate">{item.label}</span>
                        {active ? <ChevronRight className="h-3.5 w-3.5 shrink-0 opacity-60" /> : null}
                      </>
                    ) : null}
                  </button>
                );
              })}
            </div>
          );
        })}
      </nav>

      <div className="border-t border-slate-100 p-2.5 dark:border-slate-800">
        <button
          type="button"
          onClick={() => setDiagnosticsOpen(true)}
          title="Chẩn đoán môi trường & Bảo trì"
          className={`flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-xs font-medium text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 ${
            sidebarCollapsed ? 'justify-center px-0' : ''
          }`}
        >
          <Wrench className="h-4 w-4 shrink-0" />
          {!sidebarCollapsed ? <span>Bảo trì & Chẩn đoán</span> : null}
        </button>
      </div>
    </aside>
  );
}
