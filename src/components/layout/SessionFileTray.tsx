import { ArrowUpRight, FolderOpen, Printer, Trash2, X } from 'lucide-react';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useWorkspace } from '../../context/WorkspaceContext';
import { suggestToolsForExtension } from '../../config/navigation';

export function SessionFileTray() {
  const { sessionFiles, removeSessionFile, clearShelfOnly } = useSessionFiles();
  const { navigateToTool, openPrintModal } = useWorkspace();

  if (sessionFiles.length === 0) return null;

  return (
    <div className="border-b border-slate-200 bg-slate-50/90 px-4 py-2 dark:border-slate-800 dark:bg-slate-900/80">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2 overflow-x-auto">
          <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
            <FolderOpen className="h-3.5 w-3.5 text-indigo-500" />
            <span>Kho tệp phiên ({sessionFiles.length})</span>
          </span>

          {sessionFiles.map((sf) => {
            const topTool = suggestToolsForExtension(sf.ext)[0];
            return (
              <div
                key={sf.id}
                className="group inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-2.5 py-1 text-xs shadow-subtle dark:border-slate-700 dark:bg-slate-800"
              >
                <span className="rounded bg-slate-100 px-1 py-0.5 font-mono text-[10px] font-semibold uppercase text-slate-600 dark:bg-slate-700 dark:text-slate-300">
                  {sf.ext || 'FILE'}
                </span>
                <span className="max-w-[160px] truncate font-medium text-slate-800 dark:text-slate-100" title={sf.filename}>
                  {sf.filename}
                </span>
                {topTool ? (
                  <button
                    type="button"
                    onClick={() => navigateToTool(topTool.id, sf)}
                    className="inline-flex items-center gap-0.5 rounded px-1.5 py-0.5 text-[10px] font-medium text-indigo-600 hover:bg-indigo-50 dark:text-indigo-400 dark:hover:bg-indigo-950/60"
                    title={`Mở ngay trong ${topTool.label}`}
                  >
                    <span>Mở</span>
                    <ArrowUpRight className="h-3 w-3" />
                  </button>
                ) : null}
                {['pdf', 'docx', 'xlsx', 'xls', 'png', 'jpg'].includes(sf.ext) ? (
                  <button
                    type="button"
                    onClick={() => openPrintModal({ filePath: sf.filePath, filename: sf.filename })}
                    className="rounded p-0.5 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200"
                    title="In tệp này"
                  >
                    <Printer className="h-3 w-3" />
                  </button>
                ) : null}
                <button
                  type="button"
                  onClick={() => removeSessionFile(sf.id)}
                  className="rounded p-0.5 text-slate-400 hover:text-red-600"
                  title="Gỡ khỏi khay phiên"
                >
                  <X className="h-3 w-3" />
                </button>
              </div>
            );
          })}
        </div>

        <button
          type="button"
          onClick={clearShelfOnly}
          className="inline-flex items-center gap-1 rounded px-2 py-1 text-[11px] text-slate-400 hover:bg-slate-200/70 hover:text-slate-700 dark:hover:bg-slate-800"
          title="Ẩn danh sách tệp khỏi thanh phiên (không xóa file trên đĩa)"
        >
          <Trash2 className="h-3 w-3" />
          <span>Dọn khay</span>
        </button>
      </div>
    </div>
  );
}
