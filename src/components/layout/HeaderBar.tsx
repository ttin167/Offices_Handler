import { useState } from 'react';
import {
  Activity,
  HelpCircle,
  History,
  Menu,
  Moon,
  Printer,
  Search,
  ShieldCheck,
  Sparkles,
  Sun,
  Trash2,
} from 'lucide-react';
import { useTheme } from '../../context/ThemeContext';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { Modal } from '../ui/Modal';
import { Button } from '../ui/Button';

export function HeaderBar() {
  const { theme, toggleTheme } = useTheme();
  const {
    navigateToTool,
    setSidebarCollapsed,
    setCommandPaletteOpen,
    setHistoryOpen,
    openPrintModal,
    setDiagnosticsOpen,
    setGuideOpen,
    copilotOpen,
    setCopilotOpen,
  } = useWorkspace();
  const { purgeAllTraces, isPurging } = useSessionFiles();
  const toast = useToast();
  const [confirmPurgeOpen, setConfirmPurgeOpen] = useState(false);
  const [alsoClearHistory, setAlsoClearHistory] = useState(false);

  const handleConfirmPurge = async () => {
    try {
      const res = await purgeAllTraces(alsoClearHistory);
      toast.success(
        'Đã xóa sạch dấu vết phiên làm việc',
        `Đã gỡ ${res.removedFiles} tệp tạm (${res.mbFreed} MB).`
      );
      setConfirmPurgeOpen(false);
      setAlsoClearHistory(false);
    } catch (err) {
      toast.error('Lỗi dọn tệp tạm', err instanceof Error ? err.message : String(err));
    }
  };

  return (
    <>
      <header className="sticky top-0 z-30 flex h-14 items-center justify-between gap-3 border-b border-slate-200 bg-white/95 px-4 backdrop-blur dark:border-slate-800 dark:bg-slate-900/95">
        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => setSidebarCollapsed((c) => !c)}
            className="rounded-lg p-1.5 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
            title="Thu gọn / Mở rộng thanh công cụ"
          >
            <Menu className="h-4 w-4" />
          </button>

          <button
            type="button"
            onClick={() => navigateToTool('home')}
            className="flex items-center gap-2.5 text-left"
          >
            <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-slate-900 text-xs font-bold text-white dark:bg-indigo-600">
              AO
            </div>
            <div className="hidden sm:block">
              <div className="flex items-center gap-2">
                <span className="text-sm font-semibold tracking-tight text-slate-900 dark:text-slate-100">
                  Antigravity Office
                </span>
                <span className="text-[11px] font-medium text-slate-400">v5.0.0</span>
              </div>
              <div
                className="flex items-center gap-1 text-[10px] font-medium text-emerald-600 dark:text-emerald-400"
                title="Máy chủ API chỉ lắng nghe nội bộ 127.0.0.1:8000, chặn CORS từ trang web ngoài và khóa AI Copilot ở chế độ Offline/Local."
              >
                <ShieldCheck className="h-3 w-3" />
                <span>Thuần Local • 127.0.0.1</span>
              </div>
            </div>
          </button>
        </div>

        <button
          type="button"
          onClick={() => setCommandPaletteOpen(true)}
          className="hidden md:flex w-72 lg:w-96 items-center justify-between rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs text-slate-500 hover:border-slate-300 hover:bg-white dark:border-slate-700 dark:bg-slate-950 dark:text-slate-400 dark:hover:border-slate-600"
        >
          <span className="inline-flex items-center gap-2">
            <Search className="h-3.5 w-3.5 text-slate-400" />
            <span>Tìm nhanh công cụ Word, Excel, PDF...</span>
          </span>
          <kbd className="rounded border border-slate-200 bg-white px-1.5 py-0.5 font-mono text-[10px] text-slate-500 dark:border-slate-700 dark:bg-slate-800">
            Ctrl+K
          </kbd>
        </button>

        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setConfirmPurgeOpen(true)}
            className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-red-50/80 px-2.5 py-1.5 text-xs font-medium text-red-700 hover:bg-red-100 dark:border-red-900/60 dark:bg-red-950/40 dark:text-red-300 dark:hover:bg-red-950/70"
            title="Xóa ngay toàn bộ tệp tạm đã upload và xuất ra trong .web_uploads / .web_outputs"
          >
            <Trash2 className="h-3.5 w-3.5" />
            <span className="hidden xl:inline">Xóa dấu vết phiên</span>
          </button>

          <button
            type="button"
            onClick={() => openPrintModal(null)}
            className="rounded-lg p-2 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
            title="Trạm In ấn Doanh nghiệp"
          >
            <Printer className="h-4 w-4" />
          </button>

          <button
            type="button"
            onClick={() => setHistoryOpen(true)}
            className="rounded-lg p-2 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
            title="Lịch sử xử lý tệp"
          >
            <History className="h-4 w-4" />
          </button>

          <button
            type="button"
            onClick={() => setDiagnosticsOpen(true)}
            className="rounded-lg p-2 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
            title="Kiểm tra hệ thống & Bảo trì"
          >
            <Activity className="h-4 w-4" />
          </button>

          <button
            type="button"
            onClick={toggleTheme}
            className="rounded-lg p-2 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
            title={theme === 'light' ? 'Chuyển sang Giao diện Tối' : 'Chuyển sang Giao diện Sáng'}
          >
            {theme === 'light' ? <Moon className="h-4 w-4" /> : <Sun className="h-4 w-4" />}
          </button>

          <button
            type="button"
            onClick={() => setGuideOpen(true)}
            className="rounded-lg p-2 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
            title="Hướng dẫn nhanh"
          >
            <HelpCircle className="h-4 w-4" />
          </button>

          <button
            type="button"
            onClick={() => setCopilotOpen((prev) => !prev)}
            className={`ml-1 inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1.5 text-xs font-medium transition-colors ${
              copilotOpen
                ? 'border-indigo-500 bg-indigo-600 text-white'
                : 'border-indigo-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 dark:border-indigo-800 dark:bg-indigo-950/60 dark:text-indigo-300'
            }`}
            title="Mở Trợ lý AI Copilot (Chế độ Local / Offline RAG)"
          >
            <Sparkles className="h-3.5 w-3.5" />
            <span className="hidden lg:inline">Copilot</span>
          </button>
        </div>
      </header>

      <Modal
        open={confirmPurgeOpen}
        onClose={() => setConfirmPurgeOpen(false)}
        title="Xóa sạch dấu vết phiên làm việc"
        subtitle="Bảo mật dữ liệu cục bộ ngay sau khi hoàn tất công việc"
        size="md"
        footer={
          <>
            <Button variant="ghost" onClick={() => setConfirmPurgeOpen(false)}>
              Hủy
            </Button>
            <Button variant="danger" loading={isPurging} onClick={handleConfirmPurge}>
              Xóa ngay tất cả tệp tạm
            </Button>
          </>
        }
      >
        <div className="space-y-3 text-sm text-slate-600 dark:text-slate-300">
          <p>
            Thao tác này sẽ xóa ngay lập tức toàn bộ tệp đã tải lên và tệp kết quả trong{' '}
            <code className="rounded bg-slate-100 px-1 py-0.5 font-mono text-xs dark:bg-slate-800">
              .web_uploads/
            </code>{' '}
            và{' '}
            <code className="rounded bg-slate-100 px-1 py-0.5 font-mono text-xs dark:bg-slate-800">
              .web_outputs/
            </code>
            , đồng thời làm trống Kho Tệp Phiên trên giao diện.
          </p>
          <label className="flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs dark:border-slate-800 dark:bg-slate-950">
            <input
              type="checkbox"
              checked={alsoClearHistory}
              onChange={(e) => setAlsoClearHistory(e.target.checked)}
              className="rounded border-slate-300"
            />
            <span>Xóa luôn nhật ký lịch sử thao tác trong cơ sở dữ liệu SQLite cục bộ</span>
          </label>
        </div>
      </Modal>
    </>
  );
}
