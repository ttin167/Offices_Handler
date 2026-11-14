import { useEffect, useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  FilePlus2,
  Loader2,
  RefreshCw,
  RotateCw,
  Trash2,
  Undo2,
} from 'lucide-react';
import { Button, type SelectedFileInfo } from '../../components/ui';
import { apiPost } from '../../services/api';

export interface PagePlanNode {
  id: string;
  pageNumber?: number;
  rotation: number;
  deleted: boolean;
  isBlank: boolean;
  thumbnailUrl?: string;
}

interface PreviewResponse {
  success: boolean;
  total_pages: number;
  pages: Array<{ page?: number; page_num?: number; thumbnail_url?: string; preview_url?: string }>;
}

export function PdfPageOrganizerGrid({
  pdfFile,
  loadingAction,
  onApplyReorganize,
}: {
  pdfFile: SelectedFileInfo | null;
  loadingAction: boolean;
  onApplyReorganize: (pagesPlan: Array<Record<string, unknown>>) => void;
}) {
  const [pages, setPages] = useState<PagePlanNode[]>([]);
  const [fetchingThumbs, setFetchingThumbs] = useState(false);
  const [fetchError, setFetchError] = useState<string | null>(null);

  const loadThumbnails = async (filePath: string) => {
    setFetchingThumbs(true);
    setFetchError(null);
    try {
      const res = await apiPost<PreviewResponse>('/api/pdf/pages/preview', { file_path: filePath });
      const mapped: PagePlanNode[] = (res.pages || []).map((p, i) => ({
        id: `p-${p.page || p.page_num || i + 1}-${i}`,
        pageNumber: p.page || p.page_num || i + 1,
        rotation: 0,
        deleted: false,
        isBlank: false,
        thumbnailUrl: p.thumbnail_url || p.preview_url || '',
      }));
      setPages(mapped);
    } catch (err) {
      setFetchError(err instanceof Error ? err.message : String(err));
      setPages([]);
    } finally {
      setFetchingThumbs(false);
    }
  };

  useEffect(() => {
    if (pdfFile?.filePath) {
      loadThumbnails(pdfFile.filePath);
    } else {
      setPages([]);
      setFetchError(null);
    }
  }, [pdfFile?.filePath]);

  const movePage = (idx: number, dir: -1 | 1) => {
    setPages((prev) => {
      const target = idx + dir;
      if (target < 0 || target >= prev.length) return prev;
      const next = [...prev];
      const [item] = next.splice(idx, 1);
      next.splice(target, 0, item);
      return next;
    });
  };

  const rotatePage = (idx: number) => {
    setPages((prev) =>
      prev.map((p, i) => (i === idx && !p.isBlank ? { ...p, rotation: (p.rotation + 90) % 360 } : p)),
    );
  };

  const toggleDelete = (idx: number) => {
    setPages((prev) =>
      prev.flatMap((p, i) => {
        if (i !== idx) return [p];
        if (p.isBlank) return [];
        return [{ ...p, deleted: !p.deleted }];
      }),
    );
  };

  const insertBlankAfter = (idx: number) => {
    const blankNode: PagePlanNode = {
      id: `blank-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      rotation: 0,
      deleted: false,
      isBlank: true,
    };
    setPages((prev) => {
      const next = [...prev];
      next.splice(idx + 1, 0, blankNode);
      return next;
    });
  };

  const handleApply = () => {
    const plan = pages
      .filter((p) => !p.deleted)
      .map((p) => (p.isBlank ? { is_blank: true } : { page: p.pageNumber, rotation: p.rotation }));
    if (plan.length === 0) return;
    onApplyReorganize(plan);
  };

  if (!pdfFile) {
    return (
      <div className="rounded-lg border border-dashed border-slate-200 px-4 py-6 text-center text-xs text-slate-500 dark:border-slate-800 dark:text-slate-400">
        Tải lên tệp PDF ở trên để xem lưới trang 120 DPI và kéo/xoay/xóa/chèn trang trắng.
      </div>
    );
  }

  if (fetchingThumbs) {
    return (
      <div className="flex items-center justify-center gap-2 rounded-lg border border-slate-200 py-8 text-xs text-slate-600 dark:border-slate-800 dark:text-slate-300">
        <Loader2 className="h-4 w-4 animate-spin text-red-600" />
        <span>Đang kết xuất ảnh thu nhỏ 120 DPI...</span>
      </div>
    );
  }

  if (fetchError) {
    return (
      <div className="space-y-2 rounded-lg border border-red-200 bg-red-50 p-3 text-xs text-red-700 dark:border-red-900/60 dark:bg-red-950/30 dark:text-red-300">
        <p>Không tải được trang PDF: {fetchError}</p>
        <Button size="sm" variant="outline" onClick={() => loadThumbnails(pdfFile.filePath)}>
          Thử lại
        </Button>
      </div>
    );
  }

  const keptCount = pages.filter((p) => !p.deleted).length;

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <span className="text-xs font-medium text-slate-600 dark:text-slate-400">
          Tổng: <strong className="text-slate-900 dark:text-slate-100">{keptCount}</strong> trang giữ lại / {pages.length} mục
        </span>
        <div className="flex items-center gap-1.5">
          <Button
            size="sm"
            variant="ghost"
            icon={<FilePlus2 className="h-3.5 w-3.5" />}
            onClick={() => insertBlankAfter(pages.length - 1)}
          >
            + Trang trắng
          </Button>
          <Button
            size="sm"
            variant="ghost"
            icon={<RefreshCw className="h-3.5 w-3.5" />}
            onClick={() => loadThumbnails(pdfFile.filePath)}
          >
            Khôi phục gốc
          </Button>
        </div>
      </div>

      <div className="grid max-h-[380px] grid-cols-2 gap-2.5 overflow-y-auto pr-1 sm:grid-cols-3">
        {pages.map((p, idx) => (
          <div
            key={p.id}
            className={`group flex flex-col justify-between rounded-lg border p-2 transition ${
              p.deleted
                ? 'border-red-200 bg-red-50/40 opacity-50 dark:border-red-900/50 dark:bg-red-950/20'
                : p.isBlank
                  ? 'border-dashed border-sky-300 bg-sky-50/50 dark:border-sky-800 dark:bg-sky-950/20'
                  : 'border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-950'
            }`}
          >
            <div className="relative flex aspect-[3/4] items-center justify-center overflow-hidden rounded bg-slate-100 dark:bg-slate-900">
              {p.isBlank ? (
                <div className="px-2 text-center text-[11px] font-medium text-sky-700 dark:text-sky-300">
                  Trang trắng A4 chèn mới
                </div>
              ) : p.thumbnailUrl ? (
                <img
                  src={p.thumbnailUrl}
                  alt={`Trang ${p.pageNumber}`}
                  style={{ transform: `rotate(${p.rotation}deg)` }}
                  className="max-h-full max-w-full object-contain transition-transform duration-200"
                />
              ) : (
                <span className="text-xs text-slate-400">Trang {p.pageNumber}</span>
              )}
              <span className="absolute left-1 top-1 rounded bg-slate-900/75 px-1.5 py-0.5 text-[10px] font-semibold text-white">
                {p.isBlank ? 'Trống' : `Tr.${p.pageNumber}`}
                {p.rotation > 0 ? ` (${p.rotation}°)` : ''}
              </span>
            </div>

            <div className="mt-2 flex items-center justify-between gap-0.5">
              <button
                type="button"
                title="Đưa lên trước"
                disabled={idx === 0}
                onClick={() => movePage(idx, -1)}
                className="rounded p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-30 dark:hover:bg-slate-800"
              >
                <ArrowLeft className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                title="Đưa ra sau"
                disabled={idx === pages.length - 1}
                onClick={() => movePage(idx, 1)}
                className="rounded p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-30 dark:hover:bg-slate-800"
              >
                <ArrowRight className="h-3.5 w-3.5" />
              </button>
              {!p.isBlank ? (
                <button
                  type="button"
                  title="Xoay +90°"
                  onClick={() => rotatePage(idx)}
                  className="rounded p-1 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                >
                  <RotateCw className="h-3.5 w-3.5" />
                </button>
              ) : null}
              <button
                type="button"
                title="Chèn trang trắng sau trang này"
                onClick={() => insertBlankAfter(idx)}
                className="rounded p-1 text-sky-600 hover:bg-sky-50 dark:text-sky-400 dark:hover:bg-sky-950/40"
              >
                <FilePlus2 className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                title={p.deleted ? 'Khôi phục trang' : 'Xóa trang'}
                onClick={() => toggleDelete(idx)}
                className={`rounded p-1 ${
                  p.deleted
                    ? 'text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-950/40'
                    : 'text-red-500 hover:bg-red-50 dark:hover:bg-red-950/40'
                }`}
              >
                {p.deleted ? <Undo2 className="h-3.5 w-3.5" /> : <Trash2 className="h-3.5 w-3.5" />}
              </button>
            </div>
          </div>
        ))}
      </div>

      <Button
        variant="primary"
        loading={loadingAction}
        disabled={keptCount === 0}
        onClick={handleApply}
        className="w-full bg-red-600 hover:bg-red-700"
      >
        Áp dụng sắp xếp ({keptCount} trang) & Tạo PDF
      </Button>
    </div>
  );
}
