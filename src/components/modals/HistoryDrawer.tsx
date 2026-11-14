import { useCallback, useEffect, useState } from 'react';
import { Download, FolderUp, History, RefreshCw, Trash2 } from 'lucide-react';
import { Badge, Button, Drawer, EmptyState } from '../ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import {
  buildDownloadUrl,
  buildPreviewUrl,
  clearHistory,
  deleteHistoryItem,
  getExtension,
  getHistory,
} from '../../services/api';
import { suggestToolsForExtension } from '../../config/navigation';
import type { HistoryItem } from '../../types/api';

export function HistoryDrawer() {
  const { historyOpen, setHistoryOpen, navigateToTool } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [items, setItems] = useState<HistoryItem[]>([]);
  const [loading, setLoading] = useState(false);

  const loadItems = useCallback(async () => {
    setLoading(true);
    try {
      const list = await getHistory();
      setItems(list);
    } catch (err) {
      toast.error('Không tải được lịch sử xử lý', err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, [toast]);

  useEffect(() => {
    if (historyOpen) {
      loadItems();
    }
  }, [historyOpen, loadItems]);

  const handleDeleteOne = async (id: number) => {
    try {
      await deleteHistoryItem(id);
      setItems((prev) => prev.filter((i) => i.id !== id));
    } catch (err) {
      toast.error('Lỗi xóa mục lịch sử', err instanceof Error ? err.message : String(err));
    }
  };

  const handleClearAll = async () => {
    try {
      await clearHistory();
      setItems([]);
      toast.success('Đã dọn sạch lịch sử SQLite');
    } catch (err) {
      toast.error('Lỗi xóa toàn bộ lịch sử', err instanceof Error ? err.message : String(err));
    }
  };

  const handleReopen = (item: HistoryItem) => {
    const outName = item.output_name || item.filename || 'document';
    const ext = getExtension(outName);
    const sf = addSessionFile({
      filePath: item.output_path || outName,
      filename: outName,
      previewUrl: buildPreviewUrl(outName),
      downloadUrl: buildDownloadUrl(outName),
      sourceLabel: item.action || item.tool || 'Lịch sử',
    });
    const suggested = suggestToolsForExtension(ext)[0];
    setHistoryOpen(false);
    if (suggested) {
      navigateToTool(suggested.id, sf);
    } else {
      toast.success('Đã đưa tệp vào Kho tệp phiên', outName);
    }
  };

  return (
    <Drawer
      open={historyOpen}
      onClose={() => setHistoryOpen(false)}
      title="Nhật Ký Xử Lý Tài Liệu"
      subtitle="Lưu cục bộ trong SQLite trên máy cá nhân"
    >
      <div className="flex h-full flex-col gap-3">
        <div className="flex items-center justify-between">
          <Button
            size="sm"
            variant="outline"
            loading={loading}
            icon={<RefreshCw className="h-3.5 w-3.5" />}
            onClick={loadItems}
          >
            Làm mới
          </Button>
          {items.length > 0 ? (
            <Button
              size="sm"
              variant="ghost"
              icon={<Trash2 className="h-3.5 w-3.5 text-red-500" />}
              onClick={handleClearAll}
              className="text-red-600 hover:bg-red-50 dark:hover:bg-red-950/30"
            >
              Xóa toàn bộ
            </Button>
          ) : null}
        </div>

        <div className="flex-1 space-y-2 overflow-y-auto pr-1">
          {items.length === 0 ? (
            <EmptyState
              icon={<History className="h-6 w-6" />}
              title="Chưa có lịch sử xử lý"
              description="Mỗi thao tác chuyển đổi, chuẩn hóa Word, PDF Pro hoặc Excel sẽ tự động lưu nhật ký tại đây."
            />
          ) : (
            items.map((item) => {
              const outName = item.output_name || item.filename || 'Không rõ tên';
              const dlUrl = buildDownloadUrl(outName);
              return (
                <div
                  key={item.id}
                  className="rounded-xl border border-slate-200 bg-white p-3 text-xs dark:border-slate-800 dark:bg-slate-950"
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0 flex-1">
                      <div className="truncate font-semibold text-slate-900 dark:text-slate-100">{outName}</div>
                      <div className="mt-0.5 flex flex-wrap items-center gap-1.5 text-[11px] text-slate-500">
                        <Badge tone="indigo">{item.action || item.tool || 'Tác vụ'}</Badge>
                        <span>{item.created_at || item.timestamp || ''}</span>
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => handleDeleteOne(item.id)}
                      className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-red-600 dark:hover:bg-red-950/40"
                      title="Xóa bản ghi"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </div>
                  <div className="mt-2.5 flex items-center gap-2">
                    <a
                      href={dlUrl}
                      download={outName}
                      className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                    >
                      <Download className="h-3 w-3" />
                      <span>Tải tệp</span>
                    </a>
                    <button
                      type="button"
                      onClick={() => handleReopen(item)}
                      className="inline-flex items-center gap-1 rounded-md border border-indigo-200 bg-indigo-50/70 px-2.5 py-1 text-[11px] font-medium text-indigo-700 hover:bg-indigo-100 dark:border-indigo-900/60 dark:bg-indigo-950/40 dark:text-indigo-300"
                    >
                      <FolderUp className="h-3 w-3" />
                      <span>Mở xử lý tiếp</span>
                    </button>
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>
    </Drawer>
  );
}
