import { useCallback, useEffect, useState } from 'react';
import { Printer, RefreshCw } from 'lucide-react';
import {
  Badge,
  Button,
  CollapsibleSection,
  FileDropZone,
  Modal,
  sessionFileToSelected,
  type SelectedFileInfo,
} from '../ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useToast } from '../../context/ToastContext';
import { apiGet, apiPost, buildDownloadUrl, buildPreviewUrl, getExtension } from '../../services/api';

interface PrinterItem {
  name: string;
  is_default?: boolean;
  status?: string;
}

interface PrintersResponse {
  success: boolean;
  default_printer?: string;
  printers?: Array<PrinterItem | string>;
}

export function PrintStudioModal() {
  const { printModalOpen, closePrintModal, printTargetFile } = useWorkspace();
  const toast = useToast();

  const [selectedFile, setSelectedFile] = useState<SelectedFileInfo | null>(null);
  const [printers, setPrinters] = useState<PrinterItem[]>([]);
  const [printerName, setPrinterName] = useState('');
  const [printerStatusMsg, setPrinterStatusMsg] = useState('');
  const [loadingPrinters, setLoadingPrinters] = useState(false);

  const [copies, setCopies] = useState(1);
  const [pageRange, setPageRange] = useState('');
  const [pageSize, setPageSize] = useState('a4');
  const [duplex, setDuplex] = useState('none');
  const [scaling, setScaling] = useState('fit');
  const [nUp, setNUp] = useState(1);
  const [grayscale, setGrayscale] = useState(false);
  const [watermark, setWatermark] = useState('');
  const [printing, setPrinting] = useState(false);

  const loadPrinters = useCallback(async () => {
    setLoadingPrinters(true);
    try {
      const res = await apiGet<PrintersResponse>('/api/print/printers');
      const normalized: PrinterItem[] = (res.printers || []).map((p) =>
        typeof p === 'string'
          ? { name: p, is_default: p === res.default_printer }
          : p,
      );
      setPrinters(normalized);
      const def = res.default_printer || normalized[0]?.name || '';
      setPrinterName((prev) => prev || def);
    } catch (err) {
      toast.error('Không tải được danh sách máy in', err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingPrinters(false);
    }
  }, [toast]);

  useEffect(() => {
    if (!printModalOpen) return;
    loadPrinters();
    if (printTargetFile) {
      const ext = getExtension(printTargetFile.filename);
      setSelectedFile(
        sessionFileToSelected({
          id: `print-${Date.now()}`,
          filePath: printTargetFile.filePath,
          filename: printTargetFile.filename,
          ext,
          previewUrl: buildPreviewUrl(printTargetFile.filename),
          downloadUrl: buildDownloadUrl(printTargetFile.filename),
          sourceLabel: 'In ấn',
          addedAt: Date.now(),
        }),
      );
    }
  }, [printModalOpen, printTargetFile, loadPrinters]);

  const checkPrinterStatus = async () => {
    if (!printerName) return;
    try {
      const res = await apiGet<{ status_text?: string; status?: string; jobs_queued?: number }>(
        `/api/print/status?printer=${encodeURIComponent(printerName)}`,
      );
      setPrinterStatusMsg(
        `${res.status_text || res.status || 'Sẵn sàng'} · Hàng đợi: ${res.jobs_queued ?? 0} lệnh`,
      );
    } catch (err) {
      setPrinterStatusMsg(err instanceof Error ? err.message : 'Không kiểm tra được trạng thái');
    }
  };

  const handlePrint = async () => {
    if (!selectedFile) {
      toast.warning('Chưa chọn tệp tài liệu để in');
      return;
    }
    setPrinting(true);
    try {
      await apiPost('/api/print/job', {
        file_path: selectedFile.filePath,
        printer_name: printerName || undefined,
        copies,
        page_range: pageRange.trim() || undefined,
        page_size: pageSize,
        duplex: duplex === 'none' ? undefined : duplex,
        scaling,
        n_up: nUp,
        grayscale,
        vector_grayscale: true,
        watermark: watermark.trim() || undefined,
      });
      toast.success('Đã gửi lệnh in tới máy in hệ thống', `${selectedFile.filename} (${copies} bản)`);
      closePrintModal();
    } catch (err) {
      toast.error('Lỗi gửi lệnh in', err instanceof Error ? err.message : String(err));
    } finally {
      setPrinting(false);
    }
  };

  return (
    <Modal
      open={printModalOpen}
      onClose={closePrintModal}
      title="Trung Tâm In Ấn Văn Phòng"
      subtitle="Hỗ trợ in im lặng chuẩn khổ A4, in 2 mặt (Duplex), đen trắng sắc nét /DeviceGray & dồn trang N-Up."
      size="lg"
    >
      <div className="space-y-4">
        <FileDropZone
          label="Tài liệu cần in (.pdf, .docx, .xlsx, ảnh)"
          accept=".pdf,.docx,.xlsx,.png,.jpg,.jpeg"
          selectedFile={selectedFile}
          onFileSelected={setSelectedFile}
          onClear={() => setSelectedFile(null)}
          sourceLabel="In ấn"
        />

        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
          <div>
            <div className="mb-1 flex items-center justify-between">
              <label className="text-xs font-medium text-slate-700 dark:text-slate-300">Máy in hệ thống</label>
              <button
                type="button"
                onClick={loadPrinters}
                className="inline-flex items-center gap-1 text-[11px] text-indigo-600 hover:underline dark:text-indigo-400"
              >
                <RefreshCw className={`h-3 w-3 ${loadingPrinters ? 'animate-spin' : ''}`} />
                <span>Làm mới</span>
              </button>
            </div>
            <select
              value={printerName}
              onChange={(e) => {
                setPrinterName(e.target.value);
                setPrinterStatusMsg('');
              }}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            >
              {printers.length === 0 ? (
                <option value="">Máy in mặc định Windows</option>
              ) : (
                printers.map((p) => (
                  <option key={p.name} value={p.name}>
                    {p.name} {p.is_default ? '(Mặc định)' : ''}
                  </option>
                ))
              )}
            </select>
            {printerStatusMsg ? <p className="mt-1 text-[11px] text-slate-500">{printerStatusMsg}</p> : null}
          </div>

          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Số bản sao</label>
              <input
                type="number"
                min={1}
                max={999}
                value={copies}
                onChange={(e) => setCopies(Math.max(1, Number(e.target.value)))}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Khổ giấy</label>
              <select
                value={pageSize}
                onChange={(e) => setPageSize(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="a4">A4 (Tiêu chuẩn VN)</option>
                <option value="letter">Letter</option>
              </select>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              In hai mặt (Duplex)
            </label>
            <select
              value={duplex}
              onChange={(e) => setDuplex(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            >
              <option value="none">In 1 mặt (Simplex)</option>
              <option value="vertical">In 2 mặt lật cạnh dài (Đóng gáy dọc)</option>
              <option value="horizontal">In 2 mặt lật cạnh ngắn (Ngang)</option>
            </select>
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Khoảng trang in (tùy chọn)
            </label>
            <input
              type="text"
              value={pageRange}
              onChange={(e) => setPageRange(e.target.value)}
              placeholder="VD: 1-5, 8 (Để trống = in tất cả)"
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </div>
        </div>

        <CollapsibleSection title="Tùy chọn nâng cao (Tiết kiệm mực, Dồn trang N-Up, Watermark)">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
            <div>
              <label className="mb-1 block text-[11px] text-slate-500">Tỉ lệ co trang</label>
              <select
                value={scaling}
                onChange={(e) => setScaling(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="fit">Fit vừa trang in (Khuyên dùng)</option>
                <option value="actual">Đúng kích thước 100%</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-[11px] text-slate-500">Số trang / tờ (N-Up)</label>
              <select
                value={nUp}
                onChange={(e) => setNUp(Number(e.target.value))}
                className="w-full rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value={1}>1 trang / mặt</option>
                <option value={2}>2 trang / mặt</option>
                <option value={4}>4 trang / mặt</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-[11px] text-slate-500">Đóng chữ mờ khi in</label>
              <input
                type="text"
                value={watermark}
                onChange={(e) => setWatermark(e.target.value)}
                placeholder="VD: BẢN NHÁP"
                className="w-full rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>
          <label className="mt-2 flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={grayscale} onChange={(e) => setGrayscale(e.target.checked)} />
            <span>Ép chế độ In Đen Trắng (/DeviceGray) chống nhòe chữ nhỏ</span>
          </label>
        </CollapsibleSection>

        <div className="flex flex-wrap items-center justify-between gap-2 pt-2">
          <div className="flex items-center gap-2">
            <Button size="sm" variant="outline" onClick={checkPrinterStatus} disabled={!printerName}>
              Kiểm tra Spooler
            </Button>
            <Badge tone="slate">DEVMODE Guard</Badge>
          </div>
          <div className="flex items-center gap-2">
            <Button variant="ghost" onClick={closePrintModal}>
              Hủy
            </Button>
            <Button
              variant="primary"
              loading={printing}
              disabled={!selectedFile}
              icon={<Printer className="h-4 w-4" />}
              onClick={handlePrint}
            >
              Gửi lệnh in ngay
            </Button>
          </div>
        </div>
      </div>
    </Modal>
  );
}
