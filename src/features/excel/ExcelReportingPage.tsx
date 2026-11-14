import { useEffect, useState } from 'react';
import { ArrowDown, ArrowUp, FileSpreadsheet, FileText, Layers, Plus, Trash2 } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  FileDropZone,
  ResultPreviewCard,
  Tabs,
  sessionFileToSelected,
  type SelectedFileInfo,
} from '../../components/ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { apiPost, buildDownloadUrl, buildPreviewUrl, uploadFiles } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

interface ReportingResponse {
  success?: boolean;
  filename: string;
  download_url?: string;
  output_path?: string;
  details?: Record<string, unknown>;
}

export function ExcelReportingPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [mode, setMode] = useState<'to-brief' | 'consolidate'>('to-brief');

  // To Brief
  const [excelFile, setExcelFile] = useState<SelectedFileInfo | null>(null);
  const [targetFormat, setTargetFormat] = useState<'docx' | 'pdf' | 'md'>('docx');
  const [briefTitle, setBriefTitle] = useState('');
  const [includeKpi, setIncludeKpi] = useState(true);

  // Consolidate
  const [consolidateFiles, setConsolidateFiles] = useState<SelectedFileInfo[]>([]);
  const [consolidateMode, setConsolidateMode] = useState<'append_rows' | 'separate_sheets'>('append_rows');
  const [strictSchema, setStrictSchema] = useState(true);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && ['xlsx', 'xlsm', 'xls'].includes(hf.ext.toLowerCase())) {
      const sel = sessionFileToSelected(hf);
      setExcelFile(sel);
      setConsolidateFiles((prev) => (prev.some((x) => x.filePath === sel.filePath) ? prev : [...prev, sel]));
    }
  }, [consumeHandoffFile]);

  const handleMultiUpload = async (fileList: FileList | null) => {
    if (!fileList || !fileList.length) return;
    try {
      const res = await uploadFiles(Array.from(fileList));
      const added: SelectedFileInfo[] = (res.files || []).map((f) => {
        const sf = addSessionFile({
          filePath: f.file_path,
          filename: f.filename,
          size: f.size,
          previewUrl: f.preview_url,
          sourceLabel: 'Tổng hợp Excel',
        });
        return sessionFileToSelected(sf);
      });
      setConsolidateFiles((prev) => [...prev, ...added]);
      toast.success(`Đã thêm ${added.length} tệp Excel vào danh sách`);
    } catch (err) {
      toast.error('Lỗi tải tệp Excel', err instanceof Error ? err.message : String(err));
    }
  };

  const moveItem = (idx: number, dir: -1 | 1) => {
    setConsolidateFiles((prev) => {
      const target = idx + dir;
      if (target < 0 || target >= prev.length) return prev;
      const copy = [...prev];
      const [item] = copy.splice(idx, 1);
      copy.splice(target, 0, item);
      return copy;
    });
  };

  const handleRun = async () => {
    setRunning(true);
    try {
      if (mode === 'to-brief') {
        if (!excelFile) {
          toast.warning('Chưa chọn tệp Excel nguồn');
          setRunning(false);
          return;
        }
        const res = await apiPost<ReportingResponse>('/api/xlsx/to-brief', {
          file_path: excelFile.filePath,
          target_format: targetFormat,
          title: briefTitle.trim() || undefined,
          include_kpi_summary: includeKpi,
        });
        const sf = addSessionFile({
          filePath: res.output_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: `Báo cáo .${targetFormat.toUpperCase()}`,
        });
        setOutputSession(sf);
        toast.success('Đã tổng hợp báo cáo từ Excel', res.filename);
      } else {
        if (consolidateFiles.length < 2) {
          toast.warning('Cần ít nhất 2 tệp Excel để tổng hợp');
          setRunning(false);
          return;
        }
        const res = await apiPost<ReportingResponse>('/api/xlsx/consolidate', {
          inputs: consolidateFiles.map((f) => f.filePath),
          mode: consolidateMode,
          strict_schema: strictSchema,
        });
        const sf = addSessionFile({
          filePath: res.output_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: 'Tổng hợp Master Excel',
        });
        setOutputSession(sf);
        toast.success('Tổng hợp Workbooks hoàn tất', res.filename);
      }
    } catch (err) {
      toast.error('Lỗi thực thi báo cáo Excel', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const previewTarget = outputSession
    ? {
        filePath: outputSession.filePath,
        filename: outputSession.filename,
        ext: outputSession.ext,
        previewUrl: outputSession.previewUrl,
        downloadUrl: outputSession.downloadUrl,
      }
    : mode === 'to-brief' && excelFile
      ? {
          filePath: excelFile.filePath,
          filename: excelFile.filename,
          ext: excelFile.ext,
          previewUrl: excelFile.previewUrl,
          downloadUrl: excelFile.downloadUrl,
        }
      : consolidateFiles[0]
        ? {
            filePath: consolidateFiles[0].filePath,
            filename: consolidateFiles[0].filename,
            ext: consolidateFiles[0].ext,
            previewUrl: consolidateFiles[0].previewUrl,
            downloadUrl: consolidateFiles[0].downloadUrl,
          }
        : null;

  const leftPanel = (
    <div className="space-y-4">
      <Tabs
        activeId={mode}
        onChange={(id) => setMode(id as 'to-brief' | 'consolidate')}
        className="w-full"
        options={[
          { id: 'to-brief', label: 'Excel -> Báo cáo Brief (Word/PDF)', icon: <FileText className="h-3.5 w-3.5" /> },
          { id: 'consolidate', label: 'Tổng hợp nhiều File Excel', icon: <Layers className="h-3.5 w-3.5" /> },
        ]}
      />

      {mode === 'to-brief' ? (
        <div className="space-y-3">
          <FileDropZone
            label="Bảng tính Excel cần tổng hợp thành văn bản"
            accept=".xlsx,.xlsm,.xls"
            selectedFile={excelFile}
            onFileSelected={(f) => {
              setExcelFile(f);
              setOutputSession(null);
            }}
            onClear={() => setExcelFile(null)}
            sourceLabel="Excel Báo cáo"
          />
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Định dạng xuất</label>
              <select
                value={targetFormat}
                onChange={(e) => setTargetFormat(e.target.value as 'docx' | 'pdf' | 'md')}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="docx">Word Hành chính (.docx)</option>
                <option value="pdf">PDF Trình ký (.pdf)</option>
                <option value="md">Markdown (.md)</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Tiêu đề báo cáo</label>
              <input
                type="text"
                value={briefTitle}
                onChange={(e) => setBriefTitle(e.target.value)}
                placeholder="Để trống để tự đặt theo tên file"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>
          <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={includeKpi} onChange={(e) => setIncludeKpi(e.target.checked)} />
            <span>Tự động tổng hợp chỉ số KPI (Tổng / Trung bình / Min / Max) bằng Dual-Formula Evaluator</span>
          </label>
          <Button
            variant="primary"
            loading={running}
            disabled={!excelFile}
            icon={<FileSpreadsheet className="h-4 w-4" />}
            onClick={handleRun}
            className="w-full bg-emerald-600 hover:bg-emerald-700"
          >
            Xuất báo cáo từ bảng tính
          </Button>
        </div>
      ) : (
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-slate-600 dark:text-slate-400">
              Danh sách tệp đơn vị/chi nhánh ({consolidateFiles.length})
            </span>
            <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-xs font-medium text-emerald-700 hover:bg-emerald-100 dark:border-emerald-900/60 dark:bg-emerald-950/40 dark:text-emerald-300">
              <Plus className="h-3.5 w-3.5" />
              <span>Thêm nhiều Excel</span>
              <input
                type="file"
                multiple
                accept=".xlsx,.xlsm,.xls"
                className="hidden"
                onChange={(e) => {
                  handleMultiUpload(e.target.files);
                  e.currentTarget.value = '';
                }}
              />
            </label>
          </div>
          {consolidateFiles.length === 0 ? (
            <div className="rounded-lg border border-dashed border-slate-200 px-3 py-6 text-center text-xs text-slate-500 dark:border-slate-800 dark:text-slate-400">
              Chọn từ 2 tệp Excel trở lên (cùng cấu trúc cột) để nối gộp thành 1 file Master.
            </div>
          ) : (
            <ul className="max-h-48 space-y-1.5 overflow-y-auto pr-1">
              {consolidateFiles.map((f, idx) => (
                <li
                  key={`${f.filePath}-${idx}`}
                  className="flex items-center justify-between gap-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-800 dark:bg-slate-950"
                >
                  <span className="truncate font-medium text-slate-800 dark:text-slate-200">
                    {idx + 1}. {f.filename}
                  </span>
                  <div className="flex shrink-0 items-center gap-1">
                    <button
                      type="button"
                      disabled={idx === 0}
                      onClick={() => moveItem(idx, -1)}
                      className="rounded p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-30 dark:hover:bg-slate-800"
                    >
                      <ArrowUp className="h-3.5 w-3.5" />
                    </button>
                    <button
                      type="button"
                      disabled={idx === consolidateFiles.length - 1}
                      onClick={() => moveItem(idx, 1)}
                      className="rounded p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-30 dark:hover:bg-slate-800"
                    >
                      <ArrowDown className="h-3.5 w-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => setConsolidateFiles((prev) => prev.filter((_, i) => i !== idx))}
                      className="rounded p-1 text-red-500 hover:bg-red-50 dark:hover:bg-red-950/40"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </div>
                </li>
              ))}
            </ul>
          )}

          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Chế độ tổng hợp</label>
            <select
              value={consolidateMode}
              onChange={(e) => setConsolidateMode(e.target.value as 'append_rows' | 'separate_sheets')}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            >
              <option value="append_rows">Nối tiếp tất cả dòng vào 1 bảng Master (Append Rows)</option>
              <option value="separate_sheets">Giữ mỗi tệp làm 1 Sheet riêng trong cùng Workbook</option>
            </select>
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao">
            <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
              <input type="checkbox" checked={strictSchema} onChange={(e) => setStrictSchema(e.target.checked)} />
              <span>Kiểm tra khớp tên cột tiêu đề (Strict Schema Validation)</span>
            </label>
          </CollapsibleSection>

          <Button
            variant="primary"
            loading={running}
            disabled={consolidateFiles.length < 2}
            icon={<Layers className="h-4 w-4" />}
            onClick={handleRun}
            className="w-full bg-emerald-600 hover:bg-emerald-700"
          >
            Tổng hợp {consolidateFiles.length} tệp Excel thành Master
          </Button>
        </div>
      )}
    </div>
  );

  return (
    <SplitWorkspace
      title="Báo cáo & Tổng hợp Excel Đa chi nhánh"
      subtitle="Chuyển bảng tính Excel thành văn bản báo cáo Word/PDF kèm KPI hoặc hợp nhất nhiều file báo cáo phòng ban thành Master Workbook."
      badgeText="Reporting Engine"
      accent="excel"
      leftPanel={leftPanel}
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Xuất báo cáo & tổng hợp hoàn tất"
            outputFile={outputSession}
            summaryItems={[{ label: 'Định dạng', value: `.${outputSession.ext.toUpperCase()}` }]}
          />
        ) : undefined
      }
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="green">Tự tính công thức tổng hợp</Badge>}
    />
  );
}
