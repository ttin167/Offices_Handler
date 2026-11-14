import { useEffect, useState } from 'react';
import { Calculator, CheckCircle2, FileCode2, Printer, ShieldAlert } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  FileDropZone,
  ResultPreviewCard,
  Tabs,
  sessionFileToSelected,
  type SelectedFileInfo,
} from '../../components/ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { apiPost, buildDownloadUrl, buildPreviewUrl } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

type TemplateToolMode = 'eval' | 'validate' | 'print_setup' | 'mutate';

interface ValidateResponse {
  success: boolean;
  passed: boolean;
  target_sheet?: string;
  ref_sheet?: string;
  gates_evaluated?: number;
  gates_failed?: number;
  diffs?: Array<Record<string, unknown>>;
  diff_count?: number;
}

const DEFAULT_PATCH_JSON = JSON.stringify(
  {
    sheets: {
      Sheet1: {
        cells: {
          A1: 'BÁO CÁO ĐÃ CẬP NHẬT',
        },
      },
    },
  },
  null,
  2,
);

export function ExcelTemplatePage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [mode, setMode] = useState<TemplateToolMode>('eval');
  const [excelFile, setExcelFile] = useState<SelectedFileInfo | null>(null);

  // Validate
  const [targetSheet, setTargetSheet] = useState('');
  const [refSheet, setRefSheet] = useState('');
  const [validateResult, setValidateResult] = useState<ValidateResponse | null>(null);

  // Print setup
  const [printSheet, setPrintSheet] = useState('');
  const [printOrientation, setPrintOrientation] = useState('auto');

  // Mutate
  const [patchJson, setPatchJson] = useState(DEFAULT_PATCH_JSON);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [lastTitle, setLastTitle] = useState('');

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && ['xlsx', 'xlsm', 'xls'].includes(hf.ext.toLowerCase())) {
      setExcelFile(sessionFileToSelected(hf));
    }
  }, [consumeHandoffFile]);

  const handleRun = async () => {
    if (!excelFile) {
      toast.warning('Chưa có tệp Excel', 'Vui lòng tải lên tệp .xlsx để xử lý.');
      return;
    }
    setRunning(true);
    try {
      if (mode === 'eval') {
        const res = await apiPost<{ filename: string; download_url?: string; output_path?: string }>(
          '/api/xlsx/eval',
          { file_path: excelFile.filePath },
        );
        const sf = addSessionFile({
          filePath: res.output_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: 'AST Formula Eval',
        });
        setOutputSession(sf);
        setLastTitle('Đã tính trước giá trị công thức (Offline AST)');
        toast.success('Tính giá trị công thức thành công', res.filename);
      } else if (mode === 'validate') {
        const res = await apiPost<ValidateResponse>('/api/xlsx/validate', {
          file_path: excelFile.filePath,
          target_sheet: targetSheet.trim() || undefined,
          ref_sheet: refSheet.trim() || undefined,
        });
        setValidateResult(res);
        setOutputSession(null);
        if (res.passed) {
          toast.success(`Kiểm định ĐẠT (${res.gates_evaluated ?? 12} cổng)`);
        } else {
          toast.warning(`Phát hiện ${res.diff_count ?? res.gates_failed ?? 0} điểm lệch định dạng`);
        }
      } else if (mode === 'print_setup') {
        const res = await apiPost<{ filename: string }>('/api/xlsx/print-setup', {
          file_path: excelFile.filePath,
          sheet_name: printSheet.trim() || undefined,
          orientation: printOrientation === 'auto' ? undefined : printOrientation,
        });
        const sf = addSessionFile({
          filePath: excelFile.filePath,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: buildDownloadUrl(res.filename),
          sourceLabel: 'Print Setup A4',
        });
        setOutputSession(sf);
        setLastTitle('Đã chuẩn hóa Page Setup vừa khít khổ giấy in');
        toast.success('Cấu hình khổ in Excel thành công', res.filename);
      } else {
        let parsedSpec: unknown = {};
        try {
          parsedSpec = JSON.parse(patchJson);
        } catch {
          toast.error('JSON Patch không hợp lệ', 'Vui lòng kiểm tra cú pháp JSON.');
          setRunning(false);
          return;
        }
        const res = await apiPost<{ filename: string; download_url?: string; result_path?: string }>(
          '/api/xlsx/mutate',
          { template_path: excelFile.filePath, spec: parsedSpec },
        );
        const sf = addSessionFile({
          filePath: res.result_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: 'Điền Template Excel',
        });
        setOutputSession(sf);
        setLastTitle('Đã điền dữ liệu bảo toàn 100% Template');
        toast.success('Cập nhật Template Excel thành công', res.filename);
      }
    } catch (err) {
      toast.error('Lỗi thực thi Excel Engine', err instanceof Error ? err.message : String(err));
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
    : excelFile
      ? {
          filePath: excelFile.filePath,
          filename: excelFile.filename,
          ext: excelFile.ext,
          previewUrl: excelFile.previewUrl,
          downloadUrl: excelFile.downloadUrl,
        }
      : null;

  const leftPanel = (
    <div className="space-y-4">
      <FileDropZone
        label="Tệp Excel Template / Bảng tính (.xlsx)"
        accept=".xlsx,.xlsm"
        selectedFile={excelFile}
        onFileSelected={(f) => {
          setExcelFile(f);
          setOutputSession(null);
          setValidateResult(null);
        }}
        onClear={() => setExcelFile(null)}
        sourceLabel="Excel Engine"
      />

      <Tabs
        activeId={mode}
        onChange={(id) => setMode(id as TemplateToolMode)}
        className="w-full"
        options={[
          { id: 'eval', label: 'Tính công thức AST', icon: <Calculator className="h-3.5 w-3.5" /> },
          { id: 'validate', label: 'Kiểm định 12 cổng', icon: <CheckCircle2 className="h-3.5 w-3.5" /> },
          { id: 'print_setup', label: 'Khổ in vừa trang', icon: <Printer className="h-3.5 w-3.5" /> },
          { id: 'mutate', label: 'Điền JSON nâng cao', icon: <FileCode2 className="h-3.5 w-3.5" /> },
        ]}
      />

      <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
        {mode === 'eval' ? (
          <p className="text-xs leading-relaxed text-slate-600 dark:text-slate-400">
            Sử dụng bộ đánh giá cây cú pháp AST offline để tính toán và điền sẵn giá trị cached (`&lt;v&gt;`) cho mọi công thức (`SUM`, `SUMIFS`, `AVERAGE`, `IF`, `VLOOKUP`...) mà không cần mở Microsoft Excel.
          </p>
        ) : null}

        {mode === 'validate' ? (
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Sheet cần kiểm tra (tùy chọn)
              </label>
              <input
                type="text"
                value={targetSheet}
                onChange={(e) => setTargetSheet(e.target.value)}
                placeholder="Mặc định sheet đầu tiên"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Sheet mẫu chuẩn đối chiếu (tùy chọn)
              </label>
              <input
                type="text"
                value={refSheet}
                onChange={(e) => setRefSheet(e.target.value)}
                placeholder="VD: Example hoặc Template"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>
        ) : null}

        {mode === 'print_setup' ? (
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Tên Sheet áp dụng (để trống = tất cả)
              </label>
              <input
                type="text"
                value={printSheet}
                onChange={(e) => setPrintSheet(e.target.value)}
                placeholder="Tất cả các sheet"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Hướng giấy in</label>
              <select
                value={printOrientation}
                onChange={(e) => setPrintOrientation(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="auto">Tự động fit vừa chiều ngang (1 page wide)</option>
                <option value="landscape">Giấy Ngang (Landscape)</option>
                <option value="portrait">Giấy Dọc (Portrait)</option>
              </select>
            </div>
          </div>
        ) : null}

        {mode === 'mutate' ? (
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              JSON Patch Spec (Bảo toàn 100% Logo, Chart & Styles gốc)
            </label>
            <textarea
              rows={8}
              value={patchJson}
              onChange={(e) => setPatchJson(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white p-2.5 font-mono text-[11px] dark:border-slate-700 dark:bg-slate-950"
            />
          </div>
        ) : null}
      </div>

      <Button
        variant="primary"
        loading={running}
        disabled={!excelFile}
        onClick={handleRun}
        className="w-full bg-emerald-600 hover:bg-emerald-700"
      >
        Thực thi tác vụ Engine Excel
      </Button>
    </div>
  );

  const resultPanel = outputSession ? (
    <ResultPreviewCard
      title={lastTitle || 'Hoàn tất'}
      outputFile={outputSession}
      summaryItems={[{ label: 'Tệp xuất', value: outputSession.filename }]}
    />
  ) : validateResult ? (
    <div
      className={`space-y-3 rounded-xl border p-4 ${
        validateResult.passed
          ? 'border-emerald-200 bg-emerald-50/50 dark:border-emerald-900/60 dark:bg-emerald-950/30'
          : 'border-amber-200 bg-amber-50/50 dark:border-amber-900/60 dark:bg-amber-950/30'
      }`}
    >
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2 text-xs font-semibold">
          {validateResult.passed ? (
            <CheckCircle2 className="h-4 w-4 text-emerald-600" />
          ) : (
            <ShieldAlert className="h-4 w-4 text-amber-600" />
          )}
          <span>
            {validateResult.passed
              ? 'Đạt chuẩn Template OpenXML (12/12 cổng)'
              : `Phát hiện ${validateResult.diff_count ?? validateResult.gates_failed ?? 0} sai lệch cần kiểm tra`}
          </span>
        </div>
        <Badge tone={validateResult.passed ? 'green' : 'amber'}>
          {validateResult.target_sheet || 'Workbook'}
        </Badge>
      </div>
      {validateResult.diffs && validateResult.diffs.length > 0 ? (
        <ul className="max-h-52 space-y-1.5 overflow-y-auto text-[11px]">
          {validateResult.diffs.slice(0, 20).map((d, i) => (
            <li key={i} className="rounded bg-white/80 px-2.5 py-1.5 font-mono dark:bg-slate-900/80">
              {JSON.stringify(d)}
            </li>
          ))}
        </ul>
      ) : null}
    </div>
  ) : undefined;

  return (
    <SplitWorkspace
      title="Engine Excel Template, AST Formula & Khổ In"
      subtitle="Tính toán giá trị công thức offline, kiểm định 12 cổng chất lượng OpenXML, căn chỉnh trang in vừa khổ A4 và điền Template JSON nâng cao."
      badgeText="OpenXML Safe"
      accent="excel"
      leftPanel={leftPanel}
      resultPanel={resultPanel}
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="green">Bảo toàn DrawingML</Badge>}
    />
  );
}
