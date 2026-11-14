import { useEffect, useState } from 'react';
import { ArrowRightLeft, FileScan, RefreshCw } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  FileDropZone,
  ResultPreviewCard,
  sessionFileToSelected,
  type SelectedFileInfo,
} from '../../components/ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { apiPost, buildDownloadUrl, buildPreviewUrl } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

interface ConvertResponse {
  success: boolean;
  filename: string;
  download_url?: string;
  result_path?: string;
  elapsed?: number;
}

const TARGET_OPTIONS: Array<{ value: string; label: string }> = [
  { value: '.docx', label: 'Word OpenXML (.docx)' },
  { value: '.pdf', label: 'PDF Chuẩn in (.pdf)' },
  { value: '.md', label: 'Markdown Tách biệt (.md)' },
  { value: '.html', label: 'Web HTML5 (.html)' },
];

export function UniversalConverterPage() {
  const { consumeHandoffFile, navigateToTool } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [sourceFile, setSourceFile] = useState<SelectedFileInfo | null>(null);
  const [targetExt, setTargetExt] = useState('.docx');
  const [pageRange, setPageRange] = useState('');
  const [isScannedPdf, setIsScannedPdf] = useState(false);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [elapsedSec, setElapsedSec] = useState<number | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && ['pdf', 'docx', 'md', 'markdown', 'html', 'htm'].includes(hf.ext.toLowerCase())) {
      const sel = sessionFileToSelected(hf);
      setSourceFile(sel);
      suggestDefaultTarget(sel.ext);
    }
  }, [consumeHandoffFile]);

  const suggestDefaultTarget = (ext: string) => {
    const clean = ext.toLowerCase().replace('.', '');
    if (clean === 'pdf') setTargetExt('.docx');
    else if (clean === 'docx') setTargetExt('.pdf');
    else if (clean === 'md' || clean === 'markdown') setTargetExt('.docx');
    else setTargetExt('.docx');
  };

  const inspectScanned = async (file: SelectedFileInfo) => {
    setIsScannedPdf(false);
    if (file.ext.toLowerCase() !== 'pdf') return;
    try {
      const res = await apiPost<{ success: boolean; is_scanned?: boolean }>('/api/ocr/inspect-scanned', {
        file_path: file.filePath,
      });
      if (res.is_scanned) {
        setIsScannedPdf(true);
      }
    } catch {
      // Non-fatal silent probe
    }
  };

  const handleSelectFile = (f: SelectedFileInfo) => {
    setSourceFile(f);
    setOutputSession(null);
    suggestDefaultTarget(f.ext);
    inspectScanned(f);
  };

  const handleConvert = async () => {
    if (!sourceFile) {
      toast.warning('Chưa chọn tệp nguồn', 'Vui lòng tải lên PDF, Word, Markdown hoặc HTML.');
      return;
    }
    setRunning(true);
    try {
      const res = await apiPost<ConvertResponse>('/api/convert', {
        source_path: sourceFile.filePath,
        target_ext: targetExt,
        page_range: pageRange.trim() || undefined,
      });
      const sf = addSessionFile({
        filePath: res.result_path || res.filename,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: `Chuyển -> ${targetExt.toUpperCase()}`,
      });
      setOutputSession(sf);
      setElapsedSec(res.elapsed ?? null);
      toast.success('Chuyển đổi tài liệu hoàn tất', res.filename);
    } catch (err) {
      toast.error('Lỗi chuyển đổi', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleSwitchToOcr = () => {
    if (!sourceFile) return;
    navigateToTool('ocr-vietnamese', {
      id: `ocr-${Date.now()}`,
      filePath: sourceFile.filePath,
      filename: sourceFile.filename,
      ext: 'pdf',
      size: sourceFile.size,
      previewUrl: sourceFile.previewUrl,
      downloadUrl: sourceFile.downloadUrl,
      sourceLabel: 'Chuyển từ Scan PDF',
      addedAt: Date.now(),
    });
  };

  const previewTarget = outputSession
    ? {
        filePath: outputSession.filePath,
        filename: outputSession.filename,
        ext: outputSession.ext,
        previewUrl: outputSession.previewUrl,
        downloadUrl: outputSession.downloadUrl,
      }
    : sourceFile
      ? {
          filePath: sourceFile.filePath,
          filename: sourceFile.filename,
          ext: sourceFile.ext,
          previewUrl: sourceFile.previewUrl,
          downloadUrl: sourceFile.downloadUrl,
        }
      : null;

  const leftPanel = (
    <div className="space-y-4">
      <FileDropZone
        label="Tài liệu nguồn (PDF, Word DOCX, Markdown, HTML)"
        accept=".pdf,.docx,.md,.markdown,.html,.htm"
        selectedFile={sourceFile}
        onFileSelected={handleSelectFile}
        onClear={() => {
          setSourceFile(null);
          setIsScannedPdf(false);
        }}
        sourceLabel="Chuyển đổi đa năng"
      />

      {isScannedPdf ? (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-amber-300 bg-amber-50 px-3 py-2.5 text-xs text-amber-900 dark:border-amber-800 dark:bg-amber-950/40 dark:text-amber-200">
          <div className="flex items-center gap-2">
            <FileScan className="h-4 w-4 shrink-0 text-amber-600" />
            <span>Tệp PDF này là bản Scan dạng ảnh. Khuyên dùng OCR Tiếng Việt để bóc chữ chuẩn xác.</span>
          </div>
          <Button size="sm" variant="outline" onClick={handleSwitchToOcr}>
            Chuyển sang OCR ngay
          </Button>
        </div>
      ) : null}

      <div>
        <label className="mb-1.5 block text-xs font-medium text-slate-700 dark:text-slate-300">
          Định dạng đầu ra mong muốn
        </label>
        <div className="grid grid-cols-2 gap-2">
          {TARGET_OPTIONS.map((opt) => {
            const active = targetExt === opt.value;
            return (
              <button
                key={opt.value}
                type="button"
                onClick={() => setTargetExt(opt.value)}
                className={`rounded-xl border px-3 py-2.5 text-left text-xs font-medium transition ${
                  active
                    ? 'border-indigo-500 bg-indigo-50 text-indigo-800 shadow-subtle dark:border-indigo-700 dark:bg-indigo-950/40 dark:text-indigo-200'
                    : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300'
                }`}
              >
                {opt.label}
              </button>
            );
          })}
        </div>
      </div>

      <CollapsibleSection title="Tùy chọn nâng cao (Giới hạn khoảng trang PDF)">
        <div>
          <label className="mb-1 block text-[11px] text-slate-500">
            Khoảng trang (ví dụ: 1-5 hoặc 1,3,6 — để trống để chuyển toàn bộ)
          </label>
          <input
            type="text"
            value={pageRange}
            onChange={(e) => setPageRange(e.target.value)}
            placeholder="VD: 1-10"
            className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
          />
        </div>
      </CollapsibleSection>

      <Button
        variant="primary"
        loading={running}
        disabled={!sourceFile}
        icon={<ArrowRightLeft className="h-4 w-4" />}
        onClick={handleConvert}
        className="w-full bg-indigo-600 hover:bg-indigo-700"
      >
        Chuyển đổi sang {targetExt.toUpperCase()}
      </Button>
    </div>
  );

  return (
    <SplitWorkspace
      title="Chuyển Đổi Tài Liệu Đa Chiều (Lossless)"
      subtitle="Chuyển đổi hai chiều PDF ↔ Word DOCX ↔ Markdown ↔ HTML bảo toàn bảng biểu, OpenXML styles và phát hiện PDF Scan."
      badgeText="Universal Engine"
      accent="studio"
      leftPanel={leftPanel}
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Chuyển đổi định dạng thành công"
            outputFile={outputSession}
            summaryItems={[
              { label: 'Định dạng đích', value: targetExt.toUpperCase() },
              ...(elapsedSec !== null ? [{ label: 'Thời gian', value: `${elapsedSec.toFixed(2)}s` }] : []),
            ]}
            extraActions={
              <Button
                size="sm"
                variant="outline"
                icon={<RefreshCw className="h-3.5 w-3.5" />}
                onClick={() => handleSelectFile(sessionFileToSelected(outputSession))}
              >
                Chuyển đổi tiếp tệp này
              </Button>
            }
          />
        ) : undefined
      }
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="indigo">Zero-Data-Loss AST</Badge>}
    />
  );
}
