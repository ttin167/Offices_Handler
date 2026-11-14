import { useEffect, useState } from 'react';
import { ScanText } from 'lucide-react';
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

interface OcrResponse {
  success?: boolean;
  filename: string;
  download_url?: string;
  output_path?: string;
  pages_processed?: number;
  total_chars?: number;
  elapsed?: number;
}

export function OcrVietnamesePage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [scanFile, setScanFile] = useState<SelectedFileInfo | null>(null);
  const [lang, setLang] = useState('vie+eng');
  const [outFormat, setOutFormat] = useState<'docx' | 'md'>('docx');
  const [filterRedStamp, setFilterRedStamp] = useState(true);
  const [deskew, setDeskew] = useState(true);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [stats, setStats] = useState<{ pages?: number; chars?: number; elapsed?: number }>({});

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && ['pdf', 'png', 'jpg', 'jpeg', 'tiff', 'bmp', 'webp'].includes(hf.ext.toLowerCase())) {
      setScanFile(sessionFileToSelected(hf));
    }
  }, [consumeHandoffFile]);

  const handleRunOcr = async () => {
    if (!scanFile) {
      toast.warning('Chưa chọn PDF Scan hoặc ảnh', 'Vui lòng tải lên tệp scan cần nhận dạng chữ.');
      return;
    }
    setRunning(true);
    try {
      const res = await apiPost<OcrResponse>('/api/ocr/scan', {
        file_path: scanFile.filePath,
        lang,
        format: outFormat,
        filter_red_stamp: filterRedStamp,
        deskew,
      });
      const sf = addSessionFile({
        filePath: res.output_path || res.filename,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: `OCR -> .${outFormat.toUpperCase()}`,
      });
      setOutputSession(sf);
      setStats({
        pages: res.pages_processed,
        chars: res.total_chars,
        elapsed: res.elapsed,
      });
      toast.success('Nhận dạng OCR Tiếng Việt hoàn tất', res.filename);
    } catch (err) {
      toast.error('Lỗi nhận dạng OCR', err instanceof Error ? err.message : String(err));
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
    : scanFile
      ? {
          filePath: scanFile.filePath,
          filename: scanFile.filename,
          ext: scanFile.ext,
          previewUrl: scanFile.previewUrl,
          downloadUrl: scanFile.downloadUrl,
        }
      : null;

  const leftPanel = (
    <div className="space-y-4">
      <FileDropZone
        label="Tài liệu Scan / Ảnh chụp chứng từ (.pdf, .png, .jpg, .tiff)"
        accept=".pdf,.png,.jpg,.jpeg,.tiff,.bmp,.webp"
        selectedFile={scanFile}
        onFileSelected={(f) => {
          setScanFile(f);
          setOutputSession(null);
        }}
        onClear={() => setScanFile(null)}
        sourceLabel="OCR Tiếng Việt"
      />

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Định dạng xuất</label>
          <select
            value={outFormat}
            onChange={(e) => setOutFormat(e.target.value as 'docx' | 'md')}
            className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
          >
            <option value="docx">Word có thể chỉnh sửa (.docx)</option>
            <option value="md">Markdown thuần (.md)</option>
          </select>
        </div>
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Ngôn ngữ nhận dạng</label>
          <select
            value={lang}
            onChange={(e) => setLang(e.target.value)}
            className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
          >
            <option value="vie+eng">Tiếng Việt + Tiếng Anh (Khuyên dùng)</option>
            <option value="vie">Chỉ Tiếng Việt</option>
            <option value="eng">Chỉ Tiếng Anh</option>
          </select>
        </div>
      </div>

      <CollapsibleSection title="Bộ lọc tiền xử lý ảnh Scan" defaultOpen>
        <div className="space-y-2 text-xs text-slate-700 dark:text-slate-300">
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              checked={filterRedStamp}
              onChange={(e) => setFilterRedStamp(e.target.checked)}
            />
            <span>Khử mực dấu mộc đỏ HSV đè lên chữ ký (tránh sai chính tả)</span>
          </label>
          <label className="flex items-center gap-2">
            <input type="checkbox" checked={deskew} onChange={(e) => setDeskew(e.target.checked)} />
            <span>Tự động xoay thẳng trang bị nghiêng khi scan (Deskew)</span>
          </label>
        </div>
      </CollapsibleSection>

      <Button
        variant="primary"
        loading={running}
        disabled={!scanFile}
        icon={<ScanText className="h-4 w-4" />}
        onClick={handleRunOcr}
        className="w-full bg-indigo-600 hover:bg-indigo-700"
      >
        Chạy OCR Bóc tách chữ Tiếng Việt
      </Button>
    </div>
  );

  return (
    <SplitWorkspace
      title="OCR Tiếng Việt & Khôi Phục Văn Bản Scan"
      subtitle="Bóc tách văn bản từ PDF scan hoặc ảnh chụp giấy tờ, tự động xoay thẳng trang và khử dấu mộc đỏ đè lên chữ."
      badgeText="Tesseract HSV"
      accent="studio"
      leftPanel={leftPanel}
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Nhận dạng OCR hoàn tất"
            outputFile={outputSession}
            summaryItems={[
              ...(stats.pages !== undefined ? [{ label: 'Số trang', value: stats.pages }] : []),
              ...(stats.chars !== undefined ? [{ label: 'Số ký tự', value: stats.chars.toLocaleString('vi-VN') }] : []),
              ...(stats.elapsed !== undefined ? [{ label: 'Thời gian', value: `${stats.elapsed.toFixed(1)}s` }] : []),
            ]}
          />
        ) : undefined
      }
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="indigo">Xử lý ảnh 100% Offline</Badge>}
    />
  );
}
