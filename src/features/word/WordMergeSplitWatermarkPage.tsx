import { useEffect, useState } from 'react';
import { ArrowDown, ArrowUp, Layers, Plus, Scissors, Stamp, Trash2 } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
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

export function WordMergeSplitWatermarkPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [mode, setMode] = useState<'merge' | 'split' | 'watermark'>('merge');

  // Merge state
  const [mergeList, setMergeList] = useState<SelectedFileInfo[]>([]);
  const [addPageBreaks, setAddPageBreaks] = useState(true);
  const [normalizePages, setNormalizePages] = useState(true);

  // Split & Watermark single file
  const [singleFile, setSingleFile] = useState<SelectedFileInfo | null>(null);
  const [headingLevel, setHeadingLevel] = useState(1);

  // Watermark & Header/Footer
  const [watermarkText, setWatermarkText] = useState('TÀI LIỆU NỘI BỘ');
  const [headerText, setHeaderText] = useState('');
  const [footerText, setFooterText] = useState('');
  const [wmPageNumbers, setWmPageNumbers] = useState(true);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && hf.ext.toLowerCase() === 'docx') {
      const sel = sessionFileToSelected(hf);
      setSingleFile(sel);
      setMergeList((prev) => (prev.some((x) => x.filePath === sel.filePath) ? prev : [...prev, sel]));
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
          sourceLabel: 'Gộp Word',
        });
        return sessionFileToSelected(sf);
      });
      setMergeList((prev) => [...prev, ...added]);
      toast.success(`Đã thêm ${added.length} tệp Word vào danh sách ghép`);
    } catch (err) {
      toast.error('Lỗi tải tệp', err instanceof Error ? err.message : String(err));
    }
  };

  const moveItem = (idx: number, dir: -1 | 1) => {
    setMergeList((prev) => {
      const target = idx + dir;
      if (target < 0 || target >= prev.length) return prev;
      const copy = [...prev];
      const [item] = copy.splice(idx, 1);
      copy.splice(target, 0, item);
      return copy;
    });
  };

  const handleExecute = async () => {
    setRunning(true);
    try {
      if (mode === 'merge') {
        if (mergeList.length < 2) {
          toast.warning('Cần ít nhất 2 tệp .docx để gộp');
          setRunning(false);
          return;
        }
        const res = await apiPost<{ filename: string; download_url: string; output_path?: string }>(
          '/api/docx/merge',
          {
            inputs: mergeList.map((f) => f.filePath),
            add_page_breaks: addPageBreaks,
            normalize_page_numbers: normalizePages,
          }
        );
        const sf = addSessionFile({
          filePath: res.output_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: 'Gộp Word',
        });
        setOutputSession(sf);
        toast.success('Gộp Word thành công', res.filename);
      } else if (mode === 'split') {
        if (!singleFile) {
          toast.warning('Vui lòng chọn tệp Word để tách chương');
          setRunning(false);
          return;
        }
        const res = await apiPost<{ filename: string; download_url: string; output_path?: string }>(
          '/api/docx/split',
          {
            file_path: singleFile.filePath,
            heading_level: headingLevel,
          }
        );
        const sf = addSessionFile({
          filePath: res.output_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: 'Tách Word ZIP',
        });
        setOutputSession(sf);
        toast.success('Đã tách chương ra gói ZIP', res.filename);
      } else {
        if (!singleFile) {
          toast.warning('Vui lòng chọn tệp Word để đóng dấu');
          setRunning(false);
          return;
        }
        const res = await apiPost<{ filename: string; download_url: string; output_path?: string }>(
          '/api/docx/watermark-hf',
          {
            file_path: singleFile.filePath,
            watermark_text: watermarkText,
            header_text: headerText,
            footer_text: footerText,
            add_page_numbers: wmPageNumbers,
          }
        );
        const sf = addSessionFile({
          filePath: res.output_path || res.filename,
          filename: res.filename,
          previewUrl: buildPreviewUrl(res.filename),
          downloadUrl: res.download_url || buildDownloadUrl(res.filename),
          sourceLabel: 'Đóng dấu Word',
        });
        setOutputSession(sf);
        toast.success('Đã đóng dấu & Header/Footer', res.filename);
      }
    } catch (err) {
      toast.error('Thao tác thất bại', err instanceof Error ? err.message : String(err));
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
    : mode === 'merge' && mergeList[0]
    ? {
        filePath: mergeList[0].filePath,
        filename: mergeList[0].filename,
        ext: 'docx',
        previewUrl: mergeList[0].previewUrl,
        downloadUrl: mergeList[0].downloadUrl,
      }
    : singleFile
    ? {
        filePath: singleFile.filePath,
        filename: singleFile.filename,
        ext: 'docx',
        previewUrl: singleFile.previewUrl,
        downloadUrl: singleFile.downloadUrl,
      }
    : null;

  return (
    <SplitWorkspace
      title="Gộp, Tách Chương & Đóng dấu Word"
      subtitle="Ghép nối nhiều file DOCX bảo toàn bảng và ảnh, tách theo cấp Heading ra tệp ZIP, hoặc đóng dấu chéo VML Watermark."
      badgeText="Word Studio"
      accent="word"
      previewTarget={previewTarget}
      leftPanel={
        <div className="space-y-4">
          <Tabs
            activeId={mode}
            onChange={(id) => {
              setMode(id as typeof mode);
              setOutputSession(null);
            }}
            options={[
              { id: 'merge', label: 'Gộp nhiều file', icon: <Layers className="h-3.5 w-3.5" /> },
              { id: 'split', label: 'Tách theo Heading', icon: <Scissors className="h-3.5 w-3.5" /> },
              { id: 'watermark', label: 'Watermark & H/F', icon: <Stamp className="h-3.5 w-3.5" /> },
            ]}
          />

          {mode === 'merge' ? (
            <div className="space-y-3">
              <label className="flex cursor-pointer items-center justify-center gap-2 rounded-lg border-2 border-dashed border-slate-200 bg-slate-50 px-4 py-4 text-xs font-medium text-slate-700 hover:border-blue-300 dark:border-slate-700 dark:bg-slate-900/50 dark:text-slate-200">
                <Plus className="h-4 w-4 text-blue-600" />
                <span>Chọn thêm nhiều tệp Word (.docx) để ghép thứ tự</span>
                <input
                  type="file"
                  multiple
                  accept=".docx"
                  className="hidden"
                  onChange={(e) => {
                    void handleMultiUpload(e.target.files);
                    e.target.value = '';
                  }}
                />
              </label>

              {mergeList.length > 0 ? (
                <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
                  {mergeList.map((item, idx) => (
                    <div
                      key={`${item.filePath}_${idx}`}
                      className="flex items-center justify-between gap-2 rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-2 text-xs dark:border-slate-800 dark:bg-slate-950"
                    >
                      <span className="font-mono text-[11px] font-semibold text-blue-600">#{idx + 1}</span>
                      <span className="flex-1 truncate font-medium text-slate-800 dark:text-slate-100">
                        {item.filename}
                      </span>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => moveItem(idx, -1)}
                          className="rounded p-1 text-slate-400 hover:bg-white hover:text-slate-700 dark:hover:bg-slate-800"
                        >
                          <ArrowUp className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          onClick={() => moveItem(idx, 1)}
                          className="rounded p-1 text-slate-400 hover:bg-white hover:text-slate-700 dark:hover:bg-slate-800"
                        >
                          <ArrowDown className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setMergeList((p) => p.filter((_, i) => i !== idx))}
                          className="rounded p-1 text-slate-400 hover:bg-white hover:text-red-600 dark:hover:bg-slate-800"
                        >
                          <Trash2 className="h-3 w-3" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              ) : null}

              <CollapsibleSection title="Tùy chọn nâng cao">
                <div className="space-y-2 text-xs">
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={addPageBreaks}
                      onChange={(e) => setAddPageBreaks(e.target.checked)}
                    />
                    <span>Chèn ngắt trang (Page Break) giữa từng tài liệu</span>
                  </label>
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={normalizePages}
                      onChange={(e) => setNormalizePages(e.target.checked)}
                    />
                    <span>Chuẩn hóa lại số trang Page X / Y liên tục</span>
                  </label>
                </div>
              </CollapsibleSection>
            </div>
          ) : null}

          {mode === 'split' ? (
            <div className="space-y-3">
              <FileDropZone
                label="Tệp Word tổng hợp cần tách (.docx)"
                accept=".docx"
                selectedFile={singleFile}
                onFileSelected={(f) => {
                  setSingleFile(f);
                  setOutputSession(null);
                }}
                onClear={() => setSingleFile(null)}
                sourceLabel="Tách Word"
              />
              <div className="space-y-1.5">
                <label className="text-xs font-medium text-slate-700 dark:text-slate-300">
                  Cấp Heading làm mốc tách chương
                </label>
                <select
                  value={headingLevel}
                  onChange={(e) => setHeadingLevel(Number(e.target.value))}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-800"
                >
                  <option value={1}>Heading 1 (Tách theo Chương lớn)</option>
                  <option value={2}>Heading 2 (Tách theo Mục con)</option>
                  <option value={3}>Heading 3</option>
                </select>
              </div>
            </div>
          ) : null}

          {mode === 'watermark' ? (
            <div className="space-y-3">
              <FileDropZone
                label="Tệp Word cần đóng dấu (.docx)"
                accept=".docx"
                selectedFile={singleFile}
                onFileSelected={(f) => {
                  setSingleFile(f);
                  setOutputSession(null);
                }}
                onClear={() => setSingleFile(null)}
                sourceLabel="Watermark Word"
              />
              <div className="space-y-1.5">
                <label className="text-xs font-medium text-slate-700 dark:text-slate-300">
                  Chữ mờ chéo (Watermark)
                </label>
                <input
                  type="text"
                  value={watermarkText}
                  onChange={(e) => setWatermarkText(e.target.value)}
                  placeholder="Vd: MẬT / DỰ THẢO / BẢN CHÍNH THỨC"
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-800"
                />
              </div>
              <CollapsibleSection title="Tùy chọn nâng cao (Header / Footer / Số trang)">
                <div className="space-y-2.5 text-xs">
                  <input
                    type="text"
                    value={headerText}
                    onChange={(e) => setHeaderText(e.target.value)}
                    placeholder="Nội dung Header (đầu trang)..."
                    className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
                  />
                  <input
                    type="text"
                    value={footerText}
                    onChange={(e) => setFooterText(e.target.value)}
                    placeholder="Nội dung Footer (chân trang)..."
                    className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
                  />
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={wmPageNumbers}
                      onChange={(e) => setWmPageNumbers(e.target.checked)}
                    />
                    <span>Tự động gắn Page X / Y ở Footer</span>
                  </label>
                </div>
              </CollapsibleSection>
            </div>
          ) : null}

          <Button accent="word" loading={running} onClick={handleExecute} className="w-full">
            {mode === 'merge'
              ? `Gộp ${mergeList.length} tệp Word`
              : mode === 'split'
              ? 'Tách chương thành gói ZIP'
              : 'Áp dụng Watermark & Thương hiệu'}
          </Button>
        </div>
      }
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Hoàn tất xử lý tài liệu Word"
            outputFile={outputSession}
          />
        ) : undefined
      }
    />
  );
}
