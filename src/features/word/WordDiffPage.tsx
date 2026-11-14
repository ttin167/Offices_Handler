import { useEffect, useState } from 'react';
import { FileDiff } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
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

export function WordDiffPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [docA, setDocA] = useState<SelectedFileInfo | null>(null);
  const [docB, setDocB] = useState<SelectedFileInfo | null>(null);
  const [author, setAuthor] = useState('Phòng Pháp chế / Văn thư');
  const [outputName, setOutputName] = useState('');

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [diffSummary, setDiffSummary] = useState<{ insertions?: number; deletions?: number } | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && hf.ext.toLowerCase() === 'docx') {
      setDocA(sessionFileToSelected(hf));
    }
  }, [consumeHandoffFile]);

  const handleDiff = async () => {
    if (!docA || !docB) {
      toast.warning('Cần đủ 2 phiên bản', 'Vui lòng chọn cả Bản gốc (v1) và Bản chỉnh sửa (v2).');
      return;
    }
    setRunning(true);
    try {
      const res = await apiPost<{
        filename: string;
        download_url: string;
        output_path?: string;
        insertions?: number;
        deletions?: number;
        stats?: { insertions?: number; deletions?: number };
      }>('/api/docx/diff', {
        doc_a: docA.filePath,
        doc_b: docB.filePath,
        author,
        output_name: outputName.trim() || undefined,
      });
      const sf = addSessionFile({
        filePath: res.output_path || res.filename,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: 'Redline Diff',
      });
      setOutputSession(sf);
      setDiffSummary({
        insertions: res.insertions ?? res.stats?.insertions,
        deletions: res.deletions ?? res.stats?.deletions,
      });
      toast.success('Đã tạo văn bản so sánh Track Changes', res.filename);
    } catch (err) {
      toast.error('Lỗi so sánh Word', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  return (
    <SplitWorkspace
      title="So sánh & Track Changes Hai Bản Word"
      subtitle="Đối chiếu từng câu, đoạn văn và bảng biểu giữa bản gốc và bản sửa đổi, xuất file Redline mở trực tiếp trên MS Word."
      badgeText="Redline Diff"
      accent="word"
      previewTarget={
        outputSession
          ? {
              filePath: outputSession.filePath,
              filename: outputSession.filename,
              ext: 'docx',
              previewUrl: outputSession.previewUrl,
              downloadUrl: outputSession.downloadUrl,
            }
          : docB || docA
      }
      leftPanel={
        <div className="space-y-4">
          <FileDropZone
            label="1. Phiên bản gốc (v1)"
            accept=".docx"
            selectedFile={docA}
            onFileSelected={(f) => setDocA(f)}
            onClear={() => setDocA(null)}
            sourceLabel="Word v1"
          />
          <FileDropZone
            label="2. Phiên bản chỉnh sửa (v2)"
            accept=".docx"
            selectedFile={docB}
            onFileSelected={(f) => setDocB(f)}
            onClear={() => setDocB(null)}
            sourceLabel="Word v2"
          />

          <CollapsibleSection title="Tùy chọn nâng cao (Tác giả Track Changes)">
            <div className="space-y-2 text-xs">
              <input
                type="text"
                value={author}
                onChange={(e) => setAuthor(e.target.value)}
                placeholder="Tên người rà soát hiển thị trong Word"
                className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
              />
              <input
                type="text"
                value={outputName}
                onChange={(e) => setOutputName(e.target.value)}
                placeholder="Tên file kết quả tùy chọn (_redline.docx)"
                className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
              />
            </div>
          </CollapsibleSection>

          <Button
            accent="word"
            loading={running}
            disabled={!docA || !docB}
            onClick={handleDiff}
            icon={<FileDiff className="h-4 w-4" />}
            className="w-full"
          >
            Tạo File So Sánh Redline (.docx)
          </Button>
        </div>
      }
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Hoàn tất so sánh Track Changes"
            outputFile={outputSession}
            summaryItems={[
              { label: 'Chèn mới (w:ins)', value: diffSummary?.insertions ?? 'Có' },
              { label: 'Xóa bỏ (w:del)', value: diffSummary?.deletions ?? 'Có' },
              { label: 'Người duyệt', value: author },
            ]}
          />
        ) : undefined
      }
    />
  );
}
