import { useEffect, useState } from 'react';
import { Sparkles } from 'lucide-react';
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

interface PolishResponse {
  success: boolean;
  filename: string;
  download_url: string;
  output_path?: string;
  stats?: Record<string, unknown>;
}

export function WordPolishPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [inputFile, setInputFile] = useState<SelectedFileInfo | null>(null);
  const [preset, setPreset] = useState('nd30');
  const [addPageNumbers, setAddPageNumbers] = useState(true);
  const [sanitizeTypography, setSanitizeTypography] = useState(true);
  const [fixTables, setFixTables] = useState(true);
  const [normalizeCaptions, setNormalizeCaptions] = useState(false);

  const [includeAdminHeader, setIncludeAdminHeader] = useState(false);
  const [agencyParent, setAgencyParent] = useState('UBND TỈNH / CƠ QUAN CHỦ QUẢN');
  const [agencyName, setAgencyName] = useState('VĂN PHÒNG / PHÒNG HÀNH CHÍNH');
  const [documentNo, setDocumentNo] = useState('Số: 01/CV-VP');
  const [locationDate, setLocationDate] = useState('Hà Nội, ngày 14 tháng 11 năm 2026');

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && hf.ext.toLowerCase() === 'docx') {
      setInputFile(sessionFileToSelected(hf));
    }
  }, [consumeHandoffFile]);

  const handleRun = async () => {
    if (!inputFile) {
      toast.warning('Chưa chọn tệp Word', 'Vui lòng tải lên tệp .docx để chuẩn hóa.');
      return;
    }
    setRunning(true);
    try {
      const payload: Record<string, unknown> = {
        file_path: inputFile.filePath,
        preset,
        add_page_numbers: addPageNumbers,
        sanitize_typography: sanitizeTypography,
        fix_tables: fixTables,
        normalize_captions: normalizeCaptions,
      };
      if (includeAdminHeader) {
        payload.admin_header = {
          agency_parent: agencyParent,
          agency_name: agencyName,
          document_no: documentNo,
          location_date: locationDate,
        };
      }
      const res = await apiPost<PolishResponse>('/api/docx/polish', payload);
      const outPath = res.output_path || res.filename;
      const sf = addSessionFile({
        filePath: outPath,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: 'Chuẩn hóa Word',
      });
      setOutputSession(sf);
      toast.success('Chuẩn hóa Word hoàn tất', res.filename);
    } catch (err) {
      toast.error('Lỗi chuẩn hóa Word', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const previewTarget = outputSession
    ? {
        filePath: outputSession.filePath,
        filename: outputSession.filename,
        ext: 'docx',
        previewUrl: outputSession.previewUrl,
        downloadUrl: outputSession.downloadUrl,
      }
    : inputFile
    ? {
        filePath: inputFile.filePath,
        filename: inputFile.filename,
        ext: 'docx',
        previewUrl: inputFile.previewUrl,
        downloadUrl: inputFile.downloadUrl,
      }
    : null;

  return (
    <SplitWorkspace
      title="Chuẩn hóa Word theo Nghị định 30/2020"
      subtitle="Tự động căn lề chuẩn văn bản hành chính, chống rách bảng qua trang, khống chế ảnh lề in ≤ 14cm và đánh số Page X/Y."
      badgeText="Word Studio"
      accent="word"
      previewTarget={previewTarget}
      leftPanel={
        <div className="space-y-4">
          <FileDropZone
            label="1. Tài liệu Word nguồn (.docx)"
            accept=".docx"
            selectedFile={inputFile}
            onFileSelected={(f) => {
              setInputFile(f);
              setOutputSession(null);
            }}
            onClear={() => {
              setInputFile(null);
              setOutputSession(null);
            }}
            sourceLabel="Word Polish"
            hint="Chấp nhận tệp .docx"
          />

          <div className="space-y-1.5">
            <label className="text-xs font-medium text-slate-700 dark:text-slate-300">
              2. Bộ quy chuẩn trình bày
            </label>
            <select
              value={preset}
              onChange={(e) => setPreset(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs text-slate-800 focus:border-blue-500 focus:outline-none dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100"
            >
              <option value="nd30">Nghị định 30/2020/NĐ-CP (Times New Roman 13–14pt, lề chuẩn)</option>
              <option value="executive">Executive Report (Doanh nghiệp hiện đại)</option>
              <option value="corporate_report">Báo cáo Doanh nghiệp (Corporate Clean)</option>
              <option value="academic_thesis">Báo cáo / Luận văn Học thuật</option>
            </select>
          </div>

          <div className="grid grid-cols-2 gap-2 pt-1 text-xs">
            <label className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 dark:border-slate-800">
              <input
                type="checkbox"
                checked={addPageNumbers}
                onChange={(e) => setAddPageNumbers(e.target.checked)}
              />
              <span>Đánh số Trang X / Y</span>
            </label>
            <label className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 dark:border-slate-800">
              <input
                type="checkbox"
                checked={fixTables}
                onChange={(e) => setFixTables(e.target.checked)}
              />
              <span>Sửa bảng & lặp Header</span>
            </label>
          </div>

          <CollapsibleSection title="Tùy chọn nâng cao (Chính tả & Quốc hiệu)">
            <div className="space-y-3 text-xs">
              <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
                <label className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={sanitizeTypography}
                    onChange={(e) => setSanitizeTypography(e.target.checked)}
                  />
                  <span>Chuẩn hóa Unicode tiếng Việt</span>
                </label>
                <label className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={normalizeCaptions}
                    onChange={(e) => setNormalizeCaptions(e.target.checked)}
                  />
                  <span>Chuẩn hóa chú thích Hình/Bảng</span>
                </label>
              </div>

              <label className="flex items-center gap-2 font-medium text-blue-700 dark:text-blue-400">
                <input
                  type="checkbox"
                  checked={includeAdminHeader}
                  onChange={(e) => setIncludeAdminHeader(e.target.checked)}
                />
                <span>Chèn Quốc hiệu & Tiêu ngữ NĐ 30 ở đầu văn bản</span>
              </label>

              {includeAdminHeader ? (
                <div className="grid grid-cols-1 gap-2 rounded-lg border border-slate-200 bg-white p-2.5 dark:border-slate-700 dark:bg-slate-900">
                  <input
                    type="text"
                    value={agencyParent}
                    onChange={(e) => setAgencyParent(e.target.value)}
                    placeholder="Cơ quan chủ quản"
                    className="rounded border border-slate-200 px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
                  />
                  <input
                    type="text"
                    value={agencyName}
                    onChange={(e) => setAgencyName(e.target.value)}
                    placeholder="Cơ quan ban hành"
                    className="rounded border border-slate-200 px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
                  />
                  <div className="grid grid-cols-2 gap-2">
                    <input
                      type="text"
                      value={documentNo}
                      onChange={(e) => setDocumentNo(e.target.value)}
                      placeholder="Số: 01/CV"
                      className="rounded border border-slate-200 px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
                    />
                    <input
                      type="text"
                      value={locationDate}
                      onChange={(e) => setLocationDate(e.target.value)}
                      placeholder="Địa danh, ngày..."
                      className="rounded border border-slate-200 px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
                    />
                  </div>
                </div>
              ) : null}
            </div>
          </CollapsibleSection>

          <Button
            accent="word"
            loading={running}
            disabled={!inputFile}
            onClick={handleRun}
            icon={<Sparkles className="h-4 w-4" />}
            className="w-full"
          >
            Chuẩn hóa Văn bản Ngay
          </Button>
        </div>
      }
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Văn bản Word đã được chuẩn hóa"
            outputFile={outputSession}
            summaryItems={[
              { label: 'Quy chuẩn', value: preset.toUpperCase() },
              { label: 'Đánh số trang', value: addPageNumbers ? 'Page X / Y' : 'Không đổi' },
              { label: 'Sửa bảng OpenXML', value: fixTables ? 'cantSplit + tblHeader' : 'Tắt' },
            ]}
          />
        ) : undefined
      }
    />
  );
}
