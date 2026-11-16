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
  writing_style_report?: {
    style_name?: string;
    edits_count?: number;
    suitability_score?: number;
  };
}

interface QuickStyleResponse {
  success: boolean;
  style_name?: string;
  suitability_score?: number;
  refined_text?: string;
  changes_count?: number;
  findings?: Array<{ matched_text: string; suggestion: string; message: string }>;
  recommendations?: string[];
  analysis_after?: { suitability_score?: number };
}

export function WordPolishPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [inputFile, setInputFile] = useState<SelectedFileInfo | null>(null);
  const [preset, setPreset] = useState('nd30');
  const [writingStyle, setWritingStyle] = useState('');
  const [writingStyleMode, setWritingStyleMode] = useState<'safe' | 'strict'>('safe');
  const [addPageNumbers, setAddPageNumbers] = useState(true);
  const [sanitizeTypography, setSanitizeTypography] = useState(true);
  const [fixTables, setFixTables] = useState(true);
  const [normalizeCaptions, setNormalizeCaptions] = useState(false);

  const [sampleText, setSampleText] = useState('');
  const [sampleStyle, setSampleStyle] = useState('hanh_chinh');
  const [sampleResult, setSampleResult] = useState<QuickStyleResponse | null>(null);
  const [sampleLoading, setSampleLoading] = useState(false);

  const [includeAdminHeader, setIncludeAdminHeader] = useState(false);
  const [agencyParent, setAgencyParent] = useState('UBND TỈNH / CƠ QUAN CHỦ QUẢN');
  const [agencyName, setAgencyName] = useState('VĂN PHÒNG / PHÒNG HÀNH CHÍNH');
  const [documentNo, setDocumentNo] = useState('Số: 01/CV-VP');
  const [locationDate, setLocationDate] = useState('Hà Nội, ngày 14 tháng 11 năm 2026');

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [lastStyleReport, setLastStyleReport] = useState<PolishResponse['writing_style_report'] | null>(null);

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
      if (writingStyle) {
        payload.writing_style = writingStyle;
        payload.writing_style_mode = writingStyleMode;
      }
      if (includeAdminHeader) {
        payload.admin_header = {
          agency_parent: agencyParent,
          agency_name: agencyName,
          document_no: documentNo,
          location_date: locationDate,
        };
      }
      const res = await apiPost<PolishResponse>('/api/docx/polish', payload);
      setLastStyleReport(res.writing_style_report ?? null);
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

  const handleQuickStyle = async (action: 'analyze' | 'refine') => {
    if (!sampleText.trim()) {
      toast.warning('Chưa có văn bản', 'Vui lòng dán đoạn văn bản cần kiểm tra hoặc chuẩn hóa hành văn.');
      return;
    }
    setSampleLoading(true);
    try {
      const endpoint =
        action === 'analyze'
          ? '/api/docx/writing-style/analyze'
          : '/api/docx/writing-style/refine';
      const res = await apiPost<QuickStyleResponse>(endpoint, {
        text: sampleText,
        style_id: sampleStyle,
        mode: writingStyleMode,
      });
      setSampleResult(res);
      if (action === 'refine' && res.refined_text) {
        setSampleText(res.refined_text);
      }
    } catch (err) {
      toast.error('Lỗi xử lý văn phong', err instanceof Error ? err.message : String(err));
    } finally {
      setSampleLoading(false);
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

          <div className="space-y-1.5">
            <label className="text-xs font-medium text-slate-700 dark:text-slate-300">
              3. Chuẩn hóa văn phong / hành văn (Tùy chọn)
            </label>
            <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
              <select
                value={writingStyle}
                onChange={(e) => setWritingStyle(e.target.value)}
                className="sm:col-span-2 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs text-slate-800 focus:border-blue-500 focus:outline-none dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100"
              >
                <option value="">Giữ nguyên câu chữ (Mặc định)</option>
                <option value="hanh_chinh">Hành chính Công vụ (NĐ 30/2020)</option>
                <option value="hoc_thuat">Học thuật &amp; Nghiên cứu</option>
                <option value="doanh_nghiep">Doanh nghiệp &amp; Điều hành (BLUF)</option>
                <option value="phap_ly">Pháp lý &amp; Hợp đồng</option>
                <option value="ky_thuat">Đặc tả Kỹ thuật (SRS / Docs)</option>
              </select>
              <select
                value={writingStyleMode}
                disabled={!writingStyle}
                onChange={(e) => setWritingStyleMode(e.target.value as 'safe' | 'strict')}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-800 disabled:opacity-50 focus:border-blue-500 focus:outline-none dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100"
              >
                <option value="safe">An toàn (Safe)</option>
                <option value="strict">Chặt chẽ (Strict)</option>
              </select>
            </div>
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

          <CollapsibleSection title="Kiểm tra & Chuẩn hóa nhanh văn phong đoạn văn">
            <div className="space-y-2.5 text-xs">
              <div className="flex flex-wrap items-center gap-2">
                <select
                  value={sampleStyle}
                  onChange={(e) => setSampleStyle(e.target.value)}
                  className="flex-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
                >
                  <option value="hanh_chinh">Hành chính Công vụ (NĐ30)</option>
                  <option value="hoc_thuat">Học thuật &amp; Nghiên cứu</option>
                  <option value="doanh_nghiep">Doanh nghiệp &amp; Điều hành</option>
                  <option value="phap_ly">Pháp lý &amp; Hợp đồng</option>
                  <option value="ky_thuat">Đặc tả Kỹ thuật</option>
                </select>
                <Button
                  size="sm"
                  variant="secondary"
                  loading={sampleLoading}
                  onClick={() => handleQuickStyle('analyze')}
                >
                  Phân tích
                </Button>
                <Button
                  size="sm"
                  accent="word"
                  loading={sampleLoading}
                  onClick={() => handleQuickStyle('refine')}
                >
                  Chuẩn hóa
                </Button>
              </div>
              <textarea
                rows={4}
                value={sampleText}
                onChange={(e) => setSampleText(e.target.value)}
                placeholder="Dán đoạn văn bản tiếng Việt để kiểm tra khẩu ngữ, viết tắt, độ dài câu hoặc chuẩn hóa hành văn..."
                className="w-full rounded-lg border border-slate-200 bg-white p-2.5 text-xs text-slate-800 focus:border-blue-500 focus:outline-none dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100"
              />
              {sampleResult ? (
                <div className="rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-[11px] dark:border-slate-700 dark:bg-slate-900">
                  <div className="font-medium text-slate-800 dark:text-slate-200">
                    Điểm phù hợp: {sampleResult.analysis_after?.suitability_score ?? sampleResult.suitability_score ?? 100}/100
                    {typeof sampleResult.changes_count === 'number' ? ` · Đã chỉnh ${sampleResult.changes_count} mục` : ''}
                  </div>
                  {sampleResult.recommendations?.length ? (
                    <ul className="mt-1 list-disc space-y-0.5 pl-4 text-slate-600 dark:text-slate-400">
                      {sampleResult.recommendations.slice(0, 3).map((rec, i) => (
                        <li key={i}>{rec}</li>
                      ))}
                    </ul>
                  ) : null}
                </div>
              ) : null}
            </div>
          </CollapsibleSection>
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
              ...(lastStyleReport
                ? [
                    {
                      label: 'Văn phong',
                      value: `${lastStyleReport.style_name || writingStyle} (${lastStyleReport.edits_count ?? 0} sửa · ${lastStyleReport.suitability_score ?? 100}/100)`,
                    },
                  ]
                : []),
            ]}
          />
        ) : undefined
      }
    />
  );
}
