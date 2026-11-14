import { useEffect, useState } from 'react';
import { ChevronLeft, ChevronRight, FileCode2, Images, Play, Presentation } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  EmptyState,
  FileDropZone,
  Modal,
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

interface SlideSpecItem {
  title?: string;
  subtitle?: string;
  layout?: string;
  bullets?: string[];
  notes?: string;
}

interface BuildPptxResponse {
  success: boolean;
  pptx_path?: string;
  filename: string;
  download_url?: string;
  slides_count?: number;
  spec?: { slides?: SlideSpecItem[] };
  elapsed?: number;
}

interface PptxPreviewItem {
  slide?: number;
  preview_url: string;
}

const THEMES = [
  { id: 'corporate_blue', label: 'Xanh Doanh nghiệp (Corporate Blue)' },
  { id: 'thesis_blue', label: 'Xanh Học thuật (Thesis Blue)' },
  { id: 'academic_light', label: 'Sáng Tối giản (Academic Light)' },
  { id: 'minimalist_nordic', label: 'Bắc Âu Thanh lịch (Minimalist Nordic)' },
  { id: 'modern_dark', label: 'Tối Hiện đại (Modern Dark)' },
  { id: 'cyber_tech_dark', label: 'Công nghệ Tối (Cyber Tech)' },
];

const SAMPLE_SLIDE_SPEC = JSON.stringify(
  {
    title: 'BÁO CÁO TỔNG KẾT CÔNG TÁC VĂN PHÒNG',
    subtitle: 'Trình bày tóm tắt các chỉ số vận hành & kế hoạch triển khai',
    slides: [
      {
        layout: 'title_only',
        title: 'BÁO CÁO TỔNG KẾT HOẠT ĐỘNG QUÝ IV',
        subtitle: 'Ban Điều Hành Văn Phòng • Tháng 11/2026',
      },
      {
        layout: 'bullets',
        title: 'Trọng tâm 3 Kết quả Nổi bật',
        bullets: [
          'Chuẩn hóa 100% quy trình xử lý công văn theo Nghị định 30/2020/NĐ-CP.',
          'Tự động hóa đối soát báo cáo nhân sự và bảng lương bằng Excel Power Suite.',
          'Áp dụng ký số PDF/A-2b và kiểm định an toàn văn bản nội bộ hoàn toàn Offline.',
        ],
      },
    ],
  },
  null,
  2,
);

export function PresentationStudioPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [mode, setMode] = useState<'from_doc' | 'from_spec'>('from_doc');
  const [sourceDoc, setSourceDoc] = useState<SelectedFileInfo | null>(null);
  const [templatePptx, setTemplatePptx] = useState<SelectedFileInfo | null>(null);
  const [theme, setTheme] = useState('corporate_blue');
  const [deckScope, setDeckScope] = useState('auto');
  const [pdfPages, setPdfPages] = useState('');
  const [specJson, setSpecJson] = useState(SAMPLE_SLIDE_SPEC);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [generatedPptxPath, setGeneratedPptxPath] = useState<string>('');
  const [slideCards, setSlideCards] = useState<SlideSpecItem[]>([]);

  // 1080p previews & zoom modal
  const [loadingPreviews, setLoadingPreviews] = useState(false);
  const [previews, setPreviews] = useState<PptxPreviewItem[]>([]);
  const [zoomIndex, setZoomIndex] = useState<number | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && ['docx', 'pdf'].includes(hf.ext.toLowerCase())) {
      setSourceDoc(sessionFileToSelected(hf));
      setMode('from_doc');
    }
  }, [consumeHandoffFile]);

  const handleGenerateFromDoc = async () => {
    if (!sourceDoc) {
      toast.warning('Chưa chọn Word hoặc PDF nguồn');
      return;
    }
    setRunning(true);
    setPreviews([]);
    try {
      const isPdf = sourceDoc.ext.toLowerCase() === 'pdf';
      const endpoint = isPdf ? '/api/pdf-to-pptx' : '/api/docx-to-pptx';
      const payload: Record<string, unknown> = isPdf
        ? {
            pdf_path: sourceDoc.filePath,
            theme,
            deck_scope: deckScope,
            pages: pdfPages.trim() || undefined,
            template_path: templatePptx?.filePath || undefined,
          }
        : {
            docx_path: sourceDoc.filePath,
            theme,
            deck_scope: deckScope,
            template_path: templatePptx?.filePath || undefined,
          };

      const res = await apiPost<BuildPptxResponse>(endpoint, payload);
      const sf = addSessionFile({
        filePath: res.pptx_path || res.filename,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: 'Slide 16:9 PPTX',
      });
      setOutputSession(sf);
      setGeneratedPptxPath(res.pptx_path || res.filename);
      setSlideCards(res.spec?.slides || []);
      toast.success('Đã tạo bài thuyết trình 16:9', `${res.slides_count ?? res.spec?.slides?.length ?? 0} slides`);
    } catch (err) {
      toast.error('Lỗi tạo Slide PowerPoint', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleBuildFromSpec = async () => {
    let parsed: { slides?: SlideSpecItem[] };
    try {
      parsed = JSON.parse(specJson);
    } catch {
      toast.error('JSON Spec không hợp lệ');
      return;
    }
    setRunning(true);
    setPreviews([]);
    try {
      const res = await apiPost<BuildPptxResponse>('/api/pptx/build', {
        spec: parsed,
        theme,
        template_path: templatePptx?.filePath || undefined,
      });
      const sf = addSessionFile({
        filePath: res.pptx_path || res.filename,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: 'Slide từ JSON',
      });
      setOutputSession(sf);
      setGeneratedPptxPath(res.pptx_path || res.filename);
      setSlideCards(parsed.slides || []);
      toast.success('Đã kết xuất PPTX từ Spec JSON', res.filename);
    } catch (err) {
      toast.error('Lỗi kết xuất PPTX', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleExportSlideImages = async () => {
    const pathArg = generatedPptxPath || outputSession?.filePath;
    if (!pathArg) return;
    setLoadingPreviews(true);
    try {
      const res = await apiPost<{ success: boolean; previews?: PptxPreviewItem[] }>('/api/pptx/preview', {
        pptx_path: pathArg,
        slides: '1-16',
      });
      setPreviews(res.previews || []);
      toast.success(`Đã xuất ${res.previews?.length ?? 0} ảnh slide 1080p`);
    } catch (err) {
      toast.error('Không xuất được ảnh xem trước', err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingPreviews(false);
    }
  };

  const leftPanel = (
    <div className="space-y-4">
      <Tabs
        activeId={mode}
        onChange={(id) => setMode(id as 'from_doc' | 'from_spec')}
        className="w-full"
        options={[
          { id: 'from_doc', label: 'Word / PDF -> Slide 16:9', icon: <Presentation className="h-3.5 w-3.5" /> },
          { id: 'from_spec', label: 'Soạn nhanh từ JSON Spec', icon: <FileCode2 className="h-3.5 w-3.5" /> },
        ]}
      />

      {mode === 'from_doc' ? (
        <FileDropZone
          label="Tài liệu Word (.docx) hoặc PDF cần chuyển thành Slide thuyết trình"
          accept=".docx,.pdf"
          selectedFile={sourceDoc}
          onFileSelected={(f) => {
            setSourceDoc(f);
            setOutputSession(null);
            setPreviews([]);
          }}
          onClear={() => setSourceDoc(null)}
          sourceLabel="Slide Studio"
        />
      ) : (
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
            Cấu trúc JSON Slide (Tiêu đề, Bullet, Bảng biểu)
          </label>
          <textarea
            rows={9}
            value={specJson}
            onChange={(e) => setSpecJson(e.target.value)}
            className="w-full rounded-lg border border-slate-200 bg-white p-2.5 font-mono text-[11px] dark:border-slate-700 dark:bg-slate-950"
          />
        </div>
      )}

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Giao diện Slide (Theme)</label>
          <select
            value={theme}
            onChange={(e) => setTheme(e.target.value)}
            className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
          >
            {THEMES.map((t) => (
              <option key={t.id} value={t.id}>
                {t.label}
              </option>
            ))}
          </select>
        </div>
        {mode === 'from_doc' ? (
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Mật độ cô đọng nội dung</label>
            <select
              value={deckScope}
              onChange={(e) => setDeckScope(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            >
              <option value="auto">Tự động cân bằng (Khuyên dùng)</option>
              <option value="executive">Tóm tắt Lãnh đạo (Ngắn gọn)</option>
              <option value="detailed">Chi tiết Đầy đủ (Bảo toàn bảng)</option>
            </select>
          </div>
        ) : null}
      </div>

      {mode === 'from_doc' && sourceDoc?.ext.toLowerCase() === 'pdf' ? (
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
            Giới hạn trang PDF (tùy chọn, VD: 1-8)
          </label>
          <input
            type="text"
            value={pdfPages}
            onChange={(e) => setPdfPages(e.target.value)}
            placeholder="Để trống nếu lấy toàn bộ"
            className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
          />
        </div>
      ) : null}

      <CollapsibleSection title="Kế thừa Slide Master công ty (.pptx mẫu - Tùy chọn)">
        <FileDropZone
          label="Tải lên tệp .pptx mẫu có sẵn Logo & Font thương hiệu"
          accept=".pptx,.potx"
          selectedFile={templatePptx}
          onFileSelected={setTemplatePptx}
          onClear={() => setTemplatePptx(null)}
          sourceLabel="Slide Master"
        />
      </CollapsibleSection>

      <Button
        variant="primary"
        loading={running}
        disabled={mode === 'from_doc' && !sourceDoc}
        icon={<Play className="h-4 w-4" />}
        onClick={mode === 'from_doc' ? handleGenerateFromDoc : handleBuildFromSpec}
        className="w-full bg-orange-600 hover:bg-orange-700"
      >
        Tổng hợp Bài Thuyết Trình 16:9 (.pptx)
      </Button>
    </div>
  );

  const rightPanel = (
    <div className="space-y-4">
      {outputSession ? (
        <ResultPreviewCard
          title="Đã xuất bản PowerPoint 16:9"
          outputFile={outputSession}
          summaryItems={[
            { label: 'Số slide', value: slideCards.length || '-' },
            { label: 'Theme', value: theme },
          ]}
          extraActions={
            <Button
              size="sm"
              variant="outline"
              loading={loadingPreviews}
              icon={<Images className="h-3.5 w-3.5" />}
              onClick={handleExportSlideImages}
            >
              Xem thư viện ảnh Slide 1080p
            </Button>
          }
        />
      ) : null}

      {previews.length > 0 ? (
        <div className="space-y-2 rounded-xl border border-slate-200 bg-white p-3.5 dark:border-slate-800 dark:bg-slate-900">
          <div className="flex items-center justify-between">
            <h4 className="text-xs font-semibold text-slate-800 dark:text-slate-200">
              Thư viện ảnh Slide 1080p ({previews.length} trang)
            </h4>
            <Badge tone="orange">Nhấn vào slide để phóng to</Badge>
          </div>
          <div className="grid grid-cols-2 gap-2.5 sm:grid-cols-3">
            {previews.map((p, idx) => (
              <button
                key={idx}
                type="button"
                onClick={() => setZoomIndex(idx)}
                className="group overflow-hidden rounded-lg border border-slate-200 bg-slate-50 text-left transition hover:border-orange-400 dark:border-slate-800 dark:bg-slate-950"
              >
                <div className="bg-slate-900/5 px-2 py-1 text-[10px] font-semibold text-slate-600 dark:text-slate-400">
                  SLIDE {String(p.slide || idx + 1).padStart(2, '0')}
                </div>
                <img
                  src={`${p.preview_url}?t=${idx}`}
                  alt={`Slide ${p.slide || idx + 1}`}
                  className="aspect-video w-full object-cover transition group-hover:scale-[1.02]"
                />
              </button>
            ))}
          </div>
        </div>
      ) : null}

      {slideCards.length > 0 ? (
        <div className="space-y-2 rounded-xl border border-slate-200 bg-white p-3.5 dark:border-slate-800 dark:bg-slate-900">
          <div className="flex items-center justify-between">
            <h4 className="text-xs font-semibold text-slate-800 dark:text-slate-200">
              Dàn ý & Bố cục {slideCards.length} Slides đã tạo
            </h4>
            <Badge tone="slate">16:9 Zero-Overflow</Badge>
          </div>
          <div className="max-h-80 space-y-2 overflow-y-auto pr-1">
            {slideCards.map((s, idx) => (
              <div
                key={idx}
                className="rounded-lg border border-slate-100 bg-slate-50/80 px-3 py-2.5 dark:border-slate-800 dark:bg-slate-950/70"
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] font-bold text-orange-700 dark:text-orange-400">
                    SLIDE {String(idx + 1).padStart(2, '0')}
                  </span>
                  <Badge tone="slate">{s.layout || 'content'}</Badge>
                </div>
                <div className="mt-1 text-xs font-semibold text-slate-900 dark:text-slate-100">
                  {s.title || 'Không tiêu đề'}
                </div>
                {s.bullets && s.bullets.length > 0 ? (
                  <ul className="mt-1 list-disc space-y-0.5 pl-4 text-[11px] text-slate-600 dark:text-slate-400">
                    {s.bullets.slice(0, 3).map((b, bi) => (
                      <li key={bi}>{b}</li>
                    ))}
                  </ul>
                ) : null}
              </div>
            ))}
          </div>
        </div>
      ) : !outputSession ? (
        <EmptyState
          icon={<Presentation className="h-6 w-6" />}
          title="Sẵn sàng thiết kế Slide 16:9 tự động"
          description="Nạp file Word hoặc PDF bên trái để AI Engine tự tách mục, dựng bố cục thẻ và xuất PowerPoint 16:9 sắc nét."
        />
      ) : null}

      <Modal
        open={zoomIndex !== null && previews[zoomIndex] !== undefined}
        onClose={() => setZoomIndex(null)}
        title={zoomIndex !== null ? `Slide ${zoomIndex + 1} / ${previews.length}` : ''}
        size="xl"
      >
        {zoomIndex !== null && previews[zoomIndex] ? (
          <div className="space-y-3">
            <img
              src={previews[zoomIndex].preview_url}
              alt={`Slide ${zoomIndex + 1}`}
              className="w-full rounded-lg border border-slate-200 dark:border-slate-800"
            />
            <div className="flex items-center justify-between">
              <Button
                size="sm"
                variant="outline"
                disabled={zoomIndex === 0}
                icon={<ChevronLeft className="h-4 w-4" />}
                onClick={() => setZoomIndex((prev) => (prev !== null && prev > 0 ? prev - 1 : prev))}
              >
                Slide trước
              </Button>
              <Button
                size="sm"
                variant="outline"
                disabled={zoomIndex >= previews.length - 1}
                icon={<ChevronRight className="h-4 w-4" />}
                onClick={() => setZoomIndex((prev) => (prev !== null && prev < previews.length - 1 ? prev + 1 : prev))}
              >
                Slide kế tiếp
              </Button>
            </div>
          </div>
        ) : null}
      </Modal>
    </div>
  );

  return (
    <SplitWorkspace
      title="Presentation Studio — Kết Xuất Slide 16:9"
      subtitle="Tổng hợp bài thuyết trình PowerPoint chuẩn 16:9 tự động từ tài liệu Word/PDF hoặc Spec JSON, hỗ trợ kế thừa Slide Master và xem trước 1080p."
      badgeText="Widescreen 16:9"
      accent="slide"
      leftPanel={leftPanel}
      resultPanel={rightPanel}
      rightHeaderExtra={<Badge tone="orange">Zero Text Overflow</Badge>}
    />
  );
}
