import { useState } from 'react';
import { Code2, Download, Network, Play } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  Tabs,
} from '../../components/ui';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import {
  CODE_TO_DIAGRAM_SAMPLE,
  DIAGRAM_PRESETS_LIST,
  type DiagramEngineType,
} from '../../config/diagramPresets';
import { apiPost, buildDownloadUrl, buildPreviewUrl } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

interface RenderDiagramResponse {
  success: boolean;
  filename: string;
  image_path?: string;
  download_url?: string;
  preview_url?: string;
}

interface CodeToDiagramResponse {
  success: boolean;
  png_path?: string;
  drawio_path?: string;
  svg_path?: string;
  png_download_url?: string;
  drawio_download_url?: string;
  svg_download_url?: string;
  mermaid_code?: string;
}

export function DiagramStudioPage() {
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [tab, setTab] = useState<'diagram' | 'from_code'>('diagram');
  const [engine, setEngine] = useState<DiagramEngineType>('mermaid');
  const [code, setCode] = useState<string>(DIAGRAM_PRESETS_LIST[0].code);
  const [scale, setScale] = useState(3);

  // Code-to-diagram
  const [sourceCode, setSourceCode] = useState<string>(CODE_TO_DIAGRAM_SAMPLE);
  const [codeDiagType, setCodeDiagType] = useState('auto');
  const [codeOutputs, setCodeOutputs] = useState<CodeToDiagramResponse | null>(null);

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);

  const handleSelectPreset = (presetId: string) => {
    const found = DIAGRAM_PRESETS_LIST.find((p) => p.id === presetId);
    if (!found) return;
    setEngine(found.engine);
    setCode(found.code);
  };

  const handleRenderDiagram = async () => {
    if (!code.trim()) {
      toast.warning('Vui lòng nhập mã sơ đồ hoặc chọn mẫu có sẵn');
      return;
    }
    setRunning(true);
    setCodeOutputs(null);
    try {
      const res = await apiPost<RenderDiagramResponse>('/api/render-diagram', {
        engine,
        code,
        scale,
      });
      const sf = addSessionFile({
        filePath: res.image_path || res.filename,
        filename: res.filename,
        previewUrl: res.preview_url || buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: `Sơ đồ ${engine.toUpperCase()}`,
      });
      setOutputSession(sf);
      toast.success(`Đã kết xuất sơ đồ ${engine.toUpperCase()} 300+ DPI`, res.filename);
    } catch (err) {
      toast.error('Lỗi kết xuất sơ đồ', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleGenerateFromCode = async () => {
    if (!sourceCode.trim()) {
      toast.warning('Vui lòng dán mã SQL / Prisma / Python');
      return;
    }
    setRunning(true);
    try {
      const res = await apiPost<CodeToDiagramResponse>('/api/diagrams/from-code', {
        source_code: sourceCode,
        diagram_type: codeDiagType,
      });
      setCodeOutputs(res);
      if (res.png_download_url) {
        const pngName = res.png_path?.split(/[/\\]/).pop() || 'code_diagram.png';
        const sf = addSessionFile({
          filePath: res.png_path || pngName,
          filename: pngName,
          previewUrl: buildPreviewUrl(pngName),
          downloadUrl: res.png_download_url,
          sourceLabel: 'AST Code-to-Diagram',
        });
        setOutputSession(sf);
      }
      toast.success('Đã dựng sơ đồ tự động từ mã nguồn', 'Hỗ trợ PNG, Draw.io XML & SVG');
    } catch (err) {
      toast.error('Lỗi phân tích mã nguồn', err instanceof Error ? err.message : String(err));
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
    : null;

  const leftPanel = (
    <div className="space-y-4">
      <Tabs
        activeId={tab}
        onChange={(id) => setTab(id as 'diagram' | 'from_code')}
        className="w-full"
        options={[
          { id: 'diagram', label: 'Vẽ Sơ đồ Đa Engine (4 chuẩn)', icon: <Network className="h-3.5 w-3.5" /> },
          { id: 'from_code', label: 'Sinh Sơ đồ từ Code SQL/Python', icon: <Code2 className="h-3.5 w-3.5" /> },
        ]}
      />

      {tab === 'diagram' ? (
        <div className="space-y-3">
          <div>
            <label className="mb-1.5 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Chọn mẫu nhanh (1-Click nạp chuẩn kỹ thuật)
            </label>
            <div className="flex flex-wrap gap-1.5">
              {DIAGRAM_PRESETS_LIST.map((preset) => (
                <button
                  key={preset.id}
                  type="button"
                  onClick={() => handleSelectPreset(preset.id)}
                  className="rounded-lg border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-700 hover:border-indigo-400 hover:bg-indigo-50/60 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300"
                >
                  {preset.label}
                </button>
              ))}
            </div>
          </div>

          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Engine kết xuất</label>
              <select
                value={engine}
                onChange={(e) => setEngine(e.target.value as DiagramEngineType)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="mermaid">Mermaid (Flowchart / Sequence / ERD / Mindmap)</option>
                <option value="plantuml">PlantUML C4 & UML</option>
                <option value="drawio">Draw.io Native XML (mxGraph)</option>
                <option value="canvas">Canvas Precision JSON Spec</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Độ phân giải xuất</label>
              <select
                value={scale}
                onChange={(e) => setScale(Number(e.target.value))}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value={2}>2x (~200 DPI Tiêu chuẩn)</option>
                <option value={3}>3x (300+ DPI Chuẩn in/Word)</option>
              </select>
            </div>
          </div>

          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Mã mô tả sơ đồ ({engine.toUpperCase()})
            </label>
            <textarea
              rows={10}
              value={code}
              onChange={(e) => setCode(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white p-2.5 font-mono text-[11px] dark:border-slate-700 dark:bg-slate-950"
            />
          </div>

          <Button
            variant="primary"
            loading={running}
            icon={<Play className="h-4 w-4" />}
            onClick={handleRenderDiagram}
            className="w-full bg-indigo-600 hover:bg-indigo-700"
          >
            Kết xuất Hình Ảnh Sơ Đồ 300+ DPI
          </Button>
        </div>
      ) : (
        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Loại sơ đồ phân tích</label>
            <select
              value={codeDiagType}
              onChange={(e) => setCodeDiagType(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            >
              <option value="auto">Tự động nhận diện (SQL -&gt; ERD, Python -&gt; Class/Flow)</option>
              <option value="erd">Sơ đồ Quan hệ Dữ liệu (ERD)</option>
              <option value="class">Sơ đồ Lớp (Class Diagram)</option>
              <option value="flowchart">Lưu đồ Luồng xử lý (Flowchart)</option>
            </select>
          </div>

          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Dán mã nguồn SQL DDL, Prisma hoặc Python
            </label>
            <textarea
              rows={10}
              value={sourceCode}
              onChange={(e) => setSourceCode(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-white p-2.5 font-mono text-[11px] dark:border-slate-700 dark:bg-slate-950"
            />
          </div>

          <CollapsibleSection title="Mẹo xuất sang Draw.io / Vector SVG">
            <p className="text-xs text-slate-600 dark:text-slate-400">
              Tính năng tự động sinh đồng thời 3 định dạng: ảnh PNG 300 DPI, tệp `.drawio` để mở chỉnh tay trong Draw.io Desktop, và Vector SVG.
            </p>
          </CollapsibleSection>

          <Button
            variant="primary"
            loading={running}
            icon={<Code2 className="h-4 w-4" />}
            onClick={handleGenerateFromCode}
            className="w-full bg-indigo-600 hover:bg-indigo-700"
          >
            Phân tích AST & Vẽ Sơ đồ đa định dạng
          </Button>
        </div>
      )}
    </div>
  );

  const resultPanel =
    outputSession || codeOutputs ? (
      <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-emerald-200 bg-white px-4 py-3 dark:border-emerald-900/50 dark:bg-slate-900">
        <div className="text-xs font-semibold text-slate-900 dark:text-slate-100">
          {outputSession?.filename || 'Sơ đồ đã kết xuất'}
        </div>
        <div className="flex flex-wrap gap-2">
          {outputSession?.downloadUrl ? (
            <a
              href={outputSession.downloadUrl}
              download={outputSession.filename}
              className="inline-flex items-center gap-1 rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-emerald-700"
            >
              <Download className="h-3.5 w-3.5" />
              <span>Tải PNG</span>
            </a>
          ) : null}
          {codeOutputs?.drawio_download_url ? (
            <a
              href={codeOutputs.drawio_download_url}
              download="diagram.drawio"
              className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
            >
              <Download className="h-3.5 w-3.5" />
              <span>Tải .Drawio XML</span>
            </a>
          ) : null}
          {codeOutputs?.svg_download_url ? (
            <a
              href={codeOutputs.svg_download_url}
              download="diagram.svg"
              className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
            >
              <Download className="h-3.5 w-3.5" />
              <span>Tải Vector SVG</span>
            </a>
          ) : null}
        </div>
      </div>
    ) : undefined;

  return (
    <SplitWorkspace
      title="Diagram Studio — Sơ Đồ Quy Trình & Kiến Trúc"
      subtitle="Biên dịch sơ đồ quy trình Mermaid, C4 PlantUML, Draw.io mxGraph, Canvas JSON hoặc phân tích tự động từ mã SQL/Python sang 300+ DPI."
      badgeText="300+ DPI Vector"
      accent="studio"
      leftPanel={leftPanel}
      resultPanel={resultPanel}
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="indigo">Chạy Engine Local</Badge>}
    />
  );
}
