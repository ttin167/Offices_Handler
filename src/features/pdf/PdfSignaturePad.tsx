import React, { useEffect, useRef, useState } from 'react';
import { Eraser, PenTool, Upload } from 'lucide-react';
import { Button, CollapsibleSection, FileDropZone, Tabs, type SelectedFileInfo } from '../../components/ui';
import { useToast } from '../../context/ToastContext';

export interface SignaturePayload {
  page: string;
  x_pct: number;
  y_pct: number;
  w_pct: number;
  h_pct: number;
  image_base64?: string;
  image_path?: string;
  signer_title?: string;
  signer_name?: string;
  signer_date?: string;
}

const INK_COLORS = [
  { hex: '#0033aa', label: 'Xanh bút bi' },
  { hex: '#111827', label: 'Đen truyền thống' },
  { hex: '#dc2626', label: 'Đỏ con dấu' },
];

const POSITION_PRESETS = [
  { id: 'bottom-right', label: 'Dưới phải' },
  { id: 'bottom-left', label: 'Dưới trái' },
  { id: 'center', label: 'Giữa trang' },
  { id: 'top-right', label: 'Trên phải' },
] as const;

export function PdfSignaturePad({
  disabled,
  loading,
  onApplySign,
}: {
  disabled: boolean;
  loading: boolean;
  onApplySign: (sig: SignaturePayload) => void;
}) {
  const toast = useToast();
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const drawingRef = useRef(false);
  const lastPointRef = useRef<{ x: number; y: number }>({ x: 0, y: 0 });

  const [signMode, setSignMode] = useState<'draw' | 'upload'>('draw');
  const [inkColor, setInkColor] = useState('#0033aa');
  const [hasDrawn, setHasDrawn] = useState(false);
  const [stampFile, setStampFile] = useState<SelectedFileInfo | null>(null);

  const [position, setPosition] = useState<(typeof POSITION_PRESETS)[number]['id']>('bottom-right');
  const [targetPageMode, setTargetPageMode] = useState<'last' | '1' | 'all' | 'custom'>('last');
  const [customPage, setCustomPage] = useState('1');
  const [scalePct, setScalePct] = useState(25);

  const [includeMeta, setIncludeMeta] = useState(false);
  const [signerTitle, setSignerTitle] = useState('Giám đốc');
  const [signerName, setSignerName] = useState('');
  const [signerDate, setSignerDate] = useState(() => {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    return `Ngày ${d} tháng ${m} năm ${now.getFullYear()}`;
  });

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.lineWidth = 2.6;
  }, []);

  const getCoords = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current!;
    const rect = canvas.getBoundingClientRect();
    return {
      x: ((e.clientX - rect.left) * canvas.width) / rect.width,
      y: ((e.clientY - rect.top) * canvas.height) / rect.height,
    };
  };

  const handlePointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    e.currentTarget.setPointerCapture(e.pointerId);
    drawingRef.current = true;
    lastPointRef.current = getCoords(e);
  };

  const handlePointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!drawingRef.current || !canvasRef.current) return;
    const ctx = canvasRef.current.getContext('2d');
    if (!ctx) return;
    const pt = getCoords(e);
    ctx.strokeStyle = inkColor;
    ctx.lineWidth = 2.8;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    ctx.moveTo(lastPointRef.current.x, lastPointRef.current.y);
    ctx.lineTo(pt.x, pt.y);
    ctx.stroke();
    lastPointRef.current = pt;
    if (!hasDrawn) setHasDrawn(true);
  };

  const handlePointerUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    drawingRef.current = false;
    if (e.currentTarget.hasPointerCapture(e.pointerId)) {
      e.currentTarget.releasePointerCapture(e.pointerId);
    }
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  };

  const handleSubmit = () => {
    if (signMode === 'draw' && !hasDrawn) {
      toast.warning('Chưa có nét chữ ký', 'Vui lòng ký vào khung trắng trước khi đóng dấu.');
      return;
    }
    if (signMode === 'upload' && !stampFile) {
      toast.warning('Chưa có ảnh chữ ký', 'Vui lòng tải lên tệp PNG/JPG chữ ký hoặc con dấu.');
      return;
    }

    const wPct = scalePct / 100.0;
    const hPct = wPct * 0.6;
    const metaOffset = includeMeta ? 0.08 : 0.0;

    let xPct = Math.max(0.05, 0.95 - wPct);
    let yPct = Math.max(0.05, 0.95 - hPct - metaOffset);

    if (position === 'bottom-left') {
      xPct = 0.08;
      yPct = Math.max(0.05, 0.95 - hPct - metaOffset);
    } else if (position === 'center') {
      xPct = Math.max(0.05, (1.0 - wPct) / 2.0);
      yPct = 0.45;
    } else if (position === 'top-right') {
      xPct = Math.max(0.05, 0.95 - wPct);
      yPct = 0.08;
    }

    const payload: SignaturePayload = {
      page: targetPageMode === 'custom' ? customPage.trim() || '1' : targetPageMode,
      x_pct: xPct,
      y_pct: yPct,
      w_pct: wPct,
      h_pct: hPct,
    };

    if (signMode === 'draw' && canvasRef.current) {
      payload.image_base64 = canvasRef.current.toDataURL('image/png');
    } else if (stampFile) {
      payload.image_path = stampFile.filePath;
    }

    if (includeMeta) {
      payload.signer_title = signerTitle.trim();
      payload.signer_name = signerName.trim();
      payload.signer_date = signerDate.trim();
    }

    onApplySign(payload);
  };

  return (
    <div className="space-y-3">
      <Tabs
        activeId={signMode}
        onChange={(id) => setSignMode(id as 'draw' | 'upload')}
        options={[
          { id: 'draw', label: 'Vẽ chữ ký tay', icon: <PenTool className="h-3.5 w-3.5" /> },
          { id: 'upload', label: 'Tải ảnh dấu / Chữ ký PNG', icon: <Upload className="h-3.5 w-3.5" /> },
        ]}
      />

      {signMode === 'draw' ? (
        <div className="space-y-2">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="text-[11px] font-medium text-slate-600 dark:text-slate-400">Màu mực:</span>
              {INK_COLORS.map((c) => (
                <button
                  key={c.hex}
                  type="button"
                  title={c.label}
                  onClick={() => setInkColor(c.hex)}
                  style={{ backgroundColor: c.hex }}
                  className={`h-5 w-5 rounded-full border-2 transition ${
                    inkColor === c.hex
                      ? 'scale-110 border-slate-900 ring-2 ring-slate-300 dark:border-white dark:ring-slate-600'
                      : 'border-white dark:border-slate-800'
                  }`}
                />
              ))}
            </div>
            <Button size="sm" variant="ghost" icon={<Eraser className="h-3.5 w-3.5" />} onClick={clearCanvas}>
              Xóa nét vẽ
            </Button>
          </div>
          <div className="relative overflow-hidden rounded-lg border border-slate-300 bg-white dark:border-slate-700">
            <canvas
              ref={canvasRef}
              width={480}
              height={150}
              onPointerDown={handlePointerDown}
              onPointerMove={handlePointerMove}
              onPointerUp={handlePointerUp}
              onPointerLeave={handlePointerUp}
              className="h-36 w-full cursor-crosshair touch-none"
            />
            {!hasDrawn ? (
              <span className="pointer-events-none absolute inset-0 flex items-center justify-center text-xs text-slate-400">
                Ký tên bằng chuột hoặc cảm ứng tại đây...
              </span>
            ) : null}
          </div>
        </div>
      ) : (
        <FileDropZone
          label="Ảnh chữ ký tách nền hoặc con dấu (.png, .jpg)"
          accept=".png,.jpg,.jpeg,.webp"
          selectedFile={stampFile}
          onFileSelected={setStampFile}
          onClear={() => setStampFile(null)}
          sourceLabel="Chữ ký / Con dấu"
        />
      )}

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Vị trí đặt dấu</label>
          <div className="grid grid-cols-2 gap-1.5">
            {POSITION_PRESETS.map((pos) => (
              <button
                key={pos.id}
                type="button"
                onClick={() => setPosition(pos.id)}
                className={`rounded-lg border px-2 py-1.5 text-xs font-medium transition ${
                  position === pos.id
                    ? 'border-red-500 bg-red-50 text-red-700 dark:bg-red-950/40 dark:text-red-300'
                    : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300'
                }`}
              >
                {pos.label}
              </button>
            ))}
          </div>
        </div>

        <div className="space-y-2">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Trang áp dụng</label>
            <select
              value={targetPageMode}
              onChange={(e) => setTargetPageMode(e.target.value as 'last' | '1' | 'all' | 'custom')}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200"
            >
              <option value="last">Trang cuối cùng</option>
              <option value="1">Trang đầu tiên</option>
              <option value="all">Tất cả các trang</option>
              <option value="custom">Nhập số trang cụ thể</option>
            </select>
          </div>
          {targetPageMode === 'custom' ? (
            <input
              type="text"
              value={customPage}
              onChange={(e) => setCustomPage(e.target.value)}
              placeholder="Ví dụ: 1 hoặc 2"
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200"
            />
          ) : null}
          <div>
            <div className="flex items-center justify-between text-xs text-slate-600 dark:text-slate-400">
              <span>Kích thước hiển thị</span>
              <span className="font-semibold">{scalePct}%</span>
            </div>
            <input
              type="range"
              min={12}
              max={45}
              value={scalePct}
              onChange={(e) => setScalePct(Number(e.target.value))}
              className="w-full accent-red-600"
            />
          </div>
        </div>
      </div>

      <CollapsibleSection title="Kèm thông tin chức danh & ngày ký">
        <label className="flex items-center gap-2 text-xs font-medium text-slate-700 dark:text-slate-300">
          <input
            type="checkbox"
            checked={includeMeta}
            onChange={(e) => setIncludeMeta(e.target.checked)}
            className="rounded border-slate-300 text-red-600 focus:ring-red-500"
          />
          <span>Hiển thị dòng chức danh, họ tên & ngày ký dưới dấu</span>
        </label>
        {includeMeta ? (
          <div className="mt-2 grid grid-cols-1 gap-2 sm:grid-cols-3">
            <input
              type="text"
              value={signerTitle}
              onChange={(e) => setSignerTitle(e.target.value)}
              placeholder="Chức danh"
              className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
            <input
              type="text"
              value={signerName}
              onChange={(e) => setSignerName(e.target.value)}
              placeholder="Họ và tên người ký"
              className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
            <input
              type="text"
              value={signerDate}
              onChange={(e) => setSignerDate(e.target.value)}
              placeholder="Ngày ký"
              className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </div>
        ) : null}
      </CollapsibleSection>

      <Button
        variant="primary"
        loading={loading}
        disabled={disabled}
        icon={<PenTool className="h-4 w-4" />}
        onClick={handleSubmit}
        className="w-full bg-red-600 hover:bg-red-700"
      >
        Đóng chữ ký vào PDF
      </Button>
    </div>
  );
}
