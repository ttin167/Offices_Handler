import React, { useEffect, useState } from 'react';
import {
  Archive,
  ArrowDown,
  ArrowUp,
  BadgeCheck,
  Eraser,
  EyeOff,
  FileArchive,
  FileCheck2,
  FileKey2,
  FileSpreadsheet,
  FormInput,
  Image as ImageIcon,
  Layers,
  Minimize2,
  PenTool,
  Plus,
  Scissors,
  Shield,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  Stamp,
  Trash2,
} from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
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
import { PdfPageOrganizerGrid } from './PdfPageOrganizerGrid';
import { PdfSignaturePad, type SignaturePayload } from './PdfSignaturePad';

type PdfGroupId = 'organize' | 'sign_forms' | 'security_archival';

type PdfToolKey =
  | 'reorganize'
  | 'merge'
  | 'split'
  | 'images_to_pdf'
  | 'export_images_zip'
  | 'extract_tables'
  | 'compress'
  | 'normalize_a4'
  | 'clean'
  | 'sign'
  | 'giap_lai'
  | 'watermark'
  | 'forms'
  | 'redact'
  | 'protect'
  | 'flatten'
  | 'crypto_sign'
  | 'pdfa';

interface ToolPill {
  id: PdfToolKey;
  label: string;
  icon: React.ReactNode;
}

const GROUP_TOOLS: Record<PdfGroupId, ToolPill[]> = {
  organize: [
    { id: 'reorganize', label: 'Tổ chức trang', icon: <Layers className="h-3.5 w-3.5" /> },
    { id: 'merge', label: 'Ghép PDF', icon: <Plus className="h-3.5 w-3.5" /> },
    { id: 'split', label: 'Tách PDF (.ZIP)', icon: <Scissors className="h-3.5 w-3.5" /> },
    { id: 'images_to_pdf', label: 'Ảnh sang PDF', icon: <ImageIcon className="h-3.5 w-3.5" /> },
    { id: 'export_images_zip', label: 'PDF ra Ảnh ZIP', icon: <FileArchive className="h-3.5 w-3.5" /> },
    { id: 'extract_tables', label: 'Bóc bảng Excel', icon: <FileSpreadsheet className="h-3.5 w-3.5" /> },
    { id: 'compress', label: 'Nén dung lượng', icon: <Minimize2 className="h-3.5 w-3.5" /> },
    { id: 'normalize_a4', label: 'Chuẩn A4 & Bates', icon: <SlidersHorizontal className="h-3.5 w-3.5" /> },
    { id: 'clean', label: 'Lọc trang trắng', icon: <Eraser className="h-3.5 w-3.5" /> },
  ],
  sign_forms: [
    { id: 'sign', label: 'Ký tên & Con dấu', icon: <PenTool className="h-3.5 w-3.5" /> },
    { id: 'giap_lai', label: 'Đóng dấu Giáp lai', icon: <Stamp className="h-3.5 w-3.5" /> },
    { id: 'watermark', label: 'Đóng Watermark', icon: <Sparkles className="h-3.5 w-3.5" /> },
    { id: 'forms', label: 'Điền AcroForm', icon: <FormInput className="h-3.5 w-3.5" /> },
  ],
  security_archival: [
    { id: 'redact', label: 'Che đen PII', icon: <EyeOff className="h-3.5 w-3.5" /> },
    { id: 'protect', label: 'Mật khẩu AES-256', icon: <FileKey2 className="h-3.5 w-3.5" /> },
    { id: 'flatten', label: 'Làm phẳng & Xóa Meta', icon: <Shield className="h-3.5 w-3.5" /> },
    { id: 'crypto_sign', label: 'Ký số PAdES', icon: <BadgeCheck className="h-3.5 w-3.5" /> },
    { id: 'pdfa', label: 'Chuẩn hóa PDF/A', icon: <Archive className="h-3.5 w-3.5" /> },
  ],
};

interface PdfActionResponse {
  success?: boolean;
  filename?: string;
  output_file?: string;
  output_path?: string;
  download_url?: string;
  detail?: string;
  details?: Record<string, unknown>;
}

interface AcroFieldInfo {
  field_name: string;
  field_type?: string;
  field_value?: string | boolean | null;
}

export function PdfProHubPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [group, setGroup] = useState<PdfGroupId>('organize');
  const [activeTool, setActiveTool] = useState<PdfToolKey>('reorganize');

  const [primaryPdf, setPrimaryPdf] = useState<SelectedFileInfo | null>(null);
  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [lastActionTitle, setLastActionTitle] = useState('');
  const [lastDetailText, setLastDetailText] = useState('');

  // Merge & Img2Pdf multi-file lists
  const [mergeList, setMergeList] = useState<SelectedFileInfo[]>([]);
  const [imageList, setImageList] = useState<SelectedFileInfo[]>([]);
  const [imgPageSize, setImgPageSize] = useState('A4');
  const [imgOrientation, setImgOrientation] = useState('auto');

  // Split & Export images & Normalize & Clean
  const [splitRanges, setSplitRanges] = useState('');
  const [expDpi, setExpDpi] = useState(300);
  const [expFmt, setExpFmt] = useState('png');
  const [normOrientation, setNormOrientation] = useState('auto');
  const [batesPrefix, setBatesPrefix] = useState('');
  const [batesPattern, setBatesPattern] = useState('{prefix} | Trang {page:02d}/{total:02d}');
  const [batesPosition, setBatesPosition] = useState('bottom-right');
  const [cleanAngle, setCleanAngle] = useState(0);
  const [removeBlanks, setRemoveBlanks] = useState(true);

  // Giáp lai & Watermark
  const [giapLaiStamp, setGiapLaiStamp] = useState<SelectedFileInfo | null>(null);
  const [giapLaiSide, setGiapLaiSide] = useState('right');
  const [giapLaiDiameter, setGiapLaiDiameter] = useState(120);
  const [giapLaiOverlapPct, setGiapLaiOverlapPct] = useState(8);
  const [watermarkText, setWatermarkText] = useState('TÀI LIỆU NỘI BỘ');

  // Forms
  const [formFields, setFormFields] = useState<AcroFieldInfo[]>([]);
  const [formValues, setFormValues] = useState<Record<string, string | boolean>>({});
  const [formFlatten, setFormFlatten] = useState(true);
  const [scanningForm, setScanningForm] = useState(false);

  // Redact, Protect, Flatten, Crypto Sign, PDF/A
  const [redactPatterns, setRedactPatterns] = useState<string[]>(['cccd', 'phone', 'email']);
  const [redactTerms, setRedactTerms] = useState('');
  const [userPwd, setUserPwd] = useState('');
  const [ownerPwd, setOwnerPwd] = useState('');
  const [allowPrint, setAllowPrint] = useState(true);
  const [allowCopy, setAllowCopy] = useState(false);
  const [allowModify, setAllowModify] = useState(false);
  const [flatVector, setFlatVector] = useState(true);
  const [flatStrict, setFlatStrict] = useState(false);
  const [flatMeta, setFlatMeta] = useState(true);

  const [csignName, setCsignName] = useState('Antigravity Authority');
  const [csignLocation, setCsignLocation] = useState('Việt Nam');
  const [csignReason, setCsignReason] = useState('Phê duyệt & Xác thực tài liệu');
  const [pfxFile, setPfxFile] = useState<SelectedFileInfo | null>(null);
  const [pfxPassword, setPfxPassword] = useState('');

  const [pdfaPart, setPdfaPart] = useState(2);
  const [pdfaConf, setPdfaConf] = useState('B');

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (!hf) return;
    const ext = hf.ext.toLowerCase();
    if (ext === 'pdf') {
      const sel = sessionFileToSelected(hf);
      setPrimaryPdf(sel);
      setMergeList((prev) => (prev.some((x) => x.filePath === sel.filePath) ? prev : [...prev, sel]));
    } else if (['png', 'jpg', 'jpeg', 'webp'].includes(ext)) {
      const sel = sessionFileToSelected(hf);
      setImageList((prev) => (prev.some((x) => x.filePath === sel.filePath) ? prev : [...prev, sel]));
      setGroup('organize');
      setActiveTool('images_to_pdf');
    }
  }, [consumeHandoffFile]);

  const handleGroupChange = (newGroup: string) => {
    const g = newGroup as PdfGroupId;
    setGroup(g);
    setActiveTool(GROUP_TOOLS[g][0].id);
  };

  const runPdfEndpoint = async (
    endpoint: string,
    payload: Record<string, unknown>,
    title: string,
    requirePrimary = true,
  ) => {
    if (requirePrimary && !primaryPdf) {
      toast.warning('Chưa có tệp PDF chính', 'Vui lòng tải lên hoặc chọn tệp PDF ở khung phía trên.');
      return;
    }
    setRunning(true);
    try {
      const body = requirePrimary && primaryPdf ? { file_path: primaryPdf.filePath, ...payload } : payload;
      const res = await apiPost<PdfActionResponse>(endpoint, body);
      const outName = res.filename || res.output_file || 'result.pdf';
      const sf = addSessionFile({
        filePath: res.output_path || outName,
        filename: outName,
        previewUrl: buildPreviewUrl(outName),
        downloadUrl: res.download_url || buildDownloadUrl(outName),
        sourceLabel: title,
      });
      setOutputSession(sf);
      setLastActionTitle(title);
      setLastDetailText(
        res.detail ||
          (res.details ? JSON.stringify(res.details) : ''),
      );
      toast.success(`${title} thành công`, outName);
    } catch (err) {
      toast.error(`Lỗi ${title}`, err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const runPdfTool = (action: string, payload: Record<string, unknown>, title: string, requirePrimary = true) =>
    runPdfEndpoint('/api/pdf/tools', { action, ...payload }, title, requirePrimary);

  const handleMultiUpload = async (
    files: FileList | null,
    target: 'merge' | 'images',
  ) => {
    if (!files || !files.length) return;
    try {
      const res = await uploadFiles(Array.from(files));
      const added: SelectedFileInfo[] = (res.files || []).map((f) => {
        const sf = addSessionFile({
          filePath: f.file_path,
          filename: f.filename,
          size: f.size,
          previewUrl: f.preview_url,
          sourceLabel: target === 'merge' ? 'Ghép PDF' : 'Ảnh sang PDF',
        });
        return sessionFileToSelected(sf);
      });
      if (target === 'merge') {
        setMergeList((prev) => [...prev, ...added]);
      } else {
        setImageList((prev) => [...prev, ...added]);
      }
      toast.success(`Đã thêm ${added.length} tệp`);
    } catch (err) {
      toast.error('Lỗi tải lên', err instanceof Error ? err.message : String(err));
    }
  };

  const moveListItem = (
    setter: React.Dispatch<React.SetStateAction<SelectedFileInfo[]>>,
    idx: number,
    dir: -1 | 1,
  ) => {
    setter((prev) => {
      const target = idx + dir;
      if (target < 0 || target >= prev.length) return prev;
      const next = [...prev];
      const [item] = next.splice(idx, 1);
      next.splice(target, 0, item);
      return next;
    });
  };

  const handleInspectForms = async () => {
    if (!primaryPdf) {
      toast.warning('Vui lòng nạp tệp PDF trước khi quét biểu mẫu');
      return;
    }
    setScanningForm(true);
    try {
      const res = await apiPost<{ success: boolean; total_fields?: number; fields?: AcroFieldInfo[]; error?: string }>(
        '/api/pdf/inspect-form',
        { file_path: primaryPdf.filePath },
      );
      const list = res.fields || [];
      setFormFields(list);
      const initialVals: Record<string, string | boolean> = {};
      for (const f of list) {
        if (f.field_type === 'CheckBox') {
          initialVals[f.field_name] = Boolean(f.field_value);
        } else {
          initialVals[f.field_name] = f.field_value !== null && f.field_value !== undefined ? String(f.field_value) : '';
        }
      }
      setFormValues(initialVals);
      if (list.length === 0) {
        toast.info('Không tìm thấy trường AcroForm tương tác trong tệp PDF này');
      } else {
        toast.success(`Đã quét được ${list.length} trường AcroForm`);
      }
    } catch (err) {
      toast.error('Lỗi quét AcroForm', err instanceof Error ? err.message : String(err));
    } finally {
      setScanningForm(false);
    }
  };

  const toggleRedactPattern = (key: string) => {
    setRedactPatterns((prev) => (prev.includes(key) ? prev.filter((x) => x !== key) : [...prev, key]));
  };

  const previewTarget = outputSession
    ? {
        filePath: outputSession.filePath,
        filename: outputSession.filename,
        ext: outputSession.ext,
        previewUrl: outputSession.previewUrl,
        downloadUrl: outputSession.downloadUrl,
      }
    : primaryPdf
      ? {
          filePath: primaryPdf.filePath,
          filename: primaryPdf.filename,
          ext: primaryPdf.ext,
          previewUrl: primaryPdf.previewUrl,
          downloadUrl: primaryPdf.downloadUrl,
        }
      : null;

  const renderMultiQueue = (
    items: SelectedFileInfo[],
    setter: React.Dispatch<React.SetStateAction<SelectedFileInfo[]>>,
    accept: string,
    target: 'merge' | 'images',
    emptyText: string,
  ) => (
    <div className="space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-xs font-medium text-slate-600 dark:text-slate-400">
          Danh sách tệp ({items.length})
        </span>
        <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-red-200 bg-red-50 px-2.5 py-1.5 text-xs font-medium text-red-700 hover:bg-red-100 dark:border-red-900/60 dark:bg-red-950/40 dark:text-red-300">
          <Plus className="h-3.5 w-3.5" />
          <span>Thêm nhiều tệp</span>
          <input
            type="file"
            multiple
            accept={accept}
            className="hidden"
            onChange={(e) => {
              handleMultiUpload(e.target.files, target);
              e.currentTarget.value = '';
            }}
          />
        </label>
      </div>
      {items.length === 0 ? (
        <div className="rounded-lg border border-dashed border-slate-200 px-3 py-5 text-center text-xs text-slate-500 dark:border-slate-800 dark:text-slate-400">
          {emptyText}
        </div>
      ) : (
        <ul className="max-h-48 space-y-1.5 overflow-y-auto pr-1">
          {items.map((f, idx) => (
            <li
              key={`${f.filePath}-${idx}`}
              className="flex items-center justify-between gap-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-800 dark:bg-slate-950"
            >
              <span className="truncate font-medium text-slate-800 dark:text-slate-200">
                {idx + 1}. {f.filename}
              </span>
              <div className="flex shrink-0 items-center gap-1">
                <button
                  type="button"
                  disabled={idx === 0}
                  onClick={() => moveListItem(setter, idx, -1)}
                  className="rounded p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-30 dark:hover:bg-slate-800"
                >
                  <ArrowUp className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  disabled={idx === items.length - 1}
                  onClick={() => moveListItem(setter, idx, 1)}
                  className="rounded p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-30 dark:hover:bg-slate-800"
                >
                  <ArrowDown className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onClick={() => setter((prev) => prev.filter((_, i) => i !== idx))}
                  className="rounded p-1 text-red-500 hover:bg-red-50 dark:hover:bg-red-950/40"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </button>
              </div>
            </li>
          ))}
        </ul>
      )}
    </div>
  );

  const renderActiveToolPanel = () => {
    switch (activeTool) {
      case 'reorganize':
        return (
          <PdfPageOrganizerGrid
            pdfFile={primaryPdf}
            loadingAction={running}
            onApplyReorganize={(pagesPlan) =>
              runPdfTool('reorganize', { pages_plan: pagesPlan }, 'Tổ chức & sắp xếp trang PDF')
            }
          />
        );

      case 'merge':
        return (
          <div className="space-y-3">
            {renderMultiQueue(
              mergeList,
              setMergeList,
              '.pdf',
              'merge',
              'Chọn ít nhất 2 tệp PDF để nối thành một văn bản thống nhất.',
            )}
            <Button
              variant="primary"
              loading={running}
              disabled={mergeList.length < 2}
              onClick={() =>
                runPdfTool(
                  'merge',
                  { inputs: mergeList.map((m) => m.filePath) },
                  'Ghép nối nhiều PDF',
                  false,
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Ghép {mergeList.length} tệp PDF thành 1 file
            </Button>
          </div>
        );

      case 'split':
        return (
          <div className="space-y-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Khoảng trang tách (tùy chọn)
              </label>
              <input
                type="text"
                value={splitRanges}
                onChange={(e) => setSplitRanges(e.target.value)}
                placeholder="Để trống để tách từng trang, hoặc nhập: 1-3, 4-8"
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfTool('split', { ranges: splitRanges.trim() || null }, 'Tách trang PDF (.ZIP)')
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Tách PDF và đóng gói ZIP
            </Button>
          </div>
        );

      case 'images_to_pdf':
        return (
          <div className="space-y-3">
            {renderMultiQueue(
              imageList,
              setImageList,
              '.png,.jpg,.jpeg,.webp',
              'images',
              'Thêm các ảnh chụp chứng từ (.png, .jpg, .webp) để gộp thành PDF chuẩn in ấn.',
            )}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Khổ giấy</label>
                <select
                  value={imgPageSize}
                  onChange={(e) => setImgPageSize(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="A4">Chuẩn A4</option>
                  <option value="Letter">Letter (US)</option>
                  <option value="fit_image">Vừa khít kích thước ảnh</option>
                </select>
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Hướng trang</label>
                <select
                  value={imgOrientation}
                  onChange={(e) => setImgOrientation(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="auto">Tự động theo tỉ lệ ảnh</option>
                  <option value="portrait">Dọc (Portrait)</option>
                  <option value="landscape">Ngang (Landscape)</option>
                </select>
              </div>
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={imageList.length === 0}
              onClick={() =>
                runPdfTool(
                  'images_to_pdf',
                  {
                    images: imageList.map((i) => i.filePath),
                    page_size: imgPageSize,
                    orientation: imgOrientation,
                  },
                  'Gộp ảnh thành PDF',
                  false,
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Chuyển {imageList.length} ảnh thành PDF
            </Button>
          </div>
        );

      case 'export_images_zip':
        return (
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Độ phân giải</label>
                <select
                  value={expDpi}
                  onChange={(e) => setExpDpi(Number(e.target.value))}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value={150}>150 DPI (Nhẹ)</option>
                  <option value={300}>300 DPI (Sắc nét chuẩn in)</option>
                  <option value={600}>600 DPI (Siêu nét)</option>
                </select>
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Định dạng ảnh</label>
                <select
                  value={expFmt}
                  onChange={(e) => setExpFmt(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="png">PNG (Không mất nét chữ)</option>
                  <option value="jpg">JPG (Dung lượng nhỏ)</option>
                </select>
              </div>
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfTool('export_images_zip', { dpi: expDpi, format: expFmt }, 'Xuất PDF ra bộ ảnh ZIP')
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Xuất từng trang PDF ra ZIP
            </Button>
          </div>
        );

      case 'extract_tables':
        return (
          <div className="space-y-3">
            <p className="text-xs leading-relaxed text-slate-600 dark:text-slate-400">
              Tự động nhận diện toàn bộ bảng biểu trong PDF và xuất ra các sheet Excel (.xlsx) giữ nguyên cấu trúc hàng/cột.
            </p>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() => runPdfTool('extract_tables', {}, 'Bóc tách bảng PDF sang Excel')}
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Bóc toàn bộ bảng sang Excel
            </Button>
          </div>
        );

      case 'compress':
        return (
          <div className="space-y-3">
            <p className="text-xs leading-relaxed text-slate-600 dark:text-slate-400">
              Dọn rác cấu trúc đối tượng, nén luồng hình ảnh và giảm dung lượng PDF phục vụ gửi email hoặc nộp hồ sơ cổng dịch vụ công.
            </p>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() => runPdfTool('compress', {}, 'Nén tối ưu dung lượng PDF')}
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Nén PDF ngay
            </Button>
          </div>
        );

      case 'normalize_a4':
        return (
          <div className="space-y-3">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Hướng giấy A4</label>
                <select
                  value={normOrientation}
                  onChange={(e) => setNormOrientation(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="auto">Tự động (Theo trang)</option>
                  <option value="portrait">Đồng bộ Dọc (A4 Portrait)</option>
                  <option value="landscape">Đồng bộ Ngang (A4 Landscape)</option>
                </select>
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                  Tiền tố đánh số Bates (tùy chọn)
                </label>
                <input
                  type="text"
                  value={batesPrefix}
                  onChange={(e) => setBatesPrefix(e.target.value)}
                  placeholder="Ví dụ: HS-2026"
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
            </div>
            <CollapsibleSection title="Tùy chọn nâng cao (Mẫu Bates & Vị trí)">
              <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
                <input
                  type="text"
                  value={batesPattern}
                  onChange={(e) => setBatesPattern(e.target.value)}
                  placeholder="{prefix} | Trang {page:02d}/{total:02d}"
                  className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
                <select
                  value={batesPosition}
                  onChange={(e) => setBatesPosition(e.target.value)}
                  className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="bottom-right">Góc dưới phải</option>
                  <option value="bottom-center">Chính giữa dưới</option>
                  <option value="bottom-left">Góc dưới trái</option>
                  <option value="top-right">Góc trên phải</option>
                </select>
              </div>
            </CollapsibleSection>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfTool(
                  'normalize_a4',
                  {
                    orientation: normOrientation,
                    bates_prefix: batesPrefix.trim() || null,
                    bates_pattern: batesPattern,
                    position: batesPosition,
                  },
                  'Chuẩn hóa A4 & Đánh số Bates',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Chuẩn hóa A4 & Đánh số trang
            </Button>
          </div>
        );

      case 'clean':
        return (
          <div className="space-y-3">
            <label className="flex items-center gap-2 text-xs font-medium text-slate-700 dark:text-slate-300">
              <input
                type="checkbox"
                checked={removeBlanks}
                onChange={(e) => setRemoveBlanks(e.target.checked)}
                className="rounded border-slate-300 text-red-600 focus:ring-red-500"
              />
              <span>Tự động phát hiện và loại bỏ trang scan trắng</span>
            </label>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Xoay đồng loạt toàn bộ tài liệu
              </label>
              <select
                value={cleanAngle}
                onChange={(e) => setCleanAngle(Number(e.target.value))}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value={0}>Giữ nguyên hướng (0°)</option>
                <option value={90}>Xoay phải 90°</option>
                <option value={180}>Đảo ngược 180°</option>
                <option value={270}>Xoay trái 90° (270°)</option>
              </select>
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfTool(
                  'clean',
                  { angle: cleanAngle, remove_blanks: removeBlanks },
                  'Làm sạch & xoay PDF',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Thực hiện làm sạch PDF
            </Button>
          </div>
        );

      case 'sign':
        return (
          <PdfSignaturePad
            disabled={!primaryPdf}
            loading={running}
            onApplySign={(sig: SignaturePayload) =>
              runPdfTool('sign', { signatures: [sig] }, 'Ký tên & Đóng dấu điện tử')
            }
          />
        );

      case 'giap_lai':
        return (
          <div className="space-y-3">
            <FileDropZone
              label="Ảnh con dấu tròn nền trong suốt (.png)"
              accept=".png,.jpg,.jpeg,.webp"
              selectedFile={giapLaiStamp}
              onFileSelected={setGiapLaiStamp}
              onClear={() => setGiapLaiStamp(null)}
              sourceLabel="Dấu Giáp lai"
            />
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Mép đóng dấu giáp lai
              </label>
              <select
                value={giapLaiSide}
                onChange={(e) => setGiapLaiSide(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="right">Mép phải tài liệu</option>
                <option value="left">Mép trái tài liệu</option>
              </select>
            </div>
            <CollapsibleSection title="Tùy chọn nâng cao (Đường kính dấu & Độ chườm)">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="mb-1 block text-[11px] text-slate-500">Đường kính (pt)</label>
                  <input
                    type="number"
                    min={60}
                    max={220}
                    value={giapLaiDiameter}
                    onChange={(e) => setGiapLaiDiameter(Number(e.target.value))}
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-[11px] text-slate-500">Chườm mép (%)</label>
                  <input
                    type="number"
                    min={2}
                    max={25}
                    value={giapLaiOverlapPct}
                    onChange={(e) => setGiapLaiOverlapPct(Number(e.target.value))}
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              </div>
            </CollapsibleSection>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf || !giapLaiStamp}
              onClick={() =>
                runPdfTool(
                  'giap_lai',
                  {
                    stamp_image_path: giapLaiStamp?.filePath,
                    side: giapLaiSide,
                    stamp_diameter: giapLaiDiameter,
                    overlap_pct: giapLaiOverlapPct,
                  },
                  'Đóng dấu Giáp lai đa trang',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Cắt & Đóng dấu giáp lai lên tất cả các trang
            </Button>
          </div>
        );

      case 'watermark':
        return (
          <div className="space-y-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Nội dung Watermark chéo trang
              </label>
              <input
                type="text"
                value={watermarkText}
                onChange={(e) => setWatermarkText(e.target.value)}
                placeholder="CONFIDENTIAL / BẢN SAO Y BẢN CHÍNH"
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfTool(
                  'watermark',
                  { watermark_text: watermarkText.trim() || 'CONFIDENTIAL' },
                  'Đóng Watermark PDF',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Đóng Watermark lên mọi trang
            </Button>
          </div>
        );

      case 'forms':
        return (
          <div className="space-y-3">
            <Button
              variant="outline"
              loading={scanningForm}
              disabled={!primaryPdf}
              onClick={handleInspectForms}
              className="w-full"
            >
              Quét các trường AcroForm có trong PDF
            </Button>
            {formFields.length > 0 ? (
              <div className="max-h-60 space-y-2 overflow-y-auto rounded-lg border border-slate-200 bg-slate-50/70 p-2.5 dark:border-slate-800 dark:bg-slate-950/60">
                {formFields.map((field) => (
                  <div key={field.field_name} className="space-y-1">
                    <label className="block text-[11px] font-semibold text-slate-700 dark:text-slate-300">
                      {field.field_name}{' '}
                      <span className="font-normal text-slate-400">({field.field_type || 'Text'})</span>
                    </label>
                    {field.field_type === 'CheckBox' ? (
                      <label className="inline-flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                        <input
                          type="checkbox"
                          checked={Boolean(formValues[field.field_name])}
                          onChange={(e) =>
                            setFormValues((prev) => ({ ...prev, [field.field_name]: e.target.checked }))
                          }
                          className="rounded border-slate-300 text-red-600"
                        />
                        <span>Đánh dấu chọn</span>
                      </label>
                    ) : (
                      <input
                        type="text"
                        value={String(formValues[field.field_name] ?? '')}
                        onChange={(e) =>
                          setFormValues((prev) => ({ ...prev, [field.field_name]: e.target.value }))
                        }
                        className="w-full rounded border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                      />
                    )}
                  </div>
                ))}
              </div>
            ) : null}
            <label className="flex items-center gap-2 text-xs font-medium text-slate-700 dark:text-slate-300">
              <input
                type="checkbox"
                checked={formFlatten}
                onChange={(e) => setFormFlatten(e.target.checked)}
                className="rounded border-slate-300 text-red-600"
              />
              <span>Khóa cố định (Flatten) sau khi điền tránh bị sửa đổi</span>
            </label>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf || formFields.length === 0}
              onClick={() =>
                runPdfTool('fill-form', { form_data: formValues, flatten: formFlatten }, 'Điền biểu mẫu AcroForm')
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Điền dữ liệu vào PDF
            </Button>
          </div>
        );

      case 'redact':
        return (
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-2">
              {[
                { key: 'cccd', label: 'Số CCCD / CMND (12 số)' },
                { key: 'phone', label: 'Số điện thoại VN' },
                { key: 'email', label: 'Địa chỉ Email' },
                { key: 'bank_card', label: 'Thẻ ngân hàng / STK' },
                { key: 'salary', label: 'Mức lương / Số tiền' },
              ].map((opt) => (
                <label
                  key={opt.key}
                  className="flex items-center gap-2 rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-700 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-300"
                >
                  <input
                    type="checkbox"
                    checked={redactPatterns.includes(opt.key)}
                    onChange={() => toggleRedactPattern(opt.key)}
                    className="rounded border-slate-300 text-red-600"
                  />
                  <span>{opt.label}</span>
                </label>
              ))}
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                Từ khóa bổ sung (phân tách bằng dấu phẩy)
              </label>
              <input
                type="text"
                value={redactTerms}
                onChange={(e) => setRedactTerms(e.target.value)}
                placeholder="Ví dụ: Nguyễn Văn A, Mật khẩu, Số tài khoản..."
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf || (redactPatterns.length === 0 && !redactTerms.trim())}
              onClick={() =>
                runPdfTool(
                  'redact',
                  {
                    patterns: redactPatterns,
                    custom_terms: redactTerms
                      .split(',')
                      .map((s) => s.trim())
                      .filter(Boolean),
                  },
                  'Che đen vĩnh viễn PII (True Redaction)',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Xóa byte dữ liệu & Bôi đen vĩnh viễn
            </Button>
          </div>
        );

      case 'protect':
        return (
          <div className="space-y-3">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                  Mật khẩu mở tệp
                </label>
                <input
                  type="password"
                  value={userPwd}
                  onChange={(e) => setUserPwd(e.target.value)}
                  placeholder="Yêu cầu nhập để đọc..."
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
                  Mật khẩu chủ sở hữu
                </label>
                <input
                  type="password"
                  value={ownerPwd}
                  onChange={(e) => setOwnerPwd(e.target.value)}
                  placeholder="Khóa quyền in/copy..."
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
            </div>
            <div className="flex flex-wrap gap-3 text-xs text-slate-700 dark:text-slate-300">
              <label className="flex items-center gap-1.5">
                <input type="checkbox" checked={allowPrint} onChange={(e) => setAllowPrint(e.target.checked)} />
                <span>Cho phép in</span>
              </label>
              <label className="flex items-center gap-1.5">
                <input type="checkbox" checked={allowCopy} onChange={(e) => setAllowCopy(e.target.checked)} />
                <span>Cho phép sao chép chữ</span>
              </label>
              <label className="flex items-center gap-1.5">
                <input type="checkbox" checked={allowModify} onChange={(e) => setAllowModify(e.target.checked)} />
                <span>Cho phép chỉnh sửa</span>
              </label>
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf || (!userPwd && !ownerPwd)}
              onClick={() =>
                runPdfTool(
                  'protect',
                  {
                    user_pwd: userPwd,
                    owner_pwd: ownerPwd,
                    allow_print: allowPrint,
                    allow_copy: allowCopy,
                    allow_modify: allowModify,
                  },
                  'Mã hóa AES-256 PDF',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Khóa mật khẩu AES-256
            </Button>
          </div>
        );

      case 'flatten':
        return (
          <div className="space-y-3">
            <div className="space-y-2 text-xs text-slate-700 dark:text-slate-300">
              <label className="flex items-center gap-2">
                <input type="checkbox" checked={flatVector} onChange={(e) => setFlatVector(e.target.checked)} />
                <span>Làm phẳng chữ ký, ghi chú & form vector (Flatten Vector)</span>
              </label>
              <label className="flex items-center gap-2">
                <input type="checkbox" checked={flatMeta} onChange={(e) => setFlatMeta(e.target.checked)} />
                <span>Xóa sạch XMP Metadata, tác giả & lịch sử phần mềm</span>
              </label>
              <label className="flex items-center gap-2 text-amber-700 dark:text-amber-300">
                <input type="checkbox" checked={flatStrict} onChange={(e) => setFlatStrict(e.target.checked)} />
                <span>Chế độ Strict Raster (Chuyển toàn trang thành ảnh chống copy 100%)</span>
              </label>
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfTool(
                  'flatten',
                  { flatten_vector: flatVector, strict_raster: flatStrict, sanitize_metadata: flatMeta },
                  'Làm phẳng & Xóa dấu vết PDF',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Làm phẳng & Làm sạch PDF
            </Button>
          </div>
        );

      case 'crypto_sign':
        return (
          <div className="space-y-3">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Tên chủ thể ký</label>
                <input
                  type="text"
                  value={csignName}
                  onChange={(e) => setCsignName(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Địa điểm ký</label>
                <input
                  type="text"
                  value={csignLocation}
                  onChange={(e) => setCsignLocation(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Lý do ký số</label>
              <input
                type="text"
                value={csignReason}
                onChange={(e) => setCsignReason(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <CollapsibleSection title="Chứng thư số cá nhân (.pfx / .p12) - Không bắt buộc">
              <div className="space-y-2">
                <FileDropZone
                  label="Tải chứng thư số (.pfx, .p12)"
                  accept=".pfx,.p12"
                  selectedFile={pfxFile}
                  onFileSelected={setPfxFile}
                  onClear={() => setPfxFile(null)}
                  sourceLabel="Chứng thư PFX"
                />
                {pfxFile ? (
                  <input
                    type="password"
                    value={pfxPassword}
                    onChange={(e) => setPfxPassword(e.target.value)}
                    placeholder="Mật khẩu mở file .pfx..."
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                ) : null}
              </div>
            </CollapsibleSection>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfEndpoint(
                  '/api/pdf/crypto-sign',
                  {
                    signer_name: csignName.trim() || 'Antigravity Authority',
                    location: csignLocation.trim() || 'Việt Nam',
                    reason: csignReason.trim() || 'Xác thực văn bản',
                    pfx_path: pfxFile?.filePath || null,
                    pfx_password: pfxPassword || null,
                  },
                  'Ký số mật mã PKCS#7 / PAdES',
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Ký số niêm phong PAdES
            </Button>
          </div>
        );

      case 'pdfa':
        return (
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Tiêu chuẩn ISO</label>
                <select
                  value={pdfaPart}
                  onChange={(e) => setPdfaPart(Number(e.target.value))}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value={2}>PDF/A-2 (Khuyến nghị)</option>
                  <option value={1}>PDF/A-1 (Lưu trữ cổ điển)</option>
                  <option value={3}>PDF/A-3 (Cho phép nhúng tệp đính kèm)</option>
                </select>
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Mức tuân thủ</label>
                <select
                  value={pdfaConf}
                  onChange={(e) => setPdfaConf(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="B">Level B (Basic Visual Preservation)</option>
                  <option value="A">Level A (Accessible Structure)</option>
                </select>
              </div>
            </div>
            <Button
              variant="primary"
              loading={running}
              disabled={!primaryPdf}
              onClick={() =>
                runPdfEndpoint(
                  '/api/pdf/pdfa',
                  { part: pdfaPart, conformance: pdfaConf },
                  `Chuẩn hóa lưu trữ ISO PDF/A-${pdfaPart}${pdfaConf.toLowerCase()}`,
                )
              }
              className="w-full bg-red-600 hover:bg-red-700"
            >
              Chuyển đổi chuẩn lưu trữ PDF/A
            </Button>
          </div>
        );
    }
  };

  const leftPanel = (
    <div className="space-y-4">
      <FileDropZone
        label="Tệp PDF làm việc chính"
        hint="Nạp 1 lần để dùng cho hầu hết các công cụ bên dưới (tùy chọn với Ghép PDF & Ảnh sang PDF)"
        accept=".pdf"
        selectedFile={primaryPdf}
        onFileSelected={(f) => {
          setPrimaryPdf(f);
          setOutputSession(null);
          setMergeList((prev) => (prev.some((x) => x.filePath === f.filePath) ? prev : [f, ...prev]));
        }}
        onClear={() => setPrimaryPdf(null)}
        sourceLabel="PDF Pro"
      />

      <Tabs
        activeId={group}
        onChange={handleGroupChange}
        className="w-full"
        options={[
          { id: 'organize', label: 'Tổ chức & Chuyển đổi', icon: <Layers className="h-3.5 w-3.5" /> },
          { id: 'sign_forms', label: 'Ký tên, Dấu & Biểu mẫu', icon: <PenTool className="h-3.5 w-3.5" /> },
          { id: 'security_archival', label: 'Bảo mật & Lưu trữ', icon: <ShieldCheck className="h-3.5 w-3.5" /> },
        ]}
      />

      <div className="flex flex-wrap gap-1.5">
        {GROUP_TOOLS[group].map((pill) => {
          const isActive = activeTool === pill.id;
          return (
            <button
              key={pill.id}
              type="button"
              onClick={() => setActiveTool(pill.id)}
              className={`inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1.5 text-xs font-medium transition ${
                isActive
                  ? 'border-red-500 bg-red-50 text-red-700 shadow-subtle dark:border-red-700 dark:bg-red-950/50 dark:text-red-300'
                  : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300 dark:hover:bg-slate-800'
              }`}
            >
              {pill.icon}
              <span>{pill.label}</span>
            </button>
          );
        })}
      </div>

      <div className="rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
        {renderActiveToolPanel()}
      </div>
    </div>
  );

  const resultPanel = outputSession ? (
    <ResultPreviewCard
      title={lastActionTitle || 'Xử lý PDF hoàn tất'}
      outputFile={outputSession}
      summaryItems={
        lastDetailText
          ? [{ label: 'Chi tiết thực thi', value: lastDetailText.slice(0, 120) }]
          : [{ label: 'Định dạng đầu ra', value: `.${outputSession.ext.toUpperCase()}` }]
      }
      extraActions={
        outputSession.ext.toLowerCase() === 'pdf' ? (
          <Button
            size="sm"
            variant="outline"
            icon={<FileCheck2 className="h-3.5 w-3.5" />}
            onClick={() => setPrimaryPdf(sessionFileToSelected(outputSession))}
          >
            Dùng làm PDF chính tiếp theo
          </Button>
        ) : null
      }
    />
  ) : undefined;

  return (
    <SplitWorkspace
      title="PDF Pro Studio — 18 Nghiệp vụ Toàn diện"
      subtitle="Tổ chức trang 120 DPI, ghép/tách, ký vẽ tay & giáp lai mộc đỏ, điền AcroForm, che PII vĩnh viễn, mã hóa AES-256, PAdES & PDF/A-2b."
      badgeText="PyMuPDF Local"
      accent="pdf"
      leftPanel={leftPanel}
      resultPanel={resultPanel}
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="red">100% Xử lý Local</Badge>}
    />
  );
}
