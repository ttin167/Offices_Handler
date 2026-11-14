import React, { useState } from 'react';
import {
  ArrowRight,
  BookOpen,
  FileCheck2,
  FileSpreadsheet,
  FileText,
  GitBranch,
  Mail,
  Presentation,
  RefreshCw,
  ScanText,
  ShieldAlert,
  ShieldCheck,
  Sparkles,
  Stamp,
} from 'lucide-react';
import { FileDropZone, type SelectedFileInfo } from '../../components/ui/FileDropZone';
import { Badge } from '../../components/ui/Badge';
import { ACCENT_CLASSES, suggestToolsForExtension } from '../../config/navigation';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import type { ToolId } from '../../types/workspace';

interface QuickCardDef {
  toolId: ToolId;
  title: string;
  desc: string;
  tag: string;
  accent: 'word' | 'pdf' | 'excel' | 'studio' | 'slide';
  icon: React.ReactNode;
}

const FEATURED_CARDS: QuickCardDef[] = [
  {
    toolId: 'word-polish',
    title: 'Word chuẩn Nghị định 30/2020',
    desc: 'Tự động căn lề chuẩn NĐ 30, sửa bảng không rách trang, đánh số Page X/Y.',
    tag: 'Word Studio',
    accent: 'word',
    icon: <Sparkles className="h-5 w-5 text-blue-600 dark:text-blue-400" />,
  },
  {
    toolId: 'word-mail-merge',
    title: 'Trộn thư & Đọc tiền VN_WORDS',
    desc: 'Xuất hàng loạt hợp đồng/quyết định từ Excel kèm bộ lọc {{SoTien:vn_words}}.',
    tag: 'Word Studio',
    accent: 'word',
    icon: <Mail className="h-5 w-5 text-blue-600 dark:text-blue-400" />,
  },
  {
    toolId: 'pdf-pro',
    title: 'PDF Pro 17 Công cụ',
    desc: 'Ghép/tách trang, sắp xếp, ký tay, ký số PAdES, dấu giáp lai, nén & PDF/A-2b.',
    tag: 'PDF Pro',
    accent: 'pdf',
    icon: <Stamp className="h-5 w-5 text-red-600 dark:text-red-400" />,
  },
  {
    toolId: 'excel-power',
    title: 'Excel Power Suite (6 Công cụ)',
    desc: 'Làm sạch dữ liệu VN, tách Họ & Tên 3 cột, Pivot + Biểu đồ, nối XLOOKUP, Diff.',
    tag: 'Excel Suite',
    accent: 'excel',
    icon: <FileSpreadsheet className="h-5 w-5 text-emerald-600 dark:text-emerald-400" />,
  },
  {
    toolId: 'excel-tracker',
    title: 'Sổ Theo dõi Công việc Ngày',
    desc: 'Cập nhật đầu việc và đồng bộ Daily Log trực tiếp ra workbook Excel chuẩn.',
    tag: 'Excel Suite',
    accent: 'excel',
    icon: <FileCheck2 className="h-5 w-5 text-emerald-600 dark:text-emerald-400" />,
  },
  {
    toolId: 'excel-formulas',
    title: 'Cẩm nang 57 Công thức Excel',
    desc: 'Tra cứu hàm theo ngữ cảnh, copy nhanh công thức và tải file Demo chạy thực tế.',
    tag: 'Excel Suite',
    accent: 'excel',
    icon: <BookOpen className="h-5 w-5 text-emerald-600 dark:text-emerald-400" />,
  },
  {
    toolId: 'convert-universal',
    title: 'Chuyển đổi PDF ↔ Word ↔ MD',
    desc: 'Bảo toàn mục lục, bảng biểu và kiểu dáng qua lại giữa các định dạng văn phòng.',
    tag: 'Chuyển đổi',
    accent: 'studio',
    icon: <RefreshCw className="h-5 w-5 text-indigo-600 dark:text-indigo-400" />,
  },
  {
    toolId: 'ocr-vietnamese',
    title: 'OCR Tiếng Việt & Lọc Mộc đỏ',
    desc: 'Trích xuất văn bản từ PDF/ảnh scan, khử dấu mộc đỏ HSV và nắn thẳng trang.',
    tag: 'OCR',
    accent: 'studio',
    icon: <ScanText className="h-5 w-5 text-indigo-600 dark:text-indigo-400" />,
  },
  {
    toolId: 'qa-security',
    title: 'Kiểm định QA & Rà soát Bảo mật',
    desc: 'Chấm điểm OpenXML 13 cổng, gỡ Macro/OLE nguy hại và ẩn danh CCCD/STK.',
    tag: 'Bảo mật',
    accent: 'studio',
    icon: <ShieldAlert className="h-5 w-5 text-indigo-600 dark:text-indigo-400" />,
  },
  {
    toolId: 'presentation-studio',
    title: 'Trình chiếu PowerPoint 16:9',
    desc: 'Tổng hợp slide chuyên nghiệp từ Word/PDF/JSON kèm xem trước ảnh từng slide.',
    tag: 'Slide 16:9',
    accent: 'slide',
    icon: <Presentation className="h-5 w-5 text-orange-600 dark:text-orange-400" />,
  },
  {
    toolId: 'diagram-studio',
    title: 'Xưởng Sơ đồ Kỹ thuật',
    desc: 'Vẽ Mermaid, PlantUML, Draw.io XML, Canvas Spec hoặc sinh sơ đồ từ mã nguồn.',
    tag: 'Diagram',
    accent: 'studio',
    icon: <GitBranch className="h-5 w-5 text-indigo-600 dark:text-indigo-400" />,
  },
  {
    toolId: 'word-diff',
    title: 'So sánh Hợp đồng Redline Diff',
    desc: 'So sánh 2 bản Word và đánh dấu từng đoạn thêm/xóa chuẩn Track Changes.',
    tag: 'Word Studio',
    accent: 'word',
    icon: <FileText className="h-5 w-5 text-blue-600 dark:text-blue-400" />,
  },
];

export function HomeLauncherPage() {
  const { navigateToTool } = useWorkspace();
  const { sessionFiles } = useSessionFiles();
  const [droppedFile, setDroppedFile] = useState<SelectedFileInfo | null>(null);

  const recommended = droppedFile ? suggestToolsForExtension(droppedFile.ext) : [];

  const openWithDroppedFile = (toolId: ToolId) => {
    if (!droppedFile) {
      navigateToTool(toolId);
      return;
    }
    const matching = sessionFiles.find((sf) => sf.filePath === droppedFile.filePath) || {
      id: `sf_launch_${Date.now()}`,
      filename: droppedFile.filename,
      filePath: droppedFile.filePath,
      ext: droppedFile.ext,
      size: droppedFile.size,
      previewUrl: droppedFile.previewUrl,
      downloadUrl: droppedFile.downloadUrl,
      sourceLabel: 'Launcher',
      addedAt: Date.now(),
    };
    navigateToTool(toolId, matching);
  };

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-card dark:border-slate-800 dark:bg-slate-900">
        <div className="grid grid-cols-1 gap-6 lg:grid-cols-12 lg:items-center">
          <div className="space-y-3 lg:col-span-6">
            <div className="inline-flex items-center gap-1.5 rounded-full border border-emerald-200 bg-emerald-50 px-2.5 py-1 text-[11px] font-medium text-emerald-700 dark:border-emerald-900 dark:bg-emerald-950/50 dark:text-emerald-300">
              <ShieldCheck className="h-3.5 w-3.5" />
              <span>100% Thuần Local (127.0.0.1) — Dữ liệu không rời khỏi máy bạn</span>
            </div>
            <h1 className="text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
              Trạm Tự Động Hóa Tin Học Văn Phòng
            </h1>
            <p className="text-sm leading-relaxed text-slate-600 dark:text-slate-300">
              Thả bất kỳ tệp <span className="font-semibold text-blue-600">Word (.docx)</span>,{' '}
              <span className="font-semibold text-red-600">PDF</span>,{' '}
              <span className="font-semibold text-emerald-600">Excel (.xlsx)</span> hoặc{' '}
              <span className="font-semibold text-orange-600">Ảnh scan</span> vào ô bên cạnh để nhận gợi ý quy trình xử lý ngay, hoặc chọn công cụ trực tiếp bên dưới.
            </p>
          </div>

          <div className="lg:col-span-6">
            <FileDropZone
              label="Khởi chạy thông minh theo tệp"
              selectedFile={droppedFile}
              onFileSelected={(f) => setDroppedFile(f)}
              onClear={() => setDroppedFile(null)}
              sourceLabel="Trang chủ"
              hint="Hỗ trợ DOCX, PDF, XLSX, PPTX, PNG/JPG, JSON, MD"
            />
          </div>
        </div>

        {droppedFile && recommended.length > 0 ? (
          <div className="mt-4 border-t border-slate-100 pt-4 dark:border-slate-800">
            <div className="mb-2 text-xs font-semibold text-slate-700 dark:text-slate-200">
              Gợi ý xử lý cho <span className="font-mono text-indigo-600">{droppedFile.filename}</span>:
            </div>
            <div className="flex flex-wrap gap-2">
              {recommended.map((tool) => (
                <button
                  key={tool.id}
                  type="button"
                  onClick={() => openWithDroppedFile(tool.id)}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-medium text-slate-800 hover:border-indigo-400 hover:bg-indigo-50 hover:text-indigo-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100"
                >
                  <span>{tool.label}</span>
                  <ArrowRight className="h-3.5 w-3.5" />
                </button>
              ))}
            </div>
          </div>
        ) : null}
      </div>

      <div className="space-y-3">
        <div className="flex items-center justify-between">
          <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">
            Bộ Công Cụ Văn Phòng Chuyên Sâu
          </h2>
          <span className="text-xs text-slate-400">Nhấn Ctrl+K bất cứ lúc nào để nhảy nhanh</span>
        </div>

        <div className="grid grid-cols-1 gap-3.5 sm:grid-cols-2 lg:grid-cols-3">
          {FEATURED_CARDS.map((card) => {
            const accentStyle = ACCENT_CLASSES[card.accent];
            return (
              <button
                key={card.toolId}
                type="button"
                onClick={() => navigateToTool(card.toolId)}
                className="group flex flex-col justify-between rounded-xl border border-slate-200 bg-white p-4 text-left shadow-subtle transition-all hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-card dark:border-slate-800 dark:bg-slate-900 dark:hover:border-slate-700"
              >
                <div className="space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <div className={`flex h-9 w-9 items-center justify-center rounded-lg ${accentStyle.bgSoft}`}>
                      {card.icon}
                    </div>
                    <Badge tone={card.accent === 'word' ? 'blue' : card.accent === 'excel' ? 'green' : card.accent === 'pdf' ? 'red' : card.accent === 'slide' ? 'orange' : 'indigo'}>
                      {card.tag}
                    </Badge>
                  </div>
                  <div>
                    <h3 className="text-sm font-semibold text-slate-900 group-hover:text-indigo-600 dark:text-slate-100 dark:group-hover:text-indigo-400">
                      {card.title}
                    </h3>
                    <p className="mt-1 text-xs leading-relaxed text-slate-500 dark:text-slate-400">
                      {card.desc}
                    </p>
                  </div>
                </div>
                <div className="mt-3 flex items-center gap-1 text-[11px] font-medium text-slate-400 group-hover:text-indigo-600 dark:group-hover:text-indigo-400">
                  <span>Mở công cụ</span>
                  <ArrowRight className="h-3 w-3 transition-transform group-hover:translate-x-0.5" />
                </div>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}
