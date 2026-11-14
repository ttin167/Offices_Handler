import type { DomainAccent, ToolId, ToolNavItem } from '../types/workspace';

export interface DomainGroup {
  id: 'home' | 'word' | 'pdf' | 'excel' | 'convert' | 'visual';
  label: string;
  accent: DomainAccent;
}

export const DOMAIN_GROUPS: DomainGroup[] = [
  { id: 'home', label: 'Khởi đầu nhanh', accent: 'neutral' },
  { id: 'word', label: 'Word Studio (Soạn thảo)', accent: 'word' },
  { id: 'pdf', label: 'PDF Pro (17 Công cụ)', accent: 'pdf' },
  { id: 'excel', label: 'Excel Power Suite', accent: 'excel' },
  { id: 'convert', label: 'Chuyển đổi, OCR & QA', accent: 'studio' },
  { id: 'visual', label: 'Trình chiếu & Sơ đồ', accent: 'slide' },
];

export const TOOL_CATALOG: ToolNavItem[] = [
  {
    id: 'home',
    label: 'Trang chủ Launcher',
    shortDesc: 'Chọn nhanh công cụ hoặc thả tệp để hệ thống gợi ý thao tác phù hợp',
    domain: 'home',
    accent: 'neutral',
    keywords: ['home', 'launcher', 'trang chu', 'bat dau'],
    supportedExts: ['docx', 'pdf', 'xlsx', 'xls', 'xlsm', 'pptx', 'md', 'png', 'jpg'],
  },
  {
    id: 'word-polish',
    label: 'Chuẩn hóa Nghị định 30',
    shortDesc: 'Căn lề hành chính NĐ 30/2020, tự đánh số Page X/Y, sửa bảng & giới hạn ảnh 14cm',
    domain: 'word',
    accent: 'word',
    keywords: ['word', 'nd30', 'nghi dinh 30', 'polish', 'chuan hoa', 'page x/y'],
    supportedExts: ['docx'],
  },
  {
    id: 'word-merge-split',
    label: 'Gộp, Tách & Đóng dấu Word',
    shortDesc: 'Gộp nhiều tệp DOCX giữ nguyên ảnh, tách chương ra ZIP, chèn Watermark & Header/Footer',
    domain: 'word',
    accent: 'word',
    keywords: ['word', 'merge', 'split', 'gop', 'tach', 'watermark', 'header', 'footer'],
    supportedExts: ['docx'],
  },
  {
    id: 'word-mail-merge',
    label: 'Trộn thư & Ánh xạ Trường',
    shortDesc: 'Trộn hợp đồng hàng loạt từ Excel/CSV/JSON, hỗ trợ đọc tiền {{SoTien:vn_words}}',
    domain: 'word',
    accent: 'word',
    keywords: ['mail merge', 'tron thu', 'vn_words', 'doc tien', 'anh xa', 'hop dong'],
    supportedExts: ['docx', 'xlsx', 'csv', 'json'],
  },
  {
    id: 'word-quick-patch',
    label: 'Điền mẫu nhanh ({{KEY}})',
    shortDesc: 'Quét tự động biến {{KEY}} và bảng lặp trong template Word để điền biểu mẫu tốc độ cao',
    domain: 'word',
    accent: 'word',
    keywords: ['patch', 'dien mau', 'token', 'bieu mau', 'word'],
    supportedExts: ['docx'],
  },
  {
    id: 'word-diff',
    label: 'So sánh & Track Changes',
    shortDesc: 'Đối chiếu 2 phiên bản hợp đồng/văn bản Word, xuất Track Changes đỏ xanh',
    domain: 'word',
    accent: 'word',
    keywords: ['diff', 'redline', 'so sanh', 'track changes', 'word'],
    supportedExts: ['docx'],
  },
  {
    id: 'pdf-pro',
    label: 'PDF Pro 17 Công cụ',
    shortDesc: 'Ghép/tách trang, sắp xếp trang, ký tay, ký số PAdES, dấu giáp lai, mật khẩu AES-256, PDF/A-2b',
    domain: 'pdf',
    accent: 'pdf',
    keywords: ['pdf', 'ky ten', 'giap lai', 'nen', 'ghep', 'tach', 'redact', 'aes', 'pdfa', 'acroform'],
    supportedExts: ['pdf', 'png', 'jpg', 'jpeg', 'xlsx'],
  },
  {
    id: 'excel-power',
    label: 'Excel Power Suite (6 Công cụ)',
    shortDesc: 'Làm sạch dữ liệu VN, Tách Họ/Tên 3 phần, Pivot + Biểu đồ, Nối XLOOKUP, Tách file ZIP, So sánh ô',
    domain: 'excel',
    accent: 'excel',
    keywords: ['excel', 'power', 'clean', 'trim', 'tach ten', 'pivot', 'xlookup', 'diff', 'xlsx'],
    supportedExts: ['xlsx', 'xls', 'xlsm'],
  },
  {
    id: 'excel-reporting',
    label: 'Báo cáo & Hợp nhất Excel',
    shortDesc: 'Tổng hợp nhiều workbook và chuyển đổi bảng tính Excel thành báo cáo Word/PDF tóm tắt',
    domain: 'excel',
    accent: 'excel',
    keywords: ['excel', 'brief', 'bao cao', 'consolidate', 'hop nhat', 'xlsx'],
    supportedExts: ['xlsx', 'xls'],
  },
  {
    id: 'excel-tracker',
    label: 'Sổ Theo dõi Công việc',
    shortDesc: 'Khởi tạo, cập nhật đầu việc và đồng bộ tiến độ ngày trên project_tracking.xlsx & task_tracking.xlsx',
    domain: 'excel',
    accent: 'excel',
    keywords: ['tracker', 'so cong viec', 'tien do', 'task', 'project_tracking'],
    supportedExts: ['xlsx'],
  },
  {
    id: 'excel-formulas',
    label: 'Cẩm nang 57 Công thức',
    shortDesc: 'Tra cứu 57 hàm Excel thực chiến (XLOOKUP, SUMIFS, INDEX/MATCH...), copy 1-click & xuất file Demo',
    domain: 'excel',
    accent: 'excel',
    keywords: ['formula', 'cong thuc', '57', 'xlookup', 'sumifs', 'vlookup', 'playbook'],
    supportedExts: ['xlsx'],
  },
  {
    id: 'excel-template',
    label: 'Mẫu Excel, AST & Khổ in',
    shortDesc: 'Điền dữ liệu giữ nguyên 100% template, kiểm định 12 cổng, tính giá trị công thức AST & khổ in A4',
    domain: 'excel',
    accent: 'excel',
    keywords: ['template', 'mutate', 'validate', 'ast', 'eval', 'print setup', '12 gates'],
    supportedExts: ['xlsx', 'xlsm'],
  },
  {
    id: 'convert-universal',
    label: 'Chuyển đổi Đa định dạng',
    shortDesc: 'Chuyển đổi lossless 2 chiều PDF ↔ Word ↔ Markdown ↔ PowerPoint ↔ Excel',
    domain: 'convert',
    accent: 'studio',
    keywords: ['convert', 'chuyen doi', 'pdf to word', 'docx to pdf', 'markdown'],
    supportedExts: ['pdf', 'docx', 'md', 'pptx', 'xlsx', 'html'],
  },
  {
    id: 'ocr-vietnamese',
    label: 'OCR Tiếng Việt & Lọc Mộc đỏ',
    shortDesc: 'Nhận dạng chữ tiếng Việt từ PDF/ảnh scan, khử dấu mộc đỏ HSV và nắn thẳng trang nghiêng',
    domain: 'convert',
    accent: 'studio',
    keywords: ['ocr', 'scan', 'moc do', 'tieng viet', 'tesseract', 'deskew'],
    supportedExts: ['pdf', 'png', 'jpg', 'jpeg', 'tif', 'tiff'],
  },
  {
    id: 'qa-security',
    label: 'Kiểm định QA & Bảo mật CDR',
    shortDesc: 'Kiểm định 13 cổng OpenXML, rà soát VBA macro nguy hại, Deep CDR & ẩn danh PII (CCCD/SĐT/STK)',
    domain: 'convert',
    accent: 'studio',
    keywords: ['qa', 'audit', 'security', 'bao mat', 'cdr', 'macro', 'pii', 'cccd'],
    supportedExts: ['docx', 'xlsx', 'xlsm', 'docm', 'pdf', 'pptx', 'csv', 'zip'],
  },
  {
    id: 'presentation-studio',
    label: 'Trình chiếu PowerPoint 16:9',
    shortDesc: 'Tổng hợp slide 16:9 tự động từ tài liệu Word/PDF hoặc JSON Spec kèm xem trước từng trang',
    domain: 'visual',
    accent: 'slide',
    keywords: ['pptx', 'slide', 'powerpoint', 'trinh chieu', '16:9', 'deck'],
    supportedExts: ['docx', 'pdf', 'md', 'pptx', 'json'],
  },
  {
    id: 'diagram-studio',
    label: 'Xưởng Sơ đồ Kỹ thuật',
    shortDesc: 'Vẽ Mermaid, PlantUML, Draw.io XML, Canvas Spec và phân tích mã nguồn Python/SQL ra SVG/PNG 300 DPI',
    domain: 'visual',
    accent: 'studio',
    keywords: ['diagram', 'so do', 'mermaid', 'plantuml', 'drawio', 'erd', 'flowchart', 'py', 'sql'],
    supportedExts: ['json', 'mmd', 'puml', 'drawio', 'py', 'sql'],
  },
];

export function getToolMeta(id: ToolId): ToolNavItem {
  return TOOL_CATALOG.find((item) => item.id === id) || TOOL_CATALOG[0];
}

export function suggestToolsForExtension(ext: string): ToolNavItem[] {
  const clean = ext.toLowerCase().replace('.', '');
  if (!clean) return [];
  return TOOL_CATALOG.filter((t) => t.id !== 'home' && t.supportedExts.includes(clean));
}

export const ACCENT_CLASSES: Record<
  DomainAccent,
  { badge: string; text: string; bgSoft: string; border: string; button: string }
> = {
  neutral: {
    badge: 'bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300',
    text: 'text-slate-700 dark:text-slate-200',
    bgSoft: 'bg-slate-100/80 dark:bg-slate-800/70',
    border: 'border-slate-200 dark:border-slate-700',
    button: 'bg-slate-900 hover:bg-slate-800 text-white dark:bg-slate-100 dark:text-slate-900',
  },
  word: {
    badge: 'bg-blue-50 text-blue-700 border-blue-200 dark:bg-blue-950/60 dark:text-blue-300 dark:border-blue-800',
    text: 'text-blue-600 dark:text-blue-400',
    bgSoft: 'bg-blue-50 dark:bg-blue-950/40',
    border: 'border-blue-200 dark:border-blue-800',
    button: 'bg-blue-600 hover:bg-blue-700 text-white',
  },
  excel: {
    badge: 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/60 dark:text-emerald-300 dark:border-emerald-800',
    text: 'text-emerald-600 dark:text-emerald-400',
    bgSoft: 'bg-emerald-50 dark:bg-emerald-950/40',
    border: 'border-emerald-200 dark:border-emerald-800',
    button: 'bg-emerald-600 hover:bg-emerald-700 text-white',
  },
  pdf: {
    badge: 'bg-red-50 text-red-700 border-red-200 dark:bg-red-950/60 dark:text-red-300 dark:border-red-800',
    text: 'text-red-600 dark:text-red-400',
    bgSoft: 'bg-red-50 dark:bg-red-950/40',
    border: 'border-red-200 dark:border-red-800',
    button: 'bg-red-600 hover:bg-red-700 text-white',
  },
  slide: {
    badge: 'bg-orange-50 text-orange-700 border-orange-200 dark:bg-orange-950/60 dark:text-orange-300 dark:border-orange-800',
    text: 'text-orange-600 dark:text-orange-400',
    bgSoft: 'bg-orange-50 dark:bg-orange-950/40',
    border: 'border-orange-200 dark:border-orange-800',
    button: 'bg-orange-600 hover:bg-orange-700 text-white',
  },
  studio: {
    badge: 'bg-indigo-50 text-indigo-700 border-indigo-200 dark:bg-indigo-950/60 dark:text-indigo-300 dark:border-indigo-800',
    text: 'text-indigo-600 dark:text-indigo-400',
    bgSoft: 'bg-indigo-50 dark:bg-indigo-950/40',
    border: 'border-indigo-200 dark:border-indigo-800',
    button: 'bg-indigo-600 hover:bg-indigo-700 text-white',
  },
};
