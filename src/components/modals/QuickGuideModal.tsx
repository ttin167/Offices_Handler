import { Keyboard, Lock, ShieldCheck, Sparkles } from 'lucide-react';
import { Modal } from '../ui';
import { useWorkspace } from '../../context/WorkspaceContext';

export function QuickGuideModal() {
  const { guideOpen, setGuideOpen } = useWorkspace();

  return (
    <Modal
      open={guideOpen}
      onClose={() => setGuideOpen(false)}
      title="Hướng Dẫn Nhanh — Antigravity Office Studio v5.0"
      subtitle="Bộ công cụ Tin học Văn phòng Chuyên sâu chạy hoàn toàn trên máy cá nhân (127.0.0.1)"
      size="lg"
    >
      <div className="space-y-4 text-xs leading-relaxed text-slate-700 dark:text-slate-300">
        <div className="rounded-xl border border-emerald-200 bg-emerald-50/70 p-3.5 dark:border-emerald-900/50 dark:bg-emerald-950/30">
          <div className="flex items-center gap-2 font-semibold text-emerald-900 dark:text-emerald-200">
            <Lock className="h-4 w-4" />
            <span>Cam kết Bảo mật Tuyệt đối (Local-Only)</span>
          </div>
          <p className="mt-1 text-emerald-800 dark:text-emerald-300">
            Ứng dụng chỉ lắng nghe nội bộ tại <code>127.0.0.1:8000</code>, chặn truy cập chéo từ trang web lạ và khóa Trợ lý AI ở chế độ RAG cục bộ. Nhấn nút{' '}
            <strong>&ldquo;Xóa vết phiên&rdquo;</strong> trên thanh tiêu đề để quét sạch file tạm khi xong việc.
          </p>
        </div>

        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
          <div className="rounded-xl border border-slate-200 p-3 dark:border-slate-800">
            <h4 className="font-semibold text-blue-700 dark:text-blue-400">1. Word Hành Chính (Blue)</h4>
            <p className="mt-1 text-slate-600 dark:text-slate-400">
              Chuẩn hóa theo <strong>Nghị định 30/2020/NĐ-CP</strong>, trộn thư hàng loạt từ Excel kèm{' '}
              <code>vn_words</code> (đọc số thành chữ), gộp/tách văn bản, sửa biến nhanh & so sánh Track Changes.
            </p>
          </div>
          <div className="rounded-xl border border-slate-200 p-3 dark:border-slate-800">
            <h4 className="font-semibold text-red-700 dark:text-red-400">2. PDF Pro 18 Công cụ (Red)</h4>
            <p className="mt-1 text-slate-600 dark:text-slate-400">
              Sắp xếp/xoay/chèn trang trắng 120 DPI, ghép/tách, ký vẽ tay & đóng dấu{' '}
              <strong>Giáp lai mộc đỏ</strong>, điền AcroForm, che PII vĩnh viễn, mã hóa AES-256 & ISO PDF/A-2b.
            </p>
          </div>
          <div className="rounded-xl border border-slate-200 p-3 dark:border-slate-800">
            <h4 className="font-semibold text-emerald-700 dark:text-emerald-400">3. Excel Power Suite (Green)</h4>
            <p className="mt-1 text-slate-600 dark:text-slate-400">
              Làm sạch SĐT/CCCD, tách Họ & Tên chuẩn Việt Nam, Pivot + Biểu đồ, XLOOKUP không công thức, Sổ tay 57 hàm & Nhật ký công việc phòng ban.
            </p>
          </div>
          <div className="rounded-xl border border-slate-200 p-3 dark:border-slate-800">
            <h4 className="font-semibold text-orange-600 dark:text-orange-400">4. Slide 16:9, OCR & Deep CDR</h4>
            <p className="mt-1 text-slate-600 dark:text-slate-400">
              Biến Word/PDF thành Slide 16:9, OCR Tiếng Việt khử dấu đỏ, Kiểm định 13 cổng QA & Tẩy rửa Macro độc hại.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
          <div className="flex items-center gap-2">
            <Keyboard className="h-4 w-4 text-indigo-500" />
            <span>
              Mẹo nhanh: Nhấn <kbd className="rounded border px-1.5 py-0.5 font-mono text-[10px]">Ctrl + K</kbd> để tìm và nhảy tới bất kỳ công cụ nào trong 1 giây.
            </span>
          </div>
          <div className="flex items-center gap-1 text-indigo-600 dark:text-indigo-400">
            <Sparkles className="h-3.5 w-3.5" />
            <span>Kho tệp phiên giúp chuyển tiếp file liền mạch</span>
          </div>
        </div>

        <div className="flex items-center gap-2 text-[11px] text-slate-500">
          <ShieldCheck className="h-3.5 w-3.5 text-emerald-600" />
          <span>Antigravity Pure Office Studio v5.0.0 — Pure Local Edition.</span>
        </div>
      </div>
    </Modal>
  );
}
