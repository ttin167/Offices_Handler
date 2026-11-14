import { useEffect, useMemo, useState } from 'react';
import { BookOpen, Copy, Download, FileSpreadsheet, Search } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import { Badge, Button, EmptyState } from '../../components/ui';
import { useToast } from '../../context/ToastContext';
import { apiGet } from '../../services/api';

interface FormulaEntry {
  id: string;
  name: string;
  category: string;
  category_vi: string;
  template: string;
  example: string;
  when_to_use: string;
  why_choose: string;
}

interface CatalogResponse {
  success: boolean;
  count: number;
  formulas: FormulaEntry[];
}

const CATEGORIES = [
  { id: 'all', label: 'Tất cả (57)' },
  { id: 'lookup', label: 'Tra cứu & Tham chiếu' },
  { id: 'aggregation', label: 'Toán học & Thống kê' },
  { id: 'text', label: 'Xử lý Chuỗi' },
  { id: 'logic', label: 'Logic & Kiểm soát lỗi' },
  { id: 'dynamic', label: 'Mảng động (Excel 365)' },
  { id: 'financial', label: 'Thời gian & Tài chính' },
];

export function FormulaPlaybookPage() {
  const toast = useToast();

  const [formulas, setFormulas] = useState<FormulaEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [category, setCategory] = useState('all');
  const [search, setSearch] = useState('');
  const [selected, setSelected] = useState<FormulaEntry | null>(null);

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        const res = await apiGet<CatalogResponse>('/api/xlsx/formulas?category=all');
        if (!active) return;
        const list = res.formulas || [];
        setFormulas(list);
        if (list.length > 0) setSelected(list[0]);
      } catch (err) {
        if (active) toast.error('Không tải được danh mục công thức Excel', err instanceof Error ? err.message : String(err));
      } finally {
        if (active) setLoading(false);
      }
    })();
    return () => {
      active = false;
    };
  }, [toast]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return formulas.filter((f) => {
      if (category !== 'all' && f.category !== category) return false;
      if (!q) return true;
      return (
        f.name.toLowerCase().includes(q) ||
        f.template.toLowerCase().includes(q) ||
        f.when_to_use.toLowerCase().includes(q) ||
        f.category_vi.toLowerCase().includes(q)
      );
    });
  }, [formulas, category, search]);

  const copyText = async (text: string, label: string) => {
    try {
      await navigator.clipboard.writeText(text);
      toast.success(`Đã sao chép ${label}`);
    } catch {
      toast.warning('Trình duyệt không hỗ trợ clipboard tự động');
    }
  };

  const leftPanel = (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <div className="relative flex-1">
          <Search className="pointer-events-none absolute left-3 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Tìm hàm (XLOOKUP, SUMIFS, DATEDIF, FILTER...)"
            className="w-full rounded-lg border border-slate-200 bg-white py-2 pl-8 pr-3 text-xs text-slate-900 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
          />
        </div>
        <a
          href="/api/xlsx/formulas/export"
          download="Excel_Formula_Playbook.xlsx"
          className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-600 px-3 py-2 text-xs font-medium text-white hover:bg-emerald-700"
        >
          <Download className="h-3.5 w-3.5" />
          <span>Tải File Mẫu (.xlsx)</span>
        </a>
      </div>

      <div className="flex flex-wrap gap-1.5">
        {CATEGORIES.map((c) => {
          const active = category === c.id;
          return (
            <button
              key={c.id}
              type="button"
              onClick={() => setCategory(c.id)}
              className={`rounded-lg border px-2.5 py-1 text-[11px] font-medium transition ${
                active
                  ? 'border-emerald-500 bg-emerald-50 text-emerald-700 dark:border-emerald-700 dark:bg-emerald-950/50 dark:text-emerald-300'
                  : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300'
              }`}
            >
              {c.label}
            </button>
          );
        })}
      </div>

      {loading ? (
        <div className="py-8 text-center text-xs text-slate-500">Đang nạp 57 công thức Excel tinh hoa...</div>
      ) : filtered.length === 0 ? (
        <EmptyState title="Không tìm thấy công thức phù hợp" description="Thử xóa từ khóa hoặc chuyển về nhóm Tất cả." />
      ) : (
        <div className="max-h-[480px] space-y-1.5 overflow-y-auto pr-1">
          {filtered.map((item) => {
            const isSel = selected?.id === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => setSelected(item)}
                className={`flex w-full items-start justify-between gap-2 rounded-xl border px-3 py-2.5 text-left transition ${
                  isSel
                    ? 'border-emerald-500 bg-emerald-50/70 shadow-subtle dark:border-emerald-700 dark:bg-emerald-950/30'
                    : 'border-slate-200 bg-white hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:hover:bg-slate-800/70'
                }`}
              >
                <div className="min-w-0">
                  <div className="flex items-center gap-2">
                    <span className="font-mono text-xs font-bold text-emerald-700 dark:text-emerald-400">
                      {item.name}
                    </span>
                    <Badge tone="slate">{item.category_vi}</Badge>
                  </div>
                  <p className="mt-1 line-clamp-2 text-[11px] text-slate-600 dark:text-slate-400">
                    {item.when_to_use}
                  </p>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );

  const rightPanel = selected ? (
    <div className="space-y-4 rounded-xl border border-slate-200 bg-white p-4 shadow-card dark:border-slate-800 dark:bg-slate-900">
      <div className="flex flex-wrap items-start justify-between gap-2 border-b border-slate-100 pb-3 dark:border-slate-800">
        <div>
          <div className="flex items-center gap-2">
            <BookOpen className="h-4 w-4 text-emerald-600" />
            <h3 className="font-mono text-lg font-bold text-slate-900 dark:text-slate-100">{selected.name}</h3>
          </div>
          <p className="mt-0.5 text-xs text-slate-500">{selected.category_vi}</p>
        </div>
        <Badge tone="green">Playbook Văn phòng</Badge>
      </div>

      <div className="space-y-1.5">
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold text-slate-700 dark:text-slate-300">Cú pháp chuẩn</span>
          <Button
            size="sm"
            variant="ghost"
            icon={<Copy className="h-3.5 w-3.5" />}
            onClick={() => copyText(selected.template, 'cú pháp')}
          >
            Sao chép cú pháp
          </Button>
        </div>
        <pre className="overflow-x-auto rounded-lg border border-slate-200 bg-slate-50 px-3 py-2.5 font-mono text-xs text-slate-900 dark:border-slate-800 dark:bg-slate-950 dark:text-emerald-300">
          {selected.template}
        </pre>
      </div>

      <div className="space-y-1.5">
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold text-slate-700 dark:text-slate-300">Ví dụ thực chiến</span>
          <Button
            size="sm"
            variant="ghost"
            icon={<Copy className="h-3.5 w-3.5" />}
            onClick={() => copyText(selected.example, 'ví dụ công thức')}
          >
            Sao chép ví dụ
          </Button>
        </div>
        <pre className="overflow-x-auto rounded-lg border border-emerald-200 bg-emerald-50/50 px-3 py-2.5 font-mono text-xs text-emerald-950 dark:border-emerald-900/50 dark:bg-emerald-950/30 dark:text-emerald-200">
          {selected.example}
        </pre>
      </div>

      <div className="space-y-2 rounded-lg border border-slate-100 bg-slate-50/70 p-3 dark:border-slate-800 dark:bg-slate-950/60">
        <h4 className="text-xs font-semibold text-slate-800 dark:text-slate-200">Khi nào nên sử dụng?</h4>
        <p className="text-xs leading-relaxed text-slate-600 dark:text-slate-400">{selected.when_to_use}</p>
      </div>

      <div className="space-y-2 rounded-lg border border-indigo-100 bg-indigo-50/40 p-3 dark:border-indigo-900/50 dark:bg-indigo-950/20">
        <h4 className="text-xs font-semibold text-indigo-900 dark:text-indigo-200">
          Tại sao chọn hàm này thay thế cách cũ?
        </h4>
        <p className="text-xs leading-relaxed text-slate-700 dark:text-slate-300">{selected.why_choose}</p>
      </div>
    </div>
  ) : (
    <EmptyState
      icon={<FileSpreadsheet className="h-6 w-6" />}
      title="Chọn một công thức bên trái để xem chi tiết"
      description="Bao gồm cú pháp, ví dụ thực tế và lý giải tình huống áp dụng."
    />
  );

  return (
    <SplitWorkspace
      title="Sổ Tay 57 Công Thức Excel Tinh Hoa"
      subtitle="Tra cứu nhanh cú pháp, ví dụ văn phòng thực tế (XLOOKUP, SUMIFS, FILTER, PMT...) và tải file Excel mẫu kèm sẵn công thức chạy thử."
      badgeText="57 Formulas"
      accent="excel"
      leftPanel={leftPanel}
      resultPanel={rightPanel}
      rightHeaderExtra={<Badge tone="green">Có file Excel tương tác</Badge>}
    />
  );
}
