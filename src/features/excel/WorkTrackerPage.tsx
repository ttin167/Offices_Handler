import { useCallback, useEffect, useState } from 'react';
import { CalendarCheck2, CheckSquare, FileSpreadsheet, PlusCircle, RefreshCw } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  EmptyState,
  Tabs,
} from '../../components/ui';
import { useToast } from '../../context/ToastContext';
import { apiGet, apiPost } from '../../services/api';

interface TrackerEntry {
  row: number;
  stt: number | string;
  date: string;
  module: string;
  code: string;
  name: string;
  type: string;
  rules: string;
  ui: string;
  status: string;
  notes: string;
}

interface TrackerInspectResponse {
  success?: boolean;
  ok?: boolean;
  project_file: string;
  project_exists: boolean;
  total_entries: number;
  status_counts: Record<string, number>;
  recent_entries: TrackerEntry[];
  task_file: string;
  task_exists: boolean;
  task_sheets: string[];
}

interface DailySyncResponse {
  success?: boolean;
  date: string;
  entries_matched: number;
  injected_to_task_file: boolean;
  target_sheet?: string | null;
  tasks: Array<{ stt: string; module: string; description: string; status: string; notes: string }>;
}

const MODULES = [
  'Kế toán - Tài chính',
  'Hành chính - Nhân sự',
  'Vận hành & Kho',
  'Kinh doanh & CSKH',
  'Báo cáo Tổng hợp',
];

const CHANGE_TYPES = ['Tạo mới', 'Bổ sung', 'Chỉnh sửa', 'Hủy bỏ'];
const STATUSES = ['Đang soạn thảo', 'Chờ Review', 'Đã duyệt', 'Đã chốt'];

export function WorkTrackerPage() {
  const toast = useToast();

  const [tab, setTab] = useState<'add' | 'sync'>('add');
  const [inspectData, setInspectData] = useState<TrackerInspectResponse | null>(null);
  const [loadingInspect, setLoadingInspect] = useState(false);

  // Add entry state
  const [moduleName, setModuleName] = useState(MODULES[0]);
  const [code, setCode] = useState('FIN-01');
  const [name, setName] = useState('');
  const [changeType, setChangeType] = useState(CHANGE_TYPES[0]);
  const [rules, setRules] = useState('');
  const [statusVal, setStatusVal] = useState(STATUSES[0]);
  const [uiText, setUiText] = useState('');
  const [notes, setNotes] = useState('');
  const [dateStr, setDateStr] = useState('');

  // Sync daily
  const [syncDate, setSyncDate] = useState('');
  const [applyToTaskFile, setApplyToTaskFile] = useState(false);
  const [dailyResult, setDailyResult] = useState<DailySyncResponse | null>(null);

  const [submitting, setSubmitting] = useState(false);

  const loadStatus = useCallback(async () => {
    setLoadingInspect(true);
    try {
      const res = await apiGet<TrackerInspectResponse>('/api/xlsx/tracker/inspect');
      setInspectData(res);
    } catch (err) {
      toast.error('Không đọc được Sổ Theo Dõi', err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingInspect(false);
    }
  }, [toast]);

  useEffect(() => {
    loadStatus();
  }, [loadStatus]);

  const handleInitWorkbook = async (force = false) => {
    setSubmitting(true);
    try {
      const res = await apiPost<{ ok?: boolean; created?: boolean; message?: string; path?: string }>(
        '/api/xlsx/tracker/init',
        { force },
      );
      toast.success('Đã khởi tạo sổ theo dõi Excel', res.message || res.path);
      await loadStatus();
    } catch (err) {
      toast.error('Lỗi khởi tạo sổ', err instanceof Error ? err.message : String(err));
    } finally {
      setSubmitting(false);
    }
  };

  const handleAddEntry = async () => {
    if (!code.trim() || !name.trim() || !rules.trim()) {
      toast.warning('Thiếu thông tin bắt buộc', 'Vui lòng nhập đủ Mã việc, Tên đầu việc và Nội dung xử lý.');
      return;
    }
    setSubmitting(true);
    try {
      await apiPost('/api/xlsx/tracker/add', {
        module: moduleName.trim(),
        code: code.trim(),
        name: name.trim(),
        change_type: changeType,
        rules: rules.trim(),
        ui: uiText.trim(),
        status: statusVal,
        notes: notes.trim(),
        date_str: dateStr.trim() || undefined,
      });
      toast.success('Đã ghi đầu việc vào project_tracking.xlsx', `${code} — ${name}`);
      setName('');
      setRules('');
      setUiText('');
      setNotes('');
      await loadStatus();
    } catch (err) {
      toast.error('Lỗi ghi đầu việc', err instanceof Error ? err.message : String(err));
    } finally {
      setSubmitting(false);
    }
  };

  const handleSyncDaily = async () => {
    setSubmitting(true);
    try {
      const res = await apiPost<DailySyncResponse>('/api/xlsx/tracker/sync-daily', {
        date_str: syncDate.trim() || undefined,
        apply_to_task_tracking: applyToTaskFile,
      });
      setDailyResult(res);
      toast.success(
        `Đã tổng hợp ${res.entries_matched} việc ngày ${res.date}`,
        res.injected_to_task_file ? `Đã chèn vào sheet ${res.target_sheet}` : undefined,
      );
    } catch (err) {
      toast.error('Lỗi tổng hợp báo cáo ngày', err instanceof Error ? err.message : String(err));
    } finally {
      setSubmitting(false);
    }
  };

  const leftPanel = (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-emerald-200 bg-emerald-50/70 px-3.5 py-2.5 dark:border-emerald-900/50 dark:bg-emerald-950/30">
        <div className="text-xs">
          <div className="font-semibold text-emerald-900 dark:text-emerald-200">
            {inspectData?.project_exists ? 'Sổ theo dõi đã sẵn sàng (Local)' : 'Chưa có file project_tracking.xlsx'}
          </div>
          <div className="text-[11px] text-emerald-700 dark:text-emerald-400">
            Tổng số: <strong>{inspectData?.total_entries ?? 0}</strong> đầu việc đã lưu
          </div>
        </div>
        <div className="flex items-center gap-1.5">
          <Button
            size="sm"
            variant="outline"
            loading={loadingInspect}
            icon={<RefreshCw className="h-3.5 w-3.5" />}
            onClick={loadStatus}
          >
            Tải lại
          </Button>
          <Button
            size="sm"
            variant="primary"
            loading={submitting}
            onClick={() => handleInitWorkbook(false)}
            className="bg-emerald-600 hover:bg-emerald-700"
          >
            {inspectData?.project_exists ? 'Kiểm tra / Khởi tạo' : 'Tạo Sổ Excel mẫu'}
          </Button>
        </div>
      </div>

      <Tabs
        activeId={tab}
        onChange={(id) => setTab(id as 'add' | 'sync')}
        className="w-full"
        options={[
          { id: 'add', label: 'Ghi đầu việc mới', icon: <PlusCircle className="h-3.5 w-3.5" /> },
          { id: 'sync', label: 'Tổng hợp Báo cáo Ngày', icon: <CalendarCheck2 className="h-3.5 w-3.5" /> },
        ]}
      />

      {tab === 'add' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Bộ phận / Nhóm việc</label>
              <select
                value={moduleName}
                onChange={(e) => setModuleName(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                {MODULES.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Mã đầu việc</label>
              <input
                type="text"
                value={code}
                onChange={(e) => setCode(e.target.value)}
                placeholder="VD: FIN-01, HR-05"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Tên đầu việc / Hồ sơ</label>
              <input
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="VD: Tổng hợp bảng chấm công tháng 11"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Trạng thái</label>
              <select
                value={statusVal}
                onChange={(e) => setStatusVal(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                {STATUSES.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Nội dung nghiệp vụ / Kết quả cần đạt
            </label>
            <textarea
              rows={3}
              value={rules}
              onChange={(e) => setRules(e.target.value)}
              placeholder="Mô tả chi tiết nội dung đã thực hiện hoặc quy tắc cần tuân thủ..."
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </div>

          <CollapsibleSection title="Tùy chọn nâng cao (Phân loại, Ngày, Đầu ra & Ghi chú)">
            <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
              <select
                value={changeType}
                onChange={(e) => setChangeType(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                {CHANGE_TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
              <input
                type="text"
                value={dateStr}
                onChange={(e) => setDateStr(e.target.value)}
                placeholder="Ngày (DD/MM/YYYY - mặc định hôm nay)"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
              <input
                type="text"
                value={uiText}
                onChange={(e) => setUiText(e.target.value)}
                placeholder="Tệp / Chứng từ liên quan"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
              <input
                type="text"
                value={notes}
                onChange={(e) => setNotes(e.target.value)}
                placeholder="Ghi chú thêm..."
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </CollapsibleSection>

          <Button
            variant="primary"
            loading={submitting}
            icon={<PlusCircle className="h-4 w-4" />}
            onClick={handleAddEntry}
            className="w-full bg-emerald-600 hover:bg-emerald-700"
          >
            Lưu đầu việc vào Excel
          </Button>
        </div>
      ) : (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Ngày cần tổng hợp báo cáo (DD/MM/YYYY)
            </label>
            <input
              type="text"
              value={syncDate}
              onChange={(e) => setSyncDate(e.target.value)}
              placeholder="Để trống để lấy ngày hôm nay"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </div>
          <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
            <input
              type="checkbox"
              checked={applyToTaskFile}
              onChange={(e) => setApplyToTaskFile(e.target.checked)}
              disabled={!inspectData?.task_exists}
            />
            <span>
              Tự động ghi vào file chấm công/nhật ký ngày (`task_tracking.xlsx`)
              {!inspectData?.task_exists ? ' — (Chưa có file task_tracking.xlsx)' : ''}
            </span>
          </label>
          <Button
            variant="primary"
            loading={submitting}
            icon={<CalendarCheck2 className="h-4 w-4" />}
            onClick={handleSyncDaily}
            className="w-full bg-emerald-600 hover:bg-emerald-700"
          >
            Tổng hợp công việc trong ngày
          </Button>
        </div>
      )}
    </div>
  );

  const rightPanel = (
    <div className="space-y-4">
      {inspectData?.status_counts && Object.keys(inspectData.status_counts).length > 0 ? (
        <div className="grid grid-cols-2 gap-2.5 sm:grid-cols-4">
          {Object.entries(inspectData.status_counts).map(([statusLabel, count]) => (
            <div
              key={statusLabel}
              className="rounded-xl border border-slate-200 bg-white px-3 py-2.5 dark:border-slate-800 dark:bg-slate-900"
            >
              <div className="text-[11px] text-slate-500 dark:text-slate-400">{statusLabel}</div>
              <div className="mt-1 text-lg font-bold text-slate-900 dark:text-slate-100">{count}</div>
            </div>
          ))}
        </div>
      ) : null}

      {dailyResult ? (
        <div className="space-y-2 rounded-xl border border-emerald-200 bg-white p-3.5 dark:border-emerald-900/50 dark:bg-slate-900">
          <div className="flex items-center justify-between">
            <h4 className="flex items-center gap-1.5 text-xs font-semibold text-emerald-800 dark:text-emerald-300">
              <CheckSquare className="h-4 w-4" />
              <span>Báo cáo tổng hợp ngày {dailyResult.date} ({dailyResult.entries_matched} việc)</span>
            </h4>
            {dailyResult.injected_to_task_file ? <Badge tone="green">Đã đồng bộ sheet {dailyResult.target_sheet}</Badge> : null}
          </div>
          {dailyResult.tasks.length === 0 ? (
            <p className="text-xs text-slate-500">Không có đầu việc nào ghi nhận trong ngày này.</p>
          ) : (
            <ul className="space-y-1.5 text-xs">
              {dailyResult.tasks.map((t, i) => (
                <li
                  key={i}
                  className="flex items-start justify-between gap-2 rounded-lg border border-slate-100 bg-slate-50 px-2.5 py-2 dark:border-slate-800 dark:bg-slate-950"
                >
                  <div>
                    <div className="font-medium text-slate-900 dark:text-slate-100">{t.description}</div>
                    <div className="text-[11px] text-slate-500">{t.module} — {t.notes}</div>
                  </div>
                  <Badge tone={t.status === 'Done' ? 'green' : 'amber'}>{t.status}</Badge>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : null}

      <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center justify-between border-b border-slate-100 px-3.5 py-2.5 dark:border-slate-800">
          <span className="flex items-center gap-1.5 text-xs font-semibold text-slate-800 dark:text-slate-200">
            <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
            15 đầu việc gần nhất trong sổ Excel
          </span>
          <span className="text-[11px] text-slate-400">
            {inspectData?.project_file ? inspectData.project_file.split(/[/\\]/).pop() : 'project_tracking.xlsx'}
          </span>
        </div>

        {!inspectData || inspectData.recent_entries.length === 0 ? (
          <EmptyState
            title="Chưa có bản ghi công việc nào"
            description="Nhấn 'Tạo Sổ Excel mẫu' hoặc thêm đầu việc mới ở cột bên trái để bắt đầu."
          />
        ) : (
          <div className="max-h-[420px] overflow-x-auto overflow-y-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead className="sticky top-0 bg-slate-50 text-[11px] text-slate-500 dark:bg-slate-950 dark:text-slate-400">
                <tr className="border-b border-slate-200 dark:border-slate-800">
                  <th className="px-3 py-2">Ngày</th>
                  <th className="px-3 py-2">Mã</th>
                  <th className="px-3 py-2">Đầu việc & Bộ phận</th>
                  <th className="px-3 py-2">Trạng thái</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                {inspectData.recent_entries.map((entry) => (
                  <tr key={`${entry.row}-${entry.code}`} className="hover:bg-slate-50/70 dark:hover:bg-slate-800/40">
                    <td className="whitespace-nowrap px-3 py-2 text-slate-500">{entry.date}</td>
                    <td className="whitespace-nowrap px-3 py-2 font-mono text-[11px] font-semibold text-emerald-700 dark:text-emerald-400">
                      {entry.code}
                    </td>
                    <td className="px-3 py-2">
                      <div className="font-medium text-slate-900 dark:text-slate-100">{entry.name}</div>
                      <div className="line-clamp-1 text-[11px] text-slate-500">
                        {entry.module} · {entry.rules}
                      </div>
                    </td>
                    <td className="whitespace-nowrap px-3 py-2">
                      <Badge
                        tone={
                          entry.status === 'Đã duyệt' || entry.status === 'Đã chốt'
                            ? 'green'
                            : entry.status === 'Chờ Review'
                              ? 'amber'
                              : 'slate'
                        }
                      >
                        {entry.status}
                      </Badge>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );

  return (
    <SplitWorkspace
      title="Sổ Theo Dõi Công Việc & Báo Cáo Ngày"
      subtitle="Quản lý nhật ký đầu việc phòng ban trên file Excel chuẩn công thức COUNTIFS động và tổng hợp nhanh báo cáo cuối ngày."
      badgeText="Office Work Tracker"
      accent="excel"
      leftPanel={leftPanel}
      resultPanel={rightPanel}
      rightHeaderExtra={<Badge tone="green">Tự động lưu Local</Badge>}
    />
  );
}
