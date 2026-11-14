import { useCallback, useEffect, useState } from 'react';
import { Activity, Cpu, HardDrive, RefreshCw, Sparkles } from 'lucide-react';
import { Badge, Button, Modal } from '../ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useToast } from '../../context/ToastContext';
import { getDoctorReport, getSystemVitals, runSystemMaintenance } from '../../services/api';
import type { DoctorReport, MaintenanceReport, SystemVitals } from '../../types/api';

interface DoctorPayload extends DoctorReport {
  status?: string;
  platform?: string;
  python_version?: string;
  components?: Record<
    string,
    {
      status?: string;
      issues?: string[];
      version?: string;
      [key: string]: unknown;
    }
  >;
  issues_count?: number;
}

export function SystemDiagnosticsModal() {
  const { diagnosticsOpen, setDiagnosticsOpen } = useWorkspace();
  const toast = useToast();

  const [doctor, setDoctor] = useState<DoctorPayload | null>(null);
  const [vitals, setVitals] = useState<SystemVitals | null>(null);
  const [maintReport, setMaintReport] = useState<MaintenanceReport | null>(null);
  const [loading, setLoading] = useState(false);
  const [cleaning, setCleaning] = useState(false);

  const loadAll = useCallback(async () => {
    setLoading(true);
    try {
      const [docRes, vitalsRes] = await Promise.all([
        getDoctorReport() as Promise<DoctorPayload>,
        getSystemVitals().catch(() => null),
      ]);
      setDoctor(docRes);
      setVitals(vitalsRes);
    } catch (err) {
      toast.error('Không tải được báo cáo chẩn đoán', err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, [toast]);

  useEffect(() => {
    if (diagnosticsOpen) {
      loadAll();
    }
  }, [diagnosticsOpen, loadAll]);

  const handleMaintenance = async () => {
    setCleaning(true);
    try {
      const res = await runSystemMaintenance({
        dry_run: false,
        clean_locks: true,
        clean_temp: true,
        vacuum_db: true,
        clear_cache: true,
        rotate_logs: true,
      });
      setMaintReport(res);
      const freedKb = ((res.total_bytes_freed || 0) / 1024).toFixed(1);
      toast.success('Đã dọn dẹp & tối ưu hệ thống thành công', `Giải phóng ${freedKb} KB rác tạm`);
      await loadAll();
    } catch (err) {
      toast.error('Lỗi bảo trì hệ thống', err instanceof Error ? err.message : String(err));
    } finally {
      setCleaning(false);
    }
  };

  const componentEntries = Object.entries(doctor?.components || {});

  return (
    <Modal
      open={diagnosticsOpen}
      onClose={() => setDiagnosticsOpen(false)}
      title="Chẩn Đoán Môi Trường & Bảo Trì Cục Bộ"
      subtitle="Kiểm tra 12 thành phần hệ thống (PyMuPDF, LibreOffice, Java PlantUML, Tesseract OCR, Trình duyệt Headless) và tối ưu bộ nhớ."
      size="xl"
    >
      <div className="space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 dark:border-slate-800 dark:bg-slate-950">
          <div className="flex flex-wrap items-center gap-3 text-xs">
            <Badge tone={doctor?.status === 'healthy' ? 'green' : 'amber'}>
              Trạng thái: {(doctor?.status || 'Đang kiểm tra').toUpperCase()}
            </Badge>
            {doctor?.python_version ? (
              <span className="text-slate-600 dark:text-slate-400">Python {doctor.python_version}</span>
            ) : null}
            {doctor?.platform ? (
              <span className="text-slate-500 dark:text-slate-400">{doctor.platform}</span>
            ) : null}
          </div>
          <div className="flex items-center gap-2">
            <Button
              size="sm"
              variant="outline"
              loading={loading}
              icon={<RefreshCw className="h-3.5 w-3.5" />}
              onClick={loadAll}
            >
              Quét lại
            </Button>
            <Button
              size="sm"
              variant="primary"
              loading={cleaning}
              icon={<Sparkles className="h-3.5 w-3.5" />}
              onClick={handleMaintenance}
            >
              Dọn rác & Tối ưu
            </Button>
          </div>
        </div>

        {vitals ? (
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
            <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-white p-3 dark:border-slate-800 dark:bg-slate-950">
              <Cpu className="h-5 w-5 text-indigo-500" />
              <div className="text-xs">
                <div className="text-slate-500">CPU Hệ thống</div>
                <div className="font-semibold text-slate-900 dark:text-slate-100">
                  {vitals.cpu_count ?? '-'} nhân xử lý
                </div>
              </div>
            </div>
            <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-white p-3 dark:border-slate-800 dark:bg-slate-950">
              <Activity className="h-5 w-5 text-emerald-500" />
              <div className="text-xs">
                <div className="text-slate-500">RAM đang dùng</div>
                <div className="font-semibold text-slate-900 dark:text-slate-100">
                  {vitals.memory_percent !== undefined ? `${vitals.memory_percent}%` : '-'}{' '}
                  {vitals.memory_available_gb !== undefined ? `(Còn ${vitals.memory_available_gb} GB)` : ''}
                </div>
              </div>
            </div>
            <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-white p-3 dark:border-slate-800 dark:bg-slate-950">
              <HardDrive className="h-5 w-5 text-blue-500" />
              <div className="text-xs">
                <div className="text-slate-500">Ổ đĩa trống</div>
                <div className="font-semibold text-slate-900 dark:text-slate-100">
                  {vitals.disk_free_gb !== undefined ? `${vitals.disk_free_gb} GB` : '-'}
                </div>
              </div>
            </div>
          </div>
        ) : null}

        {maintReport ? (
          <div className="rounded-xl border border-emerald-200 bg-emerald-50/70 px-4 py-3 text-xs text-emerald-900 dark:border-emerald-900/50 dark:bg-emerald-950/30 dark:text-emerald-200">
            <strong>Hoàn tất bảo trì:</strong> Đã dọn {maintReport.lock_files_removed ?? 0} lock files,{' '}
            {maintReport.temp_files_removed ?? 0} file tạm, {maintReport.pycache_files_removed ?? 0} tệp cache
            ({((maintReport.total_bytes_freed || 0) / 1024).toFixed(1)} KB).
          </div>
        ) : null}

        <div className="grid max-h-72 grid-cols-1 gap-2.5 overflow-y-auto pr-1 sm:grid-cols-2">
          {componentEntries.map(([compKey, info]) => {
            const healthy = info.status === 'healthy' || info.status === 'ok';
            return (
              <div
                key={compKey}
                className="rounded-xl border border-slate-200 bg-white p-3 text-xs dark:border-slate-800 dark:bg-slate-950"
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="font-semibold text-slate-800 dark:text-slate-200">{compKey}</span>
                  <Badge tone={healthy ? 'green' : 'amber'}>{info.status || 'unknown'}</Badge>
                </div>
                {info.issues && info.issues.length > 0 ? (
                  <p className="mt-1.5 text-[11px] text-amber-700 dark:text-amber-300">{info.issues[0]}</p>
                ) : (
                  <p className="mt-1.5 text-[11px] text-slate-500">Hoạt động ổn định</p>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </Modal>
  );
}
