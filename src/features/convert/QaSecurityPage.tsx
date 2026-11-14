import { useEffect, useState } from 'react';
import { CheckCircle2, ShieldAlert, ShieldCheck, Sparkles } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  EmptyState,
  FileDropZone,
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

interface QaIssue {
  code?: string;
  severity?: string;
  category?: string;
  message?: string;
  description?: string;
  location?: string;
}

interface QaAuditResponse {
  success: boolean;
  target_file: string;
  format: string;
  elapsed: number;
  total_issues: number;
  critical_count: number;
  warning_count: number;
  passed: boolean;
  issues: QaIssue[];
  repaired: boolean;
  repair_actions: string[];
  repaired_file?: string;
  download_url?: string;
}

interface SecurityFinding {
  id: string;
  title: string;
  severity: string;
  invariant_code: string;
  attack_class: string;
  location: string;
  evidence: string;
  fix_proposal: string;
}

interface SecurityReportResponse {
  success: boolean;
  report?: {
    target: string;
    is_clean: boolean;
    critical_count: number;
    warning_count: number;
    info_count: number;
    findings: {
      confirmed: SecurityFinding[];
      needs_validation: SecurityFinding[];
      rejected: SecurityFinding[];
    };
  };
}

interface SanitizeResponse {
  success: boolean;
  sanitized_file: string;
  download_url: string;
  actions_taken?: string[];
}

export function QaSecurityPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [mode, setMode] = useState<'qa' | 'security'>('qa');
  const [targetFile, setTargetFile] = useState<SelectedFileInfo | null>(null);

  // QA flags
  const [qaRepair, setQaRepair] = useState(true);
  const [qaIncludeSec, setQaIncludeSec] = useState(true);
  const [qaRedactPii, setQaRedactPii] = useState(false);

  // Sanitize flags
  const [secDeepCdr, setSecDeepCdr] = useState(true);
  const [secDefuse, setSecDefuse] = useState(true);
  const [secStripMeta, setSecStripMeta] = useState(true);
  const [secUnhide, setSecUnhide] = useState(true);
  const [secRedactPii, setSecRedactPii] = useState(false);

  const [running, setRunning] = useState(false);
  const [qaResult, setQaResult] = useState<QaAuditResponse | null>(null);
  const [secReport, setSecReport] = useState<SecurityReportResponse['report'] | null>(null);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf) {
      setTargetFile(sessionFileToSelected(hf));
    }
  }, [consumeHandoffFile]);

  const handleRunQa = async () => {
    if (!targetFile) {
      toast.warning('Chưa chọn tệp tài liệu');
      return;
    }
    setRunning(true);
    setSecReport(null);
    setOutputSession(null);
    try {
      const res = await apiPost<QaAuditResponse>('/api/qa/audit', {
        file_path: targetFile.filePath,
        repair: qaRepair,
        security: qaIncludeSec,
        deep_cdr: true,
        redact_pii: qaRedactPii,
      });
      setQaResult(res);
      if (res.repaired && res.repaired_file) {
        const sf = addSessionFile({
          filePath: res.repaired_file,
          filename: res.repaired_file,
          previewUrl: buildPreviewUrl(res.repaired_file),
          downloadUrl: res.download_url || buildDownloadUrl(res.repaired_file),
          sourceLabel: 'QA Tự sửa OpenXML',
        });
        setOutputSession(sf);
      }
      toast.success(
        res.passed ? 'Kiểm định QA Đạt chuẩn' : `Tìm thấy ${res.total_issues} cảnh báo OpenXML`,
        res.repaired ? 'Đã xuất bản đã sửa lỗi' : undefined,
      );
    } catch (err) {
      toast.error('Lỗi kiểm định QA', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleScanSecurity = async () => {
    if (!targetFile) {
      toast.warning('Chưa chọn tệp tài liệu');
      return;
    }
    setRunning(true);
    setQaResult(null);
    setOutputSession(null);
    try {
      const res = await apiPost<SecurityReportResponse>('/api/security/scan', {
        file_path: targetFile.filePath,
      });
      setSecReport(res.report || null);
      if (res.report?.is_clean) {
        toast.success('Tệp an toàn — Không phát hiện mối nguy hại');
      } else {
        toast.warning(
          `Phát hiện ${res.report?.critical_count ?? 0} lỗ hổng nghiêm trọng & ${res.report?.warning_count ?? 0} cảnh báo`,
        );
      }
    } catch (err) {
      toast.error('Lỗi quét bảo mật', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleSanitizeSecurity = async () => {
    if (!targetFile) {
      toast.warning('Chưa chọn tệp tài liệu');
      return;
    }
    setRunning(true);
    setQaResult(null);
    try {
      const res = await apiPost<SanitizeResponse>('/api/security/sanitize', {
        file_path: targetFile.filePath,
        deep_cdr: secDeepCdr,
        defuse_formulas: secDefuse,
        strip_metadata: secStripMeta,
        unhide_sheets: secUnhide,
        redact_pii: secRedactPii,
      });
      const sf = addSessionFile({
        filePath: res.sanitized_file,
        filename: res.sanitized_file,
        previewUrl: buildPreviewUrl(res.sanitized_file),
        downloadUrl: res.download_url || buildDownloadUrl(res.sanitized_file),
        sourceLabel: 'Deep CDR Tẩy rửa sạch',
      });
      setOutputSession(sf);
      toast.success('Đã tẩy rửa sạch tài liệu (Deep CDR)', res.sanitized_file);
    } catch (err) {
      toast.error('Lỗi tẩy rửa bảo mật', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const leftPanel = (
    <div className="space-y-4">
      <FileDropZone
        label="Tài liệu cần kiểm định hoặc rà soát bảo mật (.docx, .xlsx, .pptx, .pdf, .csv)"
        accept=".docx,.xlsx,.xlsm,.pptx,.pdf,.csv,.drawio,.svg"
        selectedFile={targetFile}
        onFileSelected={(f) => {
          setTargetFile(f);
          setQaResult(null);
          setSecReport(null);
          setOutputSession(null);
        }}
        onClear={() => setTargetFile(null)}
        sourceLabel="QA & Bảo mật"
      />

      <Tabs
        activeId={mode}
        onChange={(id) => setMode(id as 'qa' | 'security')}
        className="w-full"
        options={[
          { id: 'qa', label: 'Kiểm định 13 Cổng QA', icon: <CheckCircle2 className="h-3.5 w-3.5" /> },
          { id: 'security', label: 'Quét Macro & Tẩy rửa Deep CDR', icon: <ShieldCheck className="h-3.5 w-3.5" /> },
        ]}
      />

      {mode === 'qa' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={qaRepair} onChange={(e) => setQaRepair(e.target.checked)} />
            <span>Tự động sửa lỗi OpenXML (tràn bảng Word, ngắt trang, ID số) & tạo bản sạch</span>
          </label>
          <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={qaIncludeSec} onChange={(e) => setQaIncludeSec(e.target.checked)} />
            <span>Kèm rà soát bảo mật Gate 13 (Macro / Formula Injection / OLE)</span>
          </label>
          <CollapsibleSection title="Tùy chọn nâng cao">
            <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
              <input type="checkbox" checked={qaRedactPii} onChange={(e) => setQaRedactPii(e.target.checked)} />
              <span>Ẩn thông tin nhạy cảm PII (CCCD, SĐT, Email) khi sửa lỗi</span>
            </label>
          </CollapsibleSection>
          <Button
            variant="primary"
            loading={running}
            disabled={!targetFile}
            icon={<CheckCircle2 className="h-4 w-4" />}
            onClick={handleRunQa}
            className="w-full bg-indigo-600 hover:bg-indigo-700"
          >
            Chạy Kiểm định 13 Cổng QA
          </Button>
        </div>
      ) : (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div className="space-y-2 text-xs text-slate-700 dark:text-slate-300">
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={secDeepCdr} onChange={(e) => setSecDeepCdr(e.target.checked)} />
              <span>Deep CDR: Gỡ bỏ hoàn toàn VBA Macro (vbaProject.bin) & OLE nhúng (.exe/.ps1)</span>
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={secDefuse} onChange={(e) => setSecDefuse(e.target.checked)} />
              <span>Vô hiệu hóa DDE & Công thức CSV/Excel nguy hiểm</span>
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={secStripMeta} onChange={(e) => setSecStripMeta(e.target.checked)} />
              <span>Xóa Metadata ẩn (tên tác giả, máy nội bộ, revision history)</span>
            </label>
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao">
            <div className="space-y-2 text-xs text-slate-700 dark:text-slate-300">
              <label className="flex items-center gap-2">
                <input type="checkbox" checked={secUnhide} onChange={(e) => setSecUnhide(e.target.checked)} />
                <span>Phát hiện và hiển thị các sheet Excel bị giấu kỹ (`veryHidden`)</span>
              </label>
              <label className="flex items-center gap-2">
                <input type="checkbox" checked={secRedactPii} onChange={(e) => setSecRedactPii(e.target.checked)} />
                <span>Che dấu thông tin cá nhân PII (CCCD, SĐT VN)</span>
              </label>
            </div>
          </CollapsibleSection>
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
            <Button
              variant="outline"
              loading={running}
              disabled={!targetFile}
              icon={<ShieldAlert className="h-4 w-4" />}
              onClick={handleScanSecurity}
            >
              Chỉ quét rủi ro
            </Button>
            <Button
              variant="primary"
              loading={running}
              disabled={!targetFile}
              icon={<Sparkles className="h-4 w-4" />}
              onClick={handleSanitizeSecurity}
              className="bg-indigo-600 hover:bg-indigo-700"
            >
              Tẩy rửa Sạch (Deep CDR)
            </Button>
          </div>
        </div>
      )}
    </div>
  );

  const rightPanel = (
    <div className="space-y-4">
      {outputSession ? (
        <ResultPreviewCard
          title="Đã xuất bản tài liệu Sạch & An toàn"
          outputFile={outputSession}
          summaryItems={
            qaResult?.repair_actions?.length
              ? [{ label: 'Hành động đã sửa', value: `${qaResult.repair_actions.length} mục` }]
              : [{ label: 'Trạng thái', value: 'Sẵn sàng lưu hành' }]
          }
        />
      ) : null}

      {qaResult ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-4 dark:border-slate-800 dark:bg-slate-900">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <h4 className="text-sm font-semibold text-slate-900 dark:text-slate-100">
              Báo cáo Kiểm định QA ({qaResult.target_file})
            </h4>
            <Badge tone={qaResult.passed ? 'green' : 'red'}>
              {qaResult.passed ? 'ĐẠT CHUẨN OPENXML' : `${qaResult.critical_count} Lỗi nghiêm trọng`}
            </Badge>
          </div>
          <div className="grid grid-cols-3 gap-2 text-center text-xs">
            <div className="rounded-lg bg-slate-50 p-2 dark:bg-slate-950">
              <div className="text-[11px] text-slate-500">Nghiêm trọng</div>
              <div className="font-bold text-red-600">{qaResult.critical_count}</div>
            </div>
            <div className="rounded-lg bg-slate-50 p-2 dark:bg-slate-950">
              <div className="text-[11px] text-slate-500">Cảnh báo</div>
              <div className="font-bold text-amber-600">{qaResult.warning_count}</div>
            </div>
            <div className="rounded-lg bg-slate-50 p-2 dark:bg-slate-950">
              <div className="text-[11px] text-slate-500">Thời gian</div>
              <div className="font-bold text-slate-800 dark:text-slate-200">{qaResult.elapsed}s</div>
            </div>
          </div>
          {qaResult.issues.length > 0 ? (
            <ul className="max-h-72 space-y-1.5 overflow-y-auto text-xs">
              {qaResult.issues.map((issue, i) => (
                <li
                  key={i}
                  className="rounded-lg border border-slate-100 bg-slate-50/70 px-3 py-2 dark:border-slate-800 dark:bg-slate-950/60"
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      {issue.code || issue.category || 'ISSUE'}
                    </span>
                    <Badge tone={issue.severity?.toUpperCase() === 'CRITICAL' ? 'red' : 'amber'}>
                      {issue.severity || 'INFO'}
                    </Badge>
                  </div>
                  <p className="mt-1 text-slate-600 dark:text-slate-400">{issue.message || issue.description}</p>
                </li>
              ))}
            </ul>
          ) : (
            <p className="text-xs text-emerald-700 dark:text-emerald-400">
              Không phát hiện sai lệch cấu trúc OpenXML nào.
            </p>
          )}
        </div>
      ) : null}

      {secReport ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-4 dark:border-slate-800 dark:bg-slate-900">
          <div className="flex items-center justify-between">
            <h4 className="text-sm font-semibold text-slate-900 dark:text-slate-100">Kết quả Quét An Toàn Bảo Mật</h4>
            <Badge tone={secReport.is_clean ? 'green' : 'red'}>
              {secReport.is_clean ? 'AN TOÀN' : `${secReport.critical_count} RỦI RO`}
            </Badge>
          </div>
          {secReport.findings.confirmed.length === 0 && secReport.findings.needs_validation.length === 0 ? (
            <p className="text-xs text-emerald-700 dark:text-emerald-400">
              Tệp không chứa Macro VBA nguy hiểm, đối tượng OLE hay liên kết độc hại.
            </p>
          ) : (
            <ul className="max-h-72 space-y-2 overflow-y-auto text-xs">
              {[...secReport.findings.confirmed, ...secReport.findings.needs_validation].map((f) => (
                <li
                  key={f.id}
                  className="rounded-lg border border-red-200 bg-red-50/40 p-2.5 dark:border-red-900/50 dark:bg-red-950/20"
                >
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-red-900 dark:text-red-200">
                      [{f.invariant_code}] {f.title}
                    </span>
                    <Badge tone="red">{f.severity}</Badge>
                  </div>
                  <div className="mt-1 text-[11px] text-slate-600 dark:text-slate-400">
                    Vị trí: <code>{f.location}</code> — Khuyến nghị: {f.fix_proposal}
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : null}

      {!qaResult && !secReport && !outputSession ? (
        <EmptyState
          icon={<ShieldCheck className="h-6 w-6" />}
          title="Sẵn sàng kiểm định chất lượng & bảo mật"
          description="Nạp tệp bất kỳ bên trái để chạy Kiểm định 13 cổng QA hoặc Quét/Tẩy rửa mã độc Deep CDR hoàn toàn Offline."
        />
      ) : null}
    </div>
  );

  return (
    <SplitWorkspace
      title="Kiểm Định QA 13 Cổng & Bảo Mật Deep CDR"
      subtitle="Rà soát lỗi trình bày Word/Excel/Slide, quét Macro VBA/OLE độc hại và xuất bản phiên bản đã tẩy rửa an toàn trước khi gửi đi."
      badgeText="Security & QA"
      accent="studio"
      leftPanel={leftPanel}
      resultPanel={rightPanel}
      rightHeaderExtra={<Badge tone="indigo">ERR_SEC_001–007</Badge>}
    />
  );
}
