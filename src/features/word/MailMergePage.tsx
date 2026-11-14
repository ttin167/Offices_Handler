import { useEffect, useState } from 'react';
import { ArrowRight, Link2, Mail, Sparkles } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
  Badge,
  Button,
  CollapsibleSection,
  FileDropZone,
  ResultPreviewCard,
  sessionFileToSelected,
  type SelectedFileInfo,
} from '../../components/ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { apiPost, buildDownloadUrl, buildPreviewUrl } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

function normalizeKey(str: string): string {
  return str
    .toLowerCase()
    .split(':')[0]
    .replace(/[{}\s_\-–—]/g, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

const FIELD_SYNONYMS: Record<string, string[]> = {
  name: ['ten', 'hoten', 'hovaten', 'fullname', 'nguoinhan'],
  fullname: ['ten', 'hoten', 'hovaten', 'name'],
  email: ['mail', 'thu', 'hopthu'],
  phone: ['sdt', 'so', 'dienthoai', 'mobile', 'tel'],
  address: ['diachi', 'noio', 'location'],
  id: ['ma', 'manv', 'code', 'stt', 'soid'],
  salary: ['luong', 'thunhap', 'sotien', 'tien', 'net'],
  sotien: ['sotien', 'tien', 'salary', 'tongtien', 'thanhtien'],
  date: ['ngay', 'ngaythang', 'ngayky', 'dob'],
  department: ['phongban', 'phong', 'bophan', 'dept'],
  title: ['chucvu', 'vitri', 'position'],
};

export function MailMergePage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [tplFile, setTplFile] = useState<SelectedFileInfo | null>(null);
  const [dataFile, setDataFile] = useState<SelectedFileInfo | null>(null);
  const [tokens, setTokens] = useState<string[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [sampleRecord, setSampleRecord] = useState<Record<string, unknown>>({});
  const [mapping, setMapping] = useState<Record<string, string>>({});

  const [sheetName, setSheetName] = useState('');
  const [filenamePattern, setFilenamePattern] = useState('');
  const [zipName, setZipName] = useState('KetQua_MailMerge.zip');

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [mergeStats, setMergeStats] = useState<{ generated?: number } | null>(null);

  const inspectTemplate = async (f: SelectedFileInfo) => {
    try {
      const res = await apiPost<{
        tokens?: { variables?: string[] } | string[];
      }>('/api/mail-merge/inspect', { template: f.filePath });
      const vars = Array.isArray(res.tokens)
        ? res.tokens
        : res.tokens?.variables || [];
      setTokens(vars);
    } catch (err) {
      toast.error('Không quét được biến mẫu Word', err instanceof Error ? err.message : String(err));
    }
  };

  const inspectDataset = async (f: SelectedFileInfo, sheet?: string) => {
    try {
      const res = await apiPost<{
        columns?: string[];
        sample_records?: Array<Record<string, unknown>>;
        records?: Array<Record<string, unknown>>;
      }>('/api/mail-merge/inspect-data', {
        data: f.filePath,
        sheet_name: sheet || undefined,
      });
      const cols = res.columns || [];
      setColumns(cols);
      const recs = res.sample_records || res.records || [];
      setSampleRecord(recs[0] || {});
    } catch (err) {
      toast.error('Không đọc được cột dữ liệu', err instanceof Error ? err.message : String(err));
    }
  };

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (!hf) return;
    const ext = hf.ext.toLowerCase();
    const sel = sessionFileToSelected(hf);
    if (ext === 'docx') {
      setTplFile(sel);
      void inspectTemplate(sel);
    } else if (['xlsx', 'csv', 'json'].includes(ext)) {
      setDataFile(sel);
      void inspectDataset(sel);
    }
  }, [consumeHandoffFile]);

  const handleAutoMatch = () => {
    const nextMap: Record<string, string> = {};
    tokens.forEach((tok) => {
      const nTok = normalizeKey(tok);
      let found = columns.find((c) => normalizeKey(c) === nTok);
      if (!found) {
        found = columns.find((c) => {
          const nCol = normalizeKey(c);
          return nCol.includes(nTok) || nTok.includes(nCol);
        });
      }
      if (!found) {
        const syns = FIELD_SYNONYMS[nTok] || [];
        found = columns.find((c) => syns.some((s) => normalizeKey(c).includes(s)));
      }
      nextMap[tok] = found || '';
    });
    setMapping(nextMap);
    const matched = Object.values(nextMap).filter(Boolean).length;
    toast.info(`Đã tự động khớp ${matched}/${tokens.length} trường`);
  };

  useEffect(() => {
    if (tokens.length > 0 && columns.length > 0 && Object.keys(mapping).length === 0) {
      handleAutoMatch();
    }
  }, [tokens, columns]);

  const handleRunMerge = async () => {
    if (!tplFile || !dataFile) {
      toast.warning('Thiếu tệp đầu vào', 'Cần cả mẫu Word (.docx) và dữ liệu nguồn (.xlsx/.csv/.json).');
      return;
    }
    setRunning(true);
    try {
      const cleanMap: Record<string, string> = {};
      Object.entries(mapping).forEach(([k, v]) => {
        if (v) cleanMap[k] = v;
      });
      const res = await apiPost<{
        zip_path?: string;
        download_url?: string;
        generated_count?: number;
        total?: number;
        files?: string[];
      }>('/api/mail-merge', {
        template: tplFile.filePath,
        data: dataFile.filePath,
        pattern: filenamePattern.trim() || undefined,
        zip: zipName.trim() || 'KetQua_MailMerge.zip',
        field_mapping: Object.keys(cleanMap).length ? cleanMap : undefined,
      });

      const targetPath = res.zip_path || res.files?.[0] || zipName;
      const fname = targetPath.split(/[\\/]/).pop() || zipName;
      const sf = addSessionFile({
        filePath: targetPath,
        filename: fname,
        previewUrl: buildPreviewUrl(fname),
        downloadUrl: res.download_url || buildDownloadUrl(fname),
        sourceLabel: 'Mail Merge',
      });
      setOutputSession(sf);
      setMergeStats({ generated: res.generated_count || res.total || res.files?.length || 1 });
      toast.success('Trộn thư hoàn tất', fname);
    } catch (err) {
      toast.error('Lỗi trộn thư', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  return (
    <SplitWorkspace
      title="Trộn thư Hàng loạt & Ánh xạ Trường"
      subtitle="Nối mẫu Word (.docx) với bảng Excel/CSV/JSON. Hỗ trợ đọc số tiền thành chữ ({{SoTien:vn_words}}) và ngày tháng tiếng Việt."
      badgeText="Mail Merge"
      accent="word"
      previewTarget={
        tplFile
          ? {
              filePath: tplFile.filePath,
              filename: tplFile.filename,
              ext: 'docx',
              previewUrl: tplFile.previewUrl,
              downloadUrl: tplFile.downloadUrl,
            }
          : null
      }
      leftPanel={
        <div className="space-y-4">
          <FileDropZone
            label="1. Mẫu hợp đồng / Văn bản Word (.docx)"
            accept=".docx"
            selectedFile={tplFile}
            onFileSelected={(f) => {
              setTplFile(f);
              setMapping({});
              void inspectTemplate(f);
            }}
            onClear={() => {
              setTplFile(null);
              setTokens([]);
            }}
            sourceLabel="Mẫu Mail Merge"
            hint="Chứa biến như {{HoTen}}, {{SoTien:vn_words}}"
          />

          <FileDropZone
            label="2. Tệp danh sách dữ liệu (.xlsx, .csv, .json)"
            accept=".xlsx,.xls,.csv,.json"
            selectedFile={dataFile}
            onFileSelected={(f) => {
              setDataFile(f);
              setMapping({});
              void inspectDataset(f, sheetName);
            }}
            onClear={() => {
              setDataFile(null);
              setColumns([]);
            }}
            sourceLabel="Nguồn Mail Merge"
          />

          {tokens.length > 0 ? (
            <div className="space-y-2 rounded-lg border border-slate-200 bg-slate-50 p-3 dark:border-slate-800 dark:bg-slate-950/60">
              <div className="flex items-center justify-between gap-2">
                <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-800 dark:text-slate-100">
                  <Link2 className="h-3.5 w-3.5 text-blue-600" />
                  Ánh xạ trường ({Object.values(mapping).filter(Boolean).length}/{tokens.length})
                </span>
                <Button
                  size="xs"
                  variant="secondary"
                  icon={<Sparkles className="h-3 w-3" />}
                  onClick={handleAutoMatch}
                >
                  Tự động khớp
                </Button>
              </div>

              <div className="max-h-56 space-y-1.5 overflow-y-auto pr-1">
                {tokens.map((tok) => {
                  const chosen = mapping[tok] || '';
                  const sampleVal = chosen && sampleRecord ? String(sampleRecord[chosen] ?? '') : '';
                  return (
                    <div
                      key={tok}
                      className="grid grid-cols-12 items-center gap-1.5 rounded border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-800 dark:bg-slate-900"
                    >
                      <span className="col-span-5 truncate font-mono text-[11px] font-medium text-blue-700 dark:text-blue-300">
                        {`{{${tok}}}`}
                      </span>
                      <ArrowRight className="col-span-1 h-3 w-3 text-slate-400" />
                      <div className="col-span-6 space-y-0.5">
                        <select
                          value={chosen}
                          onChange={(e) => setMapping((prev) => ({ ...prev, [tok]: e.target.value }))}
                          className="w-full rounded border border-slate-200 bg-slate-50 px-1.5 py-1 text-[11px] dark:border-slate-700 dark:bg-slate-800"
                        >
                          <option value="">-- Giữ tên gốc --</option>
                          {columns.map((c) => (
                            <option key={c} value={c}>
                              {c}
                            </option>
                          ))}
                        </select>
                        {sampleVal ? (
                          <div className="truncate text-[10px] text-emerald-600 dark:text-emerald-400">
                            Mẫu: {sampleVal}
                          </div>
                        ) : null}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : null}

          <CollapsibleSection title="Tùy chọn nâng cao (Tên file, Sheet Excel, ZIP)">
            <div className="space-y-2 text-xs">
              <div>
                <label className="mb-1 block text-[11px] text-slate-500">Tên Sheet Excel (bỏ trống lấy sheet đầu)</label>
                <input
                  type="text"
                  value={sheetName}
                  onChange={(e) => setSheetName(e.target.value)}
                  onBlur={() => {
                    if (dataFile) void inspectDataset(dataFile, sheetName);
                  }}
                  placeholder="Sheet1"
                  className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
                />
              </div>
              <div>
                <label className="mb-1 block text-[11px] text-slate-500">Mẫu đặt tên từng file đầu ra</label>
                <input
                  type="text"
                  value={filenamePattern}
                  onChange={(e) => setFilenamePattern(e.target.value)}
                  placeholder="Vd: HopDong_{{HoTen}}"
                  className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
                />
              </div>
              <div>
                <label className="mb-1 block text-[11px] text-slate-500">Tên tệp gói ZIP đầu ra</label>
                <input
                  type="text"
                  value={zipName}
                  onChange={(e) => setZipName(e.target.value)}
                  className="w-full rounded border border-slate-200 px-2.5 py-1.5 dark:border-slate-700 dark:bg-slate-800"
                />
              </div>
            </div>
          </CollapsibleSection>

          <Button
            accent="word"
            loading={running}
            disabled={!tplFile || !dataFile}
            onClick={handleRunMerge}
            icon={<Mail className="h-4 w-4" />}
            className="w-full"
          >
            Chạy Trộn thư Hàng loạt
          </Button>
        </div>
      }
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Hoàn tất trộn thư hàng loạt"
            outputFile={outputSession}
            summaryItems={[
              { label: 'Số bản ghi đã trộn', value: mergeStats?.generated ?? 'Hoàn tất' },
              { label: 'Bộ lọc tiếng Việt', value: 'vn_words, date_vn, vnd' },
            ]}
            extraActions={<Badge tone="blue">ZIP Archive Ready</Badge>}
          />
        ) : undefined
      }
    />
  );
}
