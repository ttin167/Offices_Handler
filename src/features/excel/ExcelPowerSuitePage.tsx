import { useEffect, useState } from 'react';
import {
  BarChart3,
  Columns,
  GitCompare,
  Link2,
  Split,
  Wand2,
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
import { apiPost, buildDownloadUrl, buildPreviewUrl } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

type PowerToolTab =
  | 'power_clean'
  | 'transform_cols'
  | 'pivot_chart'
  | 'xlookup_join'
  | 'split_col'
  | 'diff';

interface XlsxActionResponse {
  success?: boolean;
  filename?: string;
  download_url?: string;
  output_path?: string;
  stats?: Record<string, unknown>;
  detail?: string;
  summary?: Record<string, unknown>;
}

export function ExcelPowerSuitePage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [tab, setTab] = useState<PowerToolTab>('power_clean');
  const [primaryFile, setPrimaryFile] = useState<SelectedFileInfo | null>(null);
  const [secondaryFile, setSecondaryFile] = useState<SelectedFileInfo | null>(null);

  // Power Clean
  const [trimWhitespace, setTrimWhitespace] = useState(true);
  const [normalizeDates, setNormalizeDates] = useState(true);
  const [normalizePhonesIds, setNormalizePhonesIds] = useState(true);
  const [caseMode, setCaseMode] = useState('preserve');
  const [dedupCols, setDedupCols] = useState('');
  const [cleanSheet, setCleanSheet] = useState('');

  // Transform columns
  const [colMode, setColMode] = useState<'split_vn_fullname' | 'split_delimiter' | 'merge_columns'>('split_vn_fullname');
  const [sourceCols, setSourceCols] = useState('HoVaTen');
  const [targetCols, setTargetCols] = useState('');
  const [delimiter, setDelimiter] = useState(' ');
  const [joiner, setJoiner] = useState(' ');
  const [newColName, setNewColName] = useState('');

  // Pivot & Chart
  const [groupByCol, setGroupByCol] = useState('PhongBan');
  const [valueCols, setValueCols] = useState('Luong');
  const [aggFunc, setAggFunc] = useState('SUM');
  const [chartType, setChartType] = useState('bar');
  const [pivotSourceSheet, setPivotSourceSheet] = useState('');

  // XLOOKUP Join
  const [leftKey, setLeftKey] = useState('MaNV');
  const [rightKey, setRightKey] = useState('MaNV');
  const [returnCols, setReturnCols] = useState('HoTen,PhongBan');
  const [leftSheet, setLeftSheet] = useState('');
  const [rightSheet, setRightSheet] = useState('');
  const [notFoundVal, setNotFoundVal] = useState('N/A');

  // Split by column
  const [splitCol, setSplitCol] = useState('PhongBan');
  const [splitSheet, setSplitSheet] = useState('');

  // Diff
  const [diffSheets, setDiffSheets] = useState('');

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);
  const [lastTitle, setLastTitle] = useState('');
  const [summaryItems, setSummaryItems] = useState<Array<{ label: string; value: string | number }>>([]);

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && ['xlsx', 'xlsm', 'xls', 'csv'].includes(hf.ext.toLowerCase())) {
      setPrimaryFile(sessionFileToSelected(hf));
    }
  }, [consumeHandoffFile]);

  const executePower = async (
    endpoint: string,
    payload: Record<string, unknown>,
    title: string,
  ) => {
    setRunning(true);
    try {
      const res = await apiPost<XlsxActionResponse>(endpoint, payload);
      const outName = res.filename || 'result.xlsx';
      const sf = addSessionFile({
        filePath: res.output_path || outName,
        filename: outName,
        previewUrl: buildPreviewUrl(outName),
        downloadUrl: res.download_url || buildDownloadUrl(outName),
        sourceLabel: title,
      });
      setOutputSession(sf);
      setLastTitle(title);
      const statsObj = res.stats || res.summary;
      if (statsObj && typeof statsObj === 'object') {
        setSummaryItems(
          Object.entries(statsObj)
            .slice(0, 6)
            .map(([k, v]) => ({ label: k, value: String(v) })),
        );
      } else {
        setSummaryItems([{ label: 'Tệp xuất ra', value: outName }]);
      }
      toast.success(`${title} thành công`, outName);
    } catch (err) {
      toast.error(`Lỗi ${title}`, err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const handleRun = () => {
    if (!primaryFile) {
      toast.warning('Chưa có tệp Excel nguồn', 'Vui lòng tải lên tệp .xlsx để xử lý.');
      return;
    }

    switch (tab) {
      case 'power_clean':
        executePower(
          '/api/xlsx/power-clean',
          {
            file_path: primaryFile.filePath,
            sheet_name: cleanSheet.trim() || undefined,
            dedup_columns: dedupCols.trim() || undefined,
            normalize_dates: normalizeDates,
            normalize_phones_and_ids: normalizePhonesIds,
            trim_whitespace: trimWhitespace,
            case_mode: caseMode,
          },
          'Chuẩn hóa dữ liệu Excel',
        );
        break;

      case 'transform_cols':
        if (!sourceCols.trim()) {
          toast.warning('Thiếu tên cột nguồn');
          return;
        }
        executePower(
          '/api/xlsx/transform-columns',
          {
            file_path: primaryFile.filePath,
            mode: colMode,
            source_columns: sourceCols.trim(),
            target_columns: targetCols.trim() || undefined,
            delimiter,
            joiner,
            new_column_name: newColName.trim() || undefined,
          },
          'Tách / Gộp cột Excel',
        );
        break;

      case 'pivot_chart':
        if (!groupByCol.trim()) {
          toast.warning('Thiếu cột phân nhóm Pivot');
          return;
        }
        executePower(
          '/api/xlsx/pivot-chart',
          {
            file_path: primaryFile.filePath,
            group_by_col: groupByCol.trim(),
            value_cols: valueCols.trim() || undefined,
            agg_func: aggFunc,
            chart_type: chartType,
            source_sheet: pivotSourceSheet.trim() || undefined,
          },
          'Tạo Pivot & Biểu đồ Excel',
        );
        break;

      case 'xlookup_join':
        if (!leftKey.trim() || !returnCols.trim()) {
          toast.warning('Thiếu cột khóa hoặc cột cần lấy');
          return;
        }
        executePower(
          '/api/xlsx/xlookup-join',
          {
            left_path: primaryFile.filePath,
            right_path: secondaryFile?.filePath || undefined,
            left_key: leftKey.trim(),
            right_key: rightKey.trim() || leftKey.trim(),
            return_cols: returnCols.trim(),
            left_sheet: leftSheet.trim() || undefined,
            right_sheet: rightSheet.trim() || undefined,
            not_found_value: notFoundVal,
          },
          'Nối bảng XLOOKUP Join',
        );
        break;

      case 'split_col':
        if (!splitCol.trim()) {
          toast.warning('Thiếu tên cột dùng để tách file');
          return;
        }
        executePower(
          '/api/xlsx/split-by-column',
          {
            file_path: primaryFile.filePath,
            split_col: splitCol.trim(),
            sheet_name: splitSheet.trim() || undefined,
          },
          'Tách Workbook theo cột (.ZIP)',
        );
        break;

      case 'diff':
        if (!secondaryFile) {
          toast.warning('Cần chọn thêm Tệp Excel phiên bản V2 để đối chiếu');
          return;
        }
        executePower(
          '/api/xlsx/diff',
          {
            v1_path: primaryFile.filePath,
            v2_path: secondaryFile.filePath,
            sheets: diffSheets.trim() || undefined,
          },
          'Đối chiếu Excel Visual Diff',
        );
        break;
    }
  };

  const previewTarget = outputSession
    ? {
        filePath: outputSession.filePath,
        filename: outputSession.filename,
        ext: outputSession.ext,
        previewUrl: outputSession.previewUrl,
        downloadUrl: outputSession.downloadUrl,
      }
    : primaryFile
      ? {
          filePath: primaryFile.filePath,
          filename: primaryFile.filename,
          ext: primaryFile.ext,
          previewUrl: primaryFile.previewUrl,
          downloadUrl: primaryFile.downloadUrl,
        }
      : null;

  const leftPanel = (
    <div className="space-y-4">
      <Tabs
        activeId={tab}
        onChange={(id) => setTab(id as PowerToolTab)}
        className="w-full"
        options={[
          { id: 'power_clean', label: 'Làm sạch', icon: <Wand2 className="h-3.5 w-3.5" /> },
          { id: 'transform_cols', label: 'Tách/Gộp cột', icon: <Columns className="h-3.5 w-3.5" /> },
          { id: 'pivot_chart', label: 'Pivot & Biểu đồ', icon: <BarChart3 className="h-3.5 w-3.5" /> },
          { id: 'xlookup_join', label: 'Nối XLOOKUP', icon: <Link2 className="h-3.5 w-3.5" /> },
          { id: 'split_col', label: 'Tách File (.ZIP)', icon: <Split className="h-3.5 w-3.5" /> },
          { id: 'diff', label: 'So sánh 2 bản', icon: <GitCompare className="h-3.5 w-3.5" /> },
        ]}
      />

      <FileDropZone
        label={tab === 'diff' ? 'Tệp Excel gốc (Phiên bản 1)' : tab === 'xlookup_join' ? 'Bảng tính chính (Left Table)' : 'Tệp Excel nguồn (.xlsx)'}
        accept=".xlsx,.xlsm,.xls"
        selectedFile={primaryFile}
        onFileSelected={(f) => {
          setPrimaryFile(f);
          setOutputSession(null);
        }}
        onClear={() => setPrimaryFile(null)}
        sourceLabel="Excel Power"
      />

      {tab === 'power_clean' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
            <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
              <input type="checkbox" checked={trimWhitespace} onChange={(e) => setTrimWhitespace(e.target.checked)} />
              <span>Cắt khoảng trắng thừa</span>
            </label>
            <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
              <input type="checkbox" checked={normalizeDates} onChange={(e) => setNormalizeDates(e.target.checked)} />
              <span>Chuẩn hóa ngày DD/MM/YYYY</span>
            </label>
            <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
              <input type="checkbox" checked={normalizePhonesIds} onChange={(e) => setNormalizePhonesIds(e.target.checked)} />
              <span>Giữ số 0 đầu SĐT / CCCD (@)</span>
            </label>
          </div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Kiểu viết hoa / thường</label>
              <select
                value={caseMode}
                onChange={(e) => setCaseMode(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="preserve">Giữ nguyên gốc</option>
                <option value="title">Viết Hoa Chữ Đầu Mỗi Từ</option>
                <option value="upper">CHỮ IN HOA TOÀN BỘ</option>
                <option value="lower">chữ thường toàn bộ</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Cột lọc trùng lặp (tùy chọn)</label>
              <input
                type="text"
                value={dedupCols}
                onChange={(e) => setDedupCols(e.target.value)}
                placeholder="VD: MaNV, CCCD"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao">
            <input
              type="text"
              value={cleanSheet}
              onChange={(e) => setCleanSheet(e.target.value)}
              placeholder="Tên sheet cụ thể (để trống nếu xử lý sheet đầu)"
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </CollapsibleSection>
        </div>
      ) : null}

      {tab === 'transform_cols' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Chế độ xử lý</label>
              <select
                value={colMode}
                onChange={(e) => setColMode(e.target.value as 'split_vn_fullname' | 'split_delimiter' | 'merge_columns')}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="split_vn_fullname">Tách Họ & Tên đệm + Tên (Chuẩn VN)</option>
                <option value="split_delimiter">Tách 1 cột theo dấu phân cách</option>
                <option value="merge_columns">Gộp nhiều cột thành 1 cột mới</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Tên cột nguồn</label>
              <input
                type="text"
                value={sourceCols}
                onChange={(e) => setSourceCols(e.target.value)}
                placeholder={colMode === 'merge_columns' ? 'HoLot, Ten' : 'HoVaTen'}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao (Ký tự phân cách & Tên cột mới)">
            <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
              <input
                type="text"
                value={delimiter}
                onChange={(e) => setDelimiter(e.target.value)}
                placeholder="Dấu tách (VD: -)"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
              <input
                type="text"
                value={joiner}
                onChange={(e) => setJoiner(e.target.value)}
                placeholder="Dấu nối (VD: khoảng trắng)"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
              <input
                type="text"
                value={newColName || targetCols}
                onChange={(e) => {
                  setNewColName(e.target.value);
                  setTargetCols(e.target.value);
                }}
                placeholder="Tên cột đích"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </CollapsibleSection>
        </div>
      ) : null}

      {tab === 'pivot_chart' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Cột phân nhóm (Group By)</label>
              <input
                type="text"
                value={groupByCol}
                onChange={(e) => setGroupByCol(e.target.value)}
                placeholder="VD: PhongBan"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Cột giá trị cần tổng hợp</label>
              <input
                type="text"
                value={valueCols}
                onChange={(e) => setValueCols(e.target.value)}
                placeholder="VD: Luong, Thuong"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Hàm tổng hợp</label>
              <select
                value={aggFunc}
                onChange={(e) => setAggFunc(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="SUM">Tổng cộng (SUM)</option>
                <option value="AVERAGE">Trung bình (AVERAGE)</option>
                <option value="COUNT">Đếm số lượng (COUNT)</option>
                <option value="MAX">Lớn nhất (MAX)</option>
                <option value="MIN">Nhỏ nhất (MIN)</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Loại biểu đồ Excel</label>
              <select
                value={chartType}
                onChange={(e) => setChartType(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="bar">Biểu đồ Cột (Bar Chart)</option>
                <option value="line">Biểu đồ Đường (Line Chart)</option>
                <option value="pie">Biểu đồ Tròn (Pie Chart)</option>
              </select>
            </div>
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao">
            <input
              type="text"
              value={pivotSourceSheet}
              onChange={(e) => setPivotSourceSheet(e.target.value)}
              placeholder="Tên sheet nguồn (tùy chọn)"
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </CollapsibleSection>
        </div>
      ) : null}

      {tab === 'xlookup_join' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <FileDropZone
            label="Bảng tra cứu phụ (Tùy chọn - Để trống nếu tra ở sheet khác trong cùng tệp)"
            accept=".xlsx,.xlsm,.xls"
            selectedFile={secondaryFile}
            onFileSelected={setSecondaryFile}
            onClear={() => setSecondaryFile(null)}
            sourceLabel="XLOOKUP Right"
          />
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Khóa bảng chính</label>
              <input
                type="text"
                value={leftKey}
                onChange={(e) => setLeftKey(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Khóa bảng tra cứu</label>
              <input
                type="text"
                value={rightKey}
                onChange={(e) => setRightKey(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">Các cột cần lấy về</label>
              <input
                type="text"
                value={returnCols}
                onChange={(e) => setReturnCols(e.target.value)}
                placeholder="HoTen, PhongBan"
                className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao (Tên Sheet & Giá trị mặc định)">
            <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
              <input
                type="text"
                value={leftSheet}
                onChange={(e) => setLeftSheet(e.target.value)}
                placeholder="Sheet bảng chính"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
              <input
                type="text"
                value={rightSheet}
                onChange={(e) => setRightSheet(e.target.value)}
                placeholder="Sheet tra cứu"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
              <input
                type="text"
                value={notFoundVal}
                onChange={(e) => setNotFoundVal(e.target.value)}
                placeholder="Khi không khớp (VD: N/A)"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              />
            </div>
          </CollapsibleSection>
        </div>
      ) : null}

      {tab === 'split_col' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-700 dark:text-slate-300">
              Tên cột dùng để tách thành nhiều tệp .xlsx (đóng gói ZIP)
            </label>
            <input
              type="text"
              value={splitCol}
              onChange={(e) => setSplitCol(e.target.value)}
              placeholder="VD: PhongBan hoặc ChiNhanh"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </div>
          <CollapsibleSection title="Tùy chọn nâng cao">
            <input
              type="text"
              value={splitSheet}
              onChange={(e) => setSplitSheet(e.target.value)}
              placeholder="Tên sheet cụ thể (tùy chọn)"
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </CollapsibleSection>
        </div>
      ) : null}

      {tab === 'diff' ? (
        <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/60 p-3.5 dark:border-slate-800 dark:bg-slate-950/50">
          <FileDropZone
            label="Tệp Excel mới cần đối chiếu (Phiên bản 2)"
            accept=".xlsx,.xlsm,.xls"
            selectedFile={secondaryFile}
            onFileSelected={setSecondaryFile}
            onClear={() => setSecondaryFile(null)}
            sourceLabel="Excel Diff V2"
          />
          <CollapsibleSection title="Tùy chọn nâng cao">
            <input
              type="text"
              value={diffSheets}
              onChange={(e) => setDiffSheets(e.target.value)}
              placeholder="Chỉ đối chiếu các sheet (cách nhau dấu phẩy)"
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
            />
          </CollapsibleSection>
        </div>
      ) : null}

      <Button
        variant="primary"
        loading={running}
        disabled={!primaryFile || (tab === 'diff' && !secondaryFile)}
        onClick={handleRun}
        className="w-full bg-emerald-600 hover:bg-emerald-700"
      >
        Thực thi tác vụ Excel
      </Button>
    </div>
  );

  return (
    <SplitWorkspace
      title="Excel Power Suite — 6 Công cụ Dữ liệu"
      subtitle="Làm sạch dữ liệu VN, tách Họ & Tên, dựng Pivot + Biểu đồ, nối bảng XLOOKUP không công thức, tách file ZIP & Visual Diff."
      badgeText="6-in-1 OpenXML"
      accent="excel"
      leftPanel={leftPanel}
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title={lastTitle || 'Xử lý Excel hoàn tất'}
            outputFile={outputSession}
            summaryItems={summaryItems}
          />
        ) : undefined
      }
      previewTarget={previewTarget}
      rightHeaderExtra={<Badge tone="green">Giữ nguyên định dạng @</Badge>}
    />
  );
}
