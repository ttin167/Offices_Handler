import { useEffect, useState } from 'react';
import { Download, Eye, FileText, FunctionSquare, Table2, ZoomIn, ZoomOut, RotateCcw } from 'lucide-react';
import { apiPost, buildPreviewUrl } from '../../services/api';
import type { DocxSnapshot, XlsxCellData, XlsxSnapshot } from '../../types/api';

export interface PreviewTarget {
  filePath?: string;
  filename: string;
  ext: string;
  previewUrl?: string;
  downloadUrl?: string;
}

function ShimmerSkeleton({ isGrid }: { isGrid: boolean }) {
  return (
    <div className="space-y-3 p-5">
      <div className="flex items-center justify-between">
        <div className="shimmer-bar h-5 w-48 rounded" />
        <div className="shimmer-bar h-5 w-16 rounded" />
      </div>
      {isGrid ? (
        <div className="space-y-2 pt-2">
          <div className="shimmer-bar h-8 w-full rounded" />
          <div className="shimmer-bar h-7 w-full rounded" />
          <div className="shimmer-bar h-7 w-full rounded" />
          <div className="shimmer-bar h-7 w-5/6 rounded" />
        </div>
      ) : (
        <div className="space-y-2.5 pt-2">
          <div className="shimmer-bar h-4 w-11/12 rounded" />
          <div className="shimmer-bar h-4 w-3/4 rounded" />
          <div className="shimmer-bar h-4 w-full rounded" />
          <div className="shimmer-bar h-28 w-full rounded-lg" />
          <div className="shimmer-bar h-4 w-5/6 rounded" />
        </div>
      )}
      <p className="text-center text-xs text-slate-400">Đang nạp bản xem trước nội bộ...</p>
    </div>
  );
}

function colNumberToLetters(col: number): string {
  let c = col;
  let letters = '';
  while (c > 0) {
    const mod = (c - 1) % 26;
    letters = String.fromCharCode(65 + mod) + letters;
    c = Math.floor((c - mod) / 26);
  }
  return letters || 'A';
}

function DocxAstView({ snapshot }: { snapshot: DocxSnapshot }) {
  const body = snapshot.body || [];
  if (!body.length) {
    return (
      <div className="p-8 text-center text-xs text-slate-500">
        Tài liệu Word không có nội dung văn bản phần thân.
      </div>
    );
  }
  return (
    <div className="p-4">
      <div className="docx-paper-sheet space-y-2 text-[13px] leading-relaxed">
        {body.map((elem, idx) => {
          if (elem.type === 'table') {
            const rows = elem.rows || [];
            return (
              <div key={idx} className="my-3 overflow-x-auto">
                <table className="w-full border-collapse border border-slate-300 text-xs dark:border-slate-700">
                  <tbody>
                    {rows.map((row, rIdx) => (
                      <tr key={rIdx}>
                        {(row.cells || []).map((cell, cIdx) => {
                          const isHeader = rIdx === 0 && elem.has_header;
                          const Tag = isHeader ? 'th' : 'td';
                          return (
                            <Tag
                              key={cIdx}
                              style={{
                                backgroundColor: cell.shading ? `#${cell.shading}` : undefined,
                                verticalAlign: cell.vertical_alignment?.toLowerCase(),
                              }}
                              className={`border border-slate-300 px-2 py-1.5 text-left dark:border-slate-700 ${
                                isHeader ? 'bg-slate-100 font-semibold dark:bg-slate-800' : ''
                              }`}
                            >
                              {(cell.paragraphs || []).map((p, pIdx) => (
                                <div key={pIdx}>{p.text || ''}</div>
                              ))}
                            </Tag>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            );
          }

          const styleName = (elem.style_name || '').toLowerCase();
          const align = (elem.alignment || 'left').toLowerCase() as 'left' | 'center' | 'right' | 'justify';
          const runs = elem.runs || [];

          let className = 'text-slate-800 dark:text-slate-200';
          if (styleName.includes('heading 1') || styleName === 'title') {
            className = 'text-lg font-bold text-slate-950 dark:text-white pt-2';
          } else if (styleName.includes('heading 2')) {
            className = 'text-base font-bold text-slate-900 dark:text-slate-100 pt-1.5';
          } else if (styleName.includes('heading 3')) {
            className = 'text-sm font-semibold text-slate-900 dark:text-slate-100 pt-1';
          }

          return (
            <p key={idx} className={className} style={{ textAlign: align }}>
              {runs.map((r, rIdx) => (
                <span
                  key={rIdx}
                  style={{
                    fontWeight: r.bold ? 700 : undefined,
                    fontStyle: r.italic ? 'italic' : undefined,
                    textDecoration: r.underline ? 'underline' : undefined,
                    color: r.color ? `#${r.color}` : undefined,
                  }}
                >
                  {r.text || ''}
                </span>
              ))}
            </p>
          );
        })}
      </div>
    </div>
  );
}

function XlsxGridView({ snapshot }: { snapshot: XlsxSnapshot }) {
  const sheets = snapshot.sheets || [];
  const [sheetIdx, setSheetIdx] = useState(0);
  const [selectedCell, setSelectedCell] = useState<{ addr: string; formula: string; val: string }>({
    addr: 'A1',
    formula: '',
    val: '',
  });

  if (!sheets.length) {
    return (
      <div className="p-8 text-center text-xs text-slate-500">
        Bảng tính Excel chưa có dữ liệu sheet để hiển thị.
      </div>
    );
  }

  const active = sheets[sheetIdx] || sheets[0];
  const rows = active.rows || [];
  const maxCol = Math.min(Math.max(active.stats?.max_column || 8, 6), 26);

  return (
    <div className="flex h-full flex-col">
      <div className="flex items-center gap-2 border-b border-slate-200 bg-slate-50 px-3 py-1.5 dark:border-slate-800 dark:bg-slate-900">
        <span className="rounded border border-slate-200 bg-white px-2 py-0.5 font-mono text-[11px] font-semibold text-slate-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200">
          {selectedCell.addr}
        </span>
        <span className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700 dark:text-emerald-400">
          <FunctionSquare className="h-3.5 w-3.5" /> fx
        </span>
        <input
          type="text"
          readOnly
          value={
            selectedCell.formula
              ? `=${selectedCell.formula}  [Hiển thị: ${selectedCell.val}]`
              : selectedCell.val
          }
          placeholder="Nhấp vào một ô bên dưới để xem giá trị hoặc công thức"
          className="flex-1 bg-transparent font-mono text-xs text-slate-700 focus:outline-none dark:text-slate-200"
        />
      </div>

      <div className="flex-1 overflow-auto bg-white dark:bg-slate-950">
        <table className="min-w-full border-collapse text-xs">
          <thead className="sticky top-0 z-10 bg-slate-100 dark:bg-slate-900">
            <tr>
              <th className="w-10 border border-slate-200 px-2 py-1 text-center font-mono text-[10px] text-slate-500 dark:border-slate-800">
                #
              </th>
              {Array.from({ length: maxCol }, (_, i) => colNumberToLetters(i + 1)).map((colName) => (
                <th
                  key={colName}
                  className="min-w-[90px] border border-slate-200 px-2 py-1 text-center font-mono text-[11px] font-semibold text-slate-600 dark:border-slate-800 dark:text-slate-300"
                >
                  {colName}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.slice(0, 150).map((rowObj, rIdx) => {
              const rowNum = rowObj.row_number || rIdx + 1;
              const cellMap = new Map<number, XlsxCellData>();
              (rowObj.cells || []).forEach((c) => cellMap.set(c.column_number, c));

              return (
                <tr key={rowNum} className="hover:bg-slate-50/80 dark:hover:bg-slate-900/50">
                  <td className="border border-slate-200 bg-slate-50 px-2 py-1 text-center font-mono text-[10px] text-slate-500 dark:border-slate-800 dark:bg-slate-900">
                    {rowNum}
                  </td>
                  {Array.from({ length: maxCol }, (_, idx) => {
                    const colIdx = idx + 1;
                    const cell = cellMap.get(colIdx);
                    const val = cell ? String(cell.formatted_value ?? cell.value ?? '') : '';
                    const formula = cell?.formula || '';
                    const addr = `${colNumberToLetters(colIdx)}${rowNum}`;
                    const isSelected = selectedCell.addr === addr;

                    return (
                      <td
                        key={colIdx}
                        onClick={() => setSelectedCell({ addr, formula, val })}
                        style={{
                          fontWeight: cell?.bold ? 700 : undefined,
                          fontStyle: cell?.italic ? 'italic' : undefined,
                          color: cell?.font_color || undefined,
                          backgroundColor: cell?.fill_color || undefined,
                          textAlign: cell?.alignment?.toLowerCase() as 'left' | 'center' | 'right' | undefined,
                        }}
                        className={`cursor-cell border border-slate-200 px-2 py-1 text-slate-800 dark:border-slate-800 dark:text-slate-200 ${
                          isSelected ? 'ring-2 ring-inset ring-emerald-500' : ''
                        } ${formula ? 'underline decoration-emerald-400/60 decoration-dotted' : ''}`}
                      >
                        {val}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="flex flex-wrap items-center gap-1 border-t border-slate-200 bg-slate-50 px-2 py-1.5 dark:border-slate-800 dark:bg-slate-900">
        {sheets.map((s, idx) => (
          <button
            key={idx}
            type="button"
            onClick={() => setSheetIdx(idx)}
            className={`inline-flex items-center gap-1 rounded px-2.5 py-1 text-xs font-medium ${
              idx === sheetIdx
                ? 'bg-emerald-600 text-white'
                : 'bg-white text-slate-600 hover:bg-slate-200/70 dark:bg-slate-800 dark:text-slate-300'
            }`}
          >
            <Table2 className="h-3 w-3" />
            <span>{s.sheet_name || s.name || `Sheet ${idx + 1}`}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

export function InlinePreviewPane({
  target,
  emptyTitle = 'Khung xem trước trực tiếp',
  emptyDescription = 'Chọn tệp ở cột bên trái hoặc chạy công cụ để xem nội dung PDF, Word, Excel, Ảnh ngay tại đây.',
}: {
  target: PreviewTarget | null;
  emptyTitle?: string;
  emptyDescription?: string;
}) {
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [docxData, setDocxData] = useState<DocxSnapshot | null>(null);
  const [xlsxData, setXlsxData] = useState<XlsxSnapshot | null>(null);
  const [textContent, setTextContent] = useState<string | null>(null);
  const [zoom, setZoom] = useState(1);

  const ext = target?.ext.toLowerCase().replace('.', '') || '';

  useEffect(() => {
    let cancelled = false;
    setErrorMsg(null);
    setDocxData(null);
    setXlsxData(null);
    setTextContent(null);
    setZoom(1);

    if (!target) return;

    const loadPreview = async () => {
      if (ext === 'docx') {
        setLoading(true);
        try {
          const payload = target.filePath
            ? { filepath: target.filePath }
            : { filename: target.filename };
          const res = await apiPost<{ snapshot?: DocxSnapshot } & DocxSnapshot>('/api/docx/read', payload);
          if (!cancelled) setDocxData(res.snapshot || res);
        } catch (err) {
          if (!cancelled) setErrorMsg(err instanceof Error ? err.message : String(err));
        } finally {
          if (!cancelled) setLoading(false);
        }
      } else if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm') {
        setLoading(true);
        try {
          const payload = target.filePath
            ? { filepath: target.filePath }
            : { filename: target.filename };
          const res = await apiPost<{ snapshot?: XlsxSnapshot } & XlsxSnapshot>('/api/xlsx/read', payload);
          if (!cancelled) setXlsxData(res.snapshot || res);
        } catch (err) {
          if (!cancelled) setErrorMsg(err instanceof Error ? err.message : String(err));
        } finally {
          if (!cancelled) setLoading(false);
        }
      } else if (['md', 'txt', 'json', 'csv', 'mmd', 'puml', 'xml', 'drawio'].includes(ext)) {
        const url = target.previewUrl || buildPreviewUrl(target.filename);
        setLoading(true);
        try {
          const r = await fetch(url);
          if (!r.ok) throw new Error('Không thể đọc nội dung tệp văn bản.');
          const txt = await r.text();
          if (!cancelled) setTextContent(txt.slice(0, 60000));
        } catch (err) {
          if (!cancelled) setErrorMsg(err instanceof Error ? err.message : String(err));
        } finally {
          if (!cancelled) setLoading(false);
        }
      }
    };

    void loadPreview();
    return () => {
      cancelled = true;
    };
  }, [target?.filePath, target?.filename, ext, target?.previewUrl]);

  if (!target) {
    return (
      <div className="flex h-[520px] flex-col items-center justify-center rounded-xl border border-slate-200 bg-white p-8 text-center dark:border-slate-800 dark:bg-slate-900">
        <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-slate-100 text-slate-400 dark:bg-slate-800">
          <Eye className="h-5 w-5" />
        </div>
        <h4 className="text-sm font-semibold text-slate-800 dark:text-slate-100">{emptyTitle}</h4>
        <p className="mt-1 max-w-sm text-xs text-slate-500 dark:text-slate-400">{emptyDescription}</p>
      </div>
    );
  }

  const previewUrl = target.previewUrl || buildPreviewUrl(target.filename);
  const isImage = ['png', 'jpg', 'jpeg', 'svg', 'webp', 'gif'].includes(ext);

  return (
    <div className="flex h-[560px] flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-card dark:border-slate-800 dark:bg-slate-900">
      <div className="flex items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-3.5 py-2 dark:border-slate-800 dark:bg-slate-950/60">
        <div className="flex min-w-0 items-center gap-2">
          <FileText className="h-4 w-4 shrink-0 text-slate-400" />
          <span className="truncate text-xs font-semibold text-slate-800 dark:text-slate-100">
            {target.filename}
          </span>
          <span className="rounded bg-slate-200/80 px-1.5 py-0.5 font-mono text-[10px] font-semibold uppercase text-slate-700 dark:bg-slate-800 dark:text-slate-300">
            {ext}
          </span>
        </div>

        {(isImage || ext === 'docx') && (
          <div className="flex items-center gap-1 shrink-0">
            <button
              type="button"
              onClick={() => setZoom((z) => Math.max(0.6, +(z - 0.1).toFixed(1)))}
              className="rounded p-1 text-slate-500 hover:bg-slate-200/60 dark:hover:bg-slate-800"
              title="Thu nhỏ"
            >
              <ZoomOut className="h-3.5 w-3.5" />
            </button>
            <span className="w-10 text-center font-mono text-[11px] text-slate-600 dark:text-slate-300">
              {Math.round(zoom * 100)}%
            </span>
            <button
              type="button"
              onClick={() => setZoom((z) => Math.min(1.8, +(z + 0.1).toFixed(1)))}
              className="rounded p-1 text-slate-500 hover:bg-slate-200/60 dark:hover:bg-slate-800"
              title="Phóng to"
            >
              <ZoomIn className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onClick={() => setZoom(1)}
              className="rounded p-1 text-slate-500 hover:bg-slate-200/60 dark:hover:bg-slate-800"
              title="Đặt lại 100%"
            >
              <RotateCcw className="h-3.5 w-3.5" />
            </button>
          </div>
        )}
      </div>

      <div className="flex-1 overflow-auto bg-slate-100/70 dark:bg-slate-950">
        {loading ? (
          <ShimmerSkeleton isGrid={['xlsx', 'xls', 'xlsm'].includes(ext)} />
        ) : errorMsg ? (
          <div className="flex h-full flex-col items-center justify-center p-6 text-center">
            <p className="text-xs font-medium text-red-600 dark:text-red-400">{errorMsg}</p>
            {target.downloadUrl ? (
              <a
                href={target.downloadUrl}
                download={target.filename}
                className="mt-3 inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
              >
                <Download className="h-3.5 w-3.5" />
                Tải tệp xuống
              </a>
            ) : null}
          </div>
        ) : ext === 'pdf' ? (
          <iframe
            src={`${previewUrl}#toolbar=1&navpanes=0`}
            title={`Xem trước ${target.filename}`}
            className="h-full w-full border-0 bg-white"
          />
        ) : ext === 'docx' && docxData ? (
          <div style={{ transform: `scale(${zoom})`, transformOrigin: 'top center' }}>
            <DocxAstView snapshot={docxData} />
          </div>
        ) : ['xlsx', 'xls', 'xlsm'].includes(ext) && xlsxData ? (
          <XlsxGridView snapshot={xlsxData} />
        ) : isImage ? (
          <div className="flex min-h-full items-center justify-center p-4">
            <img
              src={previewUrl}
              alt={target.filename}
              style={{ transform: `scale(${zoom})`, transformOrigin: 'center' }}
              className="max-h-full max-w-full rounded-lg border border-slate-200 bg-white shadow-subtle transition-transform dark:border-slate-800"
            />
          </div>
        ) : textContent !== null ? (
          <pre className="whitespace-pre-wrap break-words p-4 font-mono text-xs leading-5 text-slate-800 dark:text-slate-200">
            {textContent}
          </pre>
        ) : (
          <div className="flex h-full flex-col items-center justify-center p-6 text-center">
            <p className="text-xs text-slate-500">
              Định dạng <span className="font-mono font-semibold uppercase">.{ext}</span> sẵn sàng để tải xuống.
            </p>
            {target.downloadUrl ? (
              <a
                href={target.downloadUrl}
                download={target.filename}
                className="mt-3 inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-2 text-xs font-medium text-white hover:bg-slate-800 dark:bg-slate-100 dark:text-slate-900"
              >
                <Download className="h-3.5 w-3.5" />
                Tải xuống {target.filename}
              </a>
            ) : null}
          </div>
        )}
      </div>
    </div>
  );
}
