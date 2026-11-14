import { useMemo, useRef, useState } from 'react';
import { CheckCircle2, FileUp, FolderOpen, Loader2, X } from 'lucide-react';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { uploadSingleFile, buildDownloadUrl, buildPreviewUrl, getExtension } from '../../services/api';
import type { SessionFile } from '../../types/workspace';

export interface SelectedFileInfo {
  filePath: string;
  filename: string;
  ext: string;
  size?: number;
  previewUrl: string;
  downloadUrl: string;
}

export function sessionFileToSelected(sf: SessionFile): SelectedFileInfo {
  return {
    filePath: sf.filePath,
    filename: sf.filename,
    ext: sf.ext,
    size: sf.size,
    previewUrl: sf.previewUrl || buildPreviewUrl(sf.filename),
    downloadUrl: sf.downloadUrl || buildDownloadUrl(sf.filename),
  };
}

export function FileDropZone({
  label,
  accept,
  selectedFile,
  onFileSelected,
  onClear,
  sourceLabel = 'Tải lên',
  hint,
}: {
  label: string;
  accept?: string;
  selectedFile: SelectedFileInfo | null;
  onFileSelected: (file: SelectedFileInfo) => void;
  onClear?: () => void;
  sourceLabel?: string;
  hint?: string;
}) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [dragging, setDragging] = useState(false);
  const [uploading, setUploading] = useState(false);
  const { sessionFiles, addSessionFile } = useSessionFiles();
  const toast = useToast();

  const acceptedExts = useMemo(() => {
    if (!accept) return null;
    return accept
      .split(',')
      .map((s) => s.trim().toLowerCase().replace('.', ''))
      .filter(Boolean);
  }, [accept]);

  const matchingSessionFiles = useMemo(() => {
    if (!acceptedExts) return sessionFiles.slice(0, 4);
    return sessionFiles.filter((f) => acceptedExts.includes(f.ext.toLowerCase())).slice(0, 4);
  }, [sessionFiles, acceptedExts]);

  const handleUpload = async (file: File) => {
    setUploading(true);
    try {
      const uploaded = await uploadSingleFile(file);
      const ext = getExtension(uploaded.filename);
      const sf = addSessionFile({
        filePath: uploaded.filePath,
        filename: uploaded.filename,
        size: uploaded.size,
        previewUrl: uploaded.previewUrl,
        sourceLabel,
      });
      onFileSelected({
        filePath: uploaded.filePath,
        filename: uploaded.filename,
        ext,
        size: uploaded.size,
        previewUrl: sf.previewUrl || uploaded.previewUrl,
        downloadUrl: sf.downloadUrl || buildDownloadUrl(uploaded.filename),
      });
      toast.success('Đã nạp tệp thành công', uploaded.filename);
    } catch (err) {
      toast.error('Lỗi tải tệp lên', err instanceof Error ? err.message : String(err));
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="space-y-2">
      <div className="flex items-center justify-between">
        <label className="text-xs font-medium text-slate-700 dark:text-slate-300">{label}</label>
        {accept ? <span className="text-[11px] font-mono text-slate-400">{accept}</span> : null}
      </div>

      {selectedFile ? (
        <div className="flex items-center justify-between gap-3 rounded-lg border border-emerald-200 bg-emerald-50/70 px-3 py-2.5 dark:border-emerald-900 dark:bg-emerald-950/30">
          <div className="flex min-w-0 items-center gap-2.5">
            <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-600 dark:text-emerald-400" />
            <div className="min-w-0">
              <p className="truncate text-xs font-semibold text-slate-900 dark:text-slate-100">
                {selectedFile.filename}
              </p>
              <p className="truncate text-[11px] text-slate-500 dark:text-slate-400">
                {selectedFile.ext.toUpperCase()}
                {selectedFile.size ? ` • ${(selectedFile.size / 1024).toFixed(1)} KB` : ''}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <button
              type="button"
              onClick={() => inputRef.current?.click()}
              className="rounded px-2 py-1 text-[11px] font-medium text-slate-600 hover:bg-white dark:text-slate-300 dark:hover:bg-slate-800"
            >
              Đổi tệp
            </button>
            {onClear ? (
              <button
                type="button"
                onClick={onClear}
                className="rounded p-1 text-slate-400 hover:bg-white hover:text-red-600 dark:hover:bg-slate-800"
                aria-label="Bỏ chọn tệp"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            ) : null}
          </div>
        </div>
      ) : (
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            const dropped = e.dataTransfer.files?.[0];
            if (dropped) void handleUpload(dropped);
          }}
          onClick={() => inputRef.current?.click()}
          className={`flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed px-4 py-5 text-center transition-colors ${
            dragging
              ? 'border-indigo-500 bg-indigo-50/70 dark:bg-indigo-950/30'
              : 'border-slate-200 bg-slate-50/70 hover:border-slate-300 hover:bg-slate-100/70 dark:border-slate-700 dark:bg-slate-900/50 dark:hover:border-slate-600'
          }`}
        >
          {uploading ? (
            <Loader2 className="mb-1.5 h-6 w-6 animate-spin text-indigo-600" />
          ) : (
            <FileUp className="mb-1.5 h-6 w-6 text-slate-400" />
          )}
          <p className="text-xs font-medium text-slate-700 dark:text-slate-200">
            {uploading ? 'Đang nạp tệp vào không gian local...' : 'Kéo thả tệp vào đây hoặc nhấn để chọn'}
          </p>
          {hint ? <p className="mt-0.5 text-[11px] text-slate-400">{hint}</p> : null}
        </div>
      )}

      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => {
          const picked = e.target.files?.[0];
          if (picked) void handleUpload(picked);
          e.target.value = '';
        }}
      />

      {!selectedFile && matchingSessionFiles.length > 0 ? (
        <div className="flex flex-wrap items-center gap-1.5 pt-0.5">
          <span className="inline-flex items-center gap-1 text-[11px] text-slate-400">
            <FolderOpen className="h-3 w-3" /> Dùng nhanh từ phiên:
          </span>
          {matchingSessionFiles.map((sf) => (
            <button
              key={sf.id}
              type="button"
              onClick={() => onFileSelected(sessionFileToSelected(sf))}
              className="max-w-[180px] truncate rounded border border-slate-200 bg-white px-2 py-0.5 text-[11px] font-medium text-slate-700 hover:border-indigo-300 hover:text-indigo-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-300"
              title={`Nạp lại ${sf.filename}`}
            >
              {sf.filename}
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}
