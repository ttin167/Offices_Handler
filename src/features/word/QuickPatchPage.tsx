import { useEffect, useState } from 'react';
import { Code2, FileText, Wand2 } from 'lucide-react';
import { SplitWorkspace } from '../../components/layout/SplitWorkspace';
import {
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

export function QuickPatchPage() {
  const { consumeHandoffFile } = useWorkspace();
  const { addSessionFile } = useSessionFiles();
  const toast = useToast();

  const [tplFile, setTplFile] = useState<SelectedFileInfo | null>(null);
  const [variables, setVariables] = useState<string[]>([]);
  const [valuesMap, setValuesMap] = useState<Record<string, string>>({});
  const [editorMode, setEditorMode] = useState<'form' | 'json'>('form');
  const [jsonText, setJsonText] = useState('{}');
  const [outputName, setOutputName] = useState('');

  const [running, setRunning] = useState(false);
  const [outputSession, setOutputSession] = useState<SessionFile | null>(null);

  const loadTokens = async (file: SelectedFileInfo) => {
    try {
      const res = await apiPost<{
        variables?: string[];
        conditionals?: string[];
        blocks?: string[];
      }>('/api/docx/inspect-tokens', { filepath: file.filePath });
      const vars = res.variables || [];
      setVariables(vars);
      const init: Record<string, string> = {};
      vars.forEach((v) => {
        init[v] = '';
      });
      setValuesMap(init);
      setJsonText(JSON.stringify(init, null, 2));
      toast.info(`Đã tìm thấy ${vars.length} trường {{KEY}} trong mẫu Word`);
    } catch (err) {
      toast.error('Không quét được trường mẫu', err instanceof Error ? err.message : String(err));
    }
  };

  useEffect(() => {
    const hf = consumeHandoffFile();
    if (hf && hf.ext.toLowerCase() === 'docx') {
      const sel = sessionFileToSelected(hf);
      setTplFile(sel);
      void loadTokens(sel);
    }
  }, [consumeHandoffFile]);

  const syncFormToJson = (nextMap: Record<string, string>) => {
    setValuesMap(nextMap);
    setJsonText(JSON.stringify(nextMap, null, 2));
  };

  const handleJsonChange = (txt: string) => {
    setJsonText(txt);
    try {
      const parsed = JSON.parse(txt) as Record<string, unknown>;
      const strMap: Record<string, string> = {};
      Object.entries(parsed).forEach(([k, v]) => {
        strMap[k] = String(v ?? '');
      });
      setValuesMap(strMap);
    } catch {
      // ignore while user is still typing JSON
    }
  };

  const handlePatch = async () => {
    if (!tplFile) {
      toast.warning('Vui lòng chọn tệp mẫu Word (.docx)');
      return;
    }
    let textRep: Record<string, string> = valuesMap;
    if (editorMode === 'json') {
      try {
        textRep = JSON.parse(jsonText) as Record<string, string>;
      } catch {
        toast.error('JSON không hợp lệ', 'Vui lòng kiểm tra cú pháp JSON trước khi chạy.');
        return;
      }
    }

    setRunning(true);
    try {
      const res = await apiPost<{
        filename: string;
        download_url: string;
        result_path?: string;
      }>('/api/docx/patch', {
        filepath: tplFile.filePath,
        text_replacements: textRep,
        output_name: outputName.trim() || undefined,
      });
      const sf = addSessionFile({
        filePath: res.result_path || res.filename,
        filename: res.filename,
        previewUrl: buildPreviewUrl(res.filename),
        downloadUrl: res.download_url || buildDownloadUrl(res.filename),
        sourceLabel: 'Điền mẫu Word',
      });
      setOutputSession(sf);
      toast.success('Đã điền biểu mẫu Word thành công', res.filename);
    } catch (err) {
      toast.error('Lỗi điền biểu mẫu', err instanceof Error ? err.message : String(err));
    } finally {
      setRunning(false);
    }
  };

  const previewTarget = outputSession
    ? {
        filePath: outputSession.filePath,
        filename: outputSession.filename,
        ext: 'docx',
        previewUrl: outputSession.previewUrl,
        downloadUrl: outputSession.downloadUrl,
      }
    : tplFile
    ? {
        filePath: tplFile.filePath,
        filename: tplFile.filename,
        ext: 'docx',
        previewUrl: tplFile.previewUrl,
        downloadUrl: tplFile.downloadUrl,
      }
    : null;

  return (
    <SplitWorkspace
      title="Điền Biểu Mẫu Word Nhanh ({{KEY}})"
      subtitle="Quét tự động mọi placeholder {{KEY}} trong tệp Word, đồng bộ 2 chiều giữa Form nhập liệu và JSON."
      badgeText="Quick Patch"
      accent="word"
      previewTarget={previewTarget}
      leftPanel={
        <div className="space-y-4">
          <FileDropZone
            label="1. Biểu mẫu Word chứa {{KEY}} (.docx)"
            accept=".docx"
            selectedFile={tplFile}
            onFileSelected={(f) => {
              setTplFile(f);
              setOutputSession(null);
              void loadTokens(f);
            }}
            onClear={() => {
              setTplFile(null);
              setVariables([]);
            }}
            sourceLabel="Quick Patch"
          />

          <div className="flex items-center justify-between">
            <Tabs
              activeId={editorMode}
              onChange={(id) => setEditorMode(id as 'form' | 'json')}
              options={[
                { id: 'form', label: `Form (${variables.length})`, icon: <FileText className="h-3.5 w-3.5" /> },
                { id: 'json', label: 'JSON trực tiếp', icon: <Code2 className="h-3.5 w-3.5" /> },
              ]}
            />
          </div>

          {editorMode === 'form' ? (
            variables.length > 0 ? (
              <div className="max-h-64 space-y-2 overflow-y-auto pr-1">
                {variables.map((vKey) => (
                  <div key={vKey} className="space-y-1">
                    <label className="block font-mono text-[11px] font-medium text-blue-700 dark:text-blue-300">
                      {`{{${vKey}}}`}
                    </label>
                    <input
                      type="text"
                      value={valuesMap[vKey] || ''}
                      onChange={(e) => syncFormToJson({ ...valuesMap, [vKey]: e.target.value })}
                      placeholder={`Nhập giá trị cho ${vKey}...`}
                      className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
                    />
                  </div>
                ))}
              </div>
            ) : (
              <div className="rounded-lg border border-dashed border-slate-200 p-4 text-center text-xs text-slate-400 dark:border-slate-800">
                Nạp tệp .docx để tự động hiển thị danh sách biến điền mẫu, hoặc chuyển sang tab JSON để nhập cặp key-value tùy ý.
              </div>
            )
          ) : (
            <textarea
              rows={8}
              value={jsonText}
              onChange={(e) => handleJsonChange(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-slate-50 p-3 font-mono text-xs text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
          )}

          <CollapsibleSection title="Tùy chọn nâng cao (Tên tệp đầu ra)">
            <input
              type="text"
              value={outputName}
              onChange={(e) => setOutputName(e.target.value)}
              placeholder="Vd: HopDong_DaDien.docx"
              className="w-full rounded border border-slate-200 px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-800"
            />
          </CollapsibleSection>

          <Button
            accent="word"
            loading={running}
            disabled={!tplFile}
            onClick={handlePatch}
            icon={<Wand2 className="h-4 w-4" />}
            className="w-full"
          >
            Xuất Tệp Word Đã Điền
          </Button>
        </div>
      }
      resultPanel={
        outputSession ? (
          <ResultPreviewCard
            title="Đã điền xong biểu mẫu Word"
            outputFile={outputSession}
          />
        ) : undefined
      }
    />
  );
}
