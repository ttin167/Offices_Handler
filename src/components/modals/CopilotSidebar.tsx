import { useCallback, useEffect, useState } from 'react';
import {
  ArrowRight,
  Database,
  FilePlus2,
  Lock,
  Send,
  ShieldCheck,
  Sparkles,
} from 'lucide-react';
import {
  Badge,
  Button,
  Drawer,
  FileDropZone,
  sessionFileToSelected,
  type SelectedFileInfo,
} from '../ui';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useSessionFiles } from '../../context/SessionFilesContext';
import { useToast } from '../../context/ToastContext';
import { chatCopilot, getCopilotStatus, indexFileForRag } from '../../services/api';
import type { CopilotActionCard, CopilotStatusResponse } from '../../types/api';
import type { ToolId } from '../../types/workspace';

interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  provider?: string;
  citations?: Array<{ doc_id?: string; filename?: string; section_title?: string }>;
  actions?: CopilotActionCard[];
}

const ACTION_TOOL_MAP: Record<string, ToolId> = {
  word_polish: 'word-polish',
  docx_polish: 'word-polish',
  pdf_pro: 'pdf-pro',
  excel_power: 'excel-power',
  convert: 'convert-universal',
  ocr: 'ocr-vietnamese',
  qa_audit: 'qa-security',
  pptx: 'presentation-studio',
  diagram: 'diagram-studio',
};

export function CopilotSidebar() {
  const { copilotOpen, setCopilotOpen, navigateToTool } = useWorkspace();
  const { sessionFiles } = useSessionFiles();
  const toast = useToast();

  const [status, setStatus] = useState<CopilotStatusResponse | null>(null);
  const [provider, setProvider] = useState<'local' | 'hybrid'>('local');
  const [ragFile, setRagFile] = useState<SelectedFileInfo | null>(null);
  const [indexing, setIndexing] = useState(false);

  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'welcome',
      role: 'assistant',
      content:
        'Xin chào! Tôi là Trợ lý AI Cục bộ (100% Offline RAG FTS5 / Ollama). Hãy nạp một tệp Word, Excel hoặc PDF để tra cứu nội dung và nhận gợi ý thao tác ngay trên máy cá nhân của bạn.',
    },
  ]);
  const [prompt, setPrompt] = useState('');
  const [sending, setSending] = useState(false);

  const loadStatus = useCallback(async () => {
    try {
      const res = await getCopilotStatus();
      setStatus(res);
      if (res.copilot?.local_only_enforced !== false) {
        setProvider('local');
      }
    } catch {
      // ignore initial status fetch failures
    }
  }, []);

  useEffect(() => {
    if (copilotOpen) {
      loadStatus();
    }
  }, [copilotOpen, loadStatus]);

  const handleIndexFile = async (f: SelectedFileInfo) => {
    setRagFile(f);
    setIndexing(true);
    try {
      await indexFileForRag(f.filePath);
      toast.success('Đã lập chỉ mục tìm kiếm RAG', f.filename);
      await loadStatus();
    } catch (err) {
      toast.error('Lỗi lập chỉ mục RAG', err instanceof Error ? err.message : String(err));
    } finally {
      setIndexing(false);
    }
  };

  const handleSend = async (customText?: string) => {
    const text = (customText ?? prompt).trim();
    if (!text) return;

    const userMsg: ChatMessage = { id: `u-${Date.now()}`, role: 'user', content: text };
    setMessages((prev) => [...prev, userMsg]);
    if (!customText) setPrompt('');
    setSending(true);

    try {
      const res = await chatCopilot({
        prompt: text,
        provider: localOnlyEnforced ? 'local' : provider,
        active_file: ragFile?.filePath || sessionFiles[0]?.filePath,
        top_k: 5,
      });
      setMessages((prev) => [
        ...prev,
        {
          id: `a-${Date.now()}`,
          role: 'assistant',
          content: res.answer,
          provider: `${res.provider} (${res.model})`,
          citations: res.citations,
          actions: res.action_cards,
        },
      ]);
    } catch (err) {
      toast.error('Lỗi phản hồi Copilot', err instanceof Error ? err.message : String(err));
    } finally {
      setSending(false);
    }
  };

  const localOnlyEnforced = status?.copilot?.local_only_enforced !== false;

  const handleActionClick = (card: CopilotActionCard) => {
    const mappedTool = ACTION_TOOL_MAP[card.command_type] || 'home';
    setCopilotOpen(false);
    navigateToTool(mappedTool);
    toast.info(`Đã chuyển tới: ${card.title}`);
  };

  return (
    <Drawer
      open={copilotOpen}
      onClose={() => setCopilotOpen(false)}
      title="Trợ lý AI & FTS5 RAG Cục Bộ"
      subtitle="Phân tích hồ sơ ngay trên máy cá nhân (Zero External Leakage)"
    >
      <div className="flex h-full flex-col gap-3">
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-emerald-200 bg-emerald-50/70 px-3 py-2 text-xs dark:border-emerald-900/50 dark:bg-emerald-950/30">
          <div className="flex items-center gap-1.5 font-semibold text-emerald-800 dark:text-emerald-200">
            <Lock className="h-3.5 w-3.5" />
            <span>{localOnlyEnforced ? 'Đã khóa Chế độ 100% Local' : 'Local RAG Sẵn sàng'}</span>
          </div>
          <div className="flex items-center gap-2 text-[11px] text-emerald-700 dark:text-emerald-400">
            <Database className="h-3.5 w-3.5" />
            <span>
              {status?.rag?.total_documents ?? 0} tài liệu · {status?.rag?.total_chunks ?? 0} đoạn
            </span>
          </div>
        </div>

        {!localOnlyEnforced ? (
          <div className="flex items-center justify-between text-xs text-slate-600 dark:text-slate-400">
            <span>Nguồn mô hình AI:</span>
            <div className="flex gap-1.5">
              <button
                type="button"
                onClick={() => setProvider('local')}
                className={`rounded-md px-2 py-1 text-[11px] font-medium ${
                  provider === 'local' ? 'bg-indigo-600 text-white' : 'bg-slate-100 dark:bg-slate-800'
                }`}
              >
                Local RAG
              </button>
              <button
                type="button"
                onClick={() => setProvider('hybrid')}
                className={`rounded-md px-2 py-1 text-[11px] font-medium ${
                  provider === 'hybrid' ? 'bg-indigo-600 text-white' : 'bg-slate-100 dark:bg-slate-800'
                }`}
              >
                Hybrid
              </button>
            </div>
          </div>
        ) : (
          <div className="flex items-center gap-1.5 text-[11px] text-slate-500 dark:text-slate-400">
            <ShieldCheck className="h-3.5 w-3.5 text-emerald-600" />
            <span>AI Đám mây đã tắt mặc định để bảo vệ dữ liệu văn phòng.</span>
          </div>
        )}

        <FileDropZone
          label="Nạp tài liệu vào Kho tri thức RAG cục bộ"
          accept=".pdf,.docx,.xlsx,.md,.txt"
          selectedFile={ragFile}
          onFileSelected={handleIndexFile}
          onClear={() => setRagFile(null)}
          sourceLabel="RAG Index"
          hint={indexing ? 'Đang lập chỉ mục FTS5...' : 'Hoặc chọn nhanh file từ Kho phiên làm việc'}
        />

        {sessionFiles.length > 0 && !ragFile ? (
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="text-[11px] text-slate-500">Nạp nhanh:</span>
            {sessionFiles.slice(0, 3).map((sf) => (
              <button
                key={sf.id}
                type="button"
                onClick={() => handleIndexFile(sessionFileToSelected(sf))}
                className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-[11px] text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
              >
                <FilePlus2 className="h-3 w-3 text-indigo-500" />
                <span className="max-w-[130px] truncate">{sf.filename}</span>
              </button>
            ))}
          </div>
        ) : null}

        <div className="flex-1 space-y-2.5 overflow-y-auto rounded-xl border border-slate-200 bg-slate-50/70 p-3 dark:border-slate-800 dark:bg-slate-950/60">
          {messages.map((m) => (
            <div
              key={m.id}
              className={`rounded-xl px-3 py-2 text-xs leading-relaxed ${
                m.role === 'user'
                  ? 'ml-6 bg-indigo-600 text-white'
                  : 'mr-4 border border-slate-200 bg-white text-slate-800 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-200'
              }`}
            >
              <div className="whitespace-pre-wrap">{m.content}</div>
              {m.provider ? (
                <div className="mt-1.5 text-[10px] text-slate-400">Engine: {m.provider}</div>
              ) : null}
              {m.citations && m.citations.length > 0 ? (
                <div className="mt-1.5 flex flex-wrap gap-1">
                  {m.citations.slice(0, 3).map((c, ci) => (
                    <Badge key={ci} tone="slate">
                      {c.filename} {c.section_title ? `• ${c.section_title}` : ''}
                    </Badge>
                  ))}
                </div>
              ) : null}
              {m.actions && m.actions.length > 0 ? (
                <div className="mt-2 space-y-1">
                  {m.actions.map((card) => (
                    <button
                      key={card.action_id}
                      type="button"
                      onClick={() => handleActionClick(card)}
                      className="flex w-full items-center justify-between rounded-lg border border-indigo-200 bg-indigo-50/70 px-2.5 py-1.5 text-left text-[11px] font-medium text-indigo-800 hover:bg-indigo-100 dark:border-indigo-900/60 dark:bg-indigo-950/40 dark:text-indigo-200"
                    >
                      <span>{card.title}</span>
                      <ArrowRight className="h-3 w-3 shrink-0" />
                    </button>
                  ))}
                </div>
              ) : null}
            </div>
          ))}
        </div>

        <div className="space-y-2">
          <div className="flex flex-wrap gap-1">
            {[
              'Tóm tắt nội dung chính',
              'Liệt kê số liệu & điều khoản quan trọng',
              'Đề xuất bước xử lý tiếp theo',
            ].map((q) => (
              <button
                key={q}
                type="button"
                onClick={() => handleSend(q)}
                className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-[10px] text-slate-600 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300"
              >
                <Sparkles className="h-2.5 w-2.5 text-indigo-500" />
                <span>{q}</span>
              </button>
            ))}
          </div>

          <div className="flex gap-2">
            <input
              type="text"
              value={prompt}
              onChange={(e) => setPrompt(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  handleSend();
                }
              }}
              placeholder="Đặt câu hỏi về tài liệu đã nạp..."
              className="flex-1 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs text-slate-900 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
            <Button
              variant="primary"
              loading={sending}
              disabled={!prompt.trim()}
              icon={<Send className="h-3.5 w-3.5" />}
              onClick={() => handleSend()}
              className="bg-indigo-600 hover:bg-indigo-700"
            >
              Gửi
            </Button>
          </div>
        </div>
      </div>
    </Drawer>
  );
}
