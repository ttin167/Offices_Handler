import React, { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, Command, Printer, Search, Sparkles, SunMoon, Activity, History } from 'lucide-react';
import { TOOL_CATALOG, ACCENT_CLASSES } from '../../config/navigation';
import { useWorkspace } from '../../context/WorkspaceContext';
import { useTheme } from '../../context/ThemeContext';
import type { ToolNavItem } from '../../types/workspace';

function normalizeVi(text: string): string {
  return text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

interface ActionOption {
  id: string;
  label: string;
  desc: string;
  icon: React.ReactNode;
  run: () => void;
}

export function CommandPalette() {
  const {
    commandPaletteOpen,
    setCommandPaletteOpen,
    navigateToTool,
    openPrintModal,
    setHistoryOpen,
    setDiagnosticsOpen,
    setCopilotOpen,
  } = useWorkspace();
  const { toggleTheme } = useTheme();
  const [query, setQuery] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (commandPaletteOpen) {
      setQuery('');
      setSelectedIndex(0);
      window.setTimeout(() => inputRef.current?.focus(), 20);
    }
  }, [commandPaletteOpen]);

  const quickActions: ActionOption[] = useMemo(
    () => [
      {
        id: 'act-print',
        label: 'Mở Trạm In Ấn Doanh nghiệp',
        desc: 'In tài liệu PDF/Word/Excel, chọn khổ A4, in 2 mặt, đóng quyển Booklet',
        icon: <Printer className="h-4 w-4 text-slate-500" />,
        run: () => openPrintModal(null),
      },
      {
        id: 'act-copilot',
        label: 'Bật / Tắt Trợ lý AI Copilot (Local RAG)',
        desc: 'Hỏi đáp nội bộ trên tài liệu đã lập chỉ mục',
        icon: <Sparkles className="h-4 w-4 text-indigo-500" />,
        run: () => setCopilotOpen((p) => !p),
      },
      {
        id: 'act-history',
        label: 'Xem Lịch sử xử lý tệp',
        desc: 'Mở danh sách tệp vừa tạo gần đây từ SQLite',
        icon: <History className="h-4 w-4 text-slate-500" />,
        run: () => setHistoryOpen(true),
      },
      {
        id: 'act-doctor',
        label: 'Kiểm tra hệ thống & Bảo trì 7 Trụ cột',
        desc: 'Chẩn đoán LibreOffice, Tesseract, RAM và dọn rác',
        icon: <Activity className="h-4 w-4 text-emerald-500" />,
        run: () => setDiagnosticsOpen(true),
      },
      {
        id: 'act-theme',
        label: 'Đổi chế độ Sáng / Tối',
        desc: 'Chuyển đổi giữa nền giấy Sáng và nền Slate Tối',
        icon: <SunMoon className="h-4 w-4 text-amber-500" />,
        run: () => toggleTheme(),
      },
    ],
    [openPrintModal, setCopilotOpen, setHistoryOpen, setDiagnosticsOpen, toggleTheme]
  );

  const filteredTools: ToolNavItem[] = useMemo(() => {
    const q = normalizeVi(query.trim());
    if (!q) return TOOL_CATALOG;
    return TOOL_CATALOG.filter((t) => {
      const hay = normalizeVi(`${t.label} ${t.shortDesc} ${t.keywords.join(' ')}`);
      return hay.includes(q);
    });
  }, [query]);

  const filteredActions = useMemo(() => {
    const q = normalizeVi(query.trim());
    if (!q) return quickActions;
    return quickActions.filter((a) => normalizeVi(`${a.label} ${a.desc}`).includes(q));
  }, [query, quickActions]);

  const combinedCount = filteredTools.length + filteredActions.length;

  useEffect(() => {
    if (!commandPaletteOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setCommandPaletteOpen(false);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIndex((i) => (combinedCount ? (i + 1) % combinedCount : 0));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIndex((i) => (combinedCount ? (i - 1 + combinedCount) % combinedCount : 0));
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (selectedIndex < filteredTools.length) {
          const tool = filteredTools[selectedIndex];
          if (tool) {
            navigateToTool(tool.id);
            setCommandPaletteOpen(false);
          }
        } else {
          const act = filteredActions[selectedIndex - filteredTools.length];
          if (act) {
            act.run();
            setCommandPaletteOpen(false);
          }
        }
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [commandPaletteOpen, combinedCount, selectedIndex, filteredTools, filteredActions, navigateToTool, setCommandPaletteOpen]);

  if (!commandPaletteOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center bg-slate-950/45 pt-[10vh] px-4 backdrop-blur-[1px]"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) setCommandPaletteOpen(false);
      }}
    >
      <div className="w-full max-w-xl overflow-hidden rounded-xl border border-slate-200 bg-white shadow-elevated dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center gap-2.5 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <Search className="h-4 w-4 text-slate-400" />
          <input
            ref={inputRef}
            type="text"
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              setSelectedIndex(0);
            }}
            placeholder="Nhập tên công cụ hoặc nhu cầu (vd: trộn thư, xlookup, giáp lai, ocr)..."
            className="flex-1 bg-transparent text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none dark:text-slate-100"
          />
          <kbd className="rounded border border-slate-200 px-1.5 py-0.5 text-[10px] font-mono text-slate-400 dark:border-slate-700">
            ESC
          </kbd>
        </div>

        <div className="max-h-[60vh] overflow-y-auto p-2 space-y-2">
          {filteredTools.length > 0 ? (
            <div>
              <div className="px-2 py-1 text-[10px] font-semibold uppercase tracking-wider text-slate-400">
                Công cụ Tin học Văn phòng
              </div>
              {filteredTools.map((t, idx) => {
                const active = idx === selectedIndex;
                return (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => {
                      navigateToTool(t.id);
                      setCommandPaletteOpen(false);
                    }}
                    className={`flex w-full items-center justify-between gap-3 rounded-lg px-3 py-2 text-left transition-colors ${
                      active
                        ? 'bg-indigo-50 text-indigo-950 dark:bg-indigo-950/60 dark:text-indigo-100'
                        : 'hover:bg-slate-100 dark:hover:bg-slate-800/70'
                    }`}
                  >
                    <div className="min-w-0">
                      <div className="flex items-center gap-2">
                        <span className={`text-xs font-semibold ${ACCENT_CLASSES[t.accent].text}`}>{t.label}</span>
                      </div>
                      <p className="truncate text-[11px] text-slate-500 dark:text-slate-400">{t.shortDesc}</p>
                    </div>
                    <ArrowRight className="h-3.5 w-3.5 shrink-0 text-slate-400" />
                  </button>
                );
              })}
            </div>
          ) : null}

          {filteredActions.length > 0 ? (
            <div>
              <div className="px-2 py-1 text-[10px] font-semibold uppercase tracking-wider text-slate-400">
                Tiện ích nhanh
              </div>
              {filteredActions.map((act, idx) => {
                const realIdx = filteredTools.length + idx;
                const active = realIdx === selectedIndex;
                return (
                  <button
                    key={act.id}
                    type="button"
                    onClick={() => {
                      act.run();
                      setCommandPaletteOpen(false);
                    }}
                    className={`flex w-full items-center gap-3 rounded-lg px-3 py-2 text-left transition-colors ${
                      active
                        ? 'bg-indigo-50 dark:bg-indigo-950/60'
                        : 'hover:bg-slate-100 dark:hover:bg-slate-800/70'
                    }`}
                  >
                    {act.icon}
                    <div className="min-w-0 flex-1">
                      <div className="text-xs font-medium text-slate-800 dark:text-slate-100">{act.label}</div>
                      <div className="truncate text-[11px] text-slate-500 dark:text-slate-400">{act.desc}</div>
                    </div>
                  </button>
                );
              })}
            </div>
          ) : null}

          {combinedCount === 0 ? (
            <div className="py-8 text-center text-xs text-slate-500">
              Không tìm thấy công cụ khớp với "{query}".
            </div>
          ) : null}
        </div>

        <div className="flex items-center justify-between border-t border-slate-100 bg-slate-50 px-4 py-2 text-[11px] text-slate-400 dark:border-slate-800 dark:bg-slate-950">
          <span className="inline-flex items-center gap-1">
            <Command className="h-3 w-3" /> Dùng ↑↓ để chọn và Enter để mở
          </span>
          <span>Ctrl+K</span>
        </div>
      </div>
    </div>
  );
}
