import React, { createContext, useCallback, useContext, useEffect, useState } from 'react';
import type { SessionFile, ToolId } from '../types/workspace';

interface WorkspaceContextValue {
  activeTool: ToolId;
  handoffFile: SessionFile | null;
  navigateToTool: (tool: ToolId, fileToLoad?: SessionFile | null) => void;
  consumeHandoffFile: () => SessionFile | null;
  sidebarCollapsed: boolean;
  setSidebarCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  copilotOpen: boolean;
  setCopilotOpen: React.Dispatch<React.SetStateAction<boolean>>;
  commandPaletteOpen: boolean;
  setCommandPaletteOpen: React.Dispatch<React.SetStateAction<boolean>>;
  historyOpen: boolean;
  setHistoryOpen: React.Dispatch<React.SetStateAction<boolean>>;
  diagnosticsOpen: boolean;
  setDiagnosticsOpen: React.Dispatch<React.SetStateAction<boolean>>;
  guideOpen: boolean;
  setGuideOpen: React.Dispatch<React.SetStateAction<boolean>>;
  printTargetFile: { filePath: string; filename: string } | null;
  openPrintModal: (file?: { filePath: string; filename: string } | null) => void;
  printModalOpen: boolean;
  closePrintModal: () => void;
}

const WorkspaceContext = createContext<WorkspaceContextValue | null>(null);

export function WorkspaceProvider({ children }: { children: React.ReactNode }) {
  const [activeTool, setActiveTool] = useState<ToolId>('home');
  const [handoffFile, setHandoffFile] = useState<SessionFile | null>(null);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [copilotOpen, setCopilotOpen] = useState(false);
  const [commandPaletteOpen, setCommandPaletteOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(false);
  const [guideOpen, setGuideOpen] = useState(false);
  const [printModalOpen, setPrintModalOpen] = useState(false);
  const [printTargetFile, setPrintTargetFile] = useState<{ filePath: string; filename: string } | null>(null);

  const navigateToTool = useCallback((tool: ToolId, fileToLoad?: SessionFile | null) => {
    setActiveTool(tool);
    if (fileToLoad) {
      setHandoffFile(fileToLoad);
    }
  }, []);

  const consumeHandoffFile = useCallback(() => {
    const current = handoffFile;
    setHandoffFile(null);
    return current;
  }, [handoffFile]);

  const openPrintModal = useCallback((file?: { filePath: string; filename: string } | null) => {
    setPrintTargetFile(file || null);
    setPrintModalOpen(true);
  }, []);

  const closePrintModal = useCallback(() => {
    setPrintModalOpen(false);
  }, []);

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setCommandPaletteOpen((prev) => !prev);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  return (
    <WorkspaceContext.Provider
      value={{
        activeTool,
        handoffFile,
        navigateToTool,
        consumeHandoffFile,
        sidebarCollapsed,
        setSidebarCollapsed,
        copilotOpen,
        setCopilotOpen,
        commandPaletteOpen,
        setCommandPaletteOpen,
        historyOpen,
        setHistoryOpen,
        diagnosticsOpen,
        setDiagnosticsOpen,
        guideOpen,
        setGuideOpen,
        printTargetFile,
        openPrintModal,
        printModalOpen,
        closePrintModal,
      }}
    >
      {children}
    </WorkspaceContext.Provider>
  );
}

export function useWorkspace(): WorkspaceContextValue {
  const ctx = useContext(WorkspaceContext);
  if (!ctx) throw new Error('useWorkspace must be used inside WorkspaceProvider');
  return ctx;
}
