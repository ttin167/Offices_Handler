import React, { Component, type ErrorInfo, type ReactNode } from 'react';
import { Lock, RefreshCw } from 'lucide-react';
import { HeaderBar } from './components/layout/HeaderBar';
import { LeftSidebar } from './components/layout/LeftSidebar';
import { SessionFileTray } from './components/layout/SessionFileTray';
import { CommandPalette } from './components/layout/CommandPalette';
import { CopilotSidebar } from './components/modals/CopilotSidebar';
import { PrintStudioModal } from './components/modals/PrintStudioModal';
import { HistoryDrawer } from './components/modals/HistoryDrawer';
import { SystemDiagnosticsModal } from './components/modals/SystemDiagnosticsModal';
import { QuickGuideModal } from './components/modals/QuickGuideModal';
import { useWorkspace } from './context/WorkspaceContext';
import { HomeLauncherPage } from './features/home/HomeLauncherPage';
import { WordPolishPage } from './features/word/WordPolishPage';
import { WordMergeSplitWatermarkPage } from './features/word/WordMergeSplitWatermarkPage';
import { MailMergePage } from './features/word/MailMergePage';
import { QuickPatchPage } from './features/word/QuickPatchPage';
import { WordDiffPage } from './features/word/WordDiffPage';
import { PdfProHubPage } from './features/pdf/PdfProHubPage';
import { ExcelPowerSuitePage } from './features/excel/ExcelPowerSuitePage';
import { ExcelReportingPage } from './features/excel/ExcelReportingPage';
import { WorkTrackerPage } from './features/excel/WorkTrackerPage';
import { FormulaPlaybookPage } from './features/excel/FormulaPlaybookPage';
import { ExcelTemplatePage } from './features/excel/ExcelTemplatePage';
import { UniversalConverterPage } from './features/convert/UniversalConverterPage';
import { OcrVietnamesePage } from './features/convert/OcrVietnamesePage';
import { QaSecurityPage } from './features/convert/QaSecurityPage';
import { PresentationStudioPage } from './features/visual/PresentationStudioPage';
import { DiagramStudioPage } from './features/visual/DiagramStudioPage';
import type { ToolId } from './types/workspace';

interface ErrorBoundaryState {
  hasError: boolean;
  message: string;
}

class StudioErrorBoundary extends Component<{ children: ReactNode }, ErrorBoundaryState> {
  state: ErrorBoundaryState = { hasError: false, message: '' };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, message: error.message || 'Lỗi hiển thị giao diện' };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error('Office Studio ErrorBoundary:', error, info);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="flex min-h-[320px] flex-col items-center justify-center gap-3 rounded-xl border border-red-200 bg-red-50/60 p-8 text-center dark:border-red-900/60 dark:bg-red-950/30">
          <h3 className="text-base font-semibold text-red-900 dark:text-red-200">
            Đã xảy ra sự cố cục bộ ở mô-đun này
          </h3>
          <p className="max-w-md text-xs text-red-700 dark:text-red-300">{this.state.message}</p>
          <button
            type="button"
            onClick={() => this.setState({ hasError: false, message: '' })}
            className="inline-flex items-center gap-1.5 rounded-lg bg-red-600 px-3.5 py-2 text-xs font-medium text-white hover:bg-red-700"
          >
            <RefreshCw className="h-3.5 w-3.5" />
            <span>Tải lại mô-đun</span>
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

function renderActivePage(tool: ToolId): React.ReactNode {
  switch (tool) {
    case 'home':
      return <HomeLauncherPage />;
    case 'word-polish':
      return <WordPolishPage />;
    case 'word-merge-split':
      return <WordMergeSplitWatermarkPage />;
    case 'word-mail-merge':
      return <MailMergePage />;
    case 'word-quick-patch':
      return <QuickPatchPage />;
    case 'word-diff':
      return <WordDiffPage />;
    case 'pdf-pro':
      return <PdfProHubPage />;
    case 'excel-power':
      return <ExcelPowerSuitePage />;
    case 'excel-reporting':
      return <ExcelReportingPage />;
    case 'excel-tracker':
      return <WorkTrackerPage />;
    case 'excel-formulas':
      return <FormulaPlaybookPage />;
    case 'excel-template':
      return <ExcelTemplatePage />;
    case 'convert-universal':
      return <UniversalConverterPage />;
    case 'ocr-vietnamese':
      return <OcrVietnamesePage />;
    case 'qa-security':
      return <QaSecurityPage />;
    case 'presentation-studio':
      return <PresentationStudioPage />;
    case 'diagram-studio':
      return <DiagramStudioPage />;
  }
}

export function App() {
  const { activeTool } = useWorkspace();

  return (
    <div className="flex min-h-screen flex-col bg-slate-50 text-slate-900 antialiased dark:bg-slate-950 dark:text-slate-100">
      <HeaderBar />

      <div className="flex flex-1 min-h-0">
        <LeftSidebar />

        <div className="flex min-w-0 flex-1 flex-col">
          <SessionFileTray />

          <main className="mx-auto w-full max-w-[1600px] flex-1 px-4 py-5 sm:px-6">
            <StudioErrorBoundary key={activeTool}>
              {renderActivePage(activeTool)}
            </StudioErrorBoundary>
          </main>

          <footer className="border-t border-slate-200 bg-white/80 px-4 py-2.5 text-[11px] text-slate-500 dark:border-slate-800 dark:bg-slate-900/70 dark:text-slate-400">
            <div className="mx-auto flex max-w-[1600px] flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-1.5">
                <Lock className="h-3 w-3 text-emerald-600 dark:text-emerald-400" />
                <span>
                  Antigravity Pure Office Studio v5.0.0 · Bảo mật Thuần Local (<code>127.0.0.1</code>) · Zero Cloud Leakage
                </span>
              </div>
              <span>Nhấn Ctrl+K để mở Bảng Lệnh Nhanh</span>
            </div>
          </footer>
        </div>
      </div>

      <CommandPalette />
      <CopilotSidebar />
      <PrintStudioModal />
      <HistoryDrawer />
      <SystemDiagnosticsModal />
      <QuickGuideModal />
    </div>
  );
}

export default App;
