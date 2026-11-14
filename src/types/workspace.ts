export type ToolId =
  | 'home'
  | 'word-polish'
  | 'word-merge-split'
  | 'word-mail-merge'
  | 'word-quick-patch'
  | 'word-diff'
  | 'pdf-pro'
  | 'excel-power'
  | 'excel-reporting'
  | 'excel-tracker'
  | 'excel-formulas'
  | 'excel-template'
  | 'convert-universal'
  | 'ocr-vietnamese'
  | 'qa-security'
  | 'presentation-studio'
  | 'diagram-studio';

export type DomainAccent = 'neutral' | 'word' | 'excel' | 'pdf' | 'slide' | 'studio';

export interface ToolNavItem {
  id: ToolId;
  label: string;
  shortDesc: string;
  domain: 'home' | 'word' | 'pdf' | 'excel' | 'convert' | 'visual';
  accent: DomainAccent;
  keywords: string[];
  supportedExts: string[];
}

export interface SessionFile {
  id: string;
  filename: string;
  filePath: string;
  ext: string;
  size?: number;
  previewUrl?: string;
  downloadUrl?: string;
  sourceLabel: string;
  addedAt: number;
}

export type ThemeMode = 'light' | 'dark';

export type ToastVariant = 'info' | 'success' | 'warning' | 'error';

export interface ToastItem {
  id: string;
  title: string;
  description?: string;
  variant: ToastVariant;
}
