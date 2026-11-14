export interface ApiResponseBase {
  success?: boolean;
  error?: string;
  message?: string;
}

export interface UploadedFileItem {
  file_path: string;
  filename: string;
  size: number;
  preview_url?: string;
}

export interface UploadResponse extends ApiResponseBase {
  filepath?: string;
  file_path?: string;
  filename?: string;
  output_file?: string;
  size?: number;
  preview_url?: string;
  files?: UploadedFileItem[];
  first_file?: UploadedFileItem | null;
}

export interface JobSubmitResponse extends ApiResponseBase {
  job_id: string;
  status: string;
  stream_url: string;
  status_url: string;
}

export interface JobStatusResponse extends ApiResponseBase {
  job_id: string;
  action?: string;
  status: 'pending' | 'running' | 'completed' | 'failed';
  progress?: number;
  message?: string;
  result?: Record<string, unknown>;
}

export interface HealthResponse {
  status: string;
  service: string;
  version: string;
  local_only?: boolean;
  host?: string;
  timestamp: number;
}

export interface SystemVitals {
  platform?: string;
  python_version?: string;
  cpu_count?: number;
  memory_percent?: number;
  memory_available_gb?: number;
  disk_free_gb?: number;
  disk_total_gb?: number;
  [key: string]: unknown;
}

export interface DoctorCheckItem {
  name?: string;
  component?: string;
  status: 'ok' | 'warn' | 'error' | 'missing' | string;
  version?: string;
  detail?: string;
  path?: string;
}

export interface DoctorReport {
  version?: string;
  overall_status?: string;
  score?: number;
  checks?: DoctorCheckItem[] | Record<string, DoctorCheckItem>;
  summary?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MaintenanceReport {
  timestamp?: string;
  execution_time_sec?: number;
  dry_run?: boolean;
  lock_files_removed?: number;
  temp_files_removed?: number;
  pycache_dirs_removed?: number;
  pycache_files_removed?: number;
  test_cache_dirs_removed?: number;
  total_bytes_freed?: number;
  errors?: string[];
  [key: string]: unknown;
}

export interface SessionPurgeResponse extends ApiResponseBase {
  removed_files: number;
  bytes_freed: number;
  mb_freed: number;
  history_cleared: boolean;
  details: string[];
}

export interface HistoryItem {
  id: number;
  action?: string;
  tool?: string;
  filename?: string;
  output_name?: string;
  output_path?: string;
  file_size?: number;
  created_at?: string;
  timestamp?: string;
  status?: string;
}

export interface DocxRun {
  text?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  font_size_pt?: number;
}

export interface DocxTableCell {
  shading?: string;
  vertical_alignment?: string;
  paragraphs?: Array<{ text?: string }>;
}

export interface DocxBodyElement {
  type?: 'paragraph' | 'table' | string;
  style_name?: string;
  alignment?: string;
  runs?: DocxRun[];
  has_header?: boolean;
  rows?: Array<{ cells?: DocxTableCell[] }>;
}

export interface DocxSnapshot {
  body?: DocxBodyElement[];
  stats?: Record<string, unknown>;
}

export interface XlsxCellData {
  column_number: number;
  value?: unknown;
  formatted_value?: string;
  formula?: string;
  bold?: boolean;
  italic?: boolean;
  font_color?: string;
  fill_color?: string;
  alignment?: string;
}

export interface XlsxRowData {
  row_number?: number;
  cells?: XlsxCellData[];
}

export interface XlsxSheetSnapshot {
  name?: string;
  sheet_name?: string;
  stats?: {
    max_column?: number;
    max_row?: number;
  };
  rows?: XlsxRowData[];
}

export interface XlsxSnapshot {
  sheets?: XlsxSheetSnapshot[];
}

export interface CopilotActionCard {
  action_id: string;
  title: string;
  description: string;
  command_type: string;
  payload: Record<string, unknown>;
  badge?: string;
}

export interface CopilotStatusResponse extends ApiResponseBase {
  copilot: {
    default_mode: string;
    local_only_enforced?: boolean;
    local: {
      online: boolean;
      model?: string;
      base_url?: string;
      fallback_ready?: boolean;
    };
    cloud: {
      online: boolean;
      model?: string;
      blocked_by_local_policy?: boolean;
    };
  };
  rag: {
    total_documents?: number;
    total_chunks?: number;
    db_path?: string;
  };
}

export interface CopilotChatResponse extends ApiResponseBase {
  provider: string;
  model: string;
  answer: string;
  action_cards?: CopilotActionCard[];
  citations?: Array<{
    doc_id?: string;
    filename?: string;
    section_title?: string;
  }>;
  latency_ms?: number;
  fallback_used?: boolean;
}
