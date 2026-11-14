import type {
  CopilotChatResponse,
  CopilotStatusResponse,
  DoctorReport,
  HealthResponse,
  HistoryItem,
  JobStatusResponse,
  JobSubmitResponse,
  MaintenanceReport,
  SessionPurgeResponse,
  SystemVitals,
  UploadResponse,
} from '../types/api';

async function parseErrorMessage(response: Response): Promise<string> {
  try {
    const data = (await response.json()) as { error?: string; message?: string };
    return data.error || data.message || `${response.status} ${response.statusText}`;
  } catch {
    return `${response.status} ${response.statusText}`;
  }
}

export async function apiGet<T>(url: string): Promise<T> {
  const response = await fetch(url, { method: 'GET' });
  if (!response.ok) {
    throw new Error(await parseErrorMessage(response));
  }
  return response.json() as Promise<T>;
}

export async function apiPost<T>(url: string, body: Record<string, unknown> = {}): Promise<T> {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = (await response.json().catch(() => ({}))) as T & { error?: string; success?: boolean };
  if (!response.ok || (data && data.success === false && data.error)) {
    throw new Error(data?.error || `${response.status} ${response.statusText}`);
  }
  return data;
}

export async function apiDelete<T>(url: string): Promise<T> {
  const response = await fetch(url, { method: 'DELETE' });
  if (!response.ok) {
    throw new Error(await parseErrorMessage(response));
  }
  return response.json() as Promise<T>;
}

export async function uploadFiles(files: File[]): Promise<UploadResponse> {
  if (!files.length) {
    throw new Error('Chưa chọn tệp để tải lên.');
  }
  const formData = new FormData();
  files.forEach((file) => formData.append('file', file));
  const response = await fetch('/api/upload', {
    method: 'POST',
    body: formData,
  });
  const data = (await response.json().catch(() => ({}))) as UploadResponse;
  if (!response.ok || data.success === false) {
    throw new Error(data.error || `Tải tệp lên thất bại: ${response.statusText}`);
  }
  return data;
}

export async function uploadSingleFile(file: File): Promise<{
  filePath: string;
  filename: string;
  size: number;
  previewUrl: string;
}> {
  const res = await uploadFiles([file]);
  const first = res.first_file || res.files?.[0];
  const filePath = first?.file_path || res.file_path || res.filepath;
  if (!filePath) {
    throw new Error('Máy chủ không trả về đường dẫn tệp sau khi tải lên.');
  }
  const filename = first?.filename || res.filename || file.name;
  return {
    filePath,
    filename,
    size: first?.size || res.size || file.size,
    previewUrl: first?.preview_url || res.preview_url || buildPreviewUrl(filename),
  };
}

export function buildDownloadUrl(filenameOrPath: string): string {
  const name = extractBaseFilename(filenameOrPath);
  return `/api/download/${encodeURIComponent(name)}`;
}

export function buildPreviewUrl(filenameOrPath: string): string {
  const name = extractBaseFilename(filenameOrPath);
  return `/api/preview/${encodeURIComponent(name)}`;
}

export function extractBaseFilename(input: string): string {
  if (!input) return '';
  const parts = input.split(/[\\/]/);
  return parts[parts.length - 1] || input;
}

export function getExtension(filename: string): string {
  const idx = filename.lastIndexOf('.');
  return idx >= 0 ? filename.slice(idx + 1).toLowerCase() : '';
}

export async function submitJob(action: string, payload: Record<string, unknown> = {}): Promise<JobSubmitResponse> {
  return apiPost<JobSubmitResponse>('/api/jobs/submit', { action, payload });
}

export async function getJobStatus(jobId: string): Promise<JobStatusResponse> {
  return apiGet<JobStatusResponse>(`/api/jobs/${encodeURIComponent(jobId)}`);
}

export function streamJob(
  jobId: string,
  callbacks: {
    onProgress?: (progress: number, message: string, data: JobStatusResponse) => void;
    onComplete?: (result: Record<string, unknown> | undefined, data: JobStatusResponse) => void;
    onError?: (err: Error, data?: JobStatusResponse) => void;
  } = {}
): () => void {
  let cancelled = false;
  let es: EventSource | null = null;
  let pollTimer: number | null = null;

  const cleanup = () => {
    if (es) {
      es.close();
      es = null;
    }
    if (pollTimer !== null) {
      window.clearInterval(pollTimer);
      pollTimer = null;
    }
  };

  const handlePayload = (data: JobStatusResponse) => {
    if (cancelled) return;
    if (typeof data.progress === 'number') {
      callbacks.onProgress?.(data.progress, data.message || '', data);
    }
    if (data.status === 'completed') {
      cleanup();
      callbacks.onComplete?.(data.result, data);
    } else if (data.status === 'failed') {
      cleanup();
      callbacks.onError?.(new Error(data.error || 'Tác vụ xử lý thất bại'), data);
    }
  };

  const startPolling = () => {
    if (cancelled || pollTimer !== null) return;
    pollTimer = window.setInterval(async () => {
      try {
        const job = await getJobStatus(jobId);
        handlePayload(job);
      } catch (err) {
        if (!cancelled) {
          callbacks.onError?.(err instanceof Error ? err : new Error(String(err)));
        }
        cleanup();
      }
    }, 900);
  };

  if (typeof window !== 'undefined' && 'EventSource' in window) {
    try {
      es = new EventSource(`/api/jobs/${encodeURIComponent(jobId)}/stream`);
      es.onmessage = (event) => {
        try {
          handlePayload(JSON.parse(event.data) as JobStatusResponse);
        } catch {
          // ignore malformed heartbeat/SSE chunks
        }
      };
      es.onerror = () => {
        if (es) {
          es.close();
          es = null;
        }
        startPolling();
      };
    } catch {
      startPolling();
    }
  } else {
    startPolling();
  }

  return () => {
    cancelled = true;
    cleanup();
  };
}

export async function runAsyncJob(
  action: string,
  payload: Record<string, unknown>,
  onProgress?: (pct: number, msg: string) => void
): Promise<Record<string, unknown>> {
  const submit = await submitJob(action, payload);
  return new Promise((resolve, reject) => {
    streamJob(submit.job_id, {
      onProgress: (pct, msg) => onProgress?.(pct, msg),
      onComplete: (res) => resolve(res || {}),
      onError: (err) => reject(err),
    });
  });
}

export async function getHealth(): Promise<HealthResponse> {
  return apiGet<HealthResponse>('/api/health');
}

export async function getDoctorReport(): Promise<DoctorReport> {
  return apiGet<DoctorReport>('/api/doctor');
}

export async function getSystemVitals(): Promise<SystemVitals> {
  const res = await apiGet<{ success?: boolean; vitals?: SystemVitals } & SystemVitals>('/api/system/vitals');
  return res.vitals || res;
}

export async function runSystemMaintenance(options: Record<string, unknown>): Promise<MaintenanceReport> {
  const res = await apiPost<{ success?: boolean; report?: MaintenanceReport } & MaintenanceReport>(
    '/api/system/maintenance',
    options
  );
  return res.report || res;
}

export async function purgeSessionTraces(purgeHistory = false): Promise<SessionPurgeResponse> {
  return apiPost<SessionPurgeResponse>('/api/session/purge', { purge_history: purgeHistory });
}

export async function getHistory(): Promise<HistoryItem[]> {
  const res = await apiGet<{ items?: HistoryItem[] }>('/api/history');
  return res.items || [];
}

export async function deleteHistoryItem(id: number): Promise<{ success: boolean }> {
  return apiDelete<{ success: boolean }>(`/api/history/${id}`);
}

export async function clearHistory(): Promise<{ success: boolean }> {
  return apiDelete<{ success: boolean }>('/api/history');
}

export async function getCopilotStatus(): Promise<CopilotStatusResponse> {
  return apiGet<CopilotStatusResponse>('/api/copilot/status');
}

export async function indexFileForRag(filePath: string): Promise<Record<string, unknown>> {
  return apiPost<Record<string, unknown>>('/api/copilot/index', { file_path: filePath });
}

export async function chatCopilot(payload: {
  prompt: string;
  provider?: string;
  active_file?: string;
  top_k?: number;
}): Promise<CopilotChatResponse> {
  return apiPost<CopilotChatResponse>('/api/copilot/chat', payload);
}
