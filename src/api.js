/**
 * src/api.js — Frontend API Client for Antigravity Office Studio
 *
 * Provides typed convenience methods for asynchronous job submissions,
 * real-time Server-Sent Events (SSE) streaming with automated polling fallback,
 * and system environment diagnostics.
 */

/**
 * Submit an asynchronous job to the server queue.
 * @param {string} action - Action name ('convert', 'pptx_build', 'xlsx_mutate', 'diagram_render')
 * @param {object} payload - Action-specific payload
 * @returns {Promise<{success: boolean, job_id: string, stream_url: string, status_url: string}>}
 */
export async function submitJob(action, payload = {}) {
  const response = await fetch('/api/jobs/submit', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action, payload })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error || `Failed to submit job: ${response.statusText}`);
  }

  return response.json();
}

/**
 * Retrieve current status of a specific job.
 * @param {string} jobId
 * @returns {Promise<object>}
 */
export async function getJobStatus(jobId) {
  const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`);
  if (!response.ok) {
    throw new Error(`Failed to fetch job status: ${response.statusText}`);
  }
  return response.json();
}

/**
 * List recent queued and completed jobs.
 * @returns {Promise<{jobs: Array<object>, total: number}>}
 */
export async function listJobs() {
  const response = await fetch('/api/jobs');
  if (!response.ok) {
    throw new Error(`Failed to list jobs: ${response.statusText}`);
  }
  return response.json();
}

/**
 * Run Zero-Setup Environment Doctor diagnostics.
 * @returns {Promise<object>}
 */
export async function getDoctorReport() {
  const response = await fetch('/api/doctor');
  if (!response.ok) {
    throw new Error(`Failed to fetch doctor report: ${response.statusText}`);
  }
  return response.json();
}

/**
 * Check backend health.
 * @returns {Promise<object>}
 */
export async function getHealth() {
  const response = await fetch('/api/health');
  if (!response.ok) {
    throw new Error(`Backend health check failed: ${response.statusText}`);
  }
  return response.json();
}

/**
 * Listen to job progress via Server-Sent Events (SSE) with automated polling fallback.
 *
 * @param {string} jobId
 * @param {object} callbacks
 * @param {function(number, string, object): void} [callbacks.onProgress] - (progress, message, data)
 * @param {function(any, object): void} [callbacks.onComplete] - (result, data)
 * @param {function(Error, object): void} [callbacks.onError] - (error, data)
 * @param {function(object): void} [callbacks.onMessage] - (data)
 * @returns {function(): void} cancel/unsubscribe function
 */
export function streamJob(jobId, { onProgress, onComplete, onError, onMessage } = {}) {
  let isCancelled = false;
  let activeEventSource = null;
  let pollInterval = null;

  const handlePayload = (data) => {
    if (isCancelled) return;

    if (onMessage) onMessage(data);
    if (onProgress && typeof data.progress === 'number') {
      onProgress(data.progress, data.message || '', data);
    }

    if (data.status === 'completed') {
      cleanup();
      if (onComplete) onComplete(data.result, data);
    } else if (data.status === 'failed') {
      cleanup();
      if (onError) onError(new Error(data.error || 'Job failed'), data);
    }
  };

  const cleanup = () => {
    if (activeEventSource) {
      activeEventSource.close();
      activeEventSource = null;
    }
    if (pollInterval) {
      clearInterval(pollInterval);
      pollInterval = null;
    }
  };

  const startPolling = () => {
    if (isCancelled || pollInterval) return;
    pollInterval = setInterval(async () => {
      try {
        const job = await getJobStatus(jobId);
        handlePayload(job);
      } catch (err) {
        if (!isCancelled && onError) {
          onError(err, { status: 'failed', error: err.message });
        }
        cleanup();
      }
    }, 1000);
  };

  // Attempt Native EventSource connection
  if (typeof window !== 'undefined' && 'EventSource' in window) {
    try {
      const streamUrl = `/api/jobs/${encodeURIComponent(jobId)}/stream`;
      const es = new EventSource(streamUrl);
      activeEventSource = es;

      es.onmessage = (event) => {
        try {
          const parsed = JSON.parse(event.data);
          handlePayload(parsed);
        } catch {
          // ignore malformed SSE frames
        }
      };

      es.onerror = () => {
        // Fallback to polling on network disconnection or proxy termination
        if (activeEventSource) {
          activeEventSource.close();
          activeEventSource = null;
        }
        startPolling();
      };
    } catch {
      startPolling();
    }
  } else {
    startPolling();
  }

  // Return cancel subscription closure
  return () => {
    isCancelled = true;
    cleanup();
  };
}
