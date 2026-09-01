export async function withRequestTimeout<T>(
  operation: (signal: AbortSignal) => Promise<T>,
  timeoutMs: number,
  timeoutMessage: string,
): Promise<T> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await operation(controller.signal);
  } catch (error) {
    if (controller.signal.aborted) throw new Error(timeoutMessage);
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function mergeAbortSignals(timeoutSignal: AbortSignal, requestSignal?: AbortSignal | null): AbortSignal {
  if (!requestSignal) return timeoutSignal;
  if (requestSignal.aborted) return requestSignal;

  if (typeof AbortSignal.any === "function") {
    return AbortSignal.any([timeoutSignal, requestSignal]);
  }

  const controller = new AbortController();
  const abort = () => controller.abort();
  timeoutSignal.addEventListener("abort", abort, { once: true });
  requestSignal.addEventListener("abort", abort, { once: true });
  return controller.signal;
}

/**
 * Apply a time limit to normal API requests too. Without this wrapper a stalled
 * fetch never settles, leaving callers (and their loading spinner) permanently
 * stuck.
 */
export function fetchWithTimeout(
  input: RequestInfo | URL,
  init: RequestInit,
  timeoutMs: number,
  timeoutMessage: string,
): Promise<Response> {
  return withRequestTimeout(
    (timeoutSignal) => fetch(input, { ...init, signal: mergeAbortSignals(timeoutSignal, init.signal) }),
    timeoutMs,
    timeoutMessage,
  );
}
