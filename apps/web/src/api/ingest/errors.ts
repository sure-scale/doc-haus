export async function ingestError(res: Response, fallback: string) {
  const body = await res.json().catch(() => undefined) as { error?: unknown } | undefined
  return typeof body?.error === "string" ? body.error : `${fallback} (${res.status})`
}

export const UNREACHABLE = "Can't reach the document service right now. Check that doc.haus is running, then try again."
export const TIMED_OUT = "The document service took too long to respond. Try again in a moment."

// Turn a thrown value into copy a lawyer can act on. Browsers report a refused
// connection as a bare TypeError ("Failed to fetch", "Load failed", "NetworkError
// when attempting to fetch resource"), which says nothing useful on its own.
export function humanizeError(error: unknown, fallback = "Something went wrong. Try again.") {
  if (typeof error === "string") return error.trim() || fallback
  if (!(error instanceof Error)) return fallback
  if (error.name === "AbortError" || error.name === "TimeoutError") return TIMED_OUT
  if (/failed to fetch|load failed|networkerror|network request failed|fetch failed/i.test(error.message)) return UNREACHABLE
  return error.message.trim() || fallback
}
