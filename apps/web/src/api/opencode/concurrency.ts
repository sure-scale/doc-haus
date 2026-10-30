import { OPENCODE_URL } from "../../config"

export type ConcurrencyTestResult = {
  model: string
  requested: number
  succeeded: number
  failed: number
  peak: number
  elapsed_ms: number
  errors: { kind: string; count: number }[]
}

/** Run only after the user requests a test; use the installation's saved model. */
export async function testConcurrency(count: number): Promise<ConcurrencyTestResult> {
  const response = await fetch(`${OPENCODE_URL}/api/dochaus/concurrency-test`, {
    method: "POST",
    credentials: "include",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ count }),
    signal: AbortSignal.timeout(180_000),
  })
  if (!response.ok) {
    const body: unknown = await response.json().catch(() => null)
    const value = body && typeof body === "object" && "error" in body && typeof body.error === "string" ? body.error.trim() : ""
    const safe = value.length > 0 && value.length <= 400 && !/[\r\n\x00-\x1f{}<>]/.test(value)
      && !/(?:bearer\s+\S+|(?:api[ _-]?key|token|authorization|password|secret)\s*[:=]\s*\S+|https?:\/\/|-----BEGIN|\bsk-[A-Za-z0-9_-]{10,})/i.test(value)
    throw new Error(safe ? value : `The test could not start (${response.status}). Check your saved AI setup and try again.`)
  }
  return await response.json() as ConcurrencyTestResult
}
