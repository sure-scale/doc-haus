import { INGEST_URL } from "../../config"
export type SetupStep = "models" | "details" | "research"
export type SetupProgress = { version: 1; completed: SetupStep[]; skipped: SetupStep[]; next: SetupStep | "done"; ready: boolean }
export async function getSetup(): Promise<SetupProgress> {
  const res = await fetch(`${INGEST_URL}/setup`, { signal: AbortSignal.timeout(12000) })
  if (!res.ok) throw new Error("Could not check setup. Make sure the engine is running, then retry.")
  return res.json()
}
export async function saveSetup(step: SetupStep, status: "completed" | "skipped"): Promise<SetupProgress> {
  const res = await fetch(`${INGEST_URL}/setup`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ step, status }), signal: AbortSignal.timeout(12000) })
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.error ?? "Could not save setup progress. Retry.")
  }
  return res.json()
}
