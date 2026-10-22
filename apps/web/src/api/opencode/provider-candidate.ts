import type { FormAnswer } from "@opencode/client"
import { INGEST_URL, OPENCODE_URL } from "../../config"

export type CandidateConnection =
  | { type: "key"; key: string; answer?: FormAnswer }
  | { type: "credential"; id: string }
  | { type: "env"; name: string }
  | { type: "host" }
  | { type: "reuse"; sourceId: string; connectionId: string }
  | { type: "new" }

export type CandidateInput = {
  providerID: string
  options?: Record<string, string | null>
  connection: CandidateConnection
  local?: { name: string; baseURL: string; modelIDs: string[] }
}
export type CandidateModel = { id: string; providerID: string; name: string; capabilities?: { input: string[]; output: string[]; tools: boolean }; variants?: { id: string; name?: string }[] }
export type CandidateTest = { ok: true } | {
  ok: false; kind: "auth" | "model" | "unavailable"; expired?: boolean; restriction?: "free-tier" | "region"; error: string
}
export type CandidateAuthAction = "oauth.connect" | "oauth.status" | "oauth.complete" | "oauth.cancel" | "command.connect" | "command.status" | "command.cancel"
export type CandidateAuthResult = { attemptID?: string; mode?: "auto" | "code"; url?: string; instructions?: string; message?: string; status?: string }

async function candidateRequest<T>(action: string, body: unknown, timeout = 30_000): Promise<T> {
  const response = await fetch(`${OPENCODE_URL}/api/dochaus/provider-candidate/${action}`, {
    method: "POST", credentials: "include", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body), signal: AbortSignal.timeout(timeout),
  })
  const result = await response.json().catch(() => undefined)
  if (!response.ok) throw new Error(typeof result?.error === "string" ? result.error : "Could not prepare this connection. Check the details and retry.")
  return result as T
}

export const createCandidate = (input: CandidateInput) => candidateRequest<{ sessionID: string }>("session", input)
export const candidateModels = (sessionID: string, modelIDs?: string[]) => candidateRequest<{ data: CandidateModel[]; customModels: boolean }>("models", { sessionID, ...(modelIDs ? { modelIDs } : {}) })
export const candidateAuth = (sessionID: string, action: CandidateAuthAction, input: { methodID?: string; attemptID?: string; answer?: FormAnswer; code?: string } = {}) => candidateRequest<{ data: CandidateAuthResult }>("auth", { sessionID, action, ...input })

// Model failures are expected results and retain their typed recovery message.
export async function testCandidate(sessionID: string, modelID: string): Promise<CandidateTest> {
  try {
    const response = await fetch(`${OPENCODE_URL}/api/dochaus/provider-candidate/test`, {
      method: "POST", credentials: "include", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionID, modelID }), signal: AbortSignal.timeout(30_000),
    })
    const result = await response.json()
    if (response.ok && result.ok === true) return { ok: true }
    if (["auth", "model", "unavailable"].includes(result.kind)) return result as CandidateTest
  } catch { /* An unreachable engine does not prove rejected credentials. */ }
  return { ok: false, kind: "unavailable", error: "The connection test could not complete. Check availability and try again." }
}

export async function discardCandidate(sessionID: string) {
  await fetch(`${OPENCODE_URL}/api/dochaus/provider-candidate/session?sessionID=${encodeURIComponent(sessionID)}`, {
    method: "DELETE", credentials: "include", signal: AbortSignal.timeout(10_000),
  }).catch(() => {})
}

export class SaveUnconfirmedError extends Error {}

export async function saveCandidate(sessionID: string, primary: string, fast: string) {
  let response: Response
  try {
    response = await fetch(`${INGEST_URL}/engine/provider-setup`, {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ sessionID, primary, fast }),
      signal: AbortSignal.timeout(50_000),
    })
  } catch {
    throw new SaveUnconfirmedError("Could not confirm whether AI setup saved. Reload the saved setup and check its connection before continuing.")
  }
  const result = await response.json().catch(() => undefined)
  if (!response.ok) {
    const detail = typeof result?.error === "string" ? result.error : "Could not save AI setup. Check the saved setup before continuing."
    if (result?.partialState || !result?.code) throw new SaveUnconfirmedError(detail)
    throw new Error(detail)
  }
  if (result?.ok !== true) throw new SaveUnconfirmedError("Could not confirm whether AI setup saved. Reload the saved setup before continuing.")
  return result as { ok: true; sessionClosePending?: boolean }
}
