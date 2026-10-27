import type { WordOpenDocument } from "../../../../../dochaus/lib/word-addin-contract"
import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

export type { WordOpenDocument }
export type WordConfig = { available: boolean; origin: string | null; manifestUrl: string | null; reason?: string }
export type WordSetupStatus = WordConfig & {
  platform: "macos" | "windows" | "unsupported"
  wordInstalled: boolean
  canSetup: boolean
  configured: boolean
  registered: boolean
  trusted: boolean
  running: boolean
  status: "ready" | "not_configured" | "needs_repair" | "unsupported" | "word_missing" | "external"
  activatedWithoutRestart?: boolean
  restartWord?: boolean
}

const wordUrl = (matterId: string, name?: string) => `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/word${name === undefined ? "" : `?name=${encodeURIComponent(name)}`}`

export async function fetchWordConfig(): Promise<WordConfig> {
  const response = await fetch(`${INGEST_URL}/word/config`)
  if (!response.ok) throw new Error(await ingestError(response, "Could not check Word setup"))
  return response.json()
}

export async function fetchWordSetup(): Promise<WordSetupStatus> {
  const response = await fetch(`${INGEST_URL}/word/setup`)
  if (!response.ok) throw new Error(await ingestError(response, "Could not check Word setup"))
  return response.json()
}

export async function setupWordAddin(): Promise<WordSetupStatus> {
  // OS approval can take longer than an ordinary API call. A failed response
  // is checked with GET rather than automatically repeating native setup.
  const response = await fetch(`${INGEST_URL}/word/setup`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: "{}",
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not set up Word"))
  return response.json()
}

export async function fetchWordDocuments(matterId: string): Promise<WordOpenDocument[]> {
  const response = await fetch(wordUrl(matterId))
  if (!response.ok) throw new Error(await ingestError(response, "Could not check which documents are open in Word"))
  return ((await response.json()) as { documents: WordOpenDocument[] }).documents
}

// The normal app requests a linked copy; its write-only capability stays in
// that copy. The browser never reads or stores the capability.
export async function linkedWordDocument(matterId: string, name: string, origin: string): Promise<Blob> {
  const response = await fetch(wordUrl(matterId, name), {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ origin }),
  })
  if (!response.ok) throw new Error(await ingestError(response, `Could not open ${name} in Word`))
  return response.blob()
}

export type WordOpenState = "pending" | "redeemed" | "completed" | "failed" | "expired"

// A short-lived read ticket reaches the installed desktop helper. It is not
// the document's save capability and must never be sent to an arbitrary URI.
export async function prepareWordOpen(matterId: string, name: string, origin: string, wordOrigin: string, signal?: AbortSignal): Promise<{ uri: string; id: string }> {
  const response = await fetch(`${wordUrl(matterId)}/open?name=${encodeURIComponent(name)}`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ origin }), signal,
  })
  if (!response.ok) throw new Error(await ingestError(response, `Could not open ${name} in Word`))
  const value: unknown = await response.json()
  const invalid = () => new Error("Could not prepare the desktop Word connection. Try again or download a Word copy.")
  if (!value || typeof value !== "object" || !("uri" in value) || !("id" in value) || typeof value.uri !== "string" || typeof value.id !== "string" || !/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(value.id)) throw invalid()
  try {
    const uri = new URL(value.uri), source = new URL(uri.searchParams.get("source") || "")
    if (uri.protocol !== "doc.haus:" || uri.host !== "word" || uri.pathname !== "/open" || uri.username || uri.password || uri.hash || [...uri.searchParams].length !== 1 || source.protocol !== "https:" || source.origin !== wordOrigin || source.username || source.password || source.search || source.hash || !/^\/word\/open\/[A-Za-z0-9_-]{43}$/.test(source.pathname)) throw invalid()
  } catch { throw invalid() }
  return { uri: value.uri, id: value.id }
}

export async function fetchWordOpenState(matterId: string, id: string, signal?: AbortSignal): Promise<WordOpenState> {
  const response = await fetch(`${wordUrl(matterId)}/open/status?id=${encodeURIComponent(id)}`, { signal })
  if (!response.ok) throw new Error("Could not confirm the Word launch. If Word did not open, try again or download a Word copy.")
  const value: unknown = await response.json()
  if (!value || typeof value !== "object" || !("state" in value) || !["pending", "redeemed", "completed", "failed", "expired"].includes(String(value.state))) throw new Error("Could not confirm the Word launch.")
  return value.state as WordOpenState
}

export async function revokeWordDocument(matterId: string, name: string): Promise<void> {
  const response = await fetch(wordUrl(matterId, name), { method: "DELETE" })
  if (!response.ok) throw new Error(await ingestError(response, `Could not disconnect ${name} from Word`))
}

export async function recoverWordEditing(matterId: string, document: WordOpenDocument, action: "release" | "takeover"): Promise<void> {
  const response = await fetch(`${wordUrl(matterId)}/recover?name=${encodeURIComponent(document.name)}`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action, expectedCurrent: document.currentVersion, expectedGeneration: document.lease?.generation ?? null }),
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not change Word editing ownership"))
}
