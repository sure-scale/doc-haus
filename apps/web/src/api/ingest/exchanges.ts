import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import type { LinkedDocumentBranch } from "./versions"
import type { DocumentReturnTieBreak } from "../../../../../dochaus/lib/document-return-tie-break-contract"

export type DocumentExchange = {
  id: number; kind: "sent" | "received"; documentId: string; versionId: string
  party: string | null; actor: string; occurredAt: number; recordedAt: number
  baseVersionId: string | null; requestId: string; draftNumber?: number
  payload: { sha256: string; ext: string; bytes: number } | null
  source?: { documentName?: string; attachmentId?: string; sender?: string; date?: string }
  match?: { signal: string; reason: string }
  baseCorrections?: { requestId: string; fromVersionId: string | null; toVersionId: string; actor: string; recordedAt: number }[]
  association?: { documentId: string; baseVersionId: string }
  linkedBranch?: LinkedDocumentBranch
}
export type ReturnCandidate = {
  documentId: string; documentName: string; versionId: string; versionNumber: number; sha256: string
  current: boolean; sent: boolean; signals: string[]; score: number; reason: string
}
export type ReturnInspection = {
  kind: "duplicate" | "matched" | "needs-confirmation" | "new-document"
  sha256: string; candidates: ReturnCandidate[]; selected?: ReturnCandidate
  reason: string; warnings: string[]; complete: boolean
  source?: DocumentExchange["source"]
}
export async function fetchDocumentExchanges(matter: string, name: string): Promise<DocumentExchange[]> {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/exchanges?name=${encodeURIComponent(name)}`)
  if (!response.ok) throw new Error(await ingestError(response, "Could not load send and receive history."))
  return (await response.json() as { exchanges: DocumentExchange[] }).exchanges
}
export async function downloadForSending(matter: string, version: string, input: { requestId: string; party?: string; baseVersionId?: string; marker: boolean }) {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/versions/${encodeURIComponent(version)}/send`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input),
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not record this send."))
  return { bytes: new Uint8Array(await response.arrayBuffer()), draftNumber: Number(response.headers.get("X-DocHaus-Draft-Number")) }
}
export function exchangePayloadUrl(matter: string, exchange: number) {
  return `${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/exchanges/${exchange}/content`
}
export function linkedReturnContentUrl(matter: string, exchange: number) {
  return `${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/exchanges/${exchange}/branch/content`
}
export async function correctReceivedDocument(matter: string, exchange: number, input: { targetDocumentId: string; baseVersionId: string; expectedDocumentId: string; expectedBaseVersionId: string | null; requestId: string }) {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/exchanges/${exchange}/document`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input),
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not correct the returned document."))
  return response.json() as Promise<{ exchange: DocumentExchange; changed: boolean }>
}
export async function correctReceivedBase(matter: string, exchange: number, input: { baseVersionId: string; expectedBaseVersionId: string | null; requestId: string }) {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/exchanges/${exchange}/base`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input),
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not correct the returned base."))
  return response.json() as Promise<{ exchange: DocumentExchange; changed: boolean }>
}
export async function inspectReturnedFile(matter: string, file: File, emailName?: string): Promise<ReturnInspection> {
  const form = new FormData(); form.append("file", file)
  if (emailName) form.append("emailName", emailName)
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/returns/inspect`, { method: "POST", body: form })
  if (!response.ok) throw new Error(await ingestError(response, "Could not inspect the returned file."))
  return response.json()
}
export async function identifyReturnedFile(matter: string, file: File, emailName?: string, signal?: AbortSignal): Promise<{ inspection: ReturnInspection; advisory: DocumentReturnTieBreak; source?: DocumentExchange["source"] }> {
  const form = new FormData(); form.append("file", file)
  if (emailName) form.append("emailName", emailName)
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/returns/tie-break`, { method: "POST", body: form, signal })
  if (!response.ok) throw new Error(await ingestError(response, "Could not help identify this return."))
  return response.json()
}
export async function saveReturnedFile(matter: string, file: File, input: { baseVersionId: string; priorSha256: string; requestId: string; party?: string; emailName?: string }) {
  const form = new FormData(); form.append("file", file)
  for (const [name, value] of Object.entries(input)) if (value !== undefined) form.append(name, value)
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matter)}/documents/returns`, { method: "POST", body: form })
  if (!response.ok) throw new Error(await ingestError(response, "Could not save the returned version."))
  return response.json() as Promise<{ name: string; version: { id: string; number: number }; comparison: { leftVersion: string; rightVersion: string; left?: string; right?: string } }>
}
