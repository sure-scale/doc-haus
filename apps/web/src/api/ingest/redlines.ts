import { INGEST_URL } from "../../config"

// A pending redline proposal: one tracked change awaiting accept/reject.
export type Redline = {
  id: number
  doc_name: string
  scope: "phrase" | "clause"
  find_text: string
  old_text: string
  new_text: string
  author: string
  review_comment: string
  created_at: number
  stale?: boolean
  stale_reason?: string
}

// The redlined view: the clean .docx compared against itself with every pending
// proposal applied, so the bytes carry native tracked changes the viewer paints
// green/red. Identical to the clean bytes when nothing is pending.
export async function fetchRedlinedBytes(id: string, name: string): Promise<Uint8Array> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/redlined?name=${encodeURIComponent(name)}`)
  if (!res.ok) {
    const detail = await res.json().catch(() => undefined)
    throw new Error(typeof detail?.error === "string" ? detail.error : `Could not load ${name} (${res.status})`)
  }
  return new Uint8Array(await res.arrayBuffer())
}

export async function fetchRedlines(id: string, name: string): Promise<Redline[]> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/redlines?name=${encodeURIComponent(name)}`)
  if (!res.ok) throw new Error(`Could not load changes for ${name} (${res.status})`)
  return res.json()
}

export async function prepareDocumentChanges(id:string,name:string,instructions:string):Promise<{proposed:number;open_items:string[]}> {
  const response = await fetch(`${INGEST_URL}/matters/${id}/documents/amendments`,{method:"POST",
    headers:{"content-type":"application/json"},body:JSON.stringify({name,instructions})})
  const result = await response.json()
  if (!response.ok) throw new Error(typeof result.error === "string" ? result.error : "Could not prepare changes")
  return result
}

// Present only when the accepted file is the body of a completed drafting run:
// whether that run's drafting notes recorded the accepted change, and if not, why.
export type AcceptResult = { notes_updated?: boolean; notes_skip_reason?: string }

export async function acceptRedline(id: string, redlineId: number): Promise<AcceptResult> {
  return resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/${redlineId}/accept`)
}

export async function rejectRedline(id: string, redlineId: number): Promise<void> {
  await resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/${redlineId}/reject`)
}

export async function acceptAllRedlines(id: string, name: string): Promise<AcceptResult> {
  return resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/accept-all?name=${encodeURIComponent(name)}`)
}

export async function rejectAllRedlines(id: string, name: string): Promise<void> {
  await resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/reject-all?name=${encodeURIComponent(name)}`)
}

async function resolveRedlines(url: string): Promise<AcceptResult> {
  const response = await fetch(url, { method: "POST" })
  if (response.ok) return response.json().catch(() => ({}))
  const detail: unknown = await response.json().catch(() => undefined)
  const message = detail && typeof detail === "object" && "error" in detail && typeof detail.error === "string"
    ? detail.error : `Could not resolve tracked changes (${response.status})`
  throw new Error(message)
}
