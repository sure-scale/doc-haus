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
export async function fetchRedlinedBytes(id: string, name: string, notice?: (message: string | undefined) => void): Promise<Uint8Array> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/redlined?name=${encodeURIComponent(name)}`)
  if (!res.ok) {
    const detail = await res.json().catch(() => undefined)
    throw new Error(typeof detail?.error === "string" ? detail.error : `Could not load ${name} (${res.status})`)
  }
  const encoded = res.headers.get("X-DocHaus-Revision-Notice")
  notice?.(encoded ? decodeURIComponent(encoded) : undefined)
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

// The latest Prepare's batch and open items for one document; null before any.
export type PreparedChanges = { redline_ids: number[]; open_items: string[] }
export async function fetchPreparedChanges(id: string, name: string): Promise<PreparedChanges | null> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/amendments?name=${encodeURIComponent(name)}`)
  if (!res.ok) throw new Error(`Could not load the open items for ${name} (${res.status})`)
  return res.json()
}

// Present only when the accepted file is the body of a completed drafting run:
// whether that run's drafting notes recorded the accepted change and its
// delivered redline was rebuilt from the accepted body, and if not, why. The
// combined outcome is present only when the accept changed a part of a saved
// combined file: whether that file was rebuilt, and if not, why. A reject that
// changed a part of a saved combined file reports the combined outcome too.
export type AcceptResult = {
  notes_updated?: boolean
  notes_skip_reason?: string
  redline_updated?: boolean
  redline_skip_reason?: string
  combined_updated?: boolean
  combined_skip_reason?: string
}

export async function acceptRedline(id: string, redlineId: number): Promise<AcceptResult> {
  return acceptResult(await resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/${redlineId}/accept`))
}

export async function rejectRedline(id: string, redlineId: number): Promise<AcceptResult> {
  return acceptResult(await resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/${redlineId}/reject`))
}

export async function acceptAllRedlines(id: string, name: string): Promise<AcceptResult> {
  return acceptResult(await resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/accept-all?name=${encodeURIComponent(name)}`))
}

function acceptResult(body: unknown): AcceptResult {
  if (!body || typeof body !== "object") return {}
  const outcome = (updated: "notes_updated" | "redline_updated" | "combined_updated", skipped: "notes_skip_reason" | "redline_skip_reason" | "combined_skip_reason") => {
    const value = (body as Record<string, unknown>)[updated]
    if (typeof value !== "boolean") return {}
    const reason = (body as Record<string, unknown>)[skipped]
    return { [updated]: value, ...(typeof reason === "string" ? { [skipped]: reason } : {}) }
  }
  return {
    ...outcome("notes_updated", "notes_skip_reason"),
    ...outcome("redline_updated", "redline_skip_reason"),
    ...outcome("combined_updated", "combined_skip_reason"),
  }
}

export async function rejectAllRedlines(id: string, name: string): Promise<AcceptResult> {
  return acceptResult(await resolveRedlines(`${INGEST_URL}/matters/${id}/redlines/reject-all?name=${encodeURIComponent(name)}`))
}

async function resolveRedlines(url: string): Promise<unknown> {
  const response = await fetch(url, { method: "POST" })
  if (response.ok) return response.json().catch(() => undefined)
  const detail: unknown = await response.json().catch(() => undefined)
  const message = detail && typeof detail === "object" && "error" in detail && typeof detail.error === "string"
    ? detail.error : `Could not resolve tracked changes (${response.status})`
  throw new Error(message)
}
