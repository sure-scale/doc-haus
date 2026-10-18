import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

// Reading a long PDF (OCR + embeddings) can take minutes; past this the upload
// row offers a retry instead of spinning forever.
const UPLOAD_TIMEOUT_MS = 5 * 60_000

// `role` separates the lawyer's own files ("source") from files a drafting run
// produced ("deliverable"); only sources are searched and cited.
export type DocumentLabel = "draft" | "redline" | "notes" | "final"

export type Document = {
  id: number
  name: string
  doc_path: string
  created_at: number
  role?: "source" | "deliverable"
  pending?: number
  label?: DocumentLabel | null
}

// A prompt-injection finding from ingest-time scanning: invisible Unicode, hidden
// DOCX formatting, or instruction-like text (see services/ingest/src/sanitize.ts).
export type InjectionFinding = { rule: string; detail: string; snippet?: string; charStart?: number; charEnd?: number }
export type IngestResult = {
  name: string
  docPath: string
  sections: number
  chunks: number
  // null when the document came up clean.
  injection: { findings: InjectionFinding[]; flaggedChunks: number } | null
}

// XHR rather than fetch so the row can tell "uploading" (bytes still leaving the
// browser) from "reading" (ingest parsing, OCR and indexing on the server).
export function uploadDocument(id: string, file: File, onSent?: () => void): Promise<IngestResult> {
  const form = new FormData()
  form.append("file", file)
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest()
    xhr.open("POST", `${INGEST_URL}/matters/${id}/documents`)
    xhr.timeout = UPLOAD_TIMEOUT_MS
    xhr.upload.onload = () => onSent?.()
    xhr.onerror = () => reject(new TypeError("Failed to fetch"))
    xhr.ontimeout = () => reject(new DOMException("Upload timed out", "TimeoutError"))
    xhr.onload = () => {
      const res = new Response(xhr.responseText, { status: xhr.status })
      if (!res.ok) return ingestError(res, `Could not read ${file.name}`).then((message) => reject(new Error(message)))
      return res.json().then(resolve, reject)
    }
    xhr.send(form)
  })
}

export async function deleteDocument(id: string, name: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents?name=${encodeURIComponent(name)}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, `Could not delete ${name}`))
}

// Convert a matter's uploaded .pdf into an editable .docx sibling and index it, so
// the PDF can enter the DOCX redline pipeline. Returns the new .docx document.
export async function convertPdfToDocx(id: string, name: string): Promise<IngestResult> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/convert?name=${encodeURIComponent(name)}`, {
    method: "POST",
  })
  if (!res.ok) throw new Error(`Could not convert ${name} (${res.status})`)
  return res.json()
}

// URL the PDF viewer points an <iframe> at — the browser renders the PDF natively
// from our own ingest service, so the file never leaves for a third-party service.
export function documentContentUrl(id: string, name: string): string {
  return `${INGEST_URL}/matters/${id}/documents/content?name=${encodeURIComponent(name)}`
}

// A combined document's one file (its requested memo, then the draft), which
// ingest composes from the registered draft and memo on each download.
export function combinedDocumentUrl(id: string): string {
  return `${INGEST_URL}/matters/${id}/documents/combined`
}

// The combined route's refusals, in plain words. A busy part is reported once,
// never retried here.
const SEPARATELY = "Download the draft and the memo separately."
const COMBINED_REFUSALS = new Map([
  ["no_combined_output", "this drafting run has no combined file."],
  ["combined_parts_missing", "the draft or the memo it is built from is no longer in the matter."],
  ["empty_memo", "the memo has no text to place before the draft."],
  ["unsupported_memo_dependency", `the memo holds content the combined file cannot carry. ${SEPARATELY}`],
  ["unsupported_memo_numbering", `the memo uses list numbering the combined file cannot carry. ${SEPARATELY}`],
  ["unsupported_memo_namespace", `the memo uses Word formatting the combined file cannot carry. ${SEPARATELY}`],
  ["combined_parts_busy", "the draft or the memo is being updated. Try again shortly."],
])
// invalid_docx_package / invalid_docx_xml name the part ingest could not read.
const UNREADABLE_PART = new Map([["body", "the draft"], ["memo", "the memo"]])

/** The combined file's contents; a refusal rejects with its reason. */
export async function fetchCombinedDocument(id: string, name: string): Promise<Blob> {
  const res = await fetch(combinedDocumentUrl(id))
  if (!res.ok) {
    const body = (await res.json().catch(() => undefined)) as { error?: unknown; part?: unknown } | undefined
    const error = typeof body?.error === "string" ? body.error : `Could not download ${name} (${res.status})`
    const part = typeof body?.part === "string" ? UNREADABLE_PART.get(body.part) : undefined
    const reason = ["invalid_docx_package", "invalid_docx_xml"].includes(error)
      ? `${part ?? "the draft or the memo"} could not be read as a Word document. Open it in Word, save it again, and try again.`
      : COMBINED_REFUSALS.get(error) ?? (error.startsWith("unsupported_memo_") ? COMBINED_REFUSALS.get("unsupported_memo_dependency") : undefined)
    throw new Error(reason ? `Could not download ${name}: ${reason}` : error)
  }
  return res.blob()
}

export async function setDocumentLabel(id: string, name: string, label: DocumentLabel): Promise<void> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/label`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, label }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not save the document label."))
}
