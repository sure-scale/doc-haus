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
  // On a document extracted from an email's attachment: that email, and whether
  // it has since been removed from the matter.
  attachment_of?: { name: string; removed: boolean }
  // On an email: every attachment it carried and what became of it, plus the
  // images embedded in its body, which are counted rather than listed.
  attachments?: EmailAttachment[]
  embedded_images?: number
  // On an email with extracted attachments: what removing it with its
  // attachments removes, and what it keeps and why.
  removal?: EmailRemoval
  // On a source: ingest's extraction quality result. Absent on a document
  // ingested before the quality gate, which shows nothing.
  quality?: ExtractionQuality
}

// The extraction quality gate's result for a source
// (services/ingest/src/extraction-quality.ts): ok; degraded, indexed but its
// text may be incomplete or garbled; or no_readable_text, kept in the matter
// with nothing indexed. Each signal says what tripped.
export type ExtractionQuality =
  | { status: "ok" }
  | { status: "degraded" | "no_readable_text"; signals: { code: string; detail: string }[] }

// What became of one email attachment (services/ingest/src/email-attachments.ts):
// indexed: extracted as `document`; same: the same text was already in the
// matter as `document`; removed: `document` has since been removed; unsupported:
// not a format a matter document can take; failed: could not be read; nested:
// an email nested too deep to add; pending: not processed yet;
// no_readable_text: kept as `document` but not indexed, as no readable text
// could be extracted. A zip file's entries can also be: encrypted:
// password-protected, not added; unsafe_path: a path outside the zip, not
// added; too_large: a zip past the upload's zip limits, not added.
export type AttachmentStatus =
  | "indexed"
  | "same"
  | "removed"
  | "unsupported"
  | "failed"
  | "nested"
  | "pending"
  | "no_readable_text"
  | "encrypted"
  | "unsafe_path"
  | "too_large"
export type EmailAttachment = {
  // The name, type and size as the email states them; name "" when it has none.
  name: string
  type: string
  size: number
  status: AttachmentStatus
  document?: string
  note?: "may-contain-attachments"
  flagged?: true
  // Its document's extracted text may be incomplete or garbled.
  degraded?: true
}
export type EmailRemoval = {
  attachments: string[]
  kept: { name: string; reason: "linked" | "changed" }[]
}

// The source formats ingest reads (MATTER_DOCUMENT_EXTENSIONS in
// services/ingest/src/ingest.ts). The upload picker, every drop target and the
// citation link take this one list.
export const MATTER_DOCUMENT_EXTENSIONS = [".docx", ".pdf", ".eml", ".xlsx", ".zip"] as const
export type MatterDocumentKind = "docx" | "pdf" | "eml" | "xlsx" | "zip"

// Read the extension as ingest does (path.extname of the trimmed name): from the
// base name's last dot, and none when that dot leads the name (".pdf").
export function matterDocumentKind(name: string): MatterDocumentKind | undefined {
  const trimmed = name.trim()
  const base = trimmed.slice(trimmed.lastIndexOf("/") + 1)
  const dot = base.lastIndexOf(".")
  const extension = dot > 0 ? base.slice(dot).toLowerCase() : ""
  return (MATTER_DOCUMENT_EXTENSIONS as readonly string[]).includes(extension) ? (extension.slice(1) as MatterDocumentKind) : undefined
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
  // An uploaded email's attachments, or a zip file's files, and what became of each.
  attachments?: EmailAttachment[]
  // The extraction quality gate's result for the uploaded file.
  quality?: ExtractionQuality
  // The version this upload saved (#268); created is false when the file
  // already held these bytes.
  version?: { id: string; number: number; created: boolean } | null
  // A Word file saved with pending tracked changes: kept as a version, with a
  // plain notice that doc.haus edits wait until they are resolved.
  tracked_changes?: { count: number; authors: string[] }
  notice?: string
}

// What an upload does when a source of the same name is already in the matter:
// save it as that document's next version ("replace"; earlier versions stay
// kept), or keep both by storing the upload under the next free name
// ("Name (2).docx"), returned as the result's `name`. Without a choice the
// upload is refused with DocumentExistsError, so a lawyer's file is never
// replaced silently. A generated deliverable is never replaced by an upload;
// only keep both stores beside it.
export type ExistingSourceChoice = "replace" | "keep-both"

// A document of the upload's name is already in the matter: an uploaded source
// or a generated deliverable. `existing` is its stored name.
export class DocumentExistsError extends Error {
  constructor(readonly existing: string, readonly role: "source" | "deliverable", message: string) {
    super(message)
  }
}

// The upload's bytes are already kept for the document of that name, as one of
// its versions (or, from before versions were kept, as the file itself), so
// nothing was saved. `existing` is that document's name.
export class DuplicateUploadError extends Error {
  constructor(readonly existing: string, readonly version: { id: string; number: number; current: boolean } | null, message: string) {
    super(message)
  }
}

// What an identical upload matched, said beside the file's name: plainly that
// nothing was saved.
export function duplicateMessage(name: string, version: DuplicateUploadError["version"]) {
  if (!version) return `Identical to ${name}, already in this matter. Nothing was saved.`
  return `Identical to ${version.current ? `the current version (version ${version.number})` : `version ${version.number}`} of ${name}. Nothing was saved.`
}

// XHR rather than fetch so the row can tell "uploading" (bytes still leaving the
// browser) from "reading" (ingest parsing, OCR and indexing on the server).
export function uploadDocument(
  id: string,
  file: File,
  options: { onSent?: () => void; existing?: ExistingSourceChoice } = {},
): Promise<IngestResult> {
  const form = new FormData()
  form.append("file", file)
  if (options.existing) form.append("existing", options.existing)
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest()
    xhr.open("POST", `${INGEST_URL}/matters/${id}/documents`)
    xhr.timeout = UPLOAD_TIMEOUT_MS
    xhr.upload.onload = () => options.onSent?.()
    xhr.onerror = () => reject(new TypeError("Failed to fetch"))
    xhr.ontimeout = () => reject(new DOMException("Upload timed out", "TimeoutError"))
    xhr.onload = () => {
      const res = new Response(xhr.responseText, { status: xhr.status })
      const exists = documentExists(xhr.status, xhr.responseText)
      if (exists) return reject(exists)
      const duplicate = duplicateUpload(xhr.status, xhr.responseText)
      if (duplicate) return reject(duplicate)
      if (!res.ok) return ingestError(res, `Could not read ${file.name}`).then((message) => reject(new Error(message)))
      return res.json().then(resolve, reject)
    }
    xhr.send(form)
  })
}

function documentExists(status: number, text: string) {
  if (status !== 409) return
  try {
    const body = JSON.parse(text) as { code?: unknown; name?: unknown; error?: unknown }
    const role = body.code === "source_exists" ? "source" : body.code === "deliverable_exists" ? "deliverable" : undefined
    if (!role || typeof body.name !== "string") return
    return new DocumentExistsError(body.name, role, typeof body.error === "string" ? body.error : `${body.name} is already in this matter.`)
  } catch {
    return
  }
}

function duplicateUpload(status: number, text: string) {
  if (status !== 409) return
  try {
    const body = JSON.parse(text) as { code?: unknown; name?: unknown; error?: unknown; version?: { id?: unknown; number?: unknown; current?: unknown } | null }
    if (body.code !== "duplicate" || typeof body.name !== "string") return
    const kept = body.version
    const version = kept && typeof kept.id === "string" && typeof kept.number === "number"
      ? { id: kept.id, number: kept.number, current: kept.current === true }
      : null
    return new DuplicateUploadError(body.name, version, typeof body.error === "string" ? body.error : `${body.name} is already in this matter.`)
  } catch {
    return
  }
}

// Removing an email also removes the attachments extracted from it (see
// Document.removal) unless `keepAttachments`.
export async function deleteDocument(
  id: string,
  name: string,
  options: { keepAttachments?: boolean } = {},
): Promise<{
  attachments?: { removed: string[]; kept: EmailRemoval["kept"]; failed?: { name: string; reason: string }[] }
}> {
  const keep = options.keepAttachments ? "&attachments=keep" : ""
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents?name=${encodeURIComponent(name)}${keep}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, `Could not delete ${name}`))
  return res.json()
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

// The normalised text ingest extracted from a matter document: the same text the
// assistant reads and cites. Emails and spreadsheets are viewed as this text.
export async function fetchDocumentText(id: string, name: string): Promise<string> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/text?name=${encodeURIComponent(name)}`)
  if (!res.ok) throw new Error(await ingestError(res, `Could not read ${name}`))
  return ((await res.json()) as { text: string }).text
}

export async function setDocumentLabel(id: string, name: string, label: DocumentLabel): Promise<void> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/label`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, label }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not save the document label."))
}
