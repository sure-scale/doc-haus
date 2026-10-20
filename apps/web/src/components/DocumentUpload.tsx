import { useEffect, useRef, useState } from "react"
import { deleteDocument, DocumentExistsError, uploadDocument, type Document, type ExistingSourceChoice } from "../api/ingest/documents"
import { humanizeError } from "../api/ingest/errors"
import type { LeanDraftingStatus } from "../api/ingest/lean"
import { formatDate, formatDateTime, formatTime } from "../lib/format"
import { CompareIcon } from "./CompareDialog"
import DocumentLabelBadge from "./DocumentLabel"
import DraftedList from "./drafting/DraftedList"
import { deliverableGroups, type DraftedOutput } from "./drafting/model"
import Dropzone from "./Dropzone"
import { useToast } from "./Toast"

// One row per file the lawyer added in this session, so a slow or failed upload
// is visible next to its name instead of hiding behind a single spinner.
type Upload = {
  key: number
  file: File
  status: "queued" | "uploading" | "reading" | "ready" | "failed" | "exists"
  error?: string
  // A wrong file type cannot succeed on retry, so the row offers Remove only.
  retryable?: boolean
  // A document of this name is already in the matter. The row asks the lawyer to
  // replace an uploaded source or keep both, or to keep both beside a generated
  // deliverable, and the choice is sent on retry.
  existing?: { name: string; role: "source" | "deliverable" }
  choice?: ExistingSourceChoice
}

const ACCEPTED = /\.(docx|pdf)$/i
const READY_LINGER_MS = 3000

// A newest-first document timeline with grouped deliverables and a persistent
// upload footer. The scrolling timeline also accepts dropped files.
export default function DocumentUpload({
  matterId,
  documents,
  onUploaded,
  onView,
  drafting,
  onClose,
  incoming,
  onIncoming,
  onCompare,
}: {
  matterId: string
  documents: Document[]
  onUploaded: () => void
  onView: (name: string) => void
  drafting?: LeanDraftingStatus
  onClose?: () => void
  // Files added elsewhere (dropped on the composer) to upload through this list.
  incoming?: File[]
  onIncoming?: () => void
  // Compare a Word file with another version of it.
  onCompare?: (name: string) => void
}) {
  const [uploads, setUploads] = useState<Upload[]>([])
  const [removing, setRemoving] = useState<string>()
  const [confirmId, setConfirmId] = useState<number | null>(null)
  const nextKey = useRef(0)
  const watched = useRef(false)
  // Uploads run one at a time so a batch does not swamp the reader.
  const queue = useRef(Promise.resolve())
  const toast = useToast()
  // StrictMode re-runs effects on mount; the ref keeps one batch from uploading twice.
  const taken = useRef<File[]>(undefined)

  useEffect(() => {
    if (!incoming?.length || taken.current === incoming) return
    taken.current = incoming
    onFiles(incoming)
    onIncoming?.()
  }, [incoming])

  useEffect(() => {
    if (!confirmId) return
    const dismiss = () => setConfirmId(null)
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmId])

  function update(key: number, patch: Partial<Upload>) {
    setUploads((prev) => prev.map((u) => (u.key === key ? { ...u, ...patch } : u)))
  }

  function enqueue(upload: Upload) {
    queue.current = queue.current.then(() => run(upload))
  }

  async function run(upload: Upload) {
    update(upload.key, { status: "uploading", error: undefined })
    try {
      const result = await uploadDocument(matterId, upload.file, {
        onSent: () => update(upload.key, { status: "reading" }),
        existing: upload.choice,
      })
      update(upload.key, { status: "ready" })
      setTimeout(() => setUploads((prev) => prev.filter((u) => u.key !== upload.key)), READY_LINGER_MS)
      onUploaded()
      // A flagged upload is a finding the lawyer needs to see, not a failure: the
      // document is indexed, the assistant treats the flagged text as untrusted,
      // and the warning says what ingest found.
      if (result.injection) {
        toast(
          "warning",
          `${result.name} is ready, with a warning — possible prompt injection: ${[
            ...new Set(result.injection.findings.map((f) => f.detail)),
          ].join("; ")}. The assistant will treat this content as untrusted.`,
        )
        return
      }
      toast("success", `${result.name} is ready (${result.sections} section${result.sections === 1 ? "" : "s"}).`)
    } catch (e) {
      update(upload.key, failedUpload(e, upload.file))
    }
  }

  function onFiles(files: File[]) {
    const added = files.map((file): Upload => {
      const key = nextKey.current++
      if (ACCEPTED.test(file.name)) return { key, file, status: "queued" }
      return { key, file, status: "failed", error: "Only Word (.docx) and PDF files can be added." }
    })
    setUploads((prev) => [...prev, ...added])
    added.filter((u) => u.status === "queued").forEach(enqueue)
  }

  async function onRemove(name: string) {
    setConfirmId(null)
    setRemoving(name)
    try {
      await deleteDocument(matterId, name)
      onUploaded()
      toast("success", `Removed ${name}.`)
    } catch (e) {
      toast("error", humanizeError(e, `Could not remove ${name}.`))
    } finally {
      setRemoving(undefined)
    }
  }

  if (drafting && drafting.status !== "completed") watched.current = true
  const sources = documents.filter((d) => d.role !== "deliverable")
  const generated = documents.filter((d) => d.role === "deliverable")
  const drafted = deliverableGroups(
    generated.map((d) => d.name),
    drafting,
  )
  const pending = documents.reduce((n, d) => n + (d.pending ?? 0), 0)
  const draftedPending = new Map(generated.filter((d) => d.pending).map((d) => [d.name, d.pending!]))
  const working = uploads.filter((u) => u.status === "queued" || u.status === "uploading" || u.status === "reading").length

  const timeline = [
    ...drafted.flatMap((group): DraftedOutput[] => [group.main, ...group.extras, ...(group.combined ? [{ name: group.combined, role: "combined" as const }] : [])]).map((output) => {
      const document = generated.find((document) => document.name === output.name)!
      return {
        key: `draft:${output.name}`,
        createdAt: document.created_at,
        content: <DraftedList matterId={matterId} groups={[{ main: output, extras: [] }]} onView={onView} onCompare={onCompare} pending={draftedPending} fresh={watched.current}
          renderLabel={() => <>
            <time className="document-time" dateTime={new Date(document.created_at).toISOString()} title={`Saved to matter: ${formatDateTime(document.created_at)}`}>
              {formatTime(document.created_at)}
            </time>
            <DocumentLabelBadge matterId={matterId} name={output.name}
              label={document.label ?? output.role} onChanged={onUploaded} />
          </>} />,
      }
    }),
    ...sources.map((d) => ({
      key: `source:${d.id}`,
      createdAt: d.created_at,
      content: (
      <div className="doc-row is-uploaded">
        <div className="doc-row-main">
          <span className={`doc-type is-${d.name.toLowerCase().endsWith(".pdf") ? "pdf" : "docx"}`} aria-hidden>
            <IconFile />
          </span>
          <button className="linklike" onClick={() => onView(d.name)}>
            {d.name}
          </button>
          {d.pending ? (
            <span className="redline-badge" title={`${d.pending} pending change${d.pending === 1 ? "" : "s"} to review`}>
              {d.pending}
            </span>
          ) : null}
        </div>
        <div className="doc-row-meta">
          <span className="document-origin">Uploaded</span>
          <time className="document-time" dateTime={new Date(d.created_at).toISOString()} title={`Added to matter: ${formatDateTime(d.created_at)}`}>
            {formatTime(d.created_at)}
          </time>
          {onCompare && /\.docx$/i.test(d.name) && (
            <button className="doc-compare" onClick={() => onCompare(d.name)} title="Compare with another version" aria-label={`Compare ${d.name} with another version`}>
              <CompareIcon />
            </button>
          )}
          {confirmId === d.id ? (
            <button
              className="icon-btn danger"
              title={`Remove ${d.name} — indexed text is deleted and answers can no longer cite it`}
              onClick={(e) => {
                e.stopPropagation()
                onRemove(d.name)
              }}
            >
              Confirm
            </button>
          ) : (
            <button
              className="doc-remove"
              onClick={(e) => {
                e.stopPropagation()
                setConfirmId(d.id)
              }}
              title={`Remove ${d.name}`}
              aria-label={`Remove ${d.name}`}
              disabled={removing === d.name}
            >
              <IconTrash />
            </button>
          )}
        </div>
      </div>
      ),
    })),
  ].sort((a, b) => b.createdAt - a.createdAt || a.key.localeCompare(b.key))

  const days = [...new Set(timeline.map((entry) => formatDate(entry.createdAt)))].map((date) => ({
    date,
    entries: timeline.filter((entry) => formatDate(entry.createdAt) === date),
  }))

  const files = (
    <div className="matter-files-collection">
      {uploads.length > 0 && (
        <ul className="matter-list uploaded-files" aria-label="Uploads">
          {uploads.map((u) => (
            <UploadRow
              key={u.key}
              upload={u}
              onRetry={() => {
                update(u.key, { status: "queued", error: undefined })
                enqueue(u)
              }}
              onChoose={(choice) => {
                update(u.key, { status: "queued", choice })
                enqueue({ ...u, choice })
              }}
              onDismiss={() => setUploads((prev) => prev.filter((x) => x.key !== u.key))}
            />
          ))}
        </ul>
      )}
      {documents.length === 0 && uploads.length === 0 && (
        <div className="rail-empty">
          <IconDocs />
          <strong>No documents yet</strong>
          <span>Upload the matter's documents to read, cite and mark them up here.</span>
        </div>
      )}
      {timeline.length > 0 && (
        <>
          <p className="docs-timeline-label">Newest first</p>
          <ol className="docs-timeline" aria-label="Document timeline, newest first">
            {days.map((day) => (
              <li key={day.date} className="docs-timeline-entry">
                <h3 className="docs-timeline-date">{day.date}</h3>
                <ol className="docs-timeline-files">
                  {day.entries.map((entry) => <li key={entry.key}>{entry.content}</li>)}
                </ol>
              </li>
            ))}
          </ol>
        </>
      )}
    </div>
  )

  const badge = pending > 0 && (
    <span className="redline-badge" title={`${pending} pending change${pending === 1 ? "" : "s"} to review`}>
      {pending}
    </span>
  )

  return (
    <section className="docs-drawer-panel" aria-labelledby="docs-drawer-title">
      <div className="docs-drawer-head">
        <h2 id="docs-drawer-title">
          Documents <span className="matter-files-count">{documents.length}</span>
        </h2>
        {badge}
        {onClose && (
          <button type="button" className="sidebar-icon-button docs-drawer-close" onClick={onClose} title="Close (Esc)" aria-label="Close documents">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" aria-hidden>
              <path d="M18 6 6 18M6 6l12 12" />
            </svg>
          </button>
        )}
      </div>
      <div className="docs-drawer-body" tabIndex={0} role="region" aria-label="Documents" onDragOver={(event) => event.preventDefault()} onDrop={(event) => {
        event.preventDefault()
        const added = Array.from(event.dataTransfer.files)
        if (added.length) onFiles(added)
      }}>
        {files}
      </div>
      <footer className="docs-drawer-footer">
        <Dropzone
          accept=".docx,.pdf"
          busy={working > 0}
          compact
          onFiles={onFiles}
          title={working > 0 ? `Uploading ${working} document${working === 1 ? "" : "s"}…` : "Upload documents"}
        />
        <p>Drop Word or PDF files here, or click to browse.</p>
      </footer>
    </section>
  )
}

// A same-name clash asks the lawyer what to do with the existing document; any
// other failure can be retried.
export function failedUpload(error: unknown, file: File): Partial<Upload> {
  if (error instanceof DocumentExistsError) return { status: "exists", existing: { name: error.existing, role: error.role } }
  return { status: "failed", error: humanizeError(error, `Could not read ${file.name}.`), retryable: true }
}

const UPLOAD_LABEL: Record<Upload["status"], string> = {
  queued: "Waiting…",
  uploading: "Uploading…",
  reading: "Reading the document…",
  ready: "Ready",
  failed: "Not added",
  exists: "Already in this matter. Keep both or replace it?",
}
const DELIVERABLE_EXISTS = "A document created by Doc.Haus has this name. Keep both?"

export function UploadRow({ upload, onRetry, onChoose, onDismiss }: {
  upload: Upload
  onRetry: () => void
  onChoose: (choice: ExistingSourceChoice) => void
  onDismiss: () => void
}) {
  return (
    <li className={`doc-row upload-row is-${upload.status}`} aria-live="polite">
      <div className="doc-row-main">
        <span className="upload-icon" aria-hidden>
          {upload.status === "failed" || upload.status === "exists" ? <IconAlert /> : upload.status === "ready" ? <IconCheck /> : <span className="upload-spinner" />}
        </span>
        <span className="upload-name">{upload.file.name}</span>
      </div>
      <div className="doc-row-meta">
        <span className="upload-status">
          {upload.error ?? (upload.existing?.role === "deliverable" ? DELIVERABLE_EXISTS : UPLOAD_LABEL[upload.status])}
        </span>
        {upload.status === "failed" && (
          <span className="upload-actions">
            {upload.retryable && (
              <button type="button" className="icon-btn" onClick={onRetry}>
                Retry
              </button>
            )}
            <button type="button" className="icon-btn" onClick={onDismiss}>
              Remove
            </button>
          </span>
        )}
        {upload.status === "exists" && (
          <span className="upload-actions">
            {/* Keep both loses nothing, so it leads; Replace overwrites and stays plain. */}
            <button type="button" className="icon-btn upload-keep-both" title="Add this file under a new name and keep both" onClick={() => onChoose("keep-both")}>
              Keep both
            </button>
            {upload.existing?.role === "source" && (
              <button
                type="button"
                className="icon-btn"
                title={`Replace ${upload.existing.name} with this file — the current file is overwritten`}
                onClick={() => onChoose("replace")}
              >
                Replace
              </button>
            )}
            <button type="button" className="icon-btn" onClick={onDismiss}>
              Cancel
            </button>
          </span>
        )}
      </div>
    </li>
  )
}

function IconDocs() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
    </svg>
  )
}

function IconFile() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
    </svg>
  )
}

function IconCheck() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="20 6 9 17 4 12" />
    </svg>
  )
}

function IconAlert() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="12" cy="12" r="9" />
      <line x1="12" y1="8" x2="12" y2="12.5" />
      <line x1="12" y1="16" x2="12" y2="16" />
    </svg>
  )
}

function IconTrash() {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <polyline points="3 6 5 6 21 6" />
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
    </svg>
  )
}
