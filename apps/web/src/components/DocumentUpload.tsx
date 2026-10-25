import { useEffect, useRef, useState } from "react"
import {
  deleteDocument,
  DocumentExistsError,
  duplicateMessage,
  DuplicateUploadError,
  MATTER_DOCUMENT_EXTENSIONS,
  matterDocumentKind,
  UploadFailedError,
  uploadDocument,
  type Document,
  type EmailAttachment,
  type EmailRemoval,
  type ExistingSourceChoice,
  type ExtractionQuality,
  type IngestResult,
  type MatterDocumentKind,
} from "../api/ingest/documents"
import { humanizeError } from "../api/ingest/errors"
import type { LeanDraftingStatus } from "../api/ingest/lean"
import { restoreVersion } from "../api/ingest/versions"
import { formatDate, formatDateTime, formatTime } from "../lib/format"
import { CompareIcon } from "./CompareDialog"
import DocumentLabelBadge from "./DocumentLabel"
import { restoreConfirmText, restoredMessage, unavailableReason } from "./DocumentVersions"
import DraftedList from "./drafting/DraftedList"
import { deliverableGroups, type DraftedOutput } from "./drafting/model"
import Dropzone from "./Dropzone"
import { useToast } from "./Toast"

// One row per file the lawyer added in this session, so a slow or failed upload
// is visible next to its name instead of hiding behind a single spinner.
type Upload = {
  key: number
  file: File
  // not_indexed: added, but no readable text could be extracted, so nothing
  // was indexed (the extraction quality gate). Never shown as ready.
  // duplicate: identical to a version the matter already keeps; nothing saved.
  status: "queued" | "uploading" | "reading" | "ready" | "not_indexed" | "failed" | "exists" | "duplicate"
  error?: string
  // A ready Word file saved with pending tracked changes: the notice stays in
  // its row until dismissed.
  notice?: string
  // A wrong file type cannot succeed on retry, so the row offers Remove only.
  retryable?: boolean
  // A document of this name is already in the matter. The row asks the lawyer to
  // save the upload as a new version of an uploaded source or keep both, or to
  // keep both beside a generated deliverable, and the choice is sent on retry.
  existing?: { name: string; role: "source" | "deliverable" }
  choice?: ExistingSourceChoice
  // A duplicate of an earlier version of a document: the row offers to make
  // that version current again (a restore, kept as the next version).
  match?: { name: string; version: { id: string; number: number } }
  restoring?: boolean
}

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
  // A long attachment list the lawyer opened or closed, by parent name.
  const [expanded, setExpanded] = useState<Record<string, boolean>>({})
  const nextKey = useRef(0)
  const watched = useRef(false)
  // Uploads run one at a time so a batch does not swamp the reader.
  const queue = useRef(Promise.resolve())
  const body = useRef<HTMLDivElement>(null)
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
      update(upload.key, { status: result.quality?.status === "no_readable_text" ? "not_indexed" : "ready", notice: result.notice })
      if (!result.notice) setTimeout(() => setUploads((prev) => prev.filter((u) => u.key !== upload.key)), READY_LINGER_MS)
      onUploaded()
      const { kind, message } = uploadedToast(result)
      toast(kind, message)
    } catch (e) {
      update(upload.key, failedUpload(e, upload.file))
    }
  }

  async function restoreMatch(upload: Upload) {
    const match = upload.match
    if (!match) return
    update(upload.key, { restoring: true })
    try {
      const result = await restoreVersion(matterId, match.version.id)
      // The row goes with its buttons: focus moves to the documents list.
      body.current?.focus()
      setUploads((prev) => prev.filter((u) => u.key !== upload.key))
      onUploaded()
      toast("success", restoredMessage(match.version.number, result, match.name))
      if (result.notice) toast("warning", result.notice)
    } catch (e) {
      update(upload.key, {
        restoring: false,
        match: undefined,
        error: unavailableReason(e) ?? humanizeError(e, `Could not restore version ${match.version.number} of ${match.name}.`),
      })
    }
  }

  function onFiles(files: File[]) {
    const added = files.map((file) => newUpload(nextKey.current++, file))
    setUploads((prev) => [...prev, ...added])
    added.filter((u) => u.status === "queued").forEach(enqueue)
  }

  async function onRemove(name: string, keepAttachments = false) {
    setConfirmId(null)
    setRemoving(name)
    try {
      const result = await deleteDocument(matterId, name, { keepAttachments })
      onUploaded()
      const removed = result.attachments?.removed.length ?? 0
      const kept = result.attachments?.kept ?? []
      const failed = result.attachments?.failed ?? []
      toast(
        failed.length ? "error" : "success",
        `Removed ${name}${removed ? ` and ${parentWords(name).count(removed)}` : ""}.${
          kept.length ? ` Kept ${kept.map((k) => k.name).join(", ")}.` : ""
        }${failed.map((f) => ` ${f.name}: ${f.reason}`).join("")}`,
      )
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

  // A source row, then the documents that came out of it (an email's
  // attachments, an archive's files) indented under it, each listed once.
  function renderSource(node: SourceNode, depth: number) {
    const d = node.document
    const child = depth > 0
    const parent = d.attachment_of
    const viewable = matterDocumentKind(d.name) !== undefined
    const renamed = node.attachment?.name && node.attachment.name !== d.name ? node.attachment.name : undefined
    const notes = [
      renamed && (parent && parentWords(parent.name).parent === "zip" ? `In the zip as ${renamed}` : `Attached as ${renamed}`),
      node.attachment?.note === "may-contain-attachments" && "May contain further attachments",
      node.attachment?.flagged && "Suspicious name",
    ].filter(Boolean)
    const actions = <>
      {onCompare && /\.docx$/i.test(d.name) && (
        <button className="doc-compare" onClick={() => onCompare(d.name)} title="Compare with another version" aria-label={`Compare ${d.name} with another version`}>
          <CompareIcon />
        </button>
      )}
      {confirmId === d.id && d.removal ? null : confirmId === d.id ? (
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
    </>
    return (
      <>
        <div className={`doc-row is-uploaded${child ? " is-child" : ""}`}>
          <div className="doc-row-main">
            <DocumentType name={d.name} />
            {viewable ? (
              <button className="linklike" onClick={() => onView(d.name)}>
                {d.name}
              </button>
            ) : (
              <span className="doc-row-name">{d.name}</span>
            )}
            {d.pending ? (
              <span className="redline-badge" title={`${d.pending} pending change${d.pending === 1 ? "" : "s"} to review`}>
                {d.pending}
              </span>
            ) : null}
            {child && <span className="doc-row-actions">{actions}</span>}
          </div>
          <QualityNote quality={d.quality} />
          {notes.length > 0 && <span className="doc-child-status">{notes.join(" · ")}</span>}
          {!child && (
            <div className="doc-row-meta">
              <span className="document-origin" title={parent ? `Extracted from the ${parentWords(parent.name).parent} ${parent.name}` : undefined}>
                {parent ? `From removed ${parentWords(parent.name).parent} ${parent.name}` : "Uploaded"}
              </span>
              <time className="document-time" dateTime={new Date(d.created_at).toISOString()} title={`Added to matter: ${formatDateTime(d.created_at)}`}>
                {formatTime(d.created_at)}
              </time>
              {actions}
            </div>
          )}
          {confirmId === d.id && d.removal && (
            <EmailRemoveConfirm name={d.name} removal={d.removal} onRemove={(keep) => onRemove(d.name, keep)} onCancel={() => setConfirmId(null)} />
          )}
        </div>
        {renderChildren(node, depth)}
      </>
    )
  }

  function renderChildren(node: SourceNode, depth: number) {
    const d = node.document
    const count = node.children.length
    const embedded = d.embedded_images ?? 0
    if (!count && !embedded) return null
    const long = count > EXPANDED_CHILDREN
    const open = expanded[d.name] ?? !long
    const id = `doc-children-${d.id}`
    const words = parentWords(d.name)
    return (
      <div className="doc-children-group">
        {long && (
          <button type="button" className="doc-children-toggle" aria-expanded={open} aria-controls={id}
            onClick={() => setExpanded((prev) => ({ ...prev, [d.name]: !open }))}>
            <IconChevron />
            {open ? `Hide ${words.count(count)}` : words.count(count)}
          </button>
        )}
        {open && (
          <ul id={id} className="doc-children" aria-label={`${words.list} ${d.name}`}>
            {node.children.map((item, index) => (
              <li key={"document" in item ? `doc:${item.document.id}` : `entry:${index}`}>
                {"document" in item ? renderSource(item, depth + 1) : <AttachmentEntryRow attachment={item.attachment} documents={sources} onView={onView} />}
              </li>
            ))}
            {embedded > 0 && (
              <li className="doc-children-note">
                {embedded} embedded image{embedded === 1 ? "" : "s"} in the message body
              </li>
            )}
          </ul>
        )}
      </div>
    )
  }

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
    ...documentTree(sources).map((node) => ({
      key: `source:${node.document.id}`,
      createdAt: node.document.created_at,
      content: renderSource(node, 0),
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
              onRestore={() => restoreMatch(u)}
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
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden>
              <path d="M18 6 6 18M6 6l12 12" />
            </svg>
          </button>
        )}
      </div>
      <div ref={body} className="docs-drawer-body" tabIndex={0} role="region" aria-label="Documents" onDragOver={(event) => event.preventDefault()} onDrop={(event) => {
        event.preventDefault()
        const added = Array.from(event.dataTransfer.files)
        if (added.length) onFiles(added)
      }}>
        {files}
      </div>
      <footer className="docs-drawer-footer">
        <Dropzone
          accept={MATTER_DOCUMENT_EXTENSIONS.join(",")}
          busy={working > 0}
          compact
          onFiles={onFiles}
          title={working > 0 ? `Uploading ${working} document${working === 1 ? "" : "s"}…` : "Upload documents"}
        />
        <p>Drop Word, PDF, email, Excel or zip files here, or click to browse.</p>
      </footer>
    </section>
  )
}

// The toast after an upload. A file with no readable text is added but not
// indexed, and never reported as ready. A flagged or degraded upload is a
// finding the lawyer needs to see, not a failure: the document is indexed, the
// assistant treats flagged text as untrusted and is told the text may be
// incomplete, and the warning says what ingest found.
export function uploadedToast(result: IngestResult): { kind: "success" | "warning"; message: string } {
  // A Word file kept with pending tracked changes says so after the rest; the
  // notice names the version, so the ready line does not repeat it.
  if (!result.notice) return readyToast(result)
  return { kind: "warning", message: `${readyToast({ ...result, version: undefined }).message} ${result.notice}` }
}

function readyToast(result: IngestResult): { kind: "success" | "warning"; message: string } {
  const attachments = attachmentsSummary(result)
  const quality = result.quality
  const injection = result.injection
    ? `possible prompt injection: ${[...new Set(result.injection.findings.map((f) => f.detail))].join("; ")}. The assistant will treat this content as untrusted.`
    : ""
  // An unreadable file can still carry an injection finding (hidden DOCX
  // content, invisible characters); both are shown.
  if (quality?.status === "no_readable_text")
    return {
      kind: "warning",
      message: `${result.name} was added but not indexed: no readable text (${qualityDetails(quality)}). The assistant cannot search or cite it.${
        injection ? ` It also has a warning — ${injection}` : ""
      }${attachments}`,
    }
  const degraded = quality?.status === "degraded" ? `text may be incomplete or garbled (${qualityDetails(quality)}).` : ""
  if (injection || degraded)
    return {
      kind: "warning",
      message: `${result.name} is ready, with ${injection && degraded ? "warnings" : "a warning"} — ${
        injection && degraded ? `${injection} Its ${degraded}` : injection || `its ${degraded}`
      }${attachments}`,
    }
  // A zip file's own text only lists its files, so its sections say nothing.
  if (result.name.toLowerCase().endsWith(".zip")) return { kind: "success", message: `${result.name} is ready.${attachments}` }
  // An upload saved as a later version of its document names the version.
  const version = result.version?.created && result.version.number > 1 ? ` as version ${result.version.number}` : ""
  return { kind: "success", message: `${result.name} is ready${version} (${result.sections} section${result.sections === 1 ? "" : "s"}).${attachments}` }
}

function qualityDetails(quality: Exclude<ExtractionQuality, { status: "ok" }>) {
  return quality.signals.map((signal) => signal.detail.charAt(0).toLowerCase() + signal.detail.slice(1)).join("; ")
}

// A source's extraction quality under its name: nothing when it is ok or was
// never assessed; otherwise what the lawyer and the assistant can rely on, with
// what ingest found on hover.
export function QualityNote({ quality }: { quality?: ExtractionQuality }) {
  if (!quality || quality.status === "ok") return null
  const details = quality.signals.map((signal) => signal.detail).join("; ")
  return (
    <span className={`document-quality is-${quality.status}`} title={details}>
      {quality.status === "no_readable_text" ? "Not indexed: no readable text" : "Text may be incomplete or garbled"}
    </span>
  )
}

// What became of an uploaded email's attachments, or a zip file's files, for
// the ready toast.
export function attachmentsSummary(result: Pick<IngestResult, "attachments"> & { name?: string }) {
  const attachments = result.attachments ?? []
  if (!attachments.length) return ""
  const noun = result.name?.toLowerCase().endsWith(".zip") ? "file" : "attachment"
  const added = attachments.filter((a) => a.status === "indexed").length
  const same = attachments.filter((a) => a.status === "same").length
  const unreadable = attachments.filter((a) => a.status === "no_readable_text")
  const missing = attachments.filter((a) => a.status !== "indexed" && a.status !== "same" && a.status !== "no_readable_text")
  return [
    added && ` ${added} ${noun}${added === 1 ? "" : "s"} added.`,
    same && ` ${same} already in the matter.`,
    unreadable.length && ` Added but not indexed, no readable text: ${unreadable.map(attachmentName).join(", ")}.`,
    missing.length && ` Not added: ${missing.map((a) => `${attachmentName(a)} (${[attachmentStatus(a.status).toLowerCase(), attachmentNote(a)].filter(Boolean).join(": ")})`).join(", ")}.`,
  ]
    .filter(Boolean)
    .join("")
}

// What became of an attachment, in words. A status this build does not know
// (a newer ingest's) still reads as not added rather than as nothing.
const ATTACHMENT_STATUS: Record<string, string> = {
  indexed: "Added",
  same: "Already in the matter",
  removed: "Removed from the matter",
  unsupported: "Not supported",
  failed: "Could not be read",
  nested: "Nested too deep to add",
  pending: "Not added yet. Upload the email again",
  no_readable_text: "Not indexed: no readable text",
  encrypted: "Password-protected",
  unsafe_path: "Unsafe path in the zip",
  too_large: "Too large to unpack",
} satisfies Record<EmailAttachment["status"], string> & Record<string, string>

function attachmentStatus(status: string) {
  return Object.hasOwn(ATTACHMENT_STATUS, status) ? ATTACHMENT_STATUS[status]! : "Not added"
}

function attachmentName(attachment: EmailAttachment) {
  return attachment.name || attachment.document || "Unnamed attachment"
}

// Why an Outlook attachment named like a document was not read.
function attachmentNote(attachment: EmailAttachment) {
  if (attachment.note === "linked") return "a link to a file stored elsewhere"
  if (attachment.note === "ole") return "an embedded object, not a file"
}

// The documents list as a tree. A document extracted from another (an email's
// attachment, a file from an archive) sits under that parent, keyed on the
// parent link rather than the file type, and is listed only there. Each parent
// also lists, in the order it states them, the attachments that did not become
// one of its documents: not added, removed since, or the same as a document
// already elsewhere in the matter, so nothing silently disappears. A document
// whose parent was removed, or whose links loop, is listed at the top level.
export type SourceNode = { document: Document; attachment?: EmailAttachment; children: (SourceNode | { attachment: EmailAttachment })[] }

export function documentTree(sources: Document[]): SourceNode[] {
  const byName = new Map(sources.map((d) => [d.name, d]))
  const parentOf = (d: Document) =>
    d.attachment_of && !d.attachment_of.removed && d.attachment_of.name !== d.name ? byName.get(d.attachment_of.name) : undefined
  const children = new Map<string, Document[]>()
  for (const d of sources) {
    const parent = parentOf(d)
    if (parent) children.set(parent.name, [...(children.get(parent.name) ?? []), d])
  }
  const placed = new Set<string>()
  const node = (document: Document, attachment?: EmailAttachment): SourceNode => {
    placed.add(document.name)
    const own = children.get(document.name) ?? []
    const items: SourceNode["children"] = []
    for (const entry of document.attachments ?? []) {
      const child = own.find((d) => d.name === entry.document && !placed.has(d.name))
      items.push(child ? node(child, entry) : { attachment: entry })
    }
    for (const child of own) if (!placed.has(child.name)) items.push(node(child))
    return { document, attachment, children: items }
  }
  const roots = sources.filter((d) => !parentOf(d)).map((d) => node(d))
  for (const d of sources) if (!placed.has(d.name)) roots.push(node(d))
  return roots
}

// Up to this many children show under their parent; a longer list starts
// closed behind a count.
const EXPANDED_CHILDREN = 5

// How a parent and what came out of it are named: an email's attachments, a
// zip's files, or any other parent's files.
function parentWords(name: string) {
  const parent = /\.(eml|msg)$/i.test(name) ? "email" : /\.zip$/i.test(name) ? "zip" : "file"
  const child = parent === "email" ? "attachment" : "file"
  return {
    parent,
    children: `${child}s`,
    count: (n: number) => `${n} ${child}${n === 1 ? "" : "s"}`,
    list: parent === "email" ? "Attachments of" : "Files in",
    came: parent === "email" ? "came with" : "contains",
  }
}

// An attachment that is not one of its parent's documents: its stated name and
// what became of it. One kept as, or the same as, a document elsewhere in the
// matter links to that document; the rest cannot be opened.
export function AttachmentEntryRow({ attachment, documents, onView }: {
  attachment: EmailAttachment
  documents: Document[]
  onView: (name: string) => void
}) {
  const linked = attachment.document && attachment.status !== "removed" && documents.some((d) => d.name === attachment.document)
    ? attachment.document
    : undefined
  const status = attachment.status === "same" && attachment.document ? `Same as ${attachment.document}` : attachmentStatus(attachment.status)
  const note = attachmentNote(attachment)
  return (
    <div className={`doc-row is-child is-entry is-${attachment.status}`}>
      <div className="doc-row-main">
        <DocumentType name={attachmentName(attachment)} />
        <span className="doc-row-name">{attachmentName(attachment)}</span>
      </div>
      <span className="doc-child-status">
        {linked ? (
          <button type="button" className="linklike" onClick={() => onView(linked)}>
            {status}
          </button>
        ) : (
          status
        )}
        {attachment.note === "may-contain-attachments" && " · may contain further attachments"}
        {note && ` · ${note}`}
        {attachment.flagged && " · suspicious name"}
      </span>
    </div>
  )
}

// Removing an email (or a zip) names the documents that came out of it and
// offers to remove them too (the default) or the parent only. A document
// another parent still links to, or that the lawyer changed, is always kept and
// named.
export function EmailRemoveConfirm({ name, removal, onRemove, onCancel }: {
  name: string
  removal: EmailRemoval
  onRemove: (keepAttachments: boolean) => void
  onCancel: () => void
}) {
  const count = removal.attachments.length
  const words = parentWords(name)
  return (
    <div className="email-remove-confirm" role="group" aria-label={`Remove ${name}`} onClick={(e) => e.stopPropagation()}>
      <p>
        {count
          ? `${name} ${words.came} ${removal.attachments.join(", ")}.`
          : `${name} ${words.came} ${words.children} that stay in the matter.`}
        {removal.kept.map((kept) => (
          <span key={kept.name} className="email-remove-kept">
            {" "}
            {kept.name} is kept: {kept.reason === "linked" ? "another email or zip also has it" : "it has changed since it was added"}.
          </span>
        ))}
      </p>
      <span className="upload-actions">
        <button type="button" className="icon-btn danger" autoFocus onClick={() => onRemove(false)}>
          {count ? `Remove ${words.parent} and ${words.count(count)}` : `Remove ${words.parent}`}
        </button>
        {count > 0 && (
          <button type="button" className="icon-btn" onClick={() => onRemove(true)}>
            Remove {words.parent} only
          </button>
        )}
        <button type="button" className="icon-btn" onClick={onCancel}>
          Cancel
        </button>
      </span>
    </div>
  )
}

// Every file the picker or a drop hands over (the composer's too) gets a row. A
// format ingest cannot read is refused here, with the formats it can.
export function newUpload(key: number, file: File): Upload {
  if (matterDocumentKind(file.name)) return { key, file, status: "queued" }
  return { key, file, status: "failed", error: "Only Word (.docx), PDF, email (.eml, .msg), Excel (.xlsx) and zip (.zip) files can be added." }
}

// A same-name clash asks the lawyer what to do with the existing document; an
// identical file says which version it matches. A lost connection or a service
// failure can be retried; a refusal for this file (4xx) would only repeat.
export function failedUpload(error: unknown, file: File): Partial<Upload> {
  if (error instanceof DocumentExistsError) return { status: "exists", existing: { name: error.existing, role: error.role } }
  if (error instanceof DuplicateUploadError) {
    const earlier = error.version && !error.version.current ? { name: error.existing, version: { id: error.version.id, number: error.version.number } } : undefined
    return { status: "duplicate", error: duplicateMessage(error.existing, error.version), match: earlier }
  }
  const failed = { status: "failed" as const, error: humanizeError(error, `Could not read ${file.name}.`) }
  return error instanceof UploadFailedError && error.status < 500 ? failed : { ...failed, retryable: true }
}

const UPLOAD_LABEL: Record<Upload["status"], string> = {
  queued: "Waiting…",
  uploading: "Uploading…",
  reading: "Reading the document…",
  ready: "Ready",
  not_indexed: "Added, not indexed: no readable text",
  failed: "Not added",
  exists: "Already in this matter. Save this as a new version of it, or keep both?",
  duplicate: "Already in this matter. Nothing was saved.",
}
const DELIVERABLE_EXISTS = "A document created by Doc.Haus has this name. Keep both?"

export function UploadRow({ upload, onRetry, onChoose, onDismiss, onRestore }: {
  upload: Upload
  onRetry: () => void
  onChoose: (choice: ExistingSourceChoice) => void
  onDismiss: () => void
  onRestore?: () => void
}) {
  const [confirming, setConfirming] = useState(false)
  // A cancelled confirmation hands focus back to the Restore button.
  const refocus = useRef(false)
  const dismiss = useRef<HTMLButtonElement>(null)
  const match = upload.status === "duplicate" && onRestore ? upload.match : undefined
  // A failed restore takes the confirmation with it: focus stays in the row,
  // on Dismiss, beside the reason.
  useEffect(() => {
    if (!confirming || match) return
    setConfirming(false)
    dismiss.current?.focus()
  }, [confirming, match])
  function cancel() {
    refocus.current = true
    setConfirming(false)
  }
  return (
    <li className={`doc-row upload-row is-${upload.status}${upload.notice ? " has-notice" : ""}`} aria-live="polite">
      <div className="doc-row-main">
        <span className="upload-icon" aria-hidden>
          {upload.status === "failed" || upload.status === "exists" || upload.status === "not_indexed" || upload.status === "duplicate" || upload.notice
            ? <IconAlert />
            : upload.status === "ready" ? <IconCheck /> : <span className="upload-spinner" />}
        </span>
        <span className="upload-name">{upload.file.name}</span>
      </div>
      <div className="doc-row-meta">
        <span className="upload-status">
          {match && confirming
            ? restoreConfirmText(match.version.number)
            : upload.error ?? upload.notice ?? (upload.existing?.role === "deliverable" ? DELIVERABLE_EXISTS : UPLOAD_LABEL[upload.status])}
        </span>
        {match && confirming && (
          <span className="upload-actions" role="group" aria-label={`Restore version ${match.version.number}`} onKeyDown={(event) => {
            if (event.key !== "Escape" || upload.restoring) return
            event.preventDefault()
            event.stopPropagation()
            cancel()
          }}>
            <button type="button" className="icon-btn upload-choice-lead" autoFocus disabled={upload.restoring} onClick={onRestore}>
              {upload.restoring ? "Restoring…" : `Restore version ${match.version.number}`}
            </button>
            <button type="button" className="icon-btn" disabled={upload.restoring} onClick={cancel}>
              Cancel
            </button>
          </span>
        )}
        {(upload.status === "duplicate" || upload.notice) && !(match && confirming) && (
          <span className="upload-actions">
            {/* The file is an earlier version: making it current again is
                usually what the lawyer meant by uploading it. */}
            {match && (
              <button type="button" className="icon-btn upload-choice-lead" ref={(node) => {
                if (!node || !refocus.current) return
                refocus.current = false
                node.focus()
              }}
                title={`Make version ${match.version.number} of ${match.name} current again. Every version stays in the history.`}
                onClick={() => setConfirming(true)}>
                {`Restore version ${match.version.number}`}
              </button>
            )}
            <button type="button" className="icon-btn" ref={dismiss} onClick={onDismiss}>
              Dismiss
            </button>
          </span>
        )}
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
            {/* Neither choice loses anything. A new version of the same document
                is the usual intent, so it leads; a generated document is never
                replaced by an upload, so only keep both is offered beside it. */}
            {upload.existing?.role === "source" && (
              <button
                type="button"
                className="icon-btn upload-choice-lead"
                title={`Save this file as the next version of ${upload.existing.name}. Earlier versions stay in its history.`}
                onClick={() => onChoose("replace")}
              >
                Save as new version
              </button>
            )}
            <button
              type="button"
              className={`icon-btn${upload.existing?.role === "source" ? "" : " upload-choice-lead"}`}
              title="Add this file under a new name and keep both"
              onClick={() => onChoose("keep-both")}
            >
              Keep both
            </button>
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
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
    </svg>
  )
}

function IconFile() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
    </svg>
  )
}

// The file-type glyph beside a document name: emails, spreadsheets and archives
// get their own, so they read apart from the Word and PDF files beside them.
type DocumentGlyph = MatterDocumentKind | "zip" | "file"
const DOCUMENT_TYPE_LABEL: Record<DocumentGlyph, string> = {
  docx: "Word document",
  pdf: "PDF document",
  eml: "Email",
  msg: "Outlook email",
  xlsx: "Excel spreadsheet",
  zip: "Zip file",
  file: "File",
}

function DocumentType({ name }: { name: string }) {
  const kind: DocumentGlyph = matterDocumentKind(name) ?? (/\.zip$/i.test(name.trim()) ? "zip" : "file")
  return (
    <span className={`doc-type is-${kind}`} title={DOCUMENT_TYPE_LABEL[kind]} aria-hidden>
      {kind === "eml" || kind === "msg" ? <IconEmail /> : kind === "xlsx" ? <IconSheet /> : kind === "zip" ? <IconArchive /> : <IconFile />}
    </span>
  )
}

function IconArchive() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="4" width="18" height="5" rx="1" />
      <path d="M5 9v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V9M10 13h4" />
    </svg>
  )
}

function IconChevron() {
  return (
    <svg className="doc-children-chevron" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.25" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <polyline points="9 5 16 12 9 19" />
    </svg>
  )
}

function IconEmail() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="5" width="18" height="14" rx="2" />
      <polyline points="3 7 12 13 21 7" />
    </svg>
  )
}

function IconSheet() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="M3 9h18M3 15h18M9 3v18" />
    </svg>
  )
}

function IconCheck() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="20 6 9 17 4 12" />
    </svg>
  )
}

function IconAlert() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="12" cy="12" r="9" />
      <line x1="12" y1="8" x2="12" y2="12.5" />
      <line x1="12" y1="16" x2="12" y2="16" />
    </svg>
  )
}

function IconTrash() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor">
      <polyline points="3 6 5 6 21 6" />
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
    </svg>
  )
}
