import { useRef, useState } from "react"
import { deleteDocument, uploadDocument, type Document } from "../api/ingest"
import { useToast } from "./Toast"

// The matter's documents, as a collapsible right rail beside the chat. Chat is
// the primary surface, so documents sit out of its way: expanded the rail shows
// the indexed list plus a dropzone; collapsed it shrinks to a thin tab carrying
// the document count, reclaiming the width for the conversation.
export default function DocumentUpload({
  matterId,
  documents,
  onUploaded,
  onView,
  collapsed,
  onToggle,
}: {
  matterId: string
  documents: Document[]
  onUploaded: () => void
  onView: (name: string) => void
  // Rail mode (chat surface) is collapsible; the Documents surface omits these
  // and renders the full-width manager with no collapse affordance.
  collapsed?: boolean
  onToggle?: () => void
}) {
  const input = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  const toast = useToast()

  async function onFile(file: File) {
    setBusy(true)
    const result = await uploadDocument(matterId, file)
    setBusy(false)
    onUploaded()
    toast("success", `Indexed ${result.name}: ${result.sections} sections, ${result.chunks} chunks.`)
  }

  async function onRemove(name: string) {
    if (!confirm(`Remove ${name} from this matter? Its indexed text is deleted and answers can no longer cite it.`)) return
    setBusy(true)
    await deleteDocument(matterId, name)
    setBusy(false)
    onUploaded()
    toast("success", `Removed ${name}.`)
  }

  if (collapsed) {
    return (
      <button className="docs-tab" onClick={onToggle} title="Show documents">
        <IconDocs />
        <span className="docs-tab-count">{documents.length}</span>
        <span className="docs-tab-label">Documents</span>
      </button>
    )
  }

  return (
    <aside className={onToggle ? "card docs-rail" : "card docs-surface"}>
      <div className="docs-rail-head">
        <h2>Documents</h2>
        {onToggle && (
          <button className="icon-btn" onClick={onToggle} title="Hide documents">
            <IconChevron />
          </button>
        )}
      </div>
      {documents.length === 0 ? (
        <p className="muted">No documents indexed yet.</p>
      ) : (
        <ul className="matter-list">
          {documents.map((d) => (
            <li key={d.id}>
              <button className="linklike" onClick={() => onView(d.name)}>
                {d.name}
              </button>
              <span className="muted">{new Date(d.created_at).toLocaleDateString()}</span>
              <button className="doc-remove" onClick={() => onRemove(d.name)} title={`Remove ${d.name}`} disabled={busy}>
                <IconTrash />
              </button>
            </li>
          ))}
        </ul>
      )}
      <div
        className="dropzone"
        style={{ marginTop: 12 }}
        onClick={() => input.current?.click()}
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault()
          const file = e.dataTransfer.files[0]
          if (file) onFile(file)
        }}
      >
        {busy ? "Working..." : "Drop a .docx contract here, or click to choose a file."}
      </div>
      <input
        ref={input}
        type="file"
        accept=".docx"
        hidden
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) onFile(file)
          e.target.value = ""
        }}
      />
    </aside>
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

function IconChevron() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <polyline points="9 18 15 12 9 6" />
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
