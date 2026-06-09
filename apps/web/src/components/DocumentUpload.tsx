import { useRef, useState } from "react"
import { uploadDocument, type Document } from "../api/ingest"

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
  collapsed: boolean
  onToggle: () => void
}) {
  const input = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState("")

  async function onFile(file: File) {
    setBusy(true)
    setStatus(`Ingesting ${file.name}...`)
    const result = await uploadDocument(matterId, file)
    setStatus(`Indexed ${result.name}: ${result.sections} sections, ${result.chunks} chunks.`)
    setBusy(false)
    onUploaded()
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
    <aside className="card docs-rail">
      <div className="docs-rail-head">
        <h2>Documents</h2>
        <button className="icon-btn" onClick={onToggle} title="Hide documents">
          <IconChevron />
        </button>
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
        {busy ? status : "Drop a .docx contract here, or click to choose a file."}
      </div>
      {!busy && status && (
        <p className="muted" style={{ marginTop: 8 }}>
          {status}
        </p>
      )}
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
