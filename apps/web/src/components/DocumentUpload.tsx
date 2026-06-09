import { useRef, useState } from "react"
import { uploadDocument, type Document } from "../api/ingest"

export default function DocumentUpload({
  matterId,
  documents,
  onUploaded,
  onView,
}: {
  matterId: string
  documents: Document[]
  onUploaded: () => void
  onView: (name: string) => void
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

  return (
    <div className="card">
      <h2>Documents</h2>
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
    </div>
  )
}
