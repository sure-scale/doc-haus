import { useRef, useState, type DragEvent, type ReactNode } from "react"
import { toggleDocumentsDrawer, useDocumentsDrawer } from "../matter/documents-drawer"

// The matter's documents at the composer: a Documents control beside the
// playbook that opens the same drawer as the rail, and a drop target over the
// whole composer that adds files to the matter.
export default function DocumentsDock({
  count,
  onAddFiles,
  children,
}: {
  count: number
  onAddFiles: (files: File[]) => void
  // The composer, given the Documents control to seat in its tools row.
  children: (control: ReactNode) => ReactNode
}) {
  const drawer = useDocumentsDrawer()
  const [dragging, setDragging] = useState(false)
  // Nested drag events fire enter/leave per child; a depth counter keeps the
  // highlight steady until the pointer actually leaves the dock.
  const depth = useRef(0)
  const hasFiles = (e: DragEvent) => Array.from(e.dataTransfer.types).includes("Files")

  const control = (
    <button
      type="button"
      className={`workflow-btn docs-btn${drawer.open ? " is-open" : ""}`}
      onClick={toggleDocumentsDrawer}
      aria-expanded={drawer.open}
      aria-controls="documents-drawer"
      title={drawer.open ? "Close the documents panel" : "Open this matter's documents"}
    >
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
        <polyline points="14 2 14 8 20 8" />
      </svg>
      <span>Documents</span>
      <span className="docs-btn-count">{count}</span>
    </button>
  )

  return (
    <div
      className={`composer-dock${dragging ? " is-dragover" : ""}`}
      onDragEnter={(e) => {
        if (!hasFiles(e)) return
        e.preventDefault()
        depth.current++
        setDragging(true)
      }}
      onDragOver={(e) => {
        if (!hasFiles(e)) return
        e.preventDefault()
        e.dataTransfer.dropEffect = "copy"
      }}
      onDragLeave={(e) => {
        if (!hasFiles(e)) return
        depth.current = Math.max(0, depth.current - 1)
        if (depth.current === 0) setDragging(false)
      }}
      onDrop={(e) => {
        if (!hasFiles(e)) return
        e.preventDefault()
        depth.current = 0
        setDragging(false)
        const files = Array.from(e.dataTransfer.files)
        if (files.length) onAddFiles(files)
      }}
    >
      {children(control)}
      {dragging && (
        <div className="composer-drop-hint" aria-hidden>
          Drop to add to this matter
        </div>
      )}
    </div>
  )
}
