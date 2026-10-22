import { useRef, useState, type ReactNode } from "react"

// The one file drop target: drag files onto it, click it, or focus it and press
// Enter/Space to browse. Used for matter documents, templates and skills so all
// three look and behave alike.
export default function Dropzone({
  accept,
  title,
  hint,
  busy,
  onFiles,
  compact,
}: {
  accept: string
  title: ReactNode
  hint?: ReactNode
  busy?: boolean
  onFiles: (files: File[]) => void
  compact?: boolean
}) {
  const input = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)

  return (
    <>
      <div
        className={`dropzone${dragging ? " dragover" : ""}${busy ? " is-busy" : ""}${compact ? " is-compact" : ""}`}
        role="button"
        tabIndex={0}
        aria-busy={busy || undefined}
        onClick={() => input.current?.click()}
        onKeyDown={(e) => {
          if (e.key !== "Enter" && e.key !== " ") return
          e.preventDefault()
          input.current?.click()
        }}
        onDragOver={(e) => {
          e.preventDefault()
          setDragging(true)
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault()
          setDragging(false)
          const files = Array.from(e.dataTransfer.files)
          if (files.length) onFiles(files)
        }}
      >
        <IconUpload />
        <span className="dropzone-title">{title}</span>
        {hint && <span className="dropzone-hint">{hint}</span>}
      </div>
      <input
        ref={input}
        type="file"
        accept={accept}
        multiple
        hidden
        onChange={(e) => {
          const files = Array.from(e.target.files ?? [])
          if (files.length) onFiles(files)
          e.target.value = ""
        }}
      />
    </>
  )
}

function IconUpload() {
  return (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="17 8 12 3 7 8" />
      <line x1="12" y1="3" x2="12" y2="15" />
    </svg>
  )
}
