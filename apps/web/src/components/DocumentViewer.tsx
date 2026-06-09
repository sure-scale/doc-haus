import { useEffect, useState } from "react"
import { useDocxodus } from "docxodus/react"
import { CommentRenderMode } from "docxodus"
import { fetchDocumentBytes } from "../api/ingest"

// In-browser redline viewer. Fetches a matter's .docx from ingest and converts
// it to HTML with the WASM runtime — tracked changes render as insertions and
// deletions, comments in a margin column. The bytes are converted client-side, so
// the document never leaves the browser for a third-party service.
export default function DocumentViewer({ matterId, name, onClose }: { matterId: string; name: string; onClose: () => void }) {
  const { isReady, error: wasmError, convertToHtml } = useDocxodus("/wasm/")
  const [html, setHtml] = useState<string>()
  const [error, setError] = useState<string>()

  useEffect(() => {
    if (!isReady) return
    let cancelled = false
    setHtml(undefined)
    setError(undefined)
    fetchDocumentBytes(matterId, name)
      .then((bytes) =>
        convertToHtml(bytes, {
          renderTrackedChanges: true,
          showDeletedContent: true,
          renderMoveOperations: true,
          renderHeadersAndFooters: true,
          commentRenderMode: CommentRenderMode.Margin,
        }),
      )
      .then((out) => !cancelled && setHtml(out))
      .catch((e) => !cancelled && setError(e instanceof Error ? e.message : String(e)))
    return () => {
      cancelled = true
    }
  }, [isReady, matterId, name])

  return (
    <div className="viewer-overlay" onClick={onClose}>
      <div className="viewer-panel" onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <span className="viewer-title">{name}</span>
          <button onClick={onClose}>Close</button>
        </div>
        <div className="viewer-body">
          {wasmError && <p className="muted">Viewer failed to load: {wasmError.message}</p>}
          {error && <p className="muted">{error}</p>}
          {!wasmError && !error && !html && (
            <p className="muted">{isReady ? `Rendering ${name}...` : "Loading viewer..."}</p>
          )}
          {html && <div className="docx-render" dangerouslySetInnerHTML={{ __html: html }} />}
        </div>
      </div>
    </div>
  )
}
