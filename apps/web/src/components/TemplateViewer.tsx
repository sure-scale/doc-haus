import { useEffect, useState } from "react"
import { useDocxodus } from "docxodus/react"
import { fetchTemplateBytes, type TemplatePlaceholder } from "../api/ingest/library"
import { ModalClose } from "./ModalHeader"
import backdropClose from "./backdropClose"

// In-browser template viewer. Renders a library template's .docx to HTML with the
// WASM runtime and lists the placeholders a draft would fill, so the lawyer reads
// the template body without the matter/redline machinery the document viewer
// carries — a template is a clean drafting base, never redlined. Bytes convert
// client-side, so the file never leaves the browser for a third-party service.
export default function TemplateViewer({
  name,
  placeholders,
  onClose,
}: {
  name: string
  placeholders: TemplatePlaceholder[]
  onClose: () => void
}) {
  const { isReady, error: wasmError, convertToHtml } = useDocxodus("/wasm/")
  const [html, setHtml] = useState<string>()
  const [error, setError] = useState<string>()
  // Optional-clause markers are keep-or-omit decisions, not fills — list them
  // separately under their short names.
  const fills = placeholders.filter((p) => !p.text.startsWith("[optional:"))
  const optional = placeholders.filter((p) => p.text.startsWith("[optional:"))

  useEffect(() => {
    if (!isReady) return
    let cancelled = false
    setHtml(undefined)
    setError(undefined)
    fetchTemplateBytes(name)
      .then(async (bytes) => {
        const out = await convertToHtml(bytes, { renderHeadersAndFooters: true })
        if (!cancelled) setHtml(out)
      })
      .catch((e) => !cancelled && setError(e instanceof Error ? e.message : String(e)))
    return () => {
      cancelled = true
    }
  }, [isReady, name])

  return (
    <div className="viewer-overlay" {...backdropClose(() => onClose())}>
      <div className="viewer-panel" onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <span className="viewer-title">
            <span className="viewer-doctype">DOCX</span>
            <span className="viewer-title-name">{name}</span>
          </span>
          <div className="viewer-bar-actions">
            <ModalClose onClose={onClose} />
          </div>
        </div>
        <div className="viewer-body">
          <div className="viewer-doc">
            {wasmError && <p className="muted">Viewer failed to load: {wasmError.message}</p>}
            {error && <p className="muted">{error}</p>}
            {!wasmError && !error && !html && (
              <div className="docx-render docx-skeleton" aria-label="Rendering document">
                {Array.from({ length: 14 }, (_, i) => <div key={i} className="sk-line" />)}
              </div>
            )}
            {html && <div className="docx-render" dangerouslySetInnerHTML={{ __html: html }} />}
          </div>
          <aside className="changes-panel">
            <div className="changes-head">
              <span>
                {fills.length} placeholder{fills.length === 1 ? "" : "s"}
              </span>
            </div>
            {fills.length === 0 ? (
              <p className="muted">This template exposes no placeholders.</p>
            ) : (
              <ul className="template-placeholders">
                {fills.map((p) => (
                  <li key={p.text}>{p.text}</li>
                ))}
              </ul>
            )}
            {optional.length > 0 && (
              <>
                <div className="changes-head">
                  <span>
                    {optional.length} optional clause{optional.length === 1 ? "" : "s"}
                  </span>
                </div>
                <ul className="template-placeholders">
                  {optional.map((p) => (
                    <li key={p.text}>{p.text.slice("[optional:".length, -1).trim()}</li>
                  ))}
                </ul>
              </>
            )}
          </aside>
        </div>
      </div>
    </div>
  )
}
