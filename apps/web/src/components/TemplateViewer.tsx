import { useContext, useEffect, useRef, useState, type ReactNode } from "react"
import { Closing } from "./useExit"
import { useDocxodus } from "docxodus/react"
import { fetchTemplateBytes, type TemplatePlaceholder } from "../api/ingest/library"
import { ModalClose } from "./ModalHeader"
import backdropClose from "./backdropClose"
import useModalFocus from "./useModalFocus"

// In-browser template viewer. Renders a library template's .docx to HTML with the
// WASM runtime and lists the placeholders a draft would fill, so the lawyer reads
// the template body without the matter/redline machinery the document viewer
// carries — a template is a clean drafting base, never redlined. Bytes convert
// client-side, so the file never leaves the browser for a third-party service.
export default function TemplateViewer({
  name,
  placeholders,
  onClose,
  loadBytes,
  sidebar,
  onRendered,
}: {
  name: string
  placeholders: TemplatePlaceholder[]
  onClose: () => void
  // Draft previews use immutable snapshot/proposal bytes; library templates
  // keep the normal name-based reader.
  loadBytes?: () => Promise<Uint8Array>
  sidebar?: ReactNode
  onRendered?: () => void
}) {
  const closing = useContext(Closing)
  const panel = useRef<HTMLDivElement>(null)
  useModalFocus(panel)
  const { isReady, error: wasmError, convertToHtml } = useDocxodus("/wasm/")
  const [html, setHtml] = useState<string>()
  const [error, setError] = useState<string>()
  const rendered = useRef(onRendered)
  rendered.current = onRendered
  // Optional-clause markers are keep-or-omit decisions, not fills — list them
  // separately under their short names.
  const fills = placeholders.filter((p) => !p.text.startsWith("[optional:"))
  const optional = placeholders.filter((p) => p.text.startsWith("[optional:"))

  useEffect(() => {
    if (!isReady) return
    let cancelled = false
    setHtml(undefined)
    setError(undefined)
    ;(loadBytes ? loadBytes() : fetchTemplateBytes(name))
      .then(async (bytes) => {
        const out = await convertToHtml(bytes, { renderHeadersAndFooters: true })
        if (!cancelled) {
          setHtml(out)
          rendered.current?.()
        }
      })
      .catch((e) => !cancelled && setError(e instanceof Error ? e.message : String(e)))
    return () => {
      cancelled = true
    }
  }, [isReady, name, loadBytes])

  return (
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
      <div ref={panel} className="viewer-panel template-reader" role="dialog" aria-modal="true" aria-label={name} onClick={(e) => e.stopPropagation()} onKeyDown={(event) => {
        if (event.key === "Escape" && !event.defaultPrevented) { event.stopPropagation(); onClose() }
      }}>
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
            {sidebar ?? <>
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
            </>}
          </aside>
        </div>
      </div>
    </div>
  )
}
