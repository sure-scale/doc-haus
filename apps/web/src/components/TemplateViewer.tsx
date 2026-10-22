import { useContext, useEffect, useRef, useState, type ReactNode } from "react"
import { Closing } from "./useExit"
import { createPortal } from "react-dom"
import { createViewer, type DocxViewer } from "docxodus/embed"
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
  footer,
  onRendered,
}: {
  name: string
  placeholders: TemplatePlaceholder[]
  onClose: () => void
  // Draft previews use immutable snapshot/proposal bytes; library templates
  // keep the normal name-based reader.
  loadBytes?: () => Promise<Uint8Array>
  sidebar?: ReactNode
  footer?: ReactNode
  onRendered?: () => void
}) {
  const closing = useContext(Closing)
  const panel = useRef<HTMLDivElement>(null)
  useModalFocus(panel)
  const documentHost = useRef<HTMLDivElement>(null)
  const [ready, setReady] = useState(false)
  const [error, setError] = useState<string>()
  const rendered = useRef(onRendered)
  rendered.current = onRendered
  // Optional-clause markers are keep-or-omit decisions, not fills — list them
  // separately under their short names.
  const fills = placeholders.filter((p) => !p.text.startsWith("[optional:"))
  const optional = placeholders.filter((p) => p.text.startsWith("[optional:"))

  useEffect(() => {
    let cancelled = false
    let viewer: DocxViewer | undefined
    // Render offscreen first. The embed scopes Word's body/span CSS before it
    // enters the page; raw converter HTML would restyle the surrounding app.
    const mount = document.createElement("div")
    setReady(false)
    setError(undefined)
    documentHost.current?.replaceChildren()
    ;(loadBytes ? loadBytes() : fetchTemplateBytes(name))
      .then((bytes) => createViewer(mount, bytes, { wasmBasePath: "/wasm/", renderHeadersAndFooters: true }))
      .then((result) => {
        if (cancelled) { result.destroy(); return }
        viewer = result
        documentHost.current?.replaceChildren(mount)
        setReady(true)
        rendered.current?.()
      })
      .catch((e) => !cancelled && setError(e instanceof Error ? e.message : String(e)))
    return () => {
      cancelled = true
      viewer?.destroy()
    }
  }, [name, loadBytes])

  // Keep viewport overlays outside the page's width/spacing rules.
  return createPortal(
    <div className={`viewer-overlay template-viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
      <div ref={panel} className={`viewer-panel template-reader${footer ? " has-footer" : ""}`} role="dialog" aria-modal="true" aria-label={name} onClick={(e) => e.stopPropagation()} onKeyDown={(event) => {
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
            {error && <p className="muted">{error}</p>}
            {!error && !ready && (
              <div className="docx-render docx-skeleton" aria-label="Rendering document">
                {Array.from({ length: 14 }, (_, i) => <div key={i} className="sk-line" />)}
              </div>
            )}
            <div ref={documentHost} className="docx-render" hidden={!ready} />
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
        {footer && <div className="template-reader-footer">{footer}</div>}
      </div>
    </div>,
    document.body,
  )
}
