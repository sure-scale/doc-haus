import { useContext, useEffect, useRef, useState, type ReactNode, type RefObject } from "react"
import { Closing } from "./useExit"
import { createPortal } from "react-dom"
import { createViewer, type DocxViewer } from "docxodus/embed"
import { fetchTemplateBytes, type TemplatePlaceholder } from "../api/ingest/library"
import { ModalClose } from "./ModalHeader"
import { decorateMarkers, markerEntries, type MarkerEntry } from "./document-markers"
import backdropClose from "./backdropClose"
import useModalFocus from "./useModalFocus"
import { removeUnsafeLinkTargets } from "./docx-links"
import { withSafeStyleValues } from "./docx-style-values"

// In-browser template viewer. Renders a library template's .docx to HTML with the
// WASM runtime and lists the markers it carries, so the lawyer reads
// the template body without the matter/redline machinery the document viewer
// carries — a template is a clean drafting base, never redlined. Bytes convert
// client-side, so the file never leaves the browser for a third-party service.
// The side panel lists exactly the chips decorateMarkers put in the document,
// so its counts and the document always agree. A library template's chips are
// the placeholders the server's scan listed, the set drafting fills.
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
  // The server's placeholder scan of a library template; without it the
  // viewer marks what the shared marker patterns find.
  placeholders?: TemplatePlaceholder[]
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
  const fields = useRef(placeholders)
  fields.current = placeholders
  const [markers, setMarkers] = useState<MarkerEntry[]>([])
  // Optional-clause markers are keep-or-omit decisions, not fills — list them
  // separately under their short names.
  const isOptional = (entry: MarkerEntry) => /^\[optional:/i.test(entry.text)
  const fills = markers.filter((entry) => entry.kind === "placeholder" && !isOptional(entry))
  const optional = markers.filter(isOptional)
  const redactions = markers.filter((entry) => entry.kind === "redaction")

  useEffect(() => {
    let cancelled = false
    let viewer: DocxViewer | undefined
    // Render offscreen first. The embed scopes Word's body/span CSS before it
    // enters the page; raw converter HTML would restyle the surrounding app.
    const mount = document.createElement("div")
    setReady(false)
    setError(undefined)
    setMarkers([])
    documentHost.current?.replaceChildren()
    ;(loadBytes ? loadBytes() : fetchTemplateBytes(name))
      .then((bytes) => createViewer(mount, withSafeStyleValues(bytes), { wasmBasePath: "/wasm/", renderHeadersAndFooters: true }))
      .then((result) => {
        if (cancelled) { result.destroy(); return }
        viewer = result
        removeUnsafeLinkTargets(mount)
        documentHost.current?.replaceChildren(mount)
        decorateMarkers(mount, fields.current?.map((p) => p.text))
        setMarkers(markerEntries(mount))
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
            {sidebar ?? (ready ? <MarkerPanel document={documentHost} groups={[
              { entries: fills, noun: "placeholder", empty: "This template exposes no placeholders." },
              { entries: optional, noun: "optional clause", label: (text) => text.slice("[optional:".length, -1).trim() },
              { entries: redactions, noun: "redaction" },
            ]} /> : !error && <p className="muted">Finding placeholders…</p>)}
          </aside>
        </div>
        {footer && <div className="template-reader-footer">{footer}</div>}
      </div>
    </div>,
    document.body,
  )
}

type MarkerGroup = { entries: MarkerEntry[]; noun: string; empty?: string; label?: (text: string) => string }
const LINKED = "is-linked"
const FLASH = "is-flash"

// The marker list beside the document. An entry and its chips highlight
// together on hover or keyboard focus, from either side. Clicking an entry
// scrolls to its next occurrence, so repeated clicks walk through a marker
// used several times; the count shows which one is in view.
function MarkerPanel({ document: host, groups }: { document: RefObject<HTMLDivElement | null>; groups: MarkerGroup[] }) {
  const [linked, setLinked] = useState<MarkerEntry>()
  const [shown, setShown] = useState(new Map<MarkerEntry, number>())
  const flash = useRef<number>(undefined)
  // Whether the highlight came from a chip: leaving the document clears only
  // that one, never an entry the panel just took over or holds by focus.
  const fromDocument = useRef(false)
  const link = (entry: MarkerEntry | undefined) => { fromDocument.current = false; setLinked(entry) }
  const all = useRef<MarkerEntry[]>([])
  all.current = groups.flatMap((group) => group.entries)
  const chips = (entry: MarkerEntry, id?: string) =>
    [...host.current?.querySelectorAll<HTMLElement>("span.doc-marker[data-marker-id]") ?? []].filter((span) => id ? span.dataset.markerId === id : entry.ids.includes(span.dataset.markerId!))

  useEffect(() => {
    if (!linked) return
    const spans = chips(linked)
    spans.forEach((span) => span.classList.add(LINKED))
    return () => spans.forEach((span) => span.classList.remove(LINKED))
  }, [linked])

  // Chips are decorated outside React; listen on the document host.
  useEffect(() => {
    const root = host.current
    if (!root) return
    const entryFor = (target: EventTarget | null) => {
      const id = (target as Element | null)?.closest?.<HTMLElement>("span.doc-marker[data-marker-id]")?.dataset.markerId
      return id ? all.current.find((entry) => entry.ids.includes(id)) : undefined
    }
    const over = (event: MouseEvent) => {
      const entry = entryFor(event.target)
      if (!entry && !fromDocument.current) return
      fromDocument.current = Boolean(entry)
      setLinked(entry)
    }
    const leave = () => {
      if (fromDocument.current) link(undefined)
    }
    root.addEventListener("mouseover", over)
    root.addEventListener("mouseleave", leave)
    return () => {
      root.removeEventListener("mouseover", over)
      root.removeEventListener("mouseleave", leave)
    }
  }, [host])

  useEffect(() => () => window.clearTimeout(flash.current), [])

  function reveal(entry: MarkerEntry) {
    const next = ((shown.get(entry) ?? -1) + 1) % entry.ids.length
    setShown(new Map(shown).set(entry, next))
    const spans = chips(entry, entry.ids[next])
    const still = window.matchMedia("(prefers-reduced-motion: reduce)").matches
    spans[0]?.scrollIntoView({ block: "center", behavior: still ? "auto" : "smooth" })
    host.current?.querySelectorAll(`.${FLASH}`).forEach((span) => span.classList.remove(FLASH))
    spans.forEach((span) => span.classList.add(FLASH))
    window.clearTimeout(flash.current)
    flash.current = window.setTimeout(() => spans.forEach((span) => span.classList.remove(FLASH)), 1600)
  }

  return <>
    {groups.map(({ entries, noun, empty, label }) => {
      if (!entries.length && !empty) return null
      const count = entries.reduce((total, entry) => total + entry.ids.length, 0)
      return <section key={noun} aria-label={`${noun}s`}>
        <div className="changes-head"><span>{count} {noun}{count === 1 ? "" : "s"}</span></div>
        {entries.length === 0 ? <p className="muted">{empty}</p> : <ul className="template-placeholders">
          {entries.map((entry) => {
            const at = shown.get(entry)
            const many = entry.ids.length > 1
            return <li key={entry.text}>
              <button
                type="button"
                className={`marker-entry doc-marker-${entry.kind}${linked === entry ? ` ${LINKED}` : ""}`}
                aria-label={`${entry.text}, ${noun}${many ? `, ${entry.ids.length} occurrences` : ""}. Show in document`}
                title={many ? "Show the next occurrence in the document" : "Show in the document"}
                onMouseEnter={() => link(entry)}
                onMouseLeave={() => link(undefined)}
                onFocus={() => link(entry)}
                onBlur={() => link(undefined)}
                onClick={() => reveal(entry)}
              >
                <span className="marker-entry-text">{label ? label(entry.text) : entry.text}</span>
                {many && <span className="marker-entry-count" aria-hidden="true">{at === undefined ? `×${entry.ids.length}` : `${at + 1}/${entry.ids.length}`}</span>}
              </button>
            </li>
          })}
        </ul>}
      </section>
    })}
  </>
}
