import { useEffect, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"
import type { Jurisdiction } from "../api/ingest/matters"

// A matter can be governed by more than one jurisdiction (a cross-border deal
// touches each party's law), so the picker is a multi-select: a trigger that
// summarises the chosen packs and opens a searchable checklist. With ~30 packs a
// native <select multiple> is unusable, hence the filtered popover.
export default function JurisdictionSelect({
  jurisdictions,
  selected,
  onChange,
  align = "left",
  trigger,
  id,
  "aria-describedby": ariaDescribedBy,
}: {
  jurisdictions: Jurisdiction[]
  selected: string[]
  onChange: (codes: string[]) => void
  align?: "left" | "right"
  // When given, replaces the default dropdown trigger with custom content (e.g.
  // jurisdiction badges) wrapped in a bare button that still opens the popover.
  trigger?: ReactNode
  id?: string
  "aria-describedby"?: string
}) {
  const [open, setOpen] = useState(false)
  const [filter, setFilter] = useState("")
  // Anchor position captured when the popover opens. The popover renders in a
  // fixed portal so scrolling bodies can't clip it. Inside a native modal,
  // keep the portal in that dialog's top layer so it remains interactive.
  const [pos, setPos] = useState({ top: 0, left: 0, right: 0 })
  const ref = useRef<HTMLDivElement>(null)
  const popRef = useRef<HTMLDivElement>(null)

  // Close on outside click so the popover behaves like the native select it
  // replaces, and on outside scroll since the fixed-position anchor goes stale
  // when the surrounding container scrolls. Bound only while open.
  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      const t = e.target as Node
      if (ref.current?.contains(t) || popRef.current?.contains(t)) return
      setOpen(false)
    }
    const onScroll = (e: Event) => {
      if (popRef.current?.contains(e.target as Node)) return
      setOpen(false)
    }
    window.addEventListener("mousedown", onDown)
    window.addEventListener("scroll", onScroll, true)
    return () => {
      window.removeEventListener("mousedown", onDown)
      window.removeEventListener("scroll", onScroll, true)
    }
  }, [open])

  function toggleOpen() {
    if (!open && ref.current) {
      const r = ref.current.getBoundingClientRect()
      setPos({ top: r.bottom + 6, left: r.left, right: window.innerWidth - r.right })
    }
    setOpen((o) => !o)
  }

  function toggle(code: string) {
    onChange(selected.includes(code) ? selected.filter((c) => c !== code) : [...selected, code])
  }

  const term = filter.trim().toLowerCase()
  const visible = jurisdictions.filter(
    (j) => !term || j.name.toLowerCase().includes(term) || j.code.toLowerCase().includes(term),
  )
  const label =
    selected.length === 0
      ? "Jurisdiction (optional)"
      : selected.length === 1
        ? jurisdictions.find((j) => j.code === selected[0])?.name ?? selected[0]
        : `${selected.length} jurisdictions`

  return (
    <div className="jx-select" ref={ref}>
      {trigger ? (
        <button id={id} aria-describedby={ariaDescribedBy} type="button" className="jx-trigger-bare" onClick={toggleOpen} title="Click to edit jurisdiction">
          {trigger}
        </button>
      ) : (
        <button id={id} aria-describedby={ariaDescribedBy} type="button" className={`assistant-trigger jx-trigger${selected.length === 0 ? " is-empty" : ""}`} onClick={toggleOpen} title="Jurisdiction — steers reasoning and citation style">
          <span className="assistant-trigger-label">{label}</span>
          <svg className={`assistant-caret${open ? " open" : ""}`} width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
        </button>
      )}
      {open &&
        createPortal(
          <div
            ref={popRef}
            className="jx-popover"
            onKeyDown={(event) => {
              if (event.key !== "Escape") return
              event.preventDefault()
              event.stopPropagation()
              setOpen(false)
              ref.current?.querySelector("button")?.focus({ preventScroll: true })
            }}
            style={align === "right" ? { top: pos.top, right: pos.right } : { top: pos.top, left: pos.left }}
          >
            <input
              className="jx-search"
              autoFocus
              placeholder="Search jurisdictions"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
            />
            <div className="jx-list">
              {visible.length === 0 ? (
                <p className="muted jx-empty">No match.</p>
              ) : (
                visible.map((j) => (
                  <label key={j.code} className={`jx-option${selected.includes(j.code) ? " selected" : ""}`}>
                    <input type="checkbox" checked={selected.includes(j.code)} onChange={() => toggle(j.code)} />
                    <span className="jx-name">{j.name}</span>
                    <span className="jx-code">{j.code}</span>
                  </label>
                ))
              )}
            </div>
          </div>,
          ref.current?.closest("dialog") ?? document.body,
        )}
    </div>
  )
}
