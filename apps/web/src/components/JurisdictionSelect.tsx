import { useEffect, useId, useLayoutEffect, useRef, useState, type ReactNode } from "react"
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
  const popoverId = useId()
  // A fixed portal escapes scrolling containers; a native modal must own the
  // portal so its checklist remains in the interactive top layer.
  const [pos, setPos] = useState({ top: 0, left: 0, width: 280, maxHeight: 360 })
  const ref = useRef<HTMLDivElement>(null)
  const popRef = useRef<HTMLDivElement>(null)

  useLayoutEffect(() => {
    if (!open) return
    const position = () => {
      const anchor = ref.current?.querySelector("button")
      const popover = popRef.current
      if (!anchor || !popover) return
      // Rects use viewport pixels, while fixed offsets inherit the appearance
      // preference's body zoom. Convert once, including the visible viewport
      // when a mobile keyboard reduces the available space.
      const zoom = Number.parseFloat(getComputedStyle(document.body).zoom) || 1
      const viewport = window.visualViewport
      const left = (viewport?.offsetLeft ?? 0) / zoom
      const top = (viewport?.offsetTop ?? 0) / zoom
      const width = (viewport?.width ?? window.innerWidth) / zoom
      const height = (viewport?.height ?? window.innerHeight) / zoom
      const inset = 8
      const gap = 6
      const popupWidth = Math.min(280, Math.max(0, width - inset * 2))
      // Measure wrapped option text at the final width before choosing a side.
      popover.style.width = `${popupWidth}px`
      const search = popover.querySelector<HTMLInputElement>(".jx-search")!
      const list = popover.querySelector<HTMLElement>(".jx-list")!
      const border = getComputedStyle(popover)
      const naturalHeight = search.offsetHeight + Math.min(list.scrollHeight, Number.parseFloat(getComputedStyle(list).maxHeight))
        + Number.parseFloat(border.borderTopWidth) + Number.parseFloat(border.borderBottomWidth)
      const rect = anchor.getBoundingClientRect()
      const above = Math.max(0, Math.min(rect.top / zoom - gap, top + height - inset) - top - inset)
      const below = Math.max(0, top + height - inset - Math.max(rect.bottom / zoom + gap, top + inset))
      const openAbove = below < naturalHeight && above > below
      const maxHeight = openAbove ? above : below
      const popupHeight = Math.min(naturalHeight, maxHeight)
      const preferredLeft = align === "right" ? rect.right / zoom - popupWidth : rect.left / zoom
      const preferredTop = openAbove ? rect.top / zoom - gap - popupHeight : rect.bottom / zoom + gap
      setPos({
        left: Math.max(left + inset, Math.min(preferredLeft, left + width - inset - popupWidth)),
        top: Math.max(top + inset, Math.min(preferredTop, top + height - inset - popupHeight)),
        width: popupWidth,
        maxHeight,
      })
    }
    position()
    window.addEventListener("resize", position)
    window.visualViewport?.addEventListener("resize", position)
    window.visualViewport?.addEventListener("scroll", position)
    return () => {
      window.removeEventListener("resize", position)
      window.visualViewport?.removeEventListener("resize", position)
      window.visualViewport?.removeEventListener("scroll", position)
    }
  }, [open, align, filter, jurisdictions.length])

  // Close on outside click so the popover behaves like the native select it
  // replaces, and on outside scroll since the fixed-position anchor goes stale
  // when the container scrolls. Viewport changes reposition it above instead,
  // so opening a mobile keyboard does not dismiss the search field.
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
        <button id={id} aria-describedby={ariaDescribedBy} aria-expanded={open} aria-controls={open ? popoverId : undefined} type="button" className="jx-trigger-bare" onClick={toggleOpen} title="Click to edit jurisdiction">
          {trigger}
        </button>
      ) : (
        <button id={id} aria-describedby={ariaDescribedBy} aria-expanded={open} aria-controls={open ? popoverId : undefined} type="button" className={`assistant-trigger jx-trigger${selected.length === 0 ? " is-empty" : ""}`} onClick={toggleOpen} title="Jurisdiction — steers reasoning and citation style">
          <span className="assistant-trigger-label">{label}</span>
          <svg className={`assistant-caret${open ? " open" : ""}`} width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
        </button>
      )}
      {open &&
        createPortal(
          <div
            ref={popRef}
            id={popoverId}
            className="jx-popover"
            role="group"
            aria-label="Jurisdictions"
            onKeyDown={(event) => {
              if (event.key !== "Escape") return
              event.preventDefault()
              event.stopPropagation()
              setOpen(false)
              ref.current?.querySelector("button")?.focus({ preventScroll: true })
            }}
            style={pos}
          >
            <input
              className="jx-search"
              autoFocus
              aria-label="Search jurisdictions"
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
