import { useEffect, useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from "react"
import "./ReviewAddMenu.css"

type Item = { label: string; hint: string; onSelect: () => void }

// The table's + header button: the ways to add questions, as a disclosure of
// buttons. Arrows move between them; Escape and every choice return focus to
// the button, so a dialog opened from here hands focus back to it on close.
export default function ReviewAddMenu({ disabled, items }: { disabled: boolean; items: Item[] }) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const id = useId()
  const [place, setPlace] = useState<{ top: number; left: number }>()
  // Focus the first item once the menu is shown, when opened from the keyboard's arrows.
  const enter = useRef(false)

  // Fixed, so the table's scroll area cannot clip it: under the button, its
  // right edge on the button's, kept inside the viewport. It follows the button
  // while the table scrolls, closes once the button leaves the table's view, or on resize.
  const placePanel = () => {
    if (!panel.current || !trigger.current) return
    const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1
    const anchor = trigger.current.getBoundingClientRect()
    const view = trigger.current.closest(".review-scroll")?.getBoundingClientRect()
    if (view && (anchor.right < view.left || anchor.left > view.right || anchor.bottom < view.top || anchor.top > view.bottom)) return setOpen(false)
    const width = panel.current.getBoundingClientRect().width
    const wide = document.documentElement.clientWidth
    setPlace({ top: (anchor.bottom + 6) / zoom, left: Math.max(16, Math.min(anchor.right - width, wide - 16 - width)) / zoom })
  }
  useLayoutEffect(() => {
    if (!open) return
    placePanel()
    if (enter.current) panel.current?.querySelector<HTMLButtonElement>("button")?.focus({ preventScroll: true })
    enter.current = false
  }, [open])

  useEffect(() => {
    if (!open) return
    const away = (event: PointerEvent) => { if (!root.current?.contains(event.target as Node)) setOpen(false) }
    const moved = (event: Event) => { if (!panel.current?.contains(event.target as Node)) placePanel() }
    const resized = () => setOpen(false)
    document.addEventListener("pointerdown", away)
    window.addEventListener("scroll", moved, true)
    window.addEventListener("resize", resized)
    return () => {
      document.removeEventListener("pointerdown", away)
      window.removeEventListener("scroll", moved, true)
      window.removeEventListener("resize", resized)
    }
  }, [open])

  useEffect(() => { if (disabled) setOpen(false) }, [disabled])

  function keys(event: KeyboardEvent) {
    if (event.key === "Escape" && open) {
      event.preventDefault()
      event.stopPropagation()
      setOpen(false)
      trigger.current?.focus()
    } else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      if (!open) {
        if (event.target !== trigger.current || disabled) return
        event.preventDefault()
        enter.current = true
        setOpen(true)
        return
      }
      const buttons = [...panel.current!.querySelectorAll<HTMLButtonElement>("button")]
      const at = buttons.indexOf(document.activeElement as HTMLButtonElement)
      event.preventDefault()
      buttons[at < 0 ? (event.key === "ArrowDown" ? 0 : buttons.length - 1) : (at + (event.key === "ArrowDown" ? 1 : buttons.length - 1)) % buttons.length]?.focus()
    }
  }

  return (
    <div ref={root} className="review-add" onKeyDown={keys} onBlur={(event) => {
      const next = event.relatedTarget as Node | null
      if (open && next && !root.current?.contains(next)) setOpen(false)
    }}>
      <button ref={trigger} className="review-add-button" type="button" disabled={disabled} aria-label="Ways to add questions" title="Ways to add questions" aria-expanded={open} aria-controls={id} onClick={() => setOpen((value) => !value)}>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>
      </button>
      <div id={id} ref={panel} className="review-add-menu" role="group" aria-label="Add questions" hidden={!open} style={place}>
        {items.map((item, index) => (
          <button key={item.label} type="button" aria-describedby={`${id}-${index}`} onClick={() => {
            setOpen(false)
            trigger.current?.focus()
            item.onSelect()
          }}>
            <span>{item.label}</span>
            <small id={`${id}-${index}`} aria-hidden="true">{item.hint}</small>
          </button>
        ))}
      </div>
    </div>
  )
}
