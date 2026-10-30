import { cloneElement, isValidElement, useId, useLayoutEffect, useRef, useState, type ReactElement, type ReactNode } from "react"
import { createPortal } from "react-dom"

// A press focuses what it presses, and a click handler may move focus back to
// a trigger; on a control that acts, neither is keyboard focus, so neither keeps
// help up once the pointer leaves. A tap sends the same mouse events.
let pressing = false
let watching = false
function watchPresses() {
  if (watching) return
  watching = true
  const press = () => { pressing = true; setTimeout(() => { pressing = false }) }
  document.addEventListener("mousedown", press, true)
  document.addEventListener("mouseup", press, true)
}
// A click on one of these does something, so its help has been read: it closes
// until the pointer leaves. Help-only targets (an "i", a progress bar) keep a
// tap's help up until focus leaves, and so does a control that cannot be used yet.
const ACTS = "a[href], button, input, label, select, summary, textarea, [role=button], [role=checkbox], [role=link], [role=menuitem], [role=switch], [role=tab]"
const acts = (control: Element | null | undefined) => !!control?.matches(ACTS) && control.getAttribute("aria-disabled") !== "true"

// Portaled help stays clear of scroll clipping and works with keyboard focus.
// `describes`: the bubble names itself as the child's description; off where
// the child already carries the same words in its own aria-describedby.
// `around`: the list or toolbar the control sits in; help that cannot sit
// beside the control goes above or below that whole area, never over it.
// An open disclosure's own help stays hidden, so it never covers its menu.
export function Tooltip({ label, children, describes = true, around }: { label: ReactNode; children: ReactNode; describes?: boolean; around?: string }) {
  const host = useRef<HTMLSpanElement>(null)
  const bubble = useRef<HTMLSpanElement>(null)
  const id = useId()
  const pointerWithin = useRef(false)
  // Focus that keeps help up: keyboard focus, or any focus on a help-only target.
  const held = useRef(false)
  const quiet = useRef(false)
  const dismissal = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const [anchor, setAnchor] = useState<DOMRect | null>(null)
  const [position, setPosition] = useState({ x: 12, y: 12 })
  const cancelDismissal = () => {
    clearTimeout(dismissal.current)
    dismissal.current = undefined
  }
  const hide = () => { cancelDismissal(); setAnchor(null) }
  const leave = () => {
    pointerWithin.current = false
    quiet.current = false
    cancelDismissal()
    // Allow the pointer to cross the gap into the portaled help. Held focus
    // keeps help available until blur, even after the pointer moves away.
    dismissal.current = setTimeout(() => {
      if (!held.current || !host.current?.contains(document.activeElement)) hide()
    }, 200)
  }
  const show = () => {
    cancelDismissal()
    const element = host.current?.firstElementChild
    if (element && element.getAttribute("aria-expanded") !== "true") setAnchor(element.getBoundingClientRect())
  }
  useLayoutEffect(() => { watchPresses(); return () => cancelDismissal() }, [])
  useLayoutEffect(() => {
    if (!anchor || !bubble.current) return
    const rect = bubble.current.getBoundingClientRect()
    const scale = Number(getComputedStyle(document.body).zoom) || 1
    const right = anchor.right + 8
    const left = anchor.left - rect.width - 8
    const beside = right + rect.width < innerWidth - 12 ? right : left >= 12 ? left : undefined
    // Too wide a control to sit beside (a full-width button on a phone): above
    // it, or below when there is no room above, so it never covers the control.
    const area = (around && host.current?.closest(around)?.getBoundingClientRect()) || anchor
    const above = area.top - rect.height - 8
    const x = beside ?? anchor.left + anchor.width / 2 - rect.width / 2
    const y = beside !== undefined ? anchor.top + anchor.height / 2 - rect.height / 2 : above >= 12 ? above : area.bottom + 8
    setPosition({ x: Math.max(12, Math.min(x, innerWidth - rect.width - 12)) / scale, y: Math.max(12, Math.min(y, innerHeight - rect.height - 12)) / scale })
    const onEscape = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) return
      // Help left under a newer dialog is out of sight: hide it and leave this
      // Escape to that dialog.
      const top = [...document.querySelectorAll("dialog[open]")].at(-1)
      if (top && !top.contains(host.current)) return hide()
      // Hovered help must dismiss even when keyboard focus is elsewhere, and
      // consume this Escape before a containing dialog receives it.
      event.preventDefault()
      event.stopPropagation()
      hide()
    }
    window.addEventListener("resize", hide)
    window.addEventListener("scroll", hide, true)
    document.addEventListener("keydown", onEscape, true)
    return () => { window.removeEventListener("resize", hide); window.removeEventListener("scroll", hide, true); document.removeEventListener("keydown", onEscape, true) }
  }, [anchor, around])
  const child = isValidElement(children) ? children as ReactElement<{ "aria-describedby"?: string }> : null
  // Over rather than enter: a control rendered under a pointer that has not
  // moved (Save again after saving) gets a mouseover from the control it
  // replaced and no mouseout, so React never reports an enter.
  return (
    <span
      ref={host}
      style={{ display: "contents" }}
      onMouseOver={() => { pointerWithin.current = true; if (!quiet.current) show() }}
      onMouseLeave={leave}
      onFocus={() => { held.current = !pressing || !acts(host.current?.firstElementChild); if (held.current) { quiet.current = false; show() } }}
      onBlur={() => { held.current = false; if (!pointerWithin.current) hide() }}
      onClick={(event) => {
        const control = host.current?.firstElementChild
        if (!control || !acts(control) || !control.contains(event.target as Node)) return
        quiet.current = true
        hide()
      }}
    >
      {child && describes ? cloneElement(child, { "aria-describedby": [child.props["aria-describedby"], anchor ? id : undefined].filter(Boolean).join(" ") || undefined }) : children}
      {anchor && createPortal(<span ref={bubble} id={id} role="tooltip" className="tip-bubble" onMouseEnter={() => { pointerWithin.current = true; cancelDismissal() }} onMouseLeave={leave} style={{ left: position.x, top: position.y, transform: "none", width: "max-content" }}>{label}</span>, host.current?.closest("dialog[open]") ?? document.body)}
    </span>
  )
}
