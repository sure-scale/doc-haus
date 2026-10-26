import { cloneElement, isValidElement, useId, useLayoutEffect, useRef, useState, type ReactElement, type ReactNode } from "react"
import { createPortal } from "react-dom"

// Portaled help stays clear of scroll clipping and works with keyboard focus.
// `describes`: the bubble names itself as the child's description; off where
// the child already carries the same words in its own aria-describedby.
export function Tooltip({ label, children, describes = true }: { label: string; children: ReactNode; describes?: boolean }) {
  const host = useRef<HTMLSpanElement>(null)
  const bubble = useRef<HTMLSpanElement>(null)
  const id = useId()
  const pointerWithin = useRef(false)
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
    cancelDismissal()
    // Allow the pointer to cross the gap into the portaled help. Focus keeps
    // keyboard help available until blur, even after the pointer moves away.
    dismissal.current = setTimeout(() => {
      if (!host.current?.contains(document.activeElement)) hide()
    }, 200)
  }
  const show = () => {
    cancelDismissal()
    const element = host.current?.firstElementChild
    if (element) setAnchor(element.getBoundingClientRect())
  }
  useLayoutEffect(() => () => cancelDismissal(), [])
  useLayoutEffect(() => {
    if (!anchor || !bubble.current) return
    const rect = bubble.current.getBoundingClientRect()
    const scale = Number(getComputedStyle(document.body).zoom) || 1
    const right = anchor.right + 8
    const left = anchor.left - rect.width - 8
    const beside = right + rect.width < innerWidth - 12 ? right : left >= 12 ? left : undefined
    // Too wide a control to sit beside (a full-width button on a phone): above
    // it, or below when there is no room above, so it never covers the control.
    const above = anchor.top - rect.height - 8
    const x = beside ?? anchor.left + anchor.width / 2 - rect.width / 2
    const y = beside !== undefined ? anchor.top + anchor.height / 2 - rect.height / 2 : above >= 12 ? above : anchor.bottom + 8
    setPosition({ x: Math.max(12, Math.min(x, innerWidth - rect.width - 12)) / scale, y: Math.max(12, Math.min(y, innerHeight - rect.height - 12)) / scale })
    const onEscape = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) return
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
  }, [anchor])
  const child = isValidElement(children) ? children as ReactElement<{ "aria-describedby"?: string }> : null
  // Over rather than enter: a control rendered under a pointer that has not
  // moved (Save again after saving) gets a mouseover from the control it
  // replaced and no mouseout, so React never reports an enter.
  return (
    <span ref={host} style={{ display: "contents" }} onMouseOver={() => { pointerWithin.current = true; show() }} onMouseLeave={leave} onFocus={show} onBlur={() => { if (!pointerWithin.current) hide() }}>
      {child && describes ? cloneElement(child, { "aria-describedby": [child.props["aria-describedby"], anchor ? id : undefined].filter(Boolean).join(" ") || undefined }) : children}
      {anchor && createPortal(<span ref={bubble} id={id} role="tooltip" className="tip-bubble" onMouseEnter={() => { pointerWithin.current = true; cancelDismissal() }} onMouseLeave={leave} style={{ left: position.x, top: position.y, transform: "none", width: "max-content" }}>{label}</span>, host.current?.closest("dialog[open]") ?? document.body)}
    </span>
  )
}
