import { useEffect, useId, useLayoutEffect, useRef, useState, type ReactNode } from "react"
import "./SplitButton.css"

type Item = { label: string; icon?: ReactNode; onSelect: () => void; disabled?: boolean; title?: string }

// A leading action plus a toggle for alternatives. The menu is a plain
// disclosure (a list of buttons), like the viewer's "More actions".
export function SplitButton({ label, icon, onClick, title, items, toggleLabel, disabled, primary, className }: {
  label: ReactNode
  icon?: ReactNode
  onClick: () => void
  title?: string
  items: Item[]
  toggleLabel: string
  disabled?: boolean
  primary?: boolean
  className?: string
}) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const toggle = useRef<HTMLButtonElement>(null)
  const menu = useRef<HTMLDivElement>(null)
  const menuId = useId()
  // The menu is right-anchored; where the split sits near the left edge (the
  // viewer toolbar on phones) it would run off-screen, so nudge it back inside.
  const [shift, setShift] = useState(0)

  useLayoutEffect(() => {
    if (!open || !menu.current) return
    const node = menu.current
    node.style.right = "0px"
    node.style.animation = "none"
    const rect = node.getBoundingClientRect()
    const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1
    const edge = 8
    const wide = document.documentElement.clientWidth
    const move = rect.left < edge ? edge - rect.left : rect.right > wide - edge ? wide - edge - rect.right : 0
    node.style.right = ""
    node.style.animation = ""
    setShift(move / zoom)
  }, [open])

  useEffect(() => {
    if (!open) return
    const away = (event: PointerEvent) => {
      if (!root.current?.contains(event.target as Node)) setOpen(false)
    }
    document.addEventListener("pointerdown", away)
    return () => document.removeEventListener("pointerdown", away)
  }, [open])

  useEffect(() => {
    if (disabled) setOpen(false)
  }, [disabled])

  return (
    <div
      ref={root}
      className={`split-button${open ? " is-open" : ""}${className ? ` ${className}` : ""}`}
      onKeyDown={(event) => {
        if (event.key === "Escape" && open) {
          event.preventDefault()
          event.stopPropagation()
          setOpen(false)
          toggle.current?.focus()
        } else if ((event.key === "ArrowDown" || event.key === "ArrowUp") && open) {
          const buttons = [...root.current!.querySelectorAll<HTMLButtonElement>(".split-button-menu button:not(:disabled)")]
          const at = buttons.indexOf(document.activeElement as HTMLButtonElement)
          if (at < 0 && event.key === "ArrowUp") return
          event.preventDefault()
          buttons[(at + (event.key === "ArrowDown" ? 1 : buttons.length - 1)) % buttons.length]?.focus()
        }
      }}
      onBlur={(event) => {
        // A null relatedTarget is a mouse press on a non-focusing button
        // (Safari, Firefox); the outside-pointer handler covers real outside clicks.
        const next = event.relatedTarget as Node | null
        if (open && next && !root.current?.contains(next)) setOpen(false)
      }}
    >
      <button className={`split-button-lead${primary ? " primary" : ""}`} onClick={() => { setOpen(false); onClick() }} disabled={disabled} title={title}>
        {icon}
        {label}
      </button>
      <button
        ref={toggle}
        className={`split-button-toggle${primary ? " primary" : ""}`}
        aria-label={toggleLabel}
        aria-expanded={open}
        aria-controls={menuId}
        disabled={disabled}
        onClick={() => setOpen((value) => !value)}
      >
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="m6 9 6 6 6-6" />
        </svg>
      </button>
      <div id={menuId} className="viewer-more-menu split-button-menu" hidden={!open} ref={menu} style={shift ? { right: -shift } : undefined}>
        {items.map((item) => (
          <button
            key={item.label}
            disabled={item.disabled}
            title={item.title}
            onClick={() => {
              setOpen(false)
              toggle.current?.focus()
              item.onSelect()
            }}
          >
            {item.icon}
            <span>{item.label}</span>
          </button>
        ))}
      </div>
    </div>
  )
}
