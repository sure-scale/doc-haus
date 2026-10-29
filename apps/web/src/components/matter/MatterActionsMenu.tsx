import { useEffect, useId, useLayoutEffect, useRef, useState, type ReactNode } from "react"
import "./matter-actions-menu.css"

export type MatterMenuAction = { id: string; label: string; onSelect: () => void; disabled?: boolean }

/** One disclosure for matter actions and optional editable matter context. */
export default function MatterActionsMenu({
  actions,
  children,
}: {
  actions: MatterMenuAction[]
  children?: ReactNode
}) {
  const id = useId()
  const disclosure = useRef<HTMLDetailsElement>(null)
  const [open, setOpen] = useState(false)
  useLayoutEffect(() => {
    if (!open) return
    const menu = disclosure.current
    const popup = menu?.querySelector<HTMLElement>(".matter-actions-menu-popup")
    if (!menu || !popup) return
    const position = () => {
      const scale = Number.parseFloat(getComputedStyle(document.body).zoom) || 1
      const viewport = window.visualViewport
      const top = viewport?.offsetTop ?? 0
      const bottom = (viewport?.offsetTop ?? 0) + (viewport?.height ?? window.innerHeight)
      const anchor = menu.querySelector("summary")!.getBoundingClientRect()
      const above = Math.max(0, (anchor.top - top - 19) / scale)
      const below = Math.max(0, (bottom - anchor.bottom - 19) / scale)
      const openAbove = below < Math.min(280, popup.scrollHeight) && above > below
      popup.dataset.side = openAbove ? "above" : "below"
      popup.style.setProperty("--matter-menu-available", `${openAbove ? above : below}px`)
    }
    position()
    const observer = new ResizeObserver(position)
    observer.observe(menu)
    window.addEventListener("resize", position)
    window.addEventListener("scroll", position, true)
    window.visualViewport?.addEventListener("resize", position)
    window.visualViewport?.addEventListener("scroll", position)
    return () => {
      observer.disconnect()
      window.removeEventListener("resize", position)
      window.removeEventListener("scroll", position, true)
      window.visualViewport?.removeEventListener("resize", position)
      window.visualViewport?.removeEventListener("scroll", position)
    }
  }, [open])
  useEffect(() => {
    if (!open) return
    const closeOutside = (event: PointerEvent) => {
      const target = event.target
      if (!(target instanceof Element) || disclosure.current?.contains(target)) return
      if (target.closest(".jx-popover, .client-suggestions, dialog, [role='dialog'], .viewer-overlay")) return
      if (disclosure.current) disclosure.current.open = false
    }
    document.addEventListener("pointerdown", closeOutside)
    return () => document.removeEventListener("pointerdown", closeOutside)
  }, [open])
  return (
    <details
      ref={disclosure}
      onToggle={(event) => setOpen(event.currentTarget.open)}
      className="matter-actions-menu"
      onKeyDown={(event) => {
        if (event.key !== "Escape" || event.defaultPrevented) return
        event.preventDefault()
        const menu = event.currentTarget
        // Let an edited field cancel before moving focus, so blur cannot save it.
        requestAnimationFrame(() => {
          if (!menu.isConnected) return
          menu.open = false
          menu.querySelector("summary")?.focus({ preventScroll: true })
        })
      }}
    >
      <summary aria-label="Matter actions" aria-controls={id} title="Matter actions">
        ···
      </summary>
      <div id={id} className="matter-actions-menu-popup">
        {actions.length > 0 && (
          <div className="matter-actions-menu-items">
            {actions.map((action) => (
              <button
                key={action.id}
                type="button"
                disabled={action.disabled}
                onClick={(event) => {
                  const menu = event.currentTarget.closest("details")
                  menu?.removeAttribute("open")
                  menu?.querySelector("summary")?.focus({ preventScroll: true })
                  action.onSelect()
                }}
              >
                {action.label}
              </button>
            ))}
          </div>
        )}
        {children}
      </div>
    </details>
  )
}
