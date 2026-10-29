import { useEffect, useId, useRef, useState, type ReactNode } from "react"
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
