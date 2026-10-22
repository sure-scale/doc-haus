import { useContext, useLayoutEffect, type RefObject } from "react"
import { Closing } from "./useExit"

// Div-based dialogs and responsive sheets share native-dialog keyboard behavior.
// A stack prevents a lower overlay stealing focus from a nested modal or portal.
// A closing overlay lets go as its exit starts, so focus returns at once.
const layers: HTMLElement[] = []
const targets = 'button:not(:disabled), a[href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), summary, iframe, [tabindex]:not([tabindex="-1"])'
export default function useModalFocus(ref: RefObject<HTMLElement | null>, open = true) {
  const closing = useContext(Closing)
  const active = open && !closing
  useLayoutEffect(() => {
    const panel = ref.current
    if (!active || !panel) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    layers.push(panel)
    const focusables = () => [...panel.querySelectorAll<HTMLElement>(targets)].filter(el => !el.closest('[inert]') && el.getClientRects().length > 0 && getComputedStyle(el).visibility !== "hidden")
    const initial = panel.querySelector<HTMLElement>('[autofocus]') ?? focusables()[0] ?? panel
    const previousTabIndex = panel.getAttribute("tabindex")
    if (previousTabIndex === null) panel.tabIndex = -1
    // A frame owns its keyboard events. Keep a DOM focus destination after the
    // panel so Tab from the frame's final control cannot skip to browser chrome.
    const endGuard = document.createElement("span")
    endGuard.tabIndex = 0
    endGuard.dataset.modalFocusGuard = ""
    endGuard.style.cssText = "position:fixed;width:1px;height:1px;padding:0;overflow:hidden;clip-path:inset(50%);opacity:0;pointer-events:none"
    panel.after(endGuard)
    initial.focus({ preventScroll: true })
    function ownsFocus() {
      if (layers.at(-1) !== panel) return false
      // Focus identifies the native top layer even when native dialogs nest.
      const native = document.activeElement?.closest<HTMLDialogElement>('dialog[open]')
        ?? [...document.querySelectorAll<HTMLDialogElement>('dialog[open]')].at(-1)
      return !native || native.contains(panel)
    }
    function keydown(event: KeyboardEvent) {
      if (event.key !== "Tab" || !ownsFocus()) return
      const elements = focusables()
      const first = elements[0], last = elements.at(-1)
      if (!first || !last) { event.preventDefault(); panel!.focus(); return }
      if (event.shiftKey && (document.activeElement === first || !panel!.contains(document.activeElement))) {
        event.preventDefault(); last.focus()
      } else if (!event.shiftKey && (document.activeElement === last || !panel!.contains(document.activeElement))) {
        event.preventDefault(); first.focus()
      }
    }
    function focusin(event: FocusEvent) {
      if (!ownsFocus() || panel!.contains(event.target as Node)) return
      // Keyboard events inside an embedded PDF or document do not bubble to
      // this document. Contain the focus when the browser leaves that frame.
      const target = focusables()[0] ?? panel!
      target.focus({ preventScroll: true })
    }
    document.addEventListener("keydown", keydown)
    document.addEventListener("focusin", focusin)
    return () => {
      const index = layers.lastIndexOf(panel)
      if (index >= 0) layers.splice(index, 1)
      document.removeEventListener("keydown", keydown)
      document.removeEventListener("focusin", focusin)
      endGuard.remove()
      if (previousTabIndex === null) panel.removeAttribute("tabindex")
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [ref, active])
}
