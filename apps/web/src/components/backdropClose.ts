import type { MouseEvent, PointerEvent } from "react"

// Close on a backdrop click only when the press started and ended there. A press
// that crosses the panel edge (selecting text, overshooting the release) makes
// the browser fire `click` on the shared ancestor, which looks like a backdrop
// click. A press on the element's own scrollbar, a non-primary button or a
// cancelled press never counts. Nested overlays share the recorded targets,
// each matching only its own.
let down: EventTarget | null = null
let up: EventTarget | null = null
// Outer overlays see the same bubbling pointerdown; a rejected one stays rejected.
let rejected: Event | null = null
// A scrollbar press targets the scrolling element itself, like a backdrop press,
// but lands inside its border box and outside its client (padding) area.
function onScrollbar(event: PointerEvent<HTMLElement>) {
  const el = event.currentTarget
  if (event.target !== el) return false
  const box = el.getBoundingClientRect()
  const x = event.clientX - box.left - el.clientLeft
  const y = event.clientY - box.top - el.clientTop
  const inside = event.clientX >= box.left && event.clientX <= box.right && event.clientY >= box.top && event.clientY <= box.bottom
  return inside && (x < 0 || x >= el.clientWidth || y < 0 || y >= el.clientHeight)
}
export default function backdropClose(close: (event: MouseEvent<HTMLElement>) => void) {
  return {
    onPointerDown: (event: PointerEvent<HTMLElement>) => {
      if (event.nativeEvent === rejected) return
      if (event.button !== 0 || onScrollbar(event)) { down = up = null; rejected = event.nativeEvent; return }
      down = event.target; up = null
    },
    onPointerUp: (event: PointerEvent<HTMLElement>) => { up = event.target },
    onPointerCancel: () => { down = up = null },
    onClick: (event: MouseEvent<HTMLElement>) => {
      const pressed = down === event.currentTarget && up === event.currentTarget
      down = up = null
      if (pressed && event.target === event.currentTarget) close(event)
    },
  }
}
