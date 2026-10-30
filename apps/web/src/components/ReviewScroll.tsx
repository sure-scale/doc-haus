import { useLayoutEffect, useRef, type ReactNode } from "react"

// The review table's scroll region. While columns lie beyond its right edge the
// edge fades; once scrolled sideways the sticky Document column casts a shadow.
// Both are attributes set on scroll and resize, so nothing re-renders or shifts.
export default function ReviewScroll({ children }: { children: ReactNode }) {
  const frame = useRef<HTMLDivElement>(null)
  const scroller = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    const outer = frame.current
    const view = scroller.current
    if (!outer || !view) return
    let queued = 0
    const update = () => {
      queued = 0
      const left = Math.abs(view.scrollLeft)
      outer.toggleAttribute("data-more-right", view.scrollWidth - view.clientWidth - left > 1)
      outer.toggleAttribute("data-scrolled", left > 1)
      // The fade keeps clear of classic scrollbars and matches the header's colour.
      outer.style.setProperty("--scrollbar-y", `${view.offsetWidth - view.clientWidth - 2 * view.clientLeft}px`)
      outer.style.setProperty("--scrollbar-x", `${view.offsetHeight - view.clientHeight - 2 * view.clientTop}px`)
      outer.style.setProperty("--head-h", `${view.querySelector("thead")?.offsetHeight ?? 0}px`)
    }
    const schedule = () => { if (!queued) queued = requestAnimationFrame(update) }
    update()
    view.addEventListener("scroll", schedule, { passive: true })
    const observer = new ResizeObserver(schedule)
    observer.observe(view)
    if (view.firstElementChild) observer.observe(view.firstElementChild)
    return () => {
      view.removeEventListener("scroll", schedule)
      observer.disconnect()
      cancelAnimationFrame(queued)
    }
  }, [])
  return (
    <div ref={frame} className="review-scroll-frame">
      <div ref={scroller} className="review-scroll" role="region" aria-label="Review table" tabIndex={0}>{children}</div>
    </div>
  )
}
