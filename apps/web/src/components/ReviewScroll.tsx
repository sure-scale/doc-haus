import { useLayoutEffect, useRef, type ReactNode } from "react"
import { Tooltip } from "./Tooltip"

// The review table's scroll region. While columns lie beyond its right edge the
// edge fades and a round › shows more questions; once scrolled sideways the
// sticky Document column casts a shadow and a ‹ shows earlier ones. Each step
// is one column. All are attributes set on scroll and resize, so nothing
// re-renders or shifts.
export default function ReviewScroll({ children }: { children: ReactNode }) {
  const frame = useRef<HTMLDivElement>(null)
  const scroller = useRef<HTMLDivElement>(null)
  const previous = useRef<HTMLButtonElement>(null)
  const next = useRef<HTMLButtonElement>(null)
  useLayoutEffect(() => {
    const outer = frame.current
    const view = scroller.current
    if (!outer || !view) return
    let queued = 0
    const update = () => {
      queued = 0
      const left = Math.abs(view.scrollLeft)
      const more = view.scrollWidth - view.clientWidth - left > 1
      outer.toggleAttribute("data-more-right", more)
      outer.toggleAttribute("data-scrolled", left > 1)
      // A step button that hides hands focus to the other one, else to the table.
      for (const [button, shown, other] of [[next.current, more, previous.current], [previous.current, left > 1, next.current]] as const) {
        if (!button || button.hidden === !shown) continue
        const focused = button === document.activeElement
        button.hidden = !shown
        if (focused && !shown) (other && !other.hidden ? other : view).focus()
      }
      // The fade keeps clear of classic scrollbars and matches the header's colour.
      outer.style.setProperty("--scrollbar-y", `${view.offsetWidth - view.clientWidth - 2 * view.clientLeft}px`)
      outer.style.setProperty("--scrollbar-x", `${view.offsetHeight - view.clientHeight - 2 * view.clientTop}px`)
      outer.style.setProperty("--head-h", `${view.querySelector("thead")?.offsetHeight ?? 0}px`)
      outer.style.setProperty("--doc-w", `${view.querySelector<HTMLElement>("thead th")?.offsetWidth ?? 0}px`)
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
  // Brings the next column hidden past either edge in beside the sticky Document column.
  const step = (direction: 1 | -1) => {
    const view = scroller.current
    const heads = [...(view?.querySelectorAll<HTMLElement>("thead th") ?? [])]
    if (!view || heads.length < 2) return
    const edge = view.scrollLeft + heads[0].offsetWidth
    const starts = heads.slice(1).map((head) => head.offsetLeft)
    const target = direction > 0 ? starts.find((start) => start > edge + 1) : starts.findLast((start) => start < edge - 1)
    if (target === undefined) return
    const smooth = !matchMedia("(prefers-reduced-motion: reduce)").matches
    view.scrollTo({ left: target - heads[0].offsetWidth, behavior: smooth ? "smooth" : "auto" })
  }
  const chevron = (path: string) => <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={path} /></svg>
  return (
    <div ref={frame} className="review-scroll-frame">
      <div ref={scroller} className="review-scroll" role="region" aria-label="Review table" tabIndex={0}>{children}</div>
      <Tooltip label="Show earlier questions" describes={false}>
        <button ref={previous} type="button" className="review-scroll-step is-previous" aria-label="Show earlier questions" hidden onClick={() => step(-1)}>{chevron("m15 6-6 6 6 6")}</button>
      </Tooltip>
      <Tooltip label="Show more questions" describes={false}>
        <button ref={next} type="button" className="review-scroll-step is-next" aria-label="Show more questions" hidden onClick={() => step(1)}>{chevron("m9 6 6 6-6 6")}</button>
      </Tooltip>
    </div>
  )
}
