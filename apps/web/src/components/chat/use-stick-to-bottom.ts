import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type RefObject,
  type UIEvent,
  type WheelEvent,
} from "react"

// The chat log follows new content while the lawyer is reading at the bottom,
// and stops following the moment they scroll up or start inspecting a step.
// Any change inside the log (a streamed token, a narration line, a ticking
// timer) re-pins it, so children need not re-render the panel to keep it live.
// It also reports which edges have content beyond them, for the edge fades, and
// whether newer content sits below the view, for the Jump to latest chip.

type View = { away: boolean; top: boolean; bottom: boolean }

export function useStickToBottom(logRef: RefObject<HTMLDivElement | null>, stickRef: RefObject<boolean>) {
  const [view, setView] = useState<View>({ away: false, top: false, bottom: false })
  const lastTop = useRef(0)

  const report = useCallback(() => {
    const el = logRef.current
    if (!el) return
    lastTop.current = el.scrollTop
    const gap = el.scrollHeight - el.scrollTop - el.clientHeight
    const next = { away: !stickRef.current && gap > 24, top: el.scrollTop > 1, bottom: gap > 1 }
    setView((prev) => (prev.away === next.away && prev.top === next.top && prev.bottom === next.bottom ? prev : next))
  }, [logRef, stickRef])

  const measure = useCallback(() => {
    const el = logRef.current
    if (el && stickRef.current) el.scrollTop = el.scrollHeight
    report()
  }, [logRef, stickRef, report])

  useLayoutEffect(measure)

  useEffect(() => {
    const el = logRef.current
    if (!el) return
    const mutations = new MutationObserver(measure)
    mutations.observe(el, { childList: true, subtree: true, characterData: true })
    const resize = new ResizeObserver(measure)
    resize.observe(el)
    return () => {
      mutations.disconnect()
      resize.disconnect()
    }
  }, [logRef, measure])

  const bind = {
    onWheel: (e: WheelEvent) => {
      // Unpin before the scroll event lands, so a render in between cannot snap
      // the log back to the bottom.
      if (e.deltaY < 0) stickRef.current = false
    },
    onPointerDown: () => {
      // Opening a step or starting a text selection is an inspection gesture.
      stickRef.current = false
    },
    onScroll: (e: UIEvent<HTMLDivElement>) => {
      const el = e.currentTarget
      const gap = el.scrollHeight - el.scrollTop - el.clientHeight
      // Only arriving at the bottom re-pins: the first eased frames of a smooth
      // scroll up are still within the threshold and must not snap it back.
      if (gap <= 2 && el.scrollTop >= lastTop.current) stickRef.current = true
      if (gap > 2 && el.scrollTop < lastTop.current - 1) stickRef.current = false
      report()
    },
  }

  function jump() {
    const el = logRef.current
    if (!el) return
    stickRef.current = true
    const still = window.matchMedia("(prefers-reduced-motion: reduce)").matches
    el.scrollTo({ top: el.scrollHeight, behavior: still ? "auto" : "smooth" })
  }

  return { ...view, bind, jump }
}
