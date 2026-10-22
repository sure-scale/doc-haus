import { createContext, useLayoutEffect, useState, type ReactNode } from "react"

// Keeps a closing overlay mounted while it plays its `is-closing` exit. Pass
// what opens it (a flag, or the item it shows): `shown` follows it at once on
// open and one exit later on close, keeping the last item so the exit still
// has content; `closing` is true in between. Reopening mid-exit cancels the
// exit, and a repeated close changes nothing. A timer, not animationend, ends
// the exit, so it can never stick open; reduced motion unmounts before paint.
// Only the picture lingers: a closing overlay goes inert and hands focus back
// as the exit starts (a native dialog closes then; see useModalFocus).
export default function useExit<T>(value: T): [shown: T, closing: boolean] {
  const [shown, setShown] = useState(value)
  if (value && value !== shown) setShown(value)
  const closing = !value && Boolean(shown)
  useLayoutEffect(() => {
    if (!closing) return
    if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) return setShown(value)
    const timer = window.setTimeout(() => setShown(value), exitDuration())
    return () => window.clearTimeout(timer)
  }, [closing, value])
  return [value || shown, closing]
}

// True inside an overlay that is playing its exit.
export const Closing = createContext(false)

// `{when && <Overlay />}` with an exit: renders `children` with the open item,
// and keeps the last one through the exit with `Closing` set.
export function Exit<T>({ when, children }: { when: T; children: (shown: NonNullable<T>, closing: boolean) => ReactNode }) {
  const [shown, closing] = useExit(when)
  return shown ? <Closing value={closing}>{children(shown as NonNullable<T>, closing)}</Closing> : null
}

// For a dialog that stays mounted and closes at once: true from a close until
// it opens again, so one that was never open carries no exit state.
export function useClosed(open: boolean) {
  const [opened, setOpened] = useState(open)
  if (open && !opened) setOpened(true)
  return opened && !open
}

// Exits run on the fast effects spring; its duration token times the unmount.
function exitDuration() {
  return parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--spring-effects-fast-ms")) || 150
}
