import { readBrowserValue, writeBrowserValue } from "../../browser-storage"
import { useSyncExternalStore } from "react"

// The matter's documents drawer, opened from the rail and from the composer.
// Both read one open flag (and the open matter's document count, for the rail
// badge) so either control reflects the other. Open persists per browser like
// the rail's collapse; narrow windows start closed because there the drawer
// covers the canvas.
const KEY = "dh.docs-drawer"
const listeners = new Set<() => void>()
let state = {
  open: readBrowserValue(KEY) === "1" && !window.matchMedia("(max-width: 1199px)").matches,
  count: undefined as number | undefined,
}

export function setDocumentsDrawer(open: boolean) {
  writeBrowserValue(KEY, open ? "1" : "0")
  emit({ ...state, open })
}

export function toggleDocumentsDrawer() {
  setDocumentsDrawer(!state.open)
}

export function setDocumentsCount(count: number | undefined) {
  if (count !== state.count) emit({ ...state, count })
}

export function useDocumentsDrawer() {
  return useSyncExternalStore(subscribe, () => state)
}

function emit(next: typeof state) {
  state = next
  listeners.forEach((listener) => listener())
}

function subscribe(listener: () => void) {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}
