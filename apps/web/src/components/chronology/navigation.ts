import { useSyncExternalStore } from "react"
import type { Chronology } from "../../api/ingest/chronology"
import { chronologyCounters } from "../../api/ingest/chronology"
const listeners = new Set<() => void>()
let state: { matterId?: string; visible: boolean; attention: boolean; updates: number } = {
  visible: false,
  attention: false,
  updates: 0,
}
export function setChronologyNavigation(matterId?: string, chronology?: Chronology) {
  const counters = chronology && chronologyCounters(chronology)
  state = {
    matterId,
    visible: !!chronology && (chronology.created || chronology.suggestion.status === "relevant"),
    attention: !!chronology && !chronology.created && chronology.suggestion.status === "relevant",
    updates: counters?.updates ?? 0,
  }
  listeners.forEach((listener) => listener())
}
export function useChronologyNavigation() {
  return useSyncExternalStore(
    (listener) => {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    () => state,
  )
}
