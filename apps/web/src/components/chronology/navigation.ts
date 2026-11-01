import { useSyncExternalStore } from "react"
import type { Chronology } from "../../api/ingest/chronology"
import { chronologyCounters } from "../../api/ingest/chronology"
const listeners = new Set<() => void>()
let state: { matterId?: string; visible: boolean; attention: boolean; updates: number } = {
  visible: false,
  attention: false,
  updates: 0,
}
/** `suggest` is the lawyer's "Suggest a chronology" switch: when off, a saved
 * relevance result neither reveals the rail entry nor marks it. */
export function setChronologyNavigation(matterId?: string, chronology?: Chronology, suggest = true) {
  const counters = chronology && chronologyCounters(chronology)
  const suggested = suggest && !!chronology && !chronology.created && chronology.suggestion.status === "relevant"
  state = {
    matterId,
    visible: !!chronology && (chronology.created || suggested),
    attention: suggested,
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
