import { useEffect, useRef } from "react"
import type { ClauseHistoryRequest } from "../../../../dochaus/lib/clause-history-contract"
import { mountClauseHistory } from "../../../../dochaus/lib/clause-history-view"
import { fetchClauseHistory, fetchClauseHistorySource } from "../api/ingest/clause-history"
import "../../../../dochaus/lib/clause-history-view.css"
import "./ClauseHistory.css"

export type ClauseHistoryProps = {
  matterId: string
  name: string
  request: ClauseHistoryRequest
  onClose?: () => void
  onOpenVersion?: (versionId: string) => void
}

/** Mount beside the existing document. Opening history does not replace the
 * document, restore a version, or change its selection. */
export function ClauseHistory({ matterId, name, request, onClose, onOpenVersion }: ClauseHistoryProps) {
  const root = useRef<HTMLElement>(null)
  const callbacks = useRef({ onClose, onOpenVersion }); callbacks.current = { onClose, onOpenVersion }
  const requestKey = JSON.stringify(request)
  useEffect(() => {
    if (!root.current) return
    const input = JSON.parse(requestKey) as ClauseHistoryRequest
    const history = mountClauseHistory(root.current, {
      load: signal => fetchClauseHistory(matterId, name, input, signal),
      source: (source, signal) => fetchClauseHistorySource(matterId, name, source, signal),
      onClose: onClose ? () => callbacks.current.onClose?.() : undefined,
      onOpenVersion: onOpenVersion ? id => callbacks.current.onOpenVersion?.(id) : undefined,
    })
    history.focus()
    return () => history.dispose()
  }, [matterId, name, requestKey, Boolean(onClose), Boolean(onOpenVersion)])
  return <section ref={root} className="web-clause-history" aria-label="Clause history" />
}
