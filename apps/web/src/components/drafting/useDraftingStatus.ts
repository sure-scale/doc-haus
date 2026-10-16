import { useCallback, useEffect, useRef, useState } from "react"
import { getDraftingStatus } from "../../api/ingest/drafting-status"
import { humanizeError } from "../../api/ingest/errors"
import type { LeanDraftingStatus } from "../../api/ingest/lean"

const POLL_MS = 5000

// The matter's drafting run, polled while the page is visible. One poller per
// matter page feeds the rail panel and the documents list. A failed poll keeps
// the last confirmed status and records when it was confirmed, so the panel can
// say how old it is instead of blanking.
export default function useDraftingStatus(matterId: string | undefined) {
  const [status, setStatus] = useState<LeanDraftingStatus>()
  const [error, setError] = useState("")
  const [checkedAt, setCheckedAt] = useState<number>()
  const [visible, setVisible] = useState(() => !document.hidden)
  const readVersion = useRef(0)
  // An answer response is newer than a poll already in flight. Never let that
  // older poll restore an unanswered question after its answer was accepted.
  const acceptStatus = useCallback((next: LeanDraftingStatus) => {
    readVersion.current += 1
    setStatus(next)
    setError("")
    setCheckedAt(Date.now())
  }, [])

  useEffect(() => {
    const onVisibility = () => setVisible(!document.hidden)
    document.addEventListener("visibilitychange", onVisibility)
    return () => document.removeEventListener("visibilitychange", onVisibility)
  }, [])

  useEffect(() => {
    setStatus(undefined)
    setError("")
    setCheckedAt(undefined)
  }, [matterId])

  useEffect(() => {
    if (!visible || !matterId) return
    let cancelled = false

    async function load() {
      const version = ++readVersion.current
      try {
        const next = await getDraftingStatus(matterId!)
        if (cancelled || version !== readVersion.current) return
        setError("")
        setStatus(next)
        setCheckedAt(Date.now())
      } catch (e) {
        if (!cancelled && version === readVersion.current) setError(humanizeError(e, "Could not load the drafting status."))
      }
    }

    load()
    const timer = window.setInterval(load, POLL_MS)
    return () => {
      cancelled = true
      window.clearInterval(timer)
    }
  }, [matterId, visible])

  return { status, error, checkedAt, setStatus: acceptStatus }
}

export type DraftingFeed = ReturnType<typeof useDraftingStatus>
