import { useEffect } from "react"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import { openQuestions, runState } from "./model"

// While the run waits on the lawyer's answers the browser tab says so, so a
// lawyer working in another tab sees it; the title returns once they are sent.
export default function useNeedsYouTitle(status: LeanDraftingStatus | undefined) {
  const asks = status && runState(status) !== "running" ? openQuestions(status).length : 0
  useEffect(() => {
    if (asks === 0) return
    const prior = document.title
    document.title = `Needs you (${asks}) · ${prior}`
    return () => {
      document.title = prior
    }
  }, [asks])
}
