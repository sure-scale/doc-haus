import { useRef, useState } from "react"
import { humanizeError } from "../../api/ingest/errors"
import {
  answerLeanQuestion,
  pauseLeanDraftingRun,
  resumeLeanDraftingRun,
  refineLeanChecklist,
  stopLeanDraftingRun,
  updateLeanDraftingRun,
  type LeanDraftingStatus,
} from "../../api/ingest/lean"
import { runState } from "./model"

// Actions queue work for the run's coordinator (or pause or stop it) and show
// its returned status. A synchronous guard prevents a second submission before
// React renders busy.
export default function useRunActions(
  matterId: string | undefined,
  status: LeanDraftingStatus | undefined,
  onStatus: (status: LeanDraftingStatus) => void,
) {
  const submitting = useRef(false)
  const [busy, setBusy] = useState(false)
  // Pause and Stop settle the run; neither reads as resuming.
  const [halting, setHalting] = useState<"pause" | "stop">()
  const [error, setError] = useState("")
  // The run's updated_at when a request was accepted: until the run moves on
  // from it, a stopped run reads as resuming rather than paused.
  const [resumedFrom, setResumedFrom] = useState<string>()

  async function queue(work: () => Promise<{ status: LeanDraftingStatus }>, halt?: "pause" | "stop") {
    if (submitting.current) return false
    submitting.current = true
    setBusy(true)
    setHalting(halt)
    setError("")
    try {
      const next = (await work()).status
      setResumedFrom(halt ? undefined : next.updated_at)
      onStatus(next)
      return true
    } catch (e) {
      setError(humanizeError(e, "Could not reach the drafting service. Try again."))
      return false
    } finally {
      submitting.current = false
      setBusy(false)
      setHalting(undefined)
    }
  }

  const state = status ? runState(status) : undefined
  return {
    busy,
    error,
    pausing: halting === "pause",
    stopping: halting === "stop",
    resuming:
      (busy && !halting) || (status !== undefined && resumedFrom === status.updated_at && state !== "running" && state !== "done" && state !== "input"),
    update: (text: string) =>
      matterId && status
        ? queue(() => updateLeanDraftingRun(matterId, status.owner_session_id, text))
        : Promise.resolve(false),
    refine: (parentID: string) =>
      matterId && status
        ? queue(() => refineLeanChecklist(matterId, status.owner_session_id, parentID))
        : Promise.resolve(false),
    resume: () => (matterId ? queue(() => resumeLeanDraftingRun(matterId)) : Promise.resolve(false)),
    pause: () =>
      matterId && status
        ? queue(() => pauseLeanDraftingRun(matterId, status.owner_session_id), "pause")
        : Promise.resolve(false),
    stop: () =>
      matterId && status
        ? queue(() => stopLeanDraftingRun(matterId, status.owner_session_id), "stop")
        : Promise.resolve(false),
    answer: (questionID: string, answer: string) =>
      matterId ? queue(() => answerLeanQuestion(matterId, questionID, answer)) : Promise.resolve(false),
  }
}

export type RunActions = ReturnType<typeof useRunActions>
