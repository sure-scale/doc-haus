import type { Client } from "../../api/opencode/client"
import { sendMalformedToolRecovery } from "../../api/opencode/recovery"
import { getSessionStatus } from "../../api/opencode/sessions"
import { toTurns } from "./turn-model"
import { loadConversation, bounded } from "./conversation"
import { malformedToolRecoveryCount, malformedToolRecoveryDue } from "./turn-recovery"
import { draftingTurnRecoveryMode } from "./drafting-recovery"
import type { ChatState } from "./use-chat-state"

export function useTurnSettlement({
  state,
  client,
  matterName,
  matterId,
  directory,
  onSessionCreated,
}: {
  state: ChatState
  client: Client
  matterName: string
  matterId?: string
  directory: string
  onSessionCreated?: (id: string) => void
}) {
  const {
    setTurns,
    setBusy,
    setPermissionError,
    sessionRef,
    freshRef,
    partsRef,
    rolesRef,
    childRef,
    sessionErrorRef,
    finalizingRef,
    malformedToolRecoveryRef,
    stallAbortRef,
    permissionsRef,
    finalizeRetryRef,
    seenRef,
    busyRef,
    turnRef,
    updatePermissions,
  } = state

  function clearFinalizeRetry() {
    if (!finalizeRetryRef.current) return
    clearTimeout(finalizeRetryRef.current)
    finalizeRetryRef.current = undefined
  }

  // An admitted recovery runs as a new execution, whose success settles the
  // turn. Settle here only once the engine reports the session idle; while the
  // recovery still runs, leave it to that success or the resync poll.
  function scheduleFinalizeRetry() {
    clearFinalizeRetry()
    const session = sessionRef.current
    finalizeRetryRef.current = setTimeout(async () => {
      finalizeRetryRef.current = undefined
      if (!busyRef.current || sessionRef.current !== session) return
      const status = await bounded(getSessionStatus(client, session), 5_000, "Session status").catch(() => undefined)
      if (status?.type === "idle" && busyRef.current && sessionRef.current === session) finalize()
    }, 2_000)
  }

  // Reload the settled history from the server rather than appending the
  // in-flight turn from refs. The reload gives every turn its authoritative
  // message id, which edit/retry need to revert the session to a given turn.
  function finalize() {
    if (finalizingRef.current) return finalizingRef.current
    const pending = finalizeOnce()
    finalizingRef.current = pending.finally(() => {
      finalizingRef.current = undefined
    })
    return finalizingRef.current
  }

  async function finalizeOnce() {
    if (sessionRef.current) {
      // If this reload fails the turn is not lost: leave busy set and the refs
      // intact so the 15s resync poll retries finalize, instead of clearing the
      // live view and hanging on a half-settled turn.
      const conversation = await bounded(loadConversation(client, sessionRef.current), 10_000, "Conversation finalization").catch(() => undefined)
      if (!conversation) return
      const { msgs, pool } = conversation
      const settled = toTurns(msgs, matterName, pool)
      const last = msgs[msgs.length - 1]
      malformedToolRecoveryRef.current = malformedToolRecoveryCount(msgs)
      const pendingSessionError = sessionErrorRef.current
      const malformedProgressToken = malformedToolRecoveryDue(msgs, matterId)
      // A stopped turn ends here: no recovery, and approvals it was waiting on
      // ended with it.
      const stopped = turnRef.current.stopped
      const recoverableToolCall =
        !stopped &&
        !sessionErrorRef.current &&
        permissionsRef.current.length === 0 &&
        draftingTurnRecoveryMode(malformedProgressToken) === "malformed-tool"
      if (msgs.length) setTurns(settled)
      for (const m of msgs) seenRef.current.add(m.info.id)
      persistFreshSessionRoute()
      if (permissionsRef.current.length && !stopped) return
      if (recoverableToolCall) {
        partsRef.current.clear()
        rolesRef.current.clear()
        childRef.current.clear()
        sessionErrorRef.current = undefined
        stallAbortRef.current = false
        const result = await sendMalformedToolRecovery(directory, sessionRef.current, {
          kind: "malformed-tool",
          runId: sessionRef.current,
          progressToken: malformedProgressToken!,
          messageID: `msg_recovery_${last!.info.id}`,
        }).catch((error) => {
          failTurn(error instanceof Error ? error.message : "Could not resume the interrupted workflow.")
          return undefined
        })
        if (!result) return
        if (result.admitted || (result.persisted && !result.executing))
          malformedToolRecoveryRef.current += 1
        scheduleFinalizeRetry()
        return
      }
      if (pendingSessionError) {
        setTurns([
          ...settled,
          {
            role: "assistant",
            text: "",
            citations: [],
            redlines: [],
            drafts: [],
            todos: [],
            steps: [],
            error: pendingSessionError,
          },
        ])
      }
    }
    settleDraftingTurn()
  }

  function settleDraftingTurn() {
    clearFinalizeRetry()
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    updatePermissions([])
    setPermissionError(undefined)
    busyRef.current = false
    setBusy(false)
    malformedToolRecoveryRef.current = 0
    stallAbortRef.current = false
    persistFreshSessionRoute()
  }

  function persistFreshSessionRoute() {
    if (!freshRef.current) return
    freshRef.current = false
    onSessionCreated?.(sessionRef.current)
  }

  // A send failed before the turn ever reached a settled state — routing, lazy
  // session creation, the revert, or the prompt request itself rejected. This is
  // common when the engine is briefly unreachable right after navigating back
  // into a matter. No session.idle will come to clear busy, so without this the
  // bubble hangs on "Thinking..." forever (and if session creation was what
  // failed, sessionRef is empty so the resync poll cannot recover it either).
  // Surface the failure as an error turn and settle so the composer reopens.
  function failTurn(message: string) {
    clearFinalizeRetry()
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    setTurns((prev) => [
      ...prev,
      { role: "assistant", text: "", citations: [], redlines: [], drafts: [], todos: [], steps: [], error: message },
    ])
    busyRef.current = false
    setBusy(false)
    stallAbortRef.current = false
  }

  return {
    clearFinalizeRetry,
    scheduleFinalizeRetry,
    finalize,
    failTurn,
  }
}

export type TurnSettlement = ReturnType<typeof useTurnSettlement>
