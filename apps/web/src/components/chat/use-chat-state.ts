import { useRef, useState } from "react"
import type { Part } from "../../api/session-view"
import type { PermissionRequest } from "../../api/opencode/permissions"
import type { Turn } from "./turn-model"

export function useChatState(sessionID: string | undefined) {
  const [turns, setTurns] = useState<Turn[]>([])
  const [input, setInput] = useState("")
  const [busy, setBusy] = useState(false)
  // Reopening a past conversation refetches its messages, leaving turns empty for
  // a beat. Without this the empty-starter state ("Ask a question...") flashes in
  // that gap; gate it on !loading and show a skeleton instead. A fresh chat (no
  // sessionID) has nothing to load, so it stays false and the starters show now.
  const [loading, setLoading] = useState(Boolean(sessionID))
  const [loadError, setLoadError] = useState<string>()
  // Edit-tool calls the engine parked awaiting the user's approval (see the
  // ctx.ask gates in dochaus/tool/*). Each renders an approval card above the
  // composer; the turn stays "working" until every request is answered.
  const [permissions, setPermissions] = useState<PermissionRequest[]>([])
  const [permissionError, setPermissionError] = useState<string>()
  const [editing, setEditing] = useState<{ index: number; draft: string } | null>(null)
  // While Auto is selected, the real agent the last message routed to — drives the
  // "Auto · Redline" chip and the in-flight bubble's byline. Cleared when the user
  // leaves Auto so a stale routing never lingers on the chip.
  const [resolvedAgent, setResolvedAgent] = useState<string>()
  const [, bump] = useState(0)

  const sessionRef = useRef<string>("")
  const freshRef = useRef(false)
  const partsRef = useRef<Map<string, Part>>(new Map())
  const rolesRef = useRef<Map<string, string>>(new Map())
  // Child subagent sessions a task step spawned this turn. Their parts stream on
  // their own sessionID, so we admit those into partsRef to nest under the task
  // step (see partsToSteps). Reset alongside partsRef on each new turn.
  const childRef = useRef<Set<string>>(new Set())
  // Provider failures can arrive only as session.error (with no completed
  // assistant message). Hold that event until finalize reloads stored history,
  // then attach it to the failed turn so a later session.idle cannot erase it.
  const sessionErrorRef = useRef<string | undefined>(undefined)
  const finalizingRef = useRef<Promise<void> | undefined>(undefined)
  const malformedToolRecoveryRef = useRef(0)
  const lastActivityAtRef = useRef(Date.now())
  const stallAbortRef = useRef(false)
  const permissionsRef = useRef<PermissionRequest[]>([])
  const finalizeRetryRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  // Message ids already folded into the settled `turns`. The live view aggregates
  // every part in partsRef indiscriminately, so it assumes partsRef holds only the
  // in-flight turn. Two things break that: trailing SSE part.updated events from a
  // just-finalized turn landing after the clear, and resync pulling full session
  // history. Both re-admit prior-turn parts, making the live bubble flash the
  // previous answer. Reject any part whose message is already settled to keep
  // partsRef scoped to the current turn.
  const seenRef = useRef<Set<string>>(new Set())
  const logRef = useRef<HTMLDivElement>(null)
  // Whether the log should stay pinned to the bottom as new content streams in.
  // Scrolling up beyond a small threshold releases the pin; a new send restores it.
  const stickRef = useRef(true)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  // Mirrors `busy` for the long-lived event subscription, whose resync closure is
  // created once and would otherwise capture a stale value.
  const busyRef = useRef(false)
  busyRef.current = busy

  function updatePermissions(
    update: PermissionRequest[] | ((current: PermissionRequest[]) => PermissionRequest[]),
  ) {
    const next = typeof update === "function" ? update(permissionsRef.current) : update
    permissionsRef.current = next
    setPermissions(next)
  }

  return {
    turns,
    setTurns,
    input,
    setInput,
    busy,
    setBusy,
    loading,
    setLoading,
    loadError,
    setLoadError,
    permissions,
    permissionError,
    setPermissionError,
    editing,
    setEditing,
    resolvedAgent,
    setResolvedAgent,
    bump,
    sessionRef,
    freshRef,
    partsRef,
    rolesRef,
    childRef,
    sessionErrorRef,
    finalizingRef,
    malformedToolRecoveryRef,
    lastActivityAtRef,
    stallAbortRef,
    permissionsRef,
    finalizeRetryRef,
    seenRef,
    logRef,
    stickRef,
    inputRef,
    busyRef,
    updatePermissions,
  }
}

export type ChatState = ReturnType<typeof useChatState>
