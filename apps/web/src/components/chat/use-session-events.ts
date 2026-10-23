import { useEffect, useRef, type MutableRefObject } from "react"
import type { Event, Part } from "../../api/session-view"
import { abortSession, getSessionStatus } from "../../api/opencode/sessions"
import { listPermissions, replyPermission, type PermissionRequest } from "../../api/opencode/permissions"
import type { Client } from "../../api/opencode/client"
import { subscribeEvents } from "../../api/opencode/events"
import { isAuto } from "../../agents"
import { autoApproved } from "../../prefs"
import { toTurns, errorOf, type MessageError } from "./turn-model"
import { withAcks } from "./acknowledgement"
import { loadConversation, bounded } from "./conversation"
import {
  shouldAbortStalledTurn,
  stalledAbortFailureMessage,
  malformedToolRecoveryCount,
  malformedToolRecoveryDue,
} from "./turn-recovery"
import { lastVisibleUserMessage } from "./drafting-recovery"
import type { ChatState } from "./use-chat-state"
import type { TurnSettlement } from "./use-turn-settlement"

// Tool calls that change what the hosting page lists — a new document on the
// rail, or a created/updated/deleted entry in the template, workflow, skill, or
// agent library. Their completion triggers onDocumentsChanged so the list
// refreshes mid-turn instead of waiting for a page reload.
const MUTATING_TOOLS = new Set([
  "draft-document",
  "copy-document",
  "create-template",
  "update-template",
  "prepare-template",
  "delete-template",
  "create-playbook",
  "update-playbook",
  "delete-playbook",
  "create-workflow",
  "update-workflow",
  "delete-workflow",
  "create-skill",
  "update-skill",
  "delete-skill",
  "create-agent",
  "update-agent",
  "delete-agent",
])

export function useSessionEvents({
  state,
  pendingFormsRef,
  formsRefreshRef,
  settlement,
  client,
  sessionID,
  matterId,
  matterName,
  agent,
  onAgentChange,
  directory,
  onDocumentsChanged,
  onForeignPart,
  onForeignSession,
}: {
  state: ChatState
  formsRefreshRef: MutableRefObject<() => void>
  pendingFormsRef: MutableRefObject<number>
  settlement: TurnSettlement
  client: Client
  sessionID?: string
  matterId?: string
  matterName: string
  agent: string
  onAgentChange: (name: string) => void
  directory: string
  onDocumentsChanged?: () => void
  // Parts from sessions this conversation does not own (a drafting run's stage
  // sessions). The subscription is created once, so this must be ref-stable.
  onForeignPart?: (part: Part) => void
  onForeignSession?: (sessionID: string) => boolean
}) {
  const {
    setTurns,
    setBusy,
    setLoading,
    setLoadError,
    setPermissionError,
    setResolvedAgent,
    bump,
    sessionRef,
    partsRef,
    rolesRef,
    childRef,
    sessionErrorRef,
    malformedToolRecoveryRef,
    lastActivityAtRef,
    stallAbortRef,
    permissionsRef,
    seenRef,
    busyRef,
    turnRef,
    updatePermissions,
  } = state
  const { clearFinalizeRetry, finalize, failTurn } = settlement
  const automaticAttempts = useRef(new Set<string>())
  const permissionVersion = useRef(0)
  const retiredSession = useRef<string | undefined>(undefined)

  useEffect(() => {
    const controller = new AbortController()
    let mounted = true
    automaticAttempts.current.clear()
    retiredSession.current = undefined
    permissionVersion.current += 1
    setLoadError(undefined)
    if (sessionID) {
      sessionRef.current = sessionID
      // Pending approvals are independent of active-status and transcript
      // reads. A reopened paused session must show its request immediately.
      void reconcilePermissions(controller.signal)
      setLoading(true)
      Promise.all([
        bounded(loadConversation(client, sessionID), 10_000, "Conversation load"),
        bounded(getSessionStatus(client, sessionID), 5_000, "Session status").catch(() => undefined),
      ])
        .then(([{ msgs, pool }, status]) => {
          if (!mounted || retiredSession.current === sessionID) return
          const active = status?.type === "busy" || status?.type === "retry"
          const settledMessages = active
            ? msgs.filter((message) => message.info.role !== "assistant" || Boolean(message.info.time.completed))
            : msgs
          setTurns(withAcks(toTurns(settledMessages, matterName, pool)))
          for (const m of msgs) if (m.info.role === "user" || m.info.time.completed) seenRef.current.add(m.info.id)
          for (const message of msgs.filter(
            (candidate) => candidate.info.role === "assistant" && !candidate.info.time.completed,
          )) {
            rolesRef.current.set(message.info.id, message.info.role)
            for (const part of message.parts) {
              admitPart(part)
              if (part.type === "tool" && part.tool === "task") {
                const child = (part.state as { metadata?: { sessionId?: string } }).metadata?.sessionId
                if (child) childRef.current.add(child)
              }
            }
          }
          for (const part of pool.filter((candidate) => childRef.current.has(candidate.sessionID))) admitPart(part)
          if (childRef.current.size) void reconcilePermissions(controller.signal)
          malformedToolRecoveryRef.current = malformedToolRecoveryCount(msgs)
          // Reopening a past conversation pre-selects the agent it last ran on,
          // read off the most recent user turn (the server stamps each one with
          // its agent) — but only when the user has explicitly picked an agent.
          // On Auto the stamp is just what routing chose last (and this effect
          // also fires when the first send mints the session and remounts the
          // panel), so restoring it would silently flip the picker off Auto.
          const lastUser = lastVisibleUserMessage(msgs)?.info
          if (lastUser && "agent" in lastUser && lastUser.agent) {
            if (isAuto(agent)) setResolvedAgent(lastUser.agent)
            if (!isAuto(agent)) onAgentChange(lastUser.agent)
          }
          if (active) {
            busyRef.current = true
            setBusy(true)
            lastActivityAtRef.current = Date.now()
          }
          setLoading(false)
          const lastMessage = msgs[msgs.length - 1]?.info
          // A reopened idle chat whose last turn printed a tool call settles
          // through finalize, which owns the bounded recovery.
          const idleRecoverable = !active && Boolean(malformedToolRecoveryDue(msgs, matterId))
          if (!active && ((lastMessage?.role === "assistant" && !lastMessage.time.completed) || idleRecoverable)) {
            busyRef.current = true
            setBusy(true)
            queueMicrotask(() => finalize())
          }
        })
        .catch((error) => {
          if (!mounted || retiredSession.current === sessionID) return
          setLoadError(error instanceof Error ? error.message : "Could not load this conversation.")
          setLoading(false)
        })
    }
    // No session until the first send (see onSend) — mounting the panel must not
    // mint an empty throwaway session that would clutter the conversation list.
    subscribeEvents(client, (event) => onEvent(event, controller.signal), controller.signal, {
      acceptsSession: (id) => id === sessionRef.current || childRef.current.has(id) || Boolean(onForeignSession?.(id)),
      onReconnect: () => {
        formsRefreshRef.current()
        void resync(controller.signal)
      },
      onFormsChanged: () => formsRefreshRef.current(),
    }).catch(() => {})
    // A backgrounded tab can have its SSE socket killed by the browser without
    // the stream ever ending, so the reconnect loop never fires and the turn
    // sticks on "Thinking...". Resync when the tab comes back, plus a slow poll
    // as a net for silent stalls. Approval hydration also runs for paused chats
    // whose active-status read was unavailable.
    const onVisible = () => {
      if (document.visibilityState === "visible") void resync(controller.signal)
    }
    document.addEventListener("visibilitychange", onVisible)
    window.addEventListener("focus", onVisible)
    const poll = setInterval(() => void resync(controller.signal), 15000)
    // A substantive Draft generation can legitimately be quiet for 15–20
    // minutes. After 30 minutes without any message or tool activity, abort the
    // wedged provider turn; finalize then uses the same bounded durable
    // continuation path as every other incomplete web turn.
    const stallWatch = setInterval(() => {
      if (
        !shouldAbortStalledTurn({
          busy: busyRef.current,
          sessionID: sessionRef.current,
          aborting: stallAbortRef.current,
          pendingPermissions: permissionsRef.current.length + pendingFormsRef.current,
          lastActivityAt: lastActivityAtRef.current,
          now: Date.now(),
        })
      )
        return
      stallAbortRef.current = true
      const stalledSession = sessionRef.current
      abortSession(client, stalledSession)
        .then(() => {
          if (busyRef.current && sessionRef.current === stalledSession) finalize()
        })
        .catch((error) => {
          if (!busyRef.current || sessionRef.current !== stalledSession) {
            stallAbortRef.current = false
            return
          }
          failTurn(stalledAbortFailureMessage(error))
        })
    }, 30_000)
    return () => {
      mounted = false
      controller.abort()
      document.removeEventListener("visibilitychange", onVisible)
      window.removeEventListener("focus", onVisible)
      clearInterval(poll)
      clearInterval(stallWatch)
      clearFinalizeRetry()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [client, sessionID])

  // Admit a streamed part into the live view. The moment a tool call that
  // mutates the surface's listing completes (not on its later re-deliveries),
  // tell the parent so its list (the documents rail, the template/workflow/
  // skill/agent library) picks up the change mid-turn.
  function admitPart(part: Part) {
    if (part.type === "tool" && part.tool === "question" && part.state.status === "running")
      pendingFormsRef.current = Math.max(1, pendingFormsRef.current)
    const prev = partsRef.current.get(part.id)
    partsRef.current.set(part.id, part)
    if (
      part.type === "tool" &&
      MUTATING_TOOLS.has(part.tool) &&
      part.state.status === "completed" &&
      !(prev?.type === "tool" && prev.state.status === "completed")
    )
      onDocumentsChanged?.()
  }

  function onEvent(event: Event, signal: AbortSignal) {
    if (event.type === "session.deleted") {
      childRef.current.delete(event.properties.sessionID)
      if (event.properties.sessionID !== sessionRef.current) return
      retiredSession.current = event.properties.sessionID
      permissionVersion.current += 1
      clearFinalizeRetry()
      busyRef.current = false
      setBusy(false)
      setLoading(false)
      updatePermissions([])
      setLoadError("This conversation was deleted. Start a new conversation to continue.")
      return
    }
    // projectEvent translates the native action/resources/save fields into this
    // browser's permission request shape before dispatching here.
    if (event.type === "permission.asked") {
      const req = event.properties
      if (req.sessionID !== sessionRef.current && !childRef.current.has(req.sessionID)) return
      permissionVersion.current += 1
      lastActivityAtRef.current = Date.now()
      setPermissionError(undefined)
      updatePermissions((prev) => [...prev.filter((p) => p.id !== req.id), req])
      automaticallyApprove(req, signal)
      return
    }
    if (event.type === "permission.replied") {
      if (event.properties.sessionID !== sessionRef.current && !childRef.current.has(event.properties.sessionID)) return
      permissionVersion.current += 1
      lastActivityAtRef.current = Date.now()
      setPermissionError(undefined)
      updatePermissions((prev) => prev.filter((p) => p.id !== event.properties.requestID))
      return
    }
    if (event.type === "session.error" && event.properties.sessionID === sessionRef.current) {
      // Stop fails the running step as interrupted; the stop settles the turn.
      if (turnRef.current.stopped) return
      if (stallAbortRef.current) {
        stallAbortRef.current = false
        sessionErrorRef.current = undefined
        finalize()
        return
      }
      sessionErrorRef.current = errorOf(event.properties.error as MessageError) ?? "The model returned an error."
      finalize()
      return
    }
    if (event.type === "message.updated") {
      const info = event.properties.info
      if (info.sessionID !== sessionRef.current) return
      lastActivityAtRef.current = Date.now()
      rolesRef.current.set(info.id, info.role)
      bump((n) => n + 1)
      return
    }
    if (event.type === "message.part.updated") {
      const part = event.properties.part
      const own = part.sessionID === sessionRef.current
      if (!own && !childRef.current.has(part.sessionID)) return onForeignPart?.(part)
      if (seenRef.current.has(part.messageID)) return
      lastActivityAtRef.current = Date.now()
      admitPart(part)
      // A task tool part names the child session it spawned; track that session
      // so its reasoning/tool parts get admitted and nest under this step.
      if (own && part.type === "tool" && part.tool === "task") {
        const child = (part.state as { metadata?: { sessionId?: string } }).metadata?.sessionId
        if (child) childRef.current.add(child)
      }
      bump((n) => n + 1)
      return
    }
    if (event.type === "session.status" && event.properties.sessionID === sessionRef.current) {
      if (event.properties.status.type === "busy" || event.properties.status.type === "retry") {
        lastActivityAtRef.current = Date.now()
        busyRef.current = true
        setBusy(true)
        return
      }
      if (event.properties.status.type === "idle" && busyRef.current) {
        stallAbortRef.current = false
        finalize()
        return
      }
    }
    if (event.type === "session.idle" && event.properties.sessionID === sessionRef.current && busyRef.current) {
      stallAbortRef.current = false
      finalize()
    }
  }

  // The event stream dropped and reconnected, so any parts (and the session.idle)
  // emitted during the gap were missed. If a turn is in flight, pull the active
  // session and merge its parts back into the live view so it catches up instead
  // of stalling on a partial turn; if that turn already settled while we were
  // disconnected, its session.idle was missed too, so finalize from the server's
  // completion marker. Settled views only reconcile their pending approvals.
  async function resync(signal: AbortSignal) {
    if (!sessionRef.current || signal.aborted || retiredSession.current === sessionRef.current) return
    // Approval hydration must still work when the session status or transcript
    // cannot be loaded. It also covers a request missed before subscription.
    void reconcilePermissions(signal)
    if (!busyRef.current) return
    const conversation = await bounded(
      loadConversation(client, sessionRef.current),
      10_000,
      "Conversation resync",
    ).catch(() => undefined)
    if (!conversation || signal.aborted || retiredSession.current === sessionRef.current) return
    const { msgs, pool } = conversation
    const knownChildren = childRef.current.size
    for (const m of msgs) {
      if (seenRef.current.has(m.info.id)) continue
      rolesRef.current.set(m.info.id, m.info.role)
      for (const p of m.parts) {
        admitPart(p)
        if (p.type === "tool" && p.tool === "task") {
          const child = (p.state as { metadata?: { sessionId?: string } }).metadata?.sessionId
          if (child) childRef.current.add(child)
        }
      }
    }
    for (const part of pool.filter((candidate) => childRef.current.has(candidate.sessionID)))
      if (!seenRef.current.has(part.messageID)) admitPart(part)
    if (childRef.current.size !== knownChildren) void reconcilePermissions(signal)
    bump((n) => n + 1)
    // Finalize only when THIS turn produced a NEW settled answer: the last message
    // overall is a completed assistant whose id we have not already settled, and
    // the engine reports the session idle.
    // seenRef holds every settled message id (see load/finalize). Two cases this
    // rules out, both of which otherwise dropped the optimistic in-flight turn and
    // hid the bubble until a later session.idle re-rendered it all at once:
    //  - Auto routing/choosing: the prompt is not sent yet, so the last persisted
    //    message is the PRIOR turn's completed assistant — already in seenRef.
    //  - just after dispatch, before the server has persisted this turn's messages,
    //    the last message is still that prior completed assistant.
    // A multi-step answer also completes an assistant message at the end of each
    // step while the engine runs the next one, so only an idle session settles.
    const last = msgs[msgs.length - 1]?.info
    if (last?.role !== "assistant" || !last.time.completed || seenRef.current.has(last.id)) return
    const status = await bounded(getSessionStatus(client, sessionRef.current), 5_000, "Session status").catch(
      () => undefined,
    )
    if (status?.type !== "idle" || signal.aborted || retiredSession.current === sessionRef.current) return
    if (busyRef.current && !seenRef.current.has(last.id)) finalize()
  }

  async function reconcilePermissions(signal: AbortSignal) {
    const session = sessionRef.current
    if (!session || signal.aborted || retiredSession.current === session) return
    const version = ++permissionVersion.current
    const pending = await bounded(listPermissions(directory), 5_000, "Permission reconciliation").catch(() => undefined)
    // A newer event, reply or query owns the displayed state. An older GET
    // snapshot must not resurrect a request that the engine already answered.
    if (signal.aborted || sessionRef.current !== session || version !== permissionVersion.current) return
    if (!pending) {
      setPermissionError("Could not verify pending approvals. The workflow remains paused until they can be loaded.")
      return
    }
    const mine = pending.filter((request) => request.sessionID === session || childRef.current.has(request.sessionID))
    setPermissionError(undefined)
    updatePermissions(mine)
    for (const request of mine) automaticallyApprove(request, signal)
    if (mine.length) lastActivityAtRef.current = Date.now()
  }

  function automaticallyApprove(request: PermissionRequest, signal: AbortSignal) {
    if (!autoApproved(request.permission) || automaticAttempts.current.has(request.id)) return
    automaticAttempts.current.add(request.id)
    // Keep the request pending until the engine confirms success. Repeated
    // events/reconciliation must not submit the same automatic decision twice.
    // A failed automatic attempt stays visible for a manual decision.
    replyPermission(directory, request.id, "once")
      .then(() => {
        if (signal.aborted) return
        permissionVersion.current += 1
        updatePermissions((previous) => previous.filter((pending) => pending.id !== request.id))
      })
      .catch(() => {
        if (signal.aborted || !permissionsRef.current.some((pending) => pending.id === request.id)) return
        setPermissionError("Automatic approval failed. Review this request manually.")
      })
  }
}
