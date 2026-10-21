import { abortSession, createSession, revertMessage, sendPrompt, waitForSession } from "../../api/opencode/sessions"
import { getConfig, modelSpec } from "../../api/opencode/config"
import { replyPermission, type PermissionReply } from "../../api/opencode/permissions"
import type { Client } from "../../api/opencode/client"
import { routeAgent } from "../../api/opencode/routing"
import { CHAT_ASSISTANTS, isAuto, type WorkflowMeta } from "../../agents"
import { startDraftingRun } from "../../api/ingest/drafting-run"
import type { Turn } from "./turn-model"
import { bounded } from "./conversation"
import type { ChatState } from "./use-chat-state"
import type { TurnSettlement } from "./use-turn-settlement"

// A session title from the first message: trimmed to a word boundary with an
// ellipsis, so the rail reads cleanly instead of cutting mid-word.
function titleFrom(text: string) {
  if (text.length <= 60) return text
  const cut = text.slice(0, 60)
  const space = cut.lastIndexOf(" ")
  return (space > 30 ? cut.slice(0, space) : cut).trimEnd() + "…"
}

export function useChatActions({
  state,
  settlement,
  client,
  agent,
  matterId,
  directory,
  workflows,
  onSessionPersisted,
  onSessionStarted,
  onDraftingStarted,
}: {
  state: ChatState
  settlement: TurnSettlement
  client: Client
  agent: string
  matterId?: string
  directory: string
  workflows: WorkflowMeta[]
  onSessionPersisted?: (id: string) => void
  onSessionStarted?: () => void
  // A drafting run started from this chat: its owner session and the index of
  // the user turn that asked for it, where the run's narration is anchored.
  onDraftingStarted?: (ownerSessionID: string, turnIndex: number) => void
}) {
  const {
    turns,
    setTurns,
    input,
    setInput,
    busy,
    setBusy,
    setPermissionError,
    setEditing,
    setResolvedAgent,
    sessionRef,
    freshRef,
    partsRef,
    rolesRef,
    childRef,
    sessionErrorRef,
    malformedToolRecoveryRef,
    lastActivityAtRef,
    stallAbortRef,
    stickRef,
    inputRef,
    bump,
    busyRef,
    turnRef,
    updatePermissions,
  } = state
  const { clearFinalizeRetry, finalize, failTurn } = settlement

  // A send owns its turn until its prompt returns; Stop marks the turn stopped.
  function startTurn() {
    const turn = { stopped: false, owned: true }
    turnRef.current = turn
    return turn
  }

  // Stop came before the prompt reached the engine: drop the unsent question,
  // give its text back to the composer, and settle.
  function cancelUnsent(text: string) {
    setInput((current) => (current.trim() ? current : text))
    setTurns((prev) => prev.slice(0, -1))
    void finalize()
  }

  // The engine has ended a stopped turn, so its prompt returned. Settle from the
  // stored history, which keeps the partial answer and marks it stopped.
  function settleStopped() {
    if (busyRef.current) void finalize()
  }

  // Stop the running answer through the engine's interrupt, as the stall
  // watchdog does. The send that owns the turn settles it when its prompt
  // returns; a turn no send here owns (a reopened conversation, a recovery)
  // settles once the engine is idle.
  async function onStop() {
    const turn = turnRef.current
    if (!busyRef.current || turn.stopped) return
    turn.stopped = true
    bump((n) => n + 1)
    const session = sessionRef.current
    // No session yet means nothing has reached the engine; the send cancels.
    if (!session) return
    try {
      await abortSession(client, session)
    } catch {
      // The engine did not take the stop: the answer goes on and Stop is back.
      turn.stopped = false
      bump((n) => n + 1)
      return
    }
    if (turn.owned) return
    await waitForSession(client, session).catch(() => undefined)
    if (busyRef.current && sessionRef.current === session && turnRef.current === turn) void finalize()
  }

  // Resolve which real agent answers this message. With Auto selected, a cheap
  // model routes the message to one of the real assistants; otherwise the picked
  // agent is used as-is. `prior` is the turns before this one, newest last — the
  // last few exchanges seed the router, each labelled with its role and (for
  // assistant turns) the agent that answered, so a follow-up like "I confirm"
  // routes to the assistant whose action is pending rather than whoever spoke
  // last. The result stamps the user message, so it must be a real agent id:
  // routing never returns "auto", and any failure or timeout is surfaced rather
  // than silently choosing a different assistant. The resolved agent drives the
  // "Auto · <label>" chip and the in-flight bubble's byline.
  async function resolveAgent(text: string, prior: Turn[]) {
    if (!isAuto(agent)) return agent
    const history = prior
      .filter((t) => t.text)
      .slice(-6)
      .map((t) => {
        const who = t.role === "user" ? "user" : `assistant (${t.agent ?? "qa"})`
        return `${who}: ${t.text.length > 240 ? `${t.text.slice(0, 240)}…` : t.text}`
      })
    // Like routeAgent, this read must not block the send: a stalled fetch would
    // never reject, so cap it and fall back to the default small model.
    const config = await bounded(getConfig(), 3_000, "Model configuration").catch(() => undefined)
    // Route through the engine using its configured small model, whatever provider
    // it lives on. OpenCode resolves the provider and credentials, so Auto works
    // on Vertex, Bedrock, Azure, OpenAI, or Ollama alike (see routeAgent).
    const small = config?.small_model ?? config?.model ?? "google-vertex/gemini-3.5-flash"
    const resolved = await routeAgent({
      client,
      smallModel: small,
      // On a miss (a weak small model returns an off-list name) routeAgent retries
      // the route on the primary model before surfacing failure. Omitted when it
      // equals small, so a single-model setup never pays for a duplicate call.
      primaryModel: config?.model,
      candidates: CHAT_ASSISTANTS.map((a) => ({ name: a.name, description: a.description })),
      history,
      text,
    })
    setResolvedAgent(resolved)
    return resolved
  }

  async function onSend() {
    const text = input.trim()
    if (!text || busy) return
    const turn = startTurn()
    const prior = turns
    malformedToolRecoveryRef.current = 0
    clearFinalizeRetry()
    lastActivityAtRef.current = Date.now()
    stallAbortRef.current = false
    stickRef.current = true
    setTurns((prev) => [...prev, { role: "user", text, citations: [], redlines: [], drafts: [], todos: [], steps: [] }])
    setInput("")
    if (inputRef.current) inputRef.current.style.height = "auto"
    setBusy(true)
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    let sent = false
    try {
      const resolved = await resolveAgent(text, prior)
      // Stopped while Auto chose the assistant.
      if (turn.stopped) return cancelUnsent(text)
      // Stop does not reach a drafting run once it starts; the run has its own Stop.
      if (resolved === "draft" && matterId) {
        const model = modelSpec(await bounded(getConfig(), 3_000, "Model configuration"))
        if (!model) throw new Error("The configured drafting model could not be resolved.")
        const started = await bounded(
          startDraftingRun(matterId, {
            request: text,
            agent: resolved,
            model,
          }),
          15_000,
          "Drafting run start",
        )
        // The drafting run executes in its own sessions; this conversation is
        // unchanged. The run's narration answers under the request, and the run
        // header above the chat shows its steps and what to do next.
        onDraftingStarted?.(started.ownerSessionID, prior.length)
        // A fresh chat has no conversation of its own to reopen, so a reload
        // opens the run's conversation, which shows the request and narration.
        if (!sessionRef.current) onSessionPersisted?.(started.ownerSessionID)
        onSessionStarted?.()
        busyRef.current = false
        setBusy(false)
        return
      }
      // Other chats create their ordinary conversation lazily.
      if (!sessionRef.current) {
        sessionRef.current = (await createSession(client, titleFrom(text))).id
        freshRef.current = true
        onSessionPersisted?.(sessionRef.current)
        onSessionStarted?.()
      }
      if (turn.stopped) return cancelUnsent(text)
      sent = true
      await sendPrompt(client, sessionRef.current, resolved, text)
    } catch (e) {
      if (turn.stopped) return sent ? settleStopped() : cancelUnsent(text)
      // Routing happens before a fresh session is minted, so there is no stored
      // user message to retry yet. Restore the text to the composer instead of
      // losing it when Auto cannot choose an assistant.
      setInput(text)
      failTurn(e instanceof Error ? e.message : "Could not reach the engine. Try again.")
    } finally {
      turn.owned = false
    }
  }

  // A workflow is a canned, orchestrated run that streams into the chat like any
  // answer, rather than a separate artifact surface. A matter-scoped routine spans
  // every document, so it opens its own fresh conversation; a document-scoped one
  // continues the current thread. Either way its subagent steps show on the
  // timeline and the combined report streams as the assistant turn.
  async function runWorkflow(name: string) {
    if (busy) return
    const wf = workflows.find((w) => w.name === name)
    if (!wf) return
    const turn = startTurn()
    stickRef.current = true
    setBusy(true)
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    if (wf.scope === "matter") sessionRef.current = ""
    setTurns((prev) =>
      wf.scope === "matter"
        ? [{ role: "user", text: wf.prompt, citations: [], redlines: [], drafts: [], todos: [], steps: [] }]
        : [...prev, { role: "user", text: wf.prompt, citations: [], redlines: [], drafts: [], todos: [], steps: [] }],
    )
    let sent = false
    try {
      if (!sessionRef.current) {
        sessionRef.current = (await createSession(client, wf.label)).id
        freshRef.current = true
        onSessionPersisted?.(sessionRef.current)
        onSessionStarted?.()
      }
      if (turn.stopped) return cancelUnsent("")
      sent = true
      await sendPrompt(client, sessionRef.current, wf.name, wf.prompt)
    } catch (e) {
      if (turn.stopped) return sent ? settleStopped() : cancelUnsent("")
      failTurn(e instanceof Error ? e.message : "Could not start the workflow. Try again.")
    } finally {
      turn.owned = false
    }
  }

  // Edit and retry both rewind the session to a user turn and re-ask: revert
  // drops that message and everything after it server-side, then we re-send the
  // (possibly edited) prompt so the assistant answer is regenerated. The UI is
  // truncated to before the turn optimistically; finalize reloads the truth.
  async function resendFrom(turn: Turn, index: number, text: string) {
    if (!turn.id || !text.trim() || busy) return
    const run = startTurn()
    setEditing(null)
    setBusy(true)
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    setTurns((prev) => [
      ...prev.slice(0, index),
      { role: "user", text, citations: [], redlines: [], drafts: [], todos: [], steps: [] },
    ])
    let sent = false
    try {
      const resolved = await resolveAgent(text, turns.slice(0, index))
      if (run.stopped) return cancelUnsent(text)
      await revertMessage(client, sessionRef.current, turn.id)
      if (run.stopped) return cancelUnsent(text)
      sent = true
      await sendPrompt(client, sessionRef.current, resolved, text)
    } catch (e) {
      if (run.stopped) return sent ? settleStopped() : cancelUnsent(text)
      failTurn(e instanceof Error ? e.message : "Could not reach the engine. Try again.")
    } finally {
      run.owned = false
    }
  }

  // Optimistic removal — the engine's permission.replied event confirms it, and
  // a reject cascades rejection to the session's other pending requests, whose
  // replied events clear them here too.
  async function onPermissionReply(id: string, reply: PermissionReply) {
    try {
      await replyPermission(directory, id, reply)
      lastActivityAtRef.current = Date.now()
      setPermissionError(undefined)
      updatePermissions((prev) => prev.filter((p) => p.id !== id))
    } catch {
      setPermissionError("Could not submit that approval. The request is still pending.")
    }
  }

  return {
    onSend,
    onStop,
    runWorkflow,
    resendFrom,
    onPermissionReply,
  }
}
