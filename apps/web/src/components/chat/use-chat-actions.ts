import { useState } from "react"
import { abortSession, createSession, deleteSession, revertMessage, sendPrompt, waitForSession } from "../../api/opencode/sessions"
import { getConfig, modelSpec } from "../../api/opencode/config"
import { replyPermission, type PermissionReply } from "../../api/opencode/permissions"
import { boundedRequest, type Client } from "../../api/opencode/client"
import { quickDraftEligible, routeAgent, type Route } from "../../api/opencode/routing"
import { AUTO_CANDIDATES, CHAT_ASSISTANTS, EXPLICIT_ONLY, QUICK_DRAFT, isAuto, type AssistantMeta, type WorkflowMeta } from "../../agents"
import { startDraftingRun } from "../../api/ingest/drafting-run"
import { getPreferences } from "../../api/ingest/preferences"
import type { Turn } from "./turn-model"
import { bounded } from "./conversation"
import {
  quickDraftLimit,
  rememberChoice,
  rememberedChoice,
  routeForm,
  upgradeInput,
  withMatterSources,
  type DraftChoice,
} from "./quick-draft"
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

// An explicit Quick draft pick is still checked against quick-draft scope, by
// the same router judgment Auto uses, with drafting as the only route.
const DRAFT_ONLY: AssistantMeta[] = CHAT_ASSISTANTS.filter((a) => a.name === "draft")

// The firm's quick-draft offer setting (Settings → Drafting). An unreadable
// setting keeps the default, Ask.
async function quickDraftOffer() {
  const prefs = await bounded(getPreferences(), 3_000, "Drafting preferences").catch(() => undefined)
  return prefs?.quickDraftOffer === "auto" || prefs?.quickDraftOffer === "off" ? prefs.quickDraftOffer : "ask"
}

const userTurn = (text: string): Turn => ({ role: "user", text, citations: [], redlines: [], drafts: [], todos: [], steps: [] })

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
  sourceCount,
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
  // How many source documents the matter holds, when known: with three or
  // fewer, a drafting request can only draw on few of them.
  sourceCount?: number
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
  // An explicit Quick draft pick whose request is outside quick-draft scope:
  // the reason, and the request for the full run it offers instead.
  const [quickLimit, setQuickLimit] = useState<{ text: string; reason: string }>()

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
  async function resolveAgent(text: string, prior: Turn[]): Promise<Route> {
    if (!isAuto(agent)) return { name: agent }
    const route = await routeMessage(text, prior, AUTO_CANDIDATES)
    setResolvedAgent(route.name)
    return route
  }

  async function routeMessage(text: string, prior: Turn[], candidates: AssistantMeta[]) {
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
    return routeAgent({
      client,
      smallModel: small,
      // On a miss (a weak small model returns an off-list name) routeAgent retries
      // the route on the primary model before surfacing failure. Omitted when it
      // equals small, so a single-model setup never pays for a duplicate call.
      primaryModel: config?.model,
      // Quick draft is never a candidate: Auto offers it by asking, or starts
      // it under the firm's Automatic setting, only for an eligible request.
      candidates: candidates.map((a) => ({ name: a.name, description: a.description })),
      history,
      text,
    })
  }

  // Start the matter's full drafting run for a request. The run executes in its
  // own sessions; its narration answers under the turn at `anchor`, and the run
  // header above the chat shows its steps and what to do next. An upgraded
  // quick draft rides along as an unverified working draft, never as request
  // text. Stop does not reach a drafting run once it starts; it has its own.
  async function startFullRun(request: string, anchor: number, workingDraft?: string) {
    if (!matterId) throw new Error("A drafting run needs a matter.")
    const model = modelSpec(await bounded(getConfig(), 3_000, "Model configuration"))
    if (!model) throw new Error("The configured drafting model could not be resolved.")
    const started = await bounded(
      startDraftingRun(matterId, { request, agent: "draft", model, ...(workingDraft ? { workingDraft } : {}) }),
      15_000,
      "Drafting run start",
    )
    onDraftingStarted?.(started.ownerSessionID, anchor)
    // A fresh chat has no conversation of its own to reopen, so a reload
    // opens the run's conversation, which shows the request and narration.
    if (!sessionRef.current) onSessionPersisted?.(started.ownerSessionID)
    onSessionStarted?.()
    busyRef.current = false
    setBusy(false)
  }

  // A conversation made only to hold the quick-draft offer, with nothing stored
  // in it, is removed once the offer ends without a quick draft: Full runs in
  // its own conversation, and Cancel leaves nothing to reopen. Emptiness is
  // read from the engine, so a conversation with history is never removed.
  async function retireOfferSession() {
    const id = sessionRef.current
    if (!id || turns.some((turn) => turn.id)) return
    const stored = await boundedRequest(client.message.list({ sessionID: id, limit: 1 }), 5_000, "Conversation check")
      .then((page) => page.data.length > 0)
      .catch(() => true)
    if (stored || sessionRef.current !== id) return
    sessionRef.current = ""
    freshRef.current = false
    const url = new URL(window.location.href)
    url.searchParams.delete("session")
    window.history.replaceState(window.history.state, "", url)
    await deleteSession(client, id).catch(() => undefined)
    onSessionStarted?.()
  }

  // Auto judged this request one a quick draft can serve: ask the lawyer in a
  // native session form, with their last choice on this matter pre-selected.
  // Nothing runs until they answer. The conversation exists first, because a
  // form belongs to a session; the request stays as the pending turn.
  async function offerQuickDraft(text: string, turn: { stopped: boolean }) {
    if (!sessionRef.current) {
      sessionRef.current = (await createSession(client, titleFrom(text))).id
      freshRef.current = true
      onSessionPersisted?.(sessionRef.current)
      onSessionStarted?.()
    }
    if (turn.stopped) return cancelUnsent(text)
    await boundedRequest(
      client.session.form.create(routeForm(sessionRef.current, text, rememberedChoice(matterId!))),
      10_000,
      "Quick draft offer",
    )
    busyRef.current = false
    setBusy(false)
  }

  // The lawyer answered the quick-draft offer, and the form has settled: run
  // what they chose. The settled form is the only start signal, so a second
  // answer (another tab, a double click) cannot start a second run.
  async function chooseDraft(request: string, choice: DraftChoice) {
    if (busy || !matterId) return
    rememberChoice(matterId, choice)
    const last = turns.at(-1)
    const pending = last?.role === "user" && !last.id && last.text === request
    const anchor = pending ? turns.length - 1 : turns.length
    if (!pending) setTurns((prev) => [...prev, userTurn(request)])
    const turn = startTurn()
    stickRef.current = true
    setBusy(true)
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    lastActivityAtRef.current = Date.now()
    let sent = false
    try {
      if (choice === "full") {
        await retireOfferSession()
        return await startFullRun(request, anchor)
      }
      setResolvedAgent(QUICK_DRAFT)
      if (turn.stopped) return cancelUnsent(request)
      sent = true
      await sendPrompt(client, sessionRef.current, QUICK_DRAFT, request)
    } catch (e) {
      if (turn.stopped) return sent ? settleStopped() : cancelUnsent(request)
      setInput((current) => (current.trim() ? current : request))
      failTurn(e instanceof Error ? e.message : "Could not reach the engine. Try again.")
    } finally {
      turn.owned = false
    }
  }

  // The lawyer cancelled the offer: nothing ran, so the request goes back to
  // the composer.
  function dismissDraftOffer(request: string) {
    setTurns((prev) => {
      const last = prev.at(-1)
      return last?.role === "user" && !last.id && last.text === request ? prev.slice(0, -1) : prev
    })
    setInput((current) => (current.trim() ? current : request))
    void retireOfferSession()
  }

  // "Run full drafting cycle" under a quick draft: the full run takes the
  // request and the latest quick text, and runs every stage against the sources.
  async function upgradeQuickDraft(index: number) {
    const input = upgradeInput(turns, index)
    if (busy || !matterId || !input) return
    const turn = startTurn()
    setBusy(true)
    try {
      await startFullRun(input.request, index, input.workingDraft)
    } catch (e) {
      failTurn(e instanceof Error ? e.message : "Could not start the drafting run. Try again.")
    } finally {
      turn.owned = false
    }
  }

  // Full instead of an explicit Quick draft pick the request did not fit.
  async function runFullInstead() {
    const offer = quickLimit
    if (!offer || busy || !matterId) return
    setQuickLimit(undefined)
    setInput((current) => (current.trim() === offer.text ? "" : current))
    const anchor = turns.length
    setTurns((prev) => [...prev, userTurn(offer.text)])
    const turn = startTurn()
    setBusy(true)
    try {
      await startFullRun(offer.text, anchor)
    } catch (e) {
      setInput((current) => (current.trim() ? current : offer.text))
      failTurn(e instanceof Error ? e.message : "Could not start the drafting run. Try again.")
    } finally {
      turn.owned = false
    }
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
    setQuickLimit(undefined)
    setTurns((prev) => [...prev, userTurn(text)])
    setInput("")
    if (inputRef.current) inputRef.current.style.height = "auto"
    setBusy(true)
    partsRef.current.clear()
    rolesRef.current.clear()
    childRef.current.clear()
    sessionErrorRef.current = undefined
    let sent = false
    try {
      // The Quick draft pick is the lawyer's consent; a request outside its
      // scope is explained, with the full run offered, and nothing runs.
      if (agent === QUICK_DRAFT && matterId) {
        const limit = quickDraftLimit(withMatterSources((await routeMessage(text, prior, DRAFT_ONLY)).scope, sourceCount))
        if (turn.stopped) return cancelUnsent(text)
        if (limit) {
          setTurns((prev) => prev.slice(0, -1))
          setInput((current) => (current.trim() ? current : text))
          setQuickLimit({ text, reason: limit })
          busyRef.current = false
          setBusy(false)
          return
        }
      }
      const resolved = await resolveAgent(text, prior)
      // Stopped while Auto chose the assistant.
      if (turn.stopped) return cancelUnsent(text)
      let target = resolved.name
      if (resolved.name === "draft" && matterId) {
        // A request Auto judged a quick draft can serve is offered, or started
        // under the firm's Automatic setting, unless the firm turned quick
        // drafts off; any other drafting request runs in full.
        const offer = isAuto(agent) && quickDraftEligible(withMatterSources(resolved.scope, sourceCount)) ? await quickDraftOffer() : "off"
        if (offer === "ask") return await offerQuickDraft(text, turn)
        if (offer === "off") return await startFullRun(text, prior.length)
        // Automatic is the firm's consent, as the Quick draft pick is the
        // lawyer's: the request goes to Quick draft like that pick.
        target = QUICK_DRAFT
        setResolvedAgent(target)
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
      await sendPrompt(client, sessionRef.current, target, text)
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
      // Retrying a quick draft or client summary keeps the lawyer's choice for
      // that request.
      const chosen = turns[index + 1]?.agent
      const keep = chosen && EXPLICIT_ONLY.has(chosen) && text === turn.text ? chosen : undefined
      const resolved = keep ?? (await resolveAgent(text, turns.slice(0, index))).name
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
    chooseDraft,
    dismissDraftOffer,
    upgradeQuickDraft,
    quickLimit,
    dismissQuickLimit: () => setQuickLimit(undefined),
    runFullInstead,
  }
}
