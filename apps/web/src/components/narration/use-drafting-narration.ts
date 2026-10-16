import { openQuestions } from "../drafting/model"
import { useCallback, useEffect, useRef, useState } from "react"
import type { Part } from "../../api/session-view"
import type { Client } from "../../api/opencode/client"
import { getMessages } from "../../api/opencode/sessions"
import { getDraftingStatus } from "../../api/ingest/drafting-status"
import { leanStageSessions, type LeanDraftingStatus } from "../../api/ingest/lean"
import {
  findNarratorSession,
  narrate,
  openNarratorSession,
  readNarration,
  type NarratorNote,
} from "../../api/opencode/narrator"
import { bounded } from "../chat/conversation"
import { activeStep, isWorking, partActivity, runTransition, stageMilestones, type ActivityLine } from "./activity"
import { buildExcerpt, stageSessionNotes, type StageNotes } from "./excerpt"
import { checkpointKeys, planNarration, type NarrationPlan } from "./schedule"

// Reads a matter's lean drafting run for the chat's narration: the run status
// (polled), activity lines from the stage sessions' streamed parts, and the
// narrator's updates. It only reads the run. It never continues it, answers a
// question, retries a stage or writes run state; its one write is a prompt to
// its own narrator session.

const POLL_WORKING_MS = 5_000
const POLL_SETTLED_MS = 20_000
/** A stage session silent on the event stream this long is read directly. */
const BACKFILL_MS = 30_000

type RunView = ReturnType<typeof freshRun>

function freshRun(owner: string) {
  return {
    owner,
    lines: new Map<string, ActivityLine>(),
    stageNotes: new Map<string, StageNotes>(),
    /** Stage sessions explicitly named by the current matter run. */
    sessionStep: new Map<string, string>(),
    lastPartAt: new Map<string, number>(),
    backfilledAt: new Map<string, number>(),
    narrated: new Set<string>(),
    seenAt: new Map<string, number>(),
    narrator: undefined as Promise<string> | undefined,
    historyLoaded: false,
    seenActive: false,
    inFlight: false,
    lastCallAt: 0,
    failures: 0,
  }
}

export type DraftingNarration = ReturnType<typeof useDraftingNarration>

export function useDraftingNarration({
  client,
  matterId,
  matterName,
  startedOwner,
  sessionID,
}: {
  client: Client
  matterId?: string
  matterName: string
  /** The run this panel started, if any; its narration shows even once complete. */
  startedOwner?: string
  /** The open conversation. The run's own conversation always shows its narration. */
  sessionID?: string
}) {
  const [status, setStatus] = useState<LeanDraftingStatus>()
  const [notes, setNotes] = useState<NarratorNote[]>([])
  const [narrating, setNarrating] = useState(false)
  const [failed, setFailed] = useState(false)
  const [, bump] = useState(0)
  const run = useRef<RunView>(freshRun(""))
  const statusRef = useRef<LeanDraftingStatus | undefined>(undefined)
  const startedRef = useRef(startedOwner)
  startedRef.current = startedOwner

  // A finished run stays in the conversation that started it and in its own
  // conversation; elsewhere it shows only while it is open, still has questions
  // for the lawyer, or was seen working.
  const visibleFor = (next: LeanDraftingStatus | undefined) =>
    Boolean(next) &&
    (startedRef.current === next!.owner_session_id ||
      sessionID === next!.owner_session_id ||
      next!.status !== "completed" ||
      openQuestions(next!).length > 0 ||
      run.current.seenActive)

  function admit(part: Part, view: RunView, at = Date.now()) {
    if (view !== run.current) return
    view.lastPartAt.set(part.sessionID, Date.now())
    if (part.type === "text" || part.type === "reasoning") {
      const notes = view.stageNotes.get(part.sessionID) ?? { text: "", reasoning: "" }
      view.stageNotes.set(
        part.sessionID,
        part.type === "text" ? { ...notes, text: part.text } : { ...notes, reasoning: part.text },
      )
    }
    const line = partActivity(part, at, matterName)
    if (!line) return
    const prev = view.lines.get(line.id)
    if (prev && prev.text === line.text && prev.state === line.state) return
    view.lines.set(line.id, { ...line, at: prev?.at ?? line.at, step: view.sessionStep.get(part.sessionID) ?? undefined })
    bump((n) => n + 1)
  }

  // Run status names the stage sessions. Unknown global sessions (including
  // transient routers in this matter) never need transcript or ownership reads.
  const acceptsSession = useCallback((id: string) => draftingStageSession(statusRef.current, id), [])
  const onForeignPart = useCallback((part: Part) => {
    if (acceptsSession(part.sessionID)) admit(part, run.current)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [client])

  function observe(next: LeanDraftingStatus) {
    if (next.owner_session_id !== run.current.owner) {
      run.current = freshRun(next.owner_session_id)
      setNotes([])
      setFailed(false)
      loadHistory(run.current)
    }
    const view = run.current
    const now = Date.now()
    const prev = statusRef.current
    const change = prev?.owner_session_id === next.owner_session_id ? runTransition(prev, next) : undefined
    if (change) {
      const id = `run:${change.state}:${next.updated_at}`
      view.lines.set(id, { id, at: now, kind: "milestone", ...change })
    }
    if (isWorking(next)) view.seenActive = true
    checkpointKeys(next).forEach((key) => view.seenAt.has(key) || view.seenAt.set(key, now))
    Object.entries(next.stages).forEach(([step, row]) => leanStageSessions(row).forEach(id => view.sessionStep.set(id, step)))
    const stages = stageMilestones(next)
    // A stage run that was replaced by a new run of its step before a poll saw
    // it finish stopped there; its row should not keep spinning.
    view.lines.forEach((line, id) => {
      if (line.kind !== "milestone" || !line.step || line.state !== "running") return
      const rerun = stages.find((stage) => stage.step === line.step && stage.at > line.at)
      if (rerun) view.lines.set(id, { ...line, state: "paused", until: rerun.at })
    })
    stages.forEach((line) => view.lines.set(line.id, line))
    statusRef.current = next
    setStatus(next)
  }

  function loadHistory(view: RunView) {
    findNarratorSession(client, view.owner)
      .then(async (sessionID) => {
        if (!sessionID) return
        view.narrator = Promise.resolve(sessionID)
        const history = await readNarration(client, sessionID)
        if (view !== run.current) return
        history.forEach((note) => note.keys.forEach((key) => view.narrated.add(key)))
        setNotes(history)
      })
      .catch(() => {})
      .finally(() => {
        view.historyLoaded = true
      })
  }

  // Running stage sessions are read directly when the event stream has not
  // carried them lately: after a reload, or if the stream dropped.
  async function backfill(next: LeanDraftingStatus, step: string | undefined, view: RunView) {
    const now = Date.now()
    const sessions = leanStageSessions(step ? next.stages[step] : undefined).filter(id =>
      now - (view.lastPartAt.get(id) ?? 0) >= BACKFILL_MS && now - (view.backfilledAt.get(id) ?? 0) >= BACKFILL_MS)
    sessions.forEach(id => view.backfilledAt.set(id, now))
    await Promise.all(sessions.map(async id => {
      const messages = await bounded(getMessages(client, id), 10_000, "Stage activity").catch(() => [])
      messages
        .filter((message) => message.info.role === "assistant")
        .flatMap((message) => message.parts)
        .forEach((part) => admit(part, view))
    }))
  }

  // A settled run's finished steps are read once, so its own conversation shows
  // the same per-step activity after a reload as it did live. Their lines carry
  // the step's start time, so they never count as fresh activity to narrate.
  async function backfillFinished(next: LeanDraftingStatus, view: RunView) {
    if (isWorking(next) || (startedRef.current !== next.owner_session_id && sessionID !== next.owner_session_id)) return
    const sessions = Object.values(next.stages).filter(row => row.status === "done")
      .flatMap(row => leanStageSessions(row).map(id => ({ id, at: row.started_at ? Date.parse(row.started_at) : Date.now() })))
      .filter(row => !view.backfilledAt.has(row.id))
    sessions.forEach(row => view.backfilledAt.set(row.id, Date.now()))
    await Promise.all(sessions.map(async row => {
      const messages = await bounded(getMessages(client, row.id), 10_000, "Stage activity").catch(() => [])
      messages
        .filter((message) => message.info.role === "assistant")
        .flatMap((message) => message.parts)
        .forEach((part) => admit(part, view, row.at))
    }))
  }

  async function maybeNarrate(next: LeanDraftingStatus) {
    const view = run.current
    if (!visibleFor(next) || !view.historyLoaded) return
    const now = Date.now()
    const plan = planNarration({
      status: next,
      narrated: view.narrated,
      seenAt: view.seenAt,
      now,
      inFlight: view.inFlight,
      failures: view.failures,
      lastCallAt: view.lastCallAt,
      activitySince: [...view.lines.values()].filter((line) => line.kind !== "milestone" && line.at > view.lastCallAt)
        .length,
    })
    if (!plan) return
    const since = view.lastCallAt
    view.inFlight = true
    view.lastCallAt = now
    setNarrating(true)
    try {
      const step = focusStep(next, plan)
      await backfill(next, step, view)
      const sessions = leanStageSessions(step ? next.stages[step] : undefined)
      const activity = [...view.lines.values()].filter((line) => line.at > since).sort((a, b) => a.at - b.at)
      const excerpt = buildExcerpt({
        status: next,
        plan,
        activity,
        notes: stageSessionNotes(sessions.flatMap(id => view.stageNotes.get(id) ?? [])),
        now,
      })
      view.narrator ??= openNarratorSession(client, view.owner)
      // A failed open is tried again on the next update; a failed update is not.
      const sessionID = await view.narrator.catch((error: unknown) => {
        view.narrator = undefined
        throw error
      })
      const history = await narrate(client, sessionID, excerpt, plan.keys)
      if (view !== run.current) return
      view.failures = 0
      setFailed(false)
      setNotes(history)
    } catch {
      if (view !== run.current) return
      view.failures += 1
      setFailed(true)
    } finally {
      plan.keys.forEach((key) => view.narrated.add(key))
      view.inFlight = false
      setNarrating(false)
    }
  }

  useEffect(() => {
    if (!matterId) return
    let stopped = false
    let timer: ReturnType<typeof setTimeout> | undefined
    const tick = async () => {
      const next = await getDraftingStatus(matterId).catch(() => undefined)
      if (stopped) return
      if (next) observe(next)
      const working = next ? isWorking(next) : false
      if (next && visibleFor(next)) {
        await Promise.all(Object.entries(next.stages).filter(([, row]) => row.status === "running")
          .map(([step]) => backfill(next, step, run.current))).catch(() => {})
        void backfillFinished(next, run.current)
        void maybeNarrate(next)
      }
      if (!stopped) timer = setTimeout(tick, working ? POLL_WORKING_MS : POLL_SETTLED_MS)
    }
    void tick()
    return () => {
      stopped = true
      clearTimeout(timer)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matterId, client, startedOwner])

  return {
    status,
    visible: visibleFor(status),
    lines: [...run.current.lines.values()],
    notes,
    narrating,
    failed,
    onForeignPart,
    acceptsSession,
  }
}

/** The step whose own output best explains this update: the active step for a
 * progress update, the stage named by a stage checkpoint, else the latest stage
 * that ran. */
function focusStep(status: LeanDraftingStatus, plan: NarrationPlan) {
  if (plan.reason === "progress") return plan.focus
  const named = plan.focus.split(":")[0]!
  if (status.stages[named]) return named
  return (
    activeStep(status) ??
    Object.entries(status.stages)
      .filter(([, row]) => row.finished_at)
      .sort(([, a], [, b]) => Date.parse(a.finished_at!) - Date.parse(b.finished_at!))
      .at(-1)?.[0]
  )
}

/** Only sessions named by the current run status can supply drafting activity. */
export function draftingStageSession(status: Pick<LeanDraftingStatus, "stages"> | undefined, sessionID: string) {
  return Boolean(status && Object.values(status.stages).some(row => leanStageSessions(row).includes(sessionID)))
}
