import { draftingUpdateAdmission } from "../../../dochaus/lib/drafting-update-admission"
import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { MemoryRouter } from "react-router-dom"
import answer from "../../../dochaus/workflows/lean/fixtures/answer.json"
import resume from "../../../dochaus/workflows/lean/fixtures/resume.json"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"
import refusedOpen from "../../../dochaus/workflows/lean/fixtures/start-refused-open.json"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import failed from "../../../dochaus/workflows/lean/fixtures/status-failed.json"
import orphaned from "../../../dochaus/workflows/lean/fixtures/status-orphaned.json"
import { getDraftingStatus } from "../src/api/ingest/drafting-status"
import { startDraftingRun } from "../src/api/ingest/drafting-run"
import {
  answerLeanQuestion,
  decideLeanQuestion,
  isLeanDraftingStatus,
  leanDeliveredDocuments,
  pauseLeanDraftingRun,
  resumeLeanDraftingRun,
  stopLeanDraftingRun,
  type LeanDraftingStatus,
  type LeanQuestion,
} from "../src/api/ingest/lean"
import { draftingTurnRecoveryMode } from "../src/components/chat/drafting-recovery"
import { draftingStatusLabel } from "../src/components/matter/MatterRail"
import { delegableQuestion, handedBackQuestions, leanContinueAvailable, leanRunStopped, leanRunTakesUpdates, leanSettledItems, leanSettledSummary, runPosition, stepGroups, openQuestions } from "../src/components/drafting/model"
import QuestionDock from "../src/components/drafting/QuestionDock"
import Questions from "../src/components/drafting/Questions"
import { runOutcome } from "../src/components/narration/outcome"
import RunHeader, { runHeadline } from "../src/components/drafting/RunHeader"
import type { RunActions } from "../src/components/drafting/useRunActions"

// Response bodies recorded from the real ingest lean start, status, Continue and
// answer routes with a scripted model (dochaus/workflows/lean/fixtures).

const originalFetch = globalThis.fetch

afterEach(() => {
  globalThis.fetch = originalFetch
})

function respond(body: unknown, status = 200) {
  const calls: Array<{ url: string; init?: RequestInit }> = []
  globalThis.fetch = (async (input, init) => {
    calls.push({ url: String(input), init })
    return Response.json(body, { status })
  }) as typeof fetch
  return calls
}

function lean(value: unknown): LeanDraftingStatus {
  if (!isLeanDraftingStatus(value)) throw new Error("fixture is not a lean drafting status")
  return value
}

test("a start that names no profile accepts ingest's lean start response", async () => {
  const calls = respond(start)
  const request = { request: "Draft a short supply note from the supply terms.", agent: "draft", model: { providerID: "test", modelID: "scripted" } }
  const result = await startDraftingRun("matter-a", request)
  expect(result).toMatchObject({ profile: "lean", mode: "started", ownerSessionID: "ses_child1" })
  expect(JSON.parse(String(calls[0]?.init?.body))).toEqual(request)

  respond(start)
  await expect(startDraftingRun("matter-a", { ...request, model: { providerID: "test", modelID: "other" } }))
    .rejects.toThrow("Drafting run start response is invalid.")
})

test("drafting status reads a lean run's state", async () => {
  for (const [index, fixture] of [completed, blockedGate, blockedQuestion, failed, orphaned].entries()) {
    const calls = respond(fixture)
    expect(await getDraftingStatus(`matter-${index}`)).toMatchObject({ profile: "lean", status: fixture.status })
    expect(calls[0]?.url).toEndWith(`/matters/matter-${index}/drafting-status`)
  }
  respond(null)
  expect(await getDraftingStatus("matter-none")).toBeUndefined()
  respond({ ...completed, stages: { ...completed.stages, gate: undefined } })
  await expect(getDraftingStatus("matter-x")).rejects.toThrow("Drafting workflow status response is invalid.")
})

test("a review draft delivered with unresolved checklist findings is read from ingest and named apart from a clean completion", async () => {
  const finding = {
    id: "checklist-check:checklist_unresolved:item-1",
    stage: "checklist-check",
    kind: "checklist_unresolved",
    detail: "Checklist requirement not satisfied: The note states the delivery period.",
    status: "disclosed",
    ref: { checklist_item: "item-1", checklist_class: "source_conflict", document: "d1" },
  }
  const body = { ...completed, open_items: [finding], gate: { ...completed.gate, delivery: "review_draft_with_findings", disclosed: [finding], items: [finding] } }
  expect(isLeanDraftingStatus(body)).toBe(true)
  respond(body)
  const status = (await getDraftingStatus("matter-findings")) as LeanDraftingStatus
  expect(status).toMatchObject({ status: "completed", gate: { delivery: "review_draft_with_findings", items: [finding] }, open_items: [finding] })

  expect(draftingStatusLabel(status)).toEqual({ text: "Review draft — 1 unresolved checklist finding", findings: true })
  expect(draftingStatusLabel(lean(completed))).toEqual({ text: "Drafting complete", findings: false })
  const plain = lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft" } })
  expect(draftingStatusLabel(plain)).toEqual({ text: "Drafting complete", findings: false })
  expect(draftingStatusLabel(lean(blockedGate))).toEqual({ text: "Drafting needs your attention", findings: false })
})

test("Continue posts to the lean resume route and answers to the question route", async () => {
  const resumed = respond(resume, 202)
  expect(await resumeLeanDraftingRun("matter-a")).toMatchObject({ profile: "lean", mode: "queued", status: { status: "running" } })
  expect(resumed).toEqual([{ url: expect.stringMatching(/\/matters\/matter-a\/lean\/resume$/), init: { method: "POST" } }])

  const answered = respond(answer, 202)
  const question = blockedQuestion.questions[0]!
  await answerLeanQuestion("matter-a", question.id, "No remedy.")
  expect(answered[0]?.url).toEndWith(`/matters/matter-a/lean/questions/${encodeURIComponent(question.id)}`)
  expect(JSON.parse(String(answered[0]?.init?.body))).toEqual({ answer: "No remedy." })

  respond({ error: "This matter has no lean drafting run.", code: "lean_run_not_found" }, 404)
  await expect(resumeLeanDraftingRun("matter-a")).rejects.toThrow("This matter has no lean drafting run.")
})

test("Continue is offered only for a stopped or orphaned run, and delivered documents are listed", () => {
  expect([blockedGate, blockedQuestion, failed, orphaned].map((fixture) => leanContinueAvailable(lean(fixture)))).toEqual([true, true, true, true])
  expect([completed, start.status].map((fixture) => leanContinueAvailable(lean(fixture)))).toEqual([false, false])
  expect(leanDeliveredDocuments(lean(completed))).toEqual(["Supply note.docx", "Supply note - redline.docx", "Supply note - drafting notes.docx"])
  expect(leanDeliveredDocuments(lean(blockedGate))).toEqual([])
})

test("the run header says where the run is in one line and offers the one next step", () => {
  const now = Date.parse(completed.updated_at)
  expect(runHeadline(lean(completed), now)).toMatchObject({ tone: "done", title: "Ready for review" })
  expect(runHeadline(lean(blockedQuestion), now)).toMatchObject({ tone: "input", title: "Waiting for your answers" })
  expect(runHeadline(lean(start.status), now).tone).toBe("running")
  expect([blockedGate, failed, orphaned].map((fixture) => runHeadline(lean(fixture), now).tone)).toEqual(["paused", "paused", "paused"])
  expect(runHeadline(lean(blockedGate), now).title).toBe("Paused at the final check")
  expect(runHeadline(lean(orphaned), now).title).toBe("Paused")
  expect(runHeadline(lean(failed), now, true).title).toBe("Resuming…")

  const idle: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => true, stop: async () => true }
  const render = (fixture: unknown) =>
    renderToStaticMarkup(createElement(RunHeader, { status: lean(fixture), actions: idle, own: true, onOpen: () => {}, onShowDrafts: () => {} }))
  const done = render(completed)
  expect(done).toContain("Ready for review")
  expect(done).toContain("Open draft")
  expect(done).not.toContain("Continue drafting")
  // The steps and open items are kept behind Details, folded at first.
  expect(done).toContain('aria-expanded="false"')
  expect(done).toContain("No open items remain.")
  const asked = render(blockedQuestion)
  // The question is answered, or left to the drafter with Decide for me, from
  // the dock above the composer; the header points there.
  expect(asked).toContain("1 question remaining · Drafting resumes once it is answered or left to the drafter")
  // A question Decide for me cannot take (one the run's model did not ask) waits on an answer.
  const scope = (question: (typeof blockedQuestion.questions)[number], index: number) => ({ ...question, id: `scope-${index}`, question: `Scope ${index}?` })
  expect(render({ ...blockedQuestion, questions: blockedQuestion.questions.map(scope) })).toContain("1 question remaining · Drafting resumes after all answers are received")
  const mixed = render({ ...blockedQuestion, questions: [...blockedQuestion.questions, ...blockedQuestion.questions.map(scope)] })
  expect(mixed).toContain("2 questions remaining · Drafting resumes once each is answered or, where Decide for me is offered, left to the drafter")
  expect(render({ ...blockedQuestion, questions: [...blockedQuestion.questions, { ...blockedQuestion.questions[0]!, id: "q-second", question: "Second?" }] }))
    .toContain("2 questions remaining · Drafting resumes once each is answered or left to the drafter")
  expect(asked).not.toMatch(/<button[^>]*>Answer/)
  expect(asked).not.toContain(blockedQuestion.questions[0]!.question)
  expect(asked).not.toContain("Continue drafting")
  expect(asked).toContain("held until the final check passes")
  expect(asked).toContain("3 prepared documents are held")
  // A combined file holds the memo and a draft already counted, so a stop that keeps one counts the same.
  const combined = structuredClone(blockedQuestion) as LeanDraftingStatus
  combined.documents!.documents.push({ kind: "combined", document: "d1", title: "Supply note", path: "supply-note.docx", sha256: "0".repeat(64), deliverable: "supply-note.docx" })
  expect(render(combined)).toContain("3 prepared documents are held")
  expect(asked).not.toContain("`")
  const gated = render(blockedGate)
  expect(gated).toContain("Continue drafting")
  expect(gated).toContain("Final check settled 2 open items: 1 still blocking, 1 resolved.")
  expect(gated).not.toContain("recorded in the drafting notes")
  expect(render(repairing())).toContain("Open items: 2 other items to check.")
})

function pausedRun() {
  const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
  // Paused during apply; the final check's row is still the earlier pass's.
  return lean({
    ...blockedGate,
    stop_reason: "Paused at your request.",
    paused_at: "2026-10-17T09:00:00.000Z",
    stages: { ...blockedGate.stages, "d1.apply": pending, finalize: pending },
  })
}

test("Pause posts the run's owner to the lean pause route and keeps refusals", async () => {
  const calls = respond({ profile: "lean", mode: "paused", status: pausedRun() })
  expect(await pauseLeanDraftingRun("matter-a", "ses_child1")).toMatchObject({ mode: "paused", status: { status: "blocked", stop_reason: "Paused at your request." } })
  expect(calls).toHaveLength(1)
  expect(calls[0]?.url).toEndWith("/matters/matter-a/lean/pause")
  expect(calls[0]?.init?.method).toBe("POST")
  expect(JSON.parse(String(calls[0]?.init?.body))).toEqual({ ownerSessionID: "ses_child1" })

  respond({ error: "This drafting run is not working, so there is nothing to pause.", code: "lean_run_not_working" }, 409)
  await expect(pauseLeanDraftingRun("matter-a", "ses_child1")).rejects.toThrow("This drafting run is not working, so there is nothing to pause.")
  respond({ profile: "lean", mode: "queued", request_id: "r1", status: pausedRun() })
  await expect(pauseLeanDraftingRun("matter-a", "ses_child1")).rejects.toThrow("Drafting run response is invalid.")
})

test("the header offers Pause only while the run is working; a paused run offers Continue", () => {
  const now = Date.parse(completed.updated_at)
  let paused = 0
  const actions: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => (paused++, true), stop: async () => true }
  const render = (status: LeanDraftingStatus, overrides: Partial<RunActions> = {}) =>
    renderToStaticMarkup(createElement(RunHeader, { status, actions: { ...actions, ...overrides }, own: true, onOpen: () => {}, onShowDrafts: () => {} }))
  const working = render(lean(start.status))
  expect(working).toMatch(/<button type="button" class="run-head-cta">Pause drafting<\/button>/)
  expect(working).not.toContain("Continue drafting")
  expect(render(lean(start.status), { busy: true, pausing: true })).toMatch(/<button type="button" class="run-head-cta" disabled="">Pausing…<\/button>/)
  expect(render(lean(start.status), { resuming: true })).not.toContain("Pause drafting")
  for (const fixture of [completed, blockedGate, blockedQuestion, failed, orphaned]) expect(render(lean(fixture))).not.toContain("Pause drafting")
  expect(paused).toBe(0)

  const stopped = render(pausedRun())
  expect(stopped).toContain("You paused drafting")
  expect(stopped).toContain("Continue drafting")
  expect(stopped).not.toContain("Pause drafting")
  expect(stopped).not.toContain("Continue runs a repair round")
  expect(runHeadline(pausedRun(), now)).toMatchObject({ tone: "paused", title: "Paused at apply changes" })
  expect(runOutcome(pausedRun())?.text).toBe("Paused at your request. Continue drafting (above) picks up from the step that was in progress.")
  expect(draftingUpdateAdmission({ ...pausedRun(), stages: { ...pausedRun().stages, "checklist-check": { status: "blocked" } } })).toBe("closed")
})

/** Stopped by the lawyer while a question and a request were still open. */
function stoppedRun() {
  const { paused_at: _, ...run } = pausedRun()
  return lean({ ...run, status: "failed", stopped: true, stop_reason: "Stopped at your request.", pending_requests: 1, questions: blockedQuestion.questions })
}

test("Stop posts the run's owner to the lean stop route and keeps refusals", async () => {
  const calls = respond({ profile: "lean", mode: "stopped", status: stoppedRun() })
  expect(await stopLeanDraftingRun("matter-a", "ses_child1")).toMatchObject({ mode: "stopped", status: { status: "failed", stop_reason: "Stopped at your request." } })
  expect(calls).toHaveLength(1)
  expect(calls[0]?.url).toEndWith("/matters/matter-a/lean/stop")
  expect(calls[0]?.init?.method).toBe("POST")
  expect(JSON.parse(String(calls[0]?.init?.body))).toEqual({ ownerSessionID: "ses_child1" })

  respond({ error: "This drafting run is working. Pause it before stopping it.", code: "lean_run_not_stoppable" }, 409)
  await expect(stopLeanDraftingRun("matter-a", "ses_child1")).rejects.toThrow("This drafting run is working. Pause it before stopping it.")
  respond({ profile: "lean", mode: "paused", status: pausedRun() })
  await expect(stopLeanDraftingRun("matter-a", "ses_child1")).rejects.toThrow("Drafting run response is invalid.")
})

test("a run nothing is working on offers Stop draft; a stopped run says so, asks nothing, offers no Continue and takes a new request", () => {
  const now = Date.parse(completed.updated_at)
  let stops = 0
  const actions: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => true, stop: async () => (stops++, true) }
  const render = (status: LeanDraftingStatus, overrides: Partial<RunActions> = {}, own = true) =>
    renderToStaticMarkup(createElement(MemoryRouter, null, createElement(RunHeader, { status, actions: { ...actions, ...overrides }, own, onOpen: () => {}, onShowDrafts: () => {} })))
  const paused = render(pausedRun())
  expect(paused).toMatch(/<button type="button" class="run-head-cta">Stop draft<\/button><button type="button" class="primary run-head-cta">Continue drafting<\/button>/)
  // The question is asked only once Stop is pressed.
  expect(paused).not.toContain("Stop this draft?")
  // Held for an answer, at a check, or left running by an ingest restart (nothing in flight).
  for (const fixture of [blockedQuestion, blockedGate, orphaned]) expect(render(lean(fixture))).toContain(">Stop draft<")
  for (const fixture of [start.status, completed, failed]) expect(render(lean(fixture))).not.toContain("Stop draft")
  expect(render(pausedRun(), { resuming: true })).not.toContain("Stop draft")
  expect(render(pausedRun(), { busy: true })).toMatch(/<button type="button" class="run-head-cta" disabled="">Stop draft<\/button>/)
  expect(stops).toBe(0)

  const stopped = stoppedRun()
  expect(runHeadline(stopped, now)).toMatchObject({ tone: "paused", title: "Draft stopped" })
  const html = render(stopped)
  expect(html).toContain("You stopped this draft")
  for (const text of ["Continue drafting", "Stop draft", "Pause drafting", "question for you", "Request queued"]) expect(html).not.toContain(text)
  expect(render(stopped, {}, false)).toContain("A new request starts a new draft")
  expect(openQuestions(stopped)).toEqual([])
  expect(leanContinueAvailable(stopped)).toBe(false)
  expect(runOutcome(stopped)?.text).toBe("You stopped this draft. Its work so far is kept; a new drafting request starts a new draft.")
  // The composer updates an open run; a completed or stopped run takes a new request.
  expect([pausedRun(), lean(failed), lean(blockedQuestion), stopped, lean(completed)].map(leanRunTakesUpdates)).toEqual([true, true, true, false, false])
  // Only ingest's derived flag reads as stopped; the reason is display only, and
  // any other failed run still offers Continue.
  const worded = lean({ ...failed, stop_reason: "Stopped at your request." })
  expect([lean(failed), worded, stopped].map(leanRunStopped)).toEqual([false, false, true])
  expect(leanContinueAvailable(worded)).toBe(true)
  expect(isLeanDraftingStatus({ ...stopped, stopped: "yes" })).toBe(false)
})

test("the conversation recovers only a printed tool call; a new draft request starts after a run finished and is refused while it is open", async () => {
  expect(draftingTurnRecoveryMode("token")).toBe("malformed-tool")
  expect(draftingTurnRecoveryMode(undefined)).toBe("none")
  const request = { request: "Draft a second supply note.", agent: "draft", model: { providerID: "test", modelID: "scripted" } }

  // After a completed or failed run, ingest archives it and returns a fresh start.
  const fresh = respond(start)
  expect(await startDraftingRun("matter-a", request)).toMatchObject({ profile: "lean", mode: "started", status: { status: "running" } })
  expect(fresh[0]?.url).toEndWith("/matters/matter-a/drafting-runs/start")

  // A blocked or paused run is refused with a code: the lawyer continues it or waits.
  respond(refusedOpen, 409)
  await expect(startDraftingRun("matter-a", request)).rejects.toThrow("Open its conversation to continue it, or wait until it finishes to start a new one.")
  respond({ error: "A lean drafting run is already in progress for this matter.", code: "lean_run_in_progress" }, 409)
  await expect(startDraftingRun("matter-a", request)).rejects.toThrow("Drafting is still running for this matter. Wait until it finishes to start a new one.")
})

test("open items are counted from the final check once it has run for the current pass", () => {
  // The run recorded two open items; the check resolved one and one still blocks.
  expect(blockedGate.open_items.map((item) => item.status)).toEqual(["open", "open"])
  const settled = leanSettledItems(lean(blockedGate))
  expect(settled?.map((item) => item.status)).toEqual(["resolved", "open"])
  expect(leanSettledSummary(settled!)).toBe("Final check settled 2 open items: 1 still blocking, 1 resolved.")

  const disclosed = blockedGate.open_items.map((item) => ({ ...item, status: "disclosed" }))
  const reviewDraft = lean({ ...completed, open_items: blockedGate.open_items, gate: { ...completed.gate, delivery: "review_draft", disclosed, items: disclosed } })
  expect(leanSettledSummary(leanSettledItems(reviewDraft)!)).toBe("Final check settled 2 open items: 2 disclosed in the drafting notes.")
  expect(leanSettledItems(lean(completed))).toEqual([])

  // A repair round reruns the stages; the earlier check's result no longer counts.
  expect(leanSettledItems(repairing())).toBeNull()
})

function repairing() {
  const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
  return lean({
    ...blockedGate,
    status: "running",
    in_flight: true,
    stages: { ...blockedGate.stages, "d1.apply": { ...pending, status: "running" }, finalize: pending, gate: pending },
  })
}

test("a run with several planned documents groups its steps per document", () => {
  const done = completed.stages["d1.draft"]
  const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
  const two = lean({
    ...completed,
    status: "running",
    in_flight: true,
    plan: { documents: [{ id: "d1", title: "Supply note", purpose: "" }, { id: "d2", title: "Covering letter", purpose: "" }] },
    stages: {
      "d1.draft": done, "d1.review": done, "d1.changes": done, "d1.apply": done,
      "d2.draft": done, "d2.review": { ...pending, status: "running" }, "d2.changes": pending, "d2.apply": pending, finalize: pending, gate: pending,
    },
  })
  expect(isLeanDraftingStatus(two)).toBe(true)
  expect(stepGroups(two).map((group) => [group.title, group.steps.map((step) => step.name)])).toEqual([
    ["Supply note", ["Draft", "Review", "Reconcile amendments", "Apply changes"]],
    ["Covering letter", ["Draft", "Review", "Reconcile amendments", "Apply changes"]],
    ["Finishing", ["Drafting notes", "Final check"]],
  ])
  expect(runPosition(two)).toEqual({ text: "Document 2 of 2 · Review", document: "Covering letter" })
  expect(stepGroups(lean(completed)).map((group) => group.title)).toEqual(["Supply note", "Finishing"])
  expect(isLeanDraftingStatus({ ...two, stages: { "d1.draft": done } })).toBe(false)
})

test("the question card shows how many are answered and Send waits for an answer", () => {
  const calls = respond({ answers: {}, revision: 0 })
  const open = blockedQuestion.questions.map((question) => ({ ...question, status: "open" as const }))
  const two = [open[0]!, { ...open[0]!, id: "q-second", question: "Who bears the cost of storage?" }]
  const html = renderToStaticMarkup(
    createElement(Questions, {
      matterId: "matter-q",
      ownerSessionID: blockedQuestion.owner_session_id,
      questions: two,
      answered: 1,
      titles: new Map(),
      busy: false,
      error: "",
      titleId: "q-title",
      onClose: () => {},
      onSend: async () => false,
      onDecide: async () => false,
    }),
  )
  // The heading names the dialog, which closes from its own button.
  expect(html).toContain('id="q-title"')
  expect(html).toContain('aria-label="Close"')
  expect(html).toContain("Question 1 of 2")
  expect(html).toContain("1 answered, 2 left")
  expect(html.match(/class="question-dot[ "]/g)).toHaveLength(3)
  expect(html).toContain('class="question-dot is-answered"')
  // An empty answer cannot be sent.
  expect(html).toMatch(/<button class="primary" type="button" disabled="">Send answer<\/button>/)
  expect(html).toContain("Loading saved answers…")
  expect(html).toMatch(/<textarea[^>]* disabled=""/)
  expect(calls).toHaveLength(0)
})

function questionCard(
  questions: LeanQuestion[],
  extra: { answered?: number; handedBack?: LeanQuestion[]; order?: string[]; done?: boolean; error?: string; errorFor?: string } = {},
) {
  respond({ answers: {}, revision: 0 })
  return renderToStaticMarkup(
    createElement(Questions, {
      matterId: "matter-decide",
      ownerSessionID: blockedQuestion.owner_session_id,
      questions,
      answered: extra.answered ?? 0,
      handedBack: extra.handedBack,
      order: extra.order,
      done: extra.done,
      titles: new Map(),
      busy: false,
      error: extra.error ?? "",
      errorFor: extra.errorFor,
      titleId: "q-title",
      onClose: () => {},
      onSend: async () => false,
      onDecide: async () => false,
    }),
  )
}

test("Decide for me is offered, with an empty answer, only on an open question the run's model asked", () => {
  const asked = { ...blockedQuestion.questions[0]!, status: "open" as const }
  expect(asked.id.startsWith("q-")).toBe(true)
  const html = questionCard([asked])
  // Never over a typed answer: until saved answers load the box may hold one, so
  // Decide waits like Send. It is enabled only while the box is empty (browser check).
  expect(html).toContain("Loading saved answers…")
  expect(html).toMatch(/<button type="button" aria-describedby="[^"]+" disabled="">Decide for me<\/button>/)
  expect(html).toMatch(/<button class="primary" type="button" disabled="">Send answer<\/button>/)
  // Decide and Send stay together when the actions wrap.
  expect(html).toMatch(/<span class="question-send"><button[^>]*>Decide for me<\/button><button class="primary"/)
  expect(html).toContain("You&#x27;ll get a review draft that says what was chosen and why, or what is still open.")

  // A keyed question settles the run's scope and needs the lawyer's own answer.
  const keyed = questionCard([{ ...asked, id: "package-scope-d1" }])
  expect(keyed).not.toContain("Decide for me")
  expect(keyed).not.toContain("review draft that says what was chosen")
  expect(keyed).toContain("Send answer")

  expect(delegableQuestion(asked)).toBe(true)
  expect(delegableQuestion({ ...asked, status: "answered" })).toBe(false)
  expect(delegableQuestion({ ...asked, id: "requested-files" })).toBe(false)
})

test("a hand-back posts exactly decide_for_me to the question route", async () => {
  const calls = respond(answer, 202)
  const question = blockedQuestion.questions[0]!
  expect(await decideLeanQuestion("matter-a", question.id)).toMatchObject({ profile: "lean", mode: "queued" })
  expect(calls).toHaveLength(1)
  expect(calls[0]?.url).toEndWith(`/matters/matter-a/lean/questions/${encodeURIComponent(question.id)}`)
  expect(calls[0]?.init?.method).toBe("POST")
  expect(JSON.parse(String(calls[0]?.init?.body))).toEqual({ decide_for_me: true })
})

test("a recorded hand-back reads as left to the drafter, not as an empty answer", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions = [
    { ...first, status: "answered", answer_kind: "decide_for_me" },
    { ...first, id: "q-answered", status: "answered", answer: "The buyer." },
    { ...first, id: "q-open", question: "Who bears the cost of storage?" },
  ]
  expect(handedBackQuestions(status).map((question) => question.id)).toEqual([first.id])
  expect(openQuestions(status).map((question) => question.id)).toEqual(["q-open"])

  const html = questionCard(openQuestions(status), { answered: 2, handedBack: handedBackQuestions(status) })
  expect(html).toContain("1 answered, 1 left to the drafter, 1 to answer")
  expect(html.match(/class="question-dot is-answered is-delegated"/g)).toHaveLength(1)
  expect(html.match(/class="question-dot is-answered"/g)).toHaveLength(1)
  // Without a hand-back the count reads as before.
  expect(questionCard(openQuestions(status), { answered: 2 })).toContain("2 answered, 1 left")
})

test("a refused hand-back surfaces as a plain message", async () => {
  const question = blockedQuestion.questions[0]!
  respond({ error: "This question needs the lawyer's own answer; it cannot be left to the drafter to decide.", code: "not_delegable" }, 409)
  await expect(decideLeanQuestion("matter-a", question.id)).rejects.toThrow("This question needs your own answer. It can't be left to the drafter.")
  respond({ error: "This question is already answered.", code: "already_answered" }, 409)
  await expect(decideLeanQuestion("matter-a", question.id)).rejects.toThrow("This question is already answered.")
  // Any other refusal keeps the service's own message.
  respond({ error: "This drafting run was stopped.", code: "lean_run_stopped" }, 409)
  await expect(decideLeanQuestion("matter-a", question.id)).rejects.toThrow("This drafting run was stopped.")

  // A refusal shows with the question it is about, never under another one.
  const open = { ...question, status: "open" as const }
  const other = { ...open, id: "q-other", question: "Who bears the cost of storage?" }
  const refused = questionCard([open, other], { error: "This question is already answered.", errorFor: open.id })
  expect(refused).toContain('<div class="msg-error" role="alert">This question is already answered.</div>')
  expect(questionCard([other, open], { error: "This question is already answered.", errorFor: open.id })).not.toContain("msg-error")
  expect(questionCard([open], { error: "This drafting run was stopped." })).not.toContain("msg-error")
})

test("a question left to the drafter can still be answered from the question card", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions = [{ ...first, status: "answered", answer_kind: "decide_for_me" }]
  const left = handedBackQuestions(status)
  expect(left.map((question) => question.id)).toEqual([first.id])
  expect(openQuestions(status)).toEqual([])

  // The card shows it as left to the drafter; its answer box posts the normal answer.
  const card = questionCard([], { answered: 1, handedBack: left })
  expect(card).toContain("Left to the drafter")
  expect(card).toContain(first.question)
  expect(card).toContain("Your answer")
  expect(card).toMatch(/<button class="primary" type="button" disabled="">Answer instead<\/button>/)
  expect(card).not.toContain("Decide for me")
  expect(card).not.toContain("Send answer")

  // With questions still asked, the left one follows them and is not offered back again.
  const asked = { ...first, id: "q-open", question: "Who bears the cost of storage?" }
  const mixed = questionCard([asked], { answered: 1, handedBack: left })
  expect(mixed).toContain("Question 1 of 2")
  expect(mixed).toContain("Who bears the cost of storage?")
  expect(mixed).toContain("Decide for me")
  expect(mixed).not.toContain("Left to the drafter")

  // The dock stays while a question is left to the drafter, so the lawyer can still answer it.
  const dock = (questions: LeanQuestion[], handedBack: LeanQuestion[]) =>
    renderToStaticMarkup(
      createElement(QuestionDock, {
        matterId: "m", ownerSessionID: "s", questions, answered: 1, handedBack, titles: new Map(), busy: false, error: "",
        onSend: async () => true, onDecide: async () => true,
      }),
    )
  const review = dock([], left)
  expect(review).toContain("1 question left to the drafter")
  expect(review).toMatch(/aria-haspopup="dialog"[^>]*>Review</)
  const both = dock([asked], left)
  expect(both).toContain("The drafting run has a question for you")
  expect(both).toMatch(/aria-haspopup="dialog"[^>]*>Answer</)
  expect(dock([], [])).toBe("")
  // A stopped run asks nothing and offers nothing back.
  expect(handedBackQuestions({ ...status, stopped: true })).toEqual([])
})

test("a hand-back moves no question: each keeps its number, and the current dot is its own", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions = [
    { ...first, status: "answered", answer_kind: "decide_for_me" },
    { ...first, id: "q-second", question: "Who bears the cost of storage?" },
  ]
  const order = status.questions.map((question) => question.id)
  const dots = (html: string) => [...html.matchAll(/<span class="(question-dot[^"]*)"><\/span>/g)].map((match) => match[1])
  // The first question was handed back: the second is still the second, and
  // the card opens on it, the one still asked.
  const html = questionCard(openQuestions(status), { answered: 1, handedBack: handedBackQuestions(status), order })
  expect(html).toContain("Question 2 of 2")
  expect(html).toContain("Who bears the cost of storage?")
  expect(dots(html)).toEqual(["question-dot is-answered is-delegated", "question-dot is-current"])

  // A question asked later follows those left to the drafter, under its own number.
  status.questions = [
    { ...first, status: "answered", answer_kind: "decide_for_me" },
    { ...first, id: "q-second", status: "answered", answer_kind: "decide_for_me", question: "Who bears the cost of storage?" },
    { ...first, id: "package-scope-d1", question: "Which documents should the package carry?" },
  ]
  const later = questionCard(openQuestions(status), { answered: 2, handedBack: handedBackQuestions(status), order: status.questions.map((question) => question.id) })
  expect(later).toContain("Question 3 of 3")
  expect(later).toContain("Which documents should the package carry?")
  expect(dots(later)).toEqual(["question-dot is-answered is-delegated", "question-dot is-answered is-delegated", "question-dot is-current"])

  // Answered questions leave the card as before; their dots come first.
  status.questions = [
    { ...first, status: "answered", answer: "The buyer." },
    { ...first, id: "q-second", status: "answered", answer_kind: "decide_for_me", question: "Who bears the cost of storage?" },
    { ...first, id: "q-third", question: "When is delivery due?" },
  ]
  const mixed = questionCard(openQuestions(status), { answered: 2, handedBack: handedBackQuestions(status), order: status.questions.map((question) => question.id) })
  expect(mixed).toContain("Question 2 of 2")
  expect(mixed).toContain("1 answered, 1 left to the drafter, 1 to answer")
  expect(dots(mixed)).toEqual(["question-dot is-answered", "question-dot is-answered is-delegated", "question-dot is-current"])
})

test("on a finished draft, a question left to the drafter says that answering it runs the drafting again", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions = [{ ...first, status: "answered", answer_kind: "decide_for_me" }]
  const left = handedBackQuestions(status)
  const note = "This draft is finished. Answering now runs the drafting again with your answer."
  const finished = questionCard([], { answered: 1, handedBack: left, done: true })
  expect(finished).toContain(note)
  expect(finished).toMatch(/<button class="primary" type="button" aria-describedby="[^"]+" disabled="">Answer instead<\/button>/)
  // While the run is still drafting, the answer is taken as the run goes on.
  expect(questionCard([], { answered: 1, handedBack: left })).not.toContain(note)
  // A question still asked on a finished run is answered as usual.
  expect(questionCard([{ ...first, id: "package-scope-d1" }], { answered: 1, handedBack: left, done: true })).not.toContain(note)
})

test("with nothing asked, the dock for questions left to the drafter does not call for the lawyer", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions = [{ ...first, status: "answered", answer_kind: "decide_for_me" }]
  const left = handedBackQuestions(status)
  const dock = (questions: LeanQuestion[]) =>
    renderToStaticMarkup(
      createElement(QuestionDock, {
        matterId: "m", ownerSessionID: "s", questions, answered: 1, handedBack: left, titles: new Map(), busy: false, error: "",
        onSend: async () => true, onDecide: async () => true,
      }),
    )
  expect(dock([])).toContain('<div class="question-dock is-settled">')
  expect(dock([{ ...first, id: "q-open", question: "Who bears the cost of storage?" }])).toContain('<div class="question-dock">')
})

test("a hand-back received but not yet recorded counts as left to the drafter, not answered", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions = [
    { ...first, status: "open" },
    { ...first, id: "q-later", status: "open", question: "Who bears the cost of storage?" },
    { ...first, id: "q-taken-back", status: "answered", answer_kind: "decide_for_me" },
  ]
  status.queued_answer_ids = [first.id, "q-taken-back"]
  status.queued_handed_back_ids = [first.id]
  expect(openQuestions(status).map((question) => question.id)).toEqual(["q-later"])
  // The queued hand-back is left to the drafter; the recorded one has a real answer queued.
  expect(handedBackQuestions(status).map((question) => question.id)).toEqual([first.id])
  const html = questionCard(openQuestions(status), { answered: 2, handedBack: handedBackQuestions(status) })
  expect(html).toContain("1 answered, 1 left to the drafter, 1 to answer")

  // An older ingest names no queued hand-backs: only recorded ones count, as before.
  const older = { ...status, queued_handed_back_ids: undefined }
  expect(handedBackQuestions(older).map((question) => question.id)).toEqual(["q-taken-back"])
})

test("a durably queued answer stops being asked before the coordinator applies it", () => {
  const status = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = status.questions[0]!
  status.questions.push({ ...first, id: "another-question", question: "Who should receive the summary?" })
  expect(openQuestions(status).map(question => question.id)).toEqual([first.id, "another-question"])
  status.queued_answer_ids = [first.id]
  expect(first.status).toBe("open")
  expect(openQuestions(status).map(question => question.id)).toEqual(["another-question"])
  // Reload consumes the same server projection; no optimistic browser-only flag.
  expect(openQuestions(JSON.parse(JSON.stringify(status))).map(question => question.id)).toEqual(["another-question"])
  first.status = "answered"
  first.answer = "The buyer."
  status.queued_answer_ids = []
  expect(openQuestions(status).map(question => question.id)).toEqual(["another-question"])
})


test("late drafting admission exposes only settled blocked checklist clarification", () => {
  const status = lean(structuredClone(blockedGate))
  status.stages["checklist-check"] = { ...status.stages.gate!, status: "blocked" }
  expect(draftingUpdateAdmission(status)).toBe("clarification")
  expect(draftingUpdateAdmission({ ...status, in_flight: true })).toBe("closed")
  expect(draftingUpdateAdmission({ ...status, pending_requests: 1 })).toBe("closed")
  for (const state of ["running", "completed", "failed"] as const)
    expect(draftingUpdateAdmission({ ...status, status: state })).toBe("closed")
  expect(draftingUpdateAdmission({ ...status, stages: { "d1.apply": { status: "done" } } })).toBe("closed")
  expect(draftingUpdateAdmission({ ...status, status: "running", stages: { "d1.draft": { status: "running" } } })).toBe("update")
})
