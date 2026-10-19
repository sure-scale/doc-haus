import { draftingUpdateAdmission } from "../../../dochaus/lib/drafting-update-admission"
import { afterEach, expect, test } from "bun:test"
import path from "node:path"
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
  chooseLeanOption,
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
import { delegableQuestion, leanContinueAvailable, leanRunStopped, leanRunTakesUpdates, leanSettledItems, leanSettledSummary, runPosition, stepGroups, openQuestions } from "../src/components/drafting/model"
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

  const idle: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, choose: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => true, stop: async () => true }
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
  // The question is answered from the dock above the composer; the header points there.
  expect(asked).toContain("1 question remaining · Drafting resumes after all answers are received")
  // Decide for me answers with the recommendation, so a question it can take reads the same.
  const judgment = { kind: "judgment", advice: { options: [{ id: "o1", text: "No remedy" }, { id: "o2", text: "A remedy" }], recommendation: { option: "o2", rationale: "Asked for.", basis: [{ source_id: "request", quote: "a remedy" }] } } }
  const delegable = { ...blockedQuestion, questions: [...blockedQuestion.questions, { ...blockedQuestion.questions[0]!, ...judgment, id: "q-second", question: "Second?" }] }
  expect(delegableQuestion(lean(delegable).questions[1]!)).toBe(true)
  expect(render(delegable)).toContain("2 questions remaining · Drafting resumes after all answers are received")
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
  const actions: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, choose: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => (paused++, true), stop: async () => true }
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
  const actions: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, choose: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => true, stop: async () => (stops++, true) }
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
      onChoose: async () => false,
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

function questionCard(questions: LeanQuestion[], extra: { answered?: number; error?: string; errorFor?: string } = {}) {
  respond({ answers: {}, revision: 0 })
  return renderToStaticMarkup(
    createElement(Questions, {
      matterId: "matter-decide",
      ownerSessionID: blockedQuestion.owner_session_id,
      questions,
      answered: extra.answered ?? 0,
      titles: new Map(),
      busy: false,
      error: extra.error ?? "",
      errorFor: extra.errorFor,
      titleId: "q-title",
      onClose: () => {},
      onSend: async () => false,
      onChoose: async () => false,
      onDecide: async () => false,
    }),
  )
}

// A judgment question as the run records it once advised: the advice sets out
// the options and recommends one, with the passages behind it.
const advised: LeanQuestion = {
  id: "q-remedy",
  stage: "changes",
  document: "d1",
  question: "Should late delivery carry a remedy?",
  context: "The terms are silent.",
  status: "open",
  kind: "judgment",
  advice: {
    options: [
      { id: "o1", text: "No remedy" },
      { id: "o2", text: "Liquidated damages at 1% a week" },
      { id: "o3", text: "A right to terminate after 30 days" },
    ],
    recommendation: {
      option: "o2",
      rationale: "The buyer asked for a remedy that does not need proof of loss.",
      basis: [
        { source_id: "request", quote: "a remedy that does not need proof of loss" },
        { source_id: "Supply terms.docx", quote: "Delivery is due within 14 days." },
      ],
    },
  },
}
const optionTexts = (html: string) => [...html.matchAll(/<span class="question-option-text">([^<]*)<\/span>/g)].map((match) => match[1])
const unavailable = (question: LeanQuestion, reason = "The sources do not say which remedy the buyer wants."): LeanQuestion =>
  ({ ...question, advice: { unavailable: { reason } } })

test("a question's options appear in the order the run gives them, with none chosen for the lawyer", () => {
  const html = questionCard([advised])
  expect(html).toMatch(/<span id="[^"]+" class="question-label">Your answer<\/span><div role="group" aria-labelledby="[^"]+" class="question-options">/)
  // The recommended option keeps its place: the list is never reordered around it.
  expect(optionTexts(html)).toEqual(["No remedy", "Liquidated damages at 1% a week", "A right to terminate after 30 days"])
  expect(html).not.toContain("aria-pressed")
  expect(html).not.toContain("checked")
  expect(html).not.toContain("autofocus")
  expect(html).not.toContain("is-selected")

  // A question's own options come before any the advice sets out, in their own order.
  const own: LeanQuestion = {
    ...advised,
    id: "requested-files-d1",
    kind: "file_name",
    options: [{ id: "o1", text: "As my words ask" }, { id: "o2", text: "As planned" }],
    advice: { recommendation: { option: "o2", rationale: "The plan follows the request.", basis: [{ source_id: "request", quote: "a supply note" }] } },
  }
  expect(optionTexts(questionCard([own]))).toEqual(["As my words ask", "As planned"])
})

test("the recommended option is marked, with why and the quoted passages and where each comes from", () => {
  const html = questionCard([advised])
  // Only the recommended option carries the mark, and it points at why.
  expect(html.match(/class="question-option-tag">Recommended</g)).toHaveLength(1)
  expect(html).toMatch(
    /<button type="button" class="question-option is-recommended" aria-describedby="([^"]+)" disabled=""><span class="question-option-text">Liquidated damages at 1% a week<\/span><span class="question-option-tag">Recommended<\/span><\/button>/,
  )
  const why = html.match(/class="question-option is-recommended" aria-describedby="([^"]+)"/)?.[1]
  expect(html).toContain(`<div id="${why}" class="question-advice">`)
  expect(html).toContain('<p class="question-advice-why"><span class="question-advice-label">Why</span> The buyer asked for a remedy that does not need proof of loss.</p>')
  // Each quote shows with its source: the lawyer's own instructions, or the matter file by name.
  expect(html).toContain("<blockquote>a remedy that does not need proof of loss</blockquote><cite>Your instructions</cite>")
  expect(html).toContain("<blockquote>Delivery is due within 14 days.</blockquote><cite>Supply terms.docx</cite>")
  expect(html).not.toContain("source_id")
})

test("Other, the lawyer's own words, is always offered", () => {
  // Beside options it is Other; with none, it is the answer.
  const options = questionCard([advised])
  expect(options).toMatch(/<label class="question-label" for="lean-answer-q-remedy">Other<\/label><textarea[^>]* id="lean-answer-q-remedy"/)
  expect(options).toContain("Send answer")
  for (const question of [unavailable(advised), { ...advised, advice: undefined }, { ...advised, kind: "fact" as const }]) {
    const html = questionCard([question])
    expect(html).toContain('<textarea')
    expect(html).toContain("Send answer")
  }
  expect(questionCard([unavailable(advised)])).toMatch(/<label class="question-label" for="lean-answer-q-remedy">Your answer<\/label>/)
})

test("Decide for me is offered only where the run may take the recommendation: judgment and pure scope questions", () => {
  const offered = (question: LeanQuestion) => questionCard([question]).includes(">Decide for me</button>")
  expect(offered(advised)).toBe(true)
  // Decided by the question's kind and advice, never by its id.
  expect(offered({ ...advised, id: "package-scope-d1", kind: "scope" })).toBe(true)
  for (const kind of ["fact", "instruction", "conflict", "file_name"] as const) expect(offered({ ...advised, kind })).toBe(false)
  // A scope card that also asks for facts or approvals needs the lawyer's own answer.
  expect(offered({ ...advised, kind: "scope", mixed: true })).toBe(false)
  expect(offered(unavailable(advised))).toBe(false)
  expect(offered({ ...advised, kind: undefined })).toBe(false)
  expect(offered({ ...advised, advice: undefined })).toBe(false)

  expect(delegableQuestion(advised)).toBe(true)
  expect(delegableQuestion({ ...advised, id: "intake-1" })).toBe(true)
  expect(delegableQuestion({ ...advised, kind: "fact" })).toBe(false)
  expect(delegableQuestion({ ...advised, status: "answered" })).toBe(false)
  const advice = advised.advice as Extract<NonNullable<LeanQuestion["advice"]>, { recommendation: unknown }>
  // A recommendation of an option the question does not offer cannot be taken.
  expect(delegableQuestion({ ...advised, advice: { ...advice, recommendation: { ...advice.recommendation, option: "o9" } } })).toBe(false)

  const html = questionCard([advised])
  // Never over a typed answer: until saved answers load the box may hold one, so
  // Decide for me and the options wait like Send (enabled while it is empty).
  expect(html).toContain("Loading saved answers…")
  expect(html).toMatch(/<button type="button" aria-describedby="[^"]+" disabled="">Decide for me<\/button>/)
  expect(html).toMatch(/<button type="button" class="question-option" disabled="">/)
  // Decide for me and Send stay together when the actions wrap.
  expect(html).toMatch(/<span class="question-send"><button[^>]*>Decide for me<\/button><button class="primary"/)
  expect(html).toContain('class="question-delegate-hint">Decide for me answers with the recommended option.</p>')
  expect(questionCard([{ ...advised, kind: "fact" }])).not.toContain("recommended option.</p>")
})

test("advice with no recommendation says why, marks no option and offers no Decide for me", () => {
  const html = questionCard([unavailable(advised)])
  expect(html).toContain('<p class="question-advice is-unavailable">No recommendation: The sources do not say which remedy the buyer wants.</p>')
  expect(html).not.toContain("Decide for me")
  expect(html).not.toContain("Recommended")
  expect(html).not.toContain("question-options")
  // Options the advice set out are still offered, none of them marked.
  const advice = advised.advice as Extract<NonNullable<LeanQuestion["advice"]>, { recommendation: unknown }>
  const listed = questionCard([{ ...advised, advice: { options: advice.options, unavailable: { reason: "No source settles it." } } }])
  expect(optionTexts(listed)).toEqual(["No remedy", "Liquidated damages at 1% a week", "A right to terminate after 30 days"])
  expect(listed).toContain("No recommendation: No source settles it.")
  expect(listed).not.toContain("Recommended")
  expect(listed).not.toContain("Decide for me")
})

test("choosing an option posts its id, and Decide for me posts decide_for_me: true", async () => {
  const route = `/matters/matter-a/lean/questions/${encodeURIComponent(advised.id)}`
  const chosen = respond(answer, 202)
  expect(await chooseLeanOption("matter-a", advised.id, "o2")).toMatchObject({ profile: "lean", mode: "queued" })
  expect(chosen).toHaveLength(1)
  expect(chosen[0]?.url).toEndWith(route)
  expect(chosen[0]?.init?.method).toBe("POST")
  expect(JSON.parse(String(chosen[0]?.init?.body))).toEqual({ option_id: "o2" })

  const decided = respond(answer, 202)
  expect(await decideLeanQuestion("matter-a", advised.id)).toMatchObject({ profile: "lean", mode: "queued" })
  expect(decided[0]?.url).toEndWith(route)
  expect(JSON.parse(String(decided[0]?.init?.body))).toEqual({ decide_for_me: true })
})

test("a refused answer surfaces as a plain message, with the question it is about", async () => {
  respond({ error: "This question has no recommendation to adopt; it needs the lawyer's own answer.", code: "not_delegable" }, 409)
  await expect(decideLeanQuestion("matter-a", advised.id)).rejects.toThrow("Decide for me isn't available for this question. Choose an option or write your own answer.")
  respond({ error: "This question is already answered.", code: "already_answered" }, 409)
  await expect(decideLeanQuestion("matter-a", advised.id)).rejects.toThrow("This question is already answered.")
  respond({ error: "That option is not one this question offers.", code: "invalid_answer" }, 400)
  await expect(chooseLeanOption("matter-a", advised.id, "o9")).rejects.toThrow("This answer can't be sent. Choose one of the options shown, or write an answer of up to 20,000 characters.")
  respond({ error: "An answer must be non-empty text of at most 20,000 characters.", code: "invalid_answer" }, 400)
  await expect(answerLeanQuestion("matter-a", advised.id, " ")).rejects.toThrow("This answer can't be sent. Choose one of the options shown, or write an answer of up to 20,000 characters.")
  // Any other refusal keeps the service's own message.
  respond({ error: "This drafting run was stopped.", code: "lean_run_stopped" }, 409)
  await expect(decideLeanQuestion("matter-a", advised.id)).rejects.toThrow("This drafting run was stopped.")

  // A refusal shows with the question it is about, never under another one.
  const other = { ...advised, id: "q-other", question: "Who bears the cost of storage?" }
  const refused = questionCard([advised, other], { error: "This question is already answered.", errorFor: advised.id })
  expect(refused).toContain('<div class="msg-error" role="alert">This question is already answered.</div>')
  expect(questionCard([other, advised], { error: "This question is already answered.", errorFor: advised.id })).not.toContain("msg-error")
  expect(questionCard([advised], { error: "This drafting run was stopped." })).not.toContain("msg-error")
})

test("nothing in the web app says a question is left to the drafter or handed back", async () => {
  const root = path.resolve(import.meta.dir, "../src")
  const found: string[] = []
  for await (const file of new Bun.Glob("**/*.{ts,tsx,css,html,json,md}").scan({ cwd: root }))
    if (/left to the drafter|hand(?:ed|ing)?[ -]?back/i.test(await Bun.file(path.join(root, file)).text())) found.push(file)
  expect(found).toEqual([])
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
