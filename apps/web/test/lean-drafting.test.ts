import { draftingUpdateAdmission } from "../../../dochaus/lib/drafting-update-admission"
import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
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
  isLeanDraftingStatus,
  leanDeliveredDocuments,
  resumeLeanDraftingRun,
  type LeanDraftingStatus,
} from "../src/api/ingest/lean"
import { draftingTurnRecoveryMode } from "../src/components/chat/drafting-recovery"
import { leanContinueAvailable, leanSettledItems, leanSettledSummary, runPosition, stepGroups, openQuestions } from "../src/components/drafting/model"
import Questions from "../src/components/drafting/Questions"
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

  const idle: RunActions = { busy: false, error: "", resuming: false, resume: async () => true, answer: async () => true }
  const render = (fixture: unknown) =>
    renderToStaticMarkup(createElement(RunHeader, { status: lean(fixture), actions: idle, own: true, onOpen: () => {} }))
  const done = render(completed)
  expect(done).toContain("Ready for review")
  expect(done).toContain("Open draft")
  expect(done).not.toContain("Continue drafting")
  // The steps and open items are kept behind Details, folded at first.
  expect(done).toContain('aria-expanded="false"')
  expect(done).toContain("No open items remain.")
  const asked = render(blockedQuestion)
  // The question is answered from the dock above the composer; the header points there.
  expect(asked).toContain("Drafting resumes after all answers are received")
  expect(asked).not.toMatch(/<button[^>]*>Answer/)
  expect(asked).not.toContain(blockedQuestion.questions[0]!.question)
  expect(asked).not.toContain("Continue drafting")
  expect(asked).toContain("held until the final check passes")
  expect(asked).not.toContain("`")
  const gated = render(blockedGate)
  expect(gated).toContain("Continue drafting")
  expect(gated).toContain("Final check settled 2 open items: 1 still blocking, 1 resolved.")
  expect(gated).not.toContain("recorded in the drafting notes")
  expect(render(repairing())).toContain("Open items: 2 other items to check.")
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
