import { afterEach, expect, test } from "bun:test"
import { draftingRunDrafted, draftingUpdateAdmission } from "../../../dochaus/lib/drafting-update-admission"
import answer from "../../../dochaus/workflows/lean/fixtures/answer.json"
import resume from "../../../dochaus/workflows/lean/fixtures/resume.json"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"
import refusedOpen from "../../../dochaus/workflows/lean/fixtures/start-refused-open.json"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import failed from "../../../dochaus/workflows/lean/fixtures/status-failed.json"
import orphaned from "../../../dochaus/workflows/lean/fixtures/status-orphaned.json"
import { startDraftingRun } from "../src/api/ingest/drafting-run"
import { getDraftingStatus } from "../src/api/ingest/drafting-status"
import {
  answerLeanQuestion,
  chooseLeanOption,
  decideLeanQuestion,
  isLeanDraftingStatus,
  leanDeliveredDocuments,
  pauseLeanDraftingRun,
  refineLeanChecklist,
  resumeLeanDraftingRun,
  stopLeanDraftingRun,
  type LeanDraftingStatus,
  type LeanQuestion,
} from "../src/api/ingest/lean"
import { draftingTurnRecoveryMode } from "../src/components/chat/drafting-recovery"
import {
  delegableQuestion,
  leanContinueAvailable,
  leanRunStopped,
  leanRunTakesUpdates,
  leanSettledItems,
  openQuestions,
  runPosition,
  stepGroups,
} from "../src/components/drafting/model"
import { runHeadline } from "../src/components/drafting/RunHeader"
import { continueModel } from "../src/components/drafting/useRunActions"
import { draftingStatusLabel } from "../src/components/matter/MatterRail"

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

function pausedRun() {
  const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
  return lean({
    ...blockedGate,
    stop_reason: "Paused at your request.",
    paused_at: "2026-10-17T09:00:00.000Z",
    stages: { ...blockedGate.stages, "d1.apply": pending, finalize: pending },
  })
}

function stoppedRun() {
  const { paused_at: _, ...run } = pausedRun()
  return lean({ ...run, status: "failed", stopped: true, stop_reason: "Stopped at your request.", pending_requests: 1, questions: blockedQuestion.questions })
}

function repairing() {
  const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
  return lean({
    ...blockedGate,
    status: "running",
    in_flight: true,
    stages: { ...blockedGate.stages, "d1.apply": { ...pending, status: "running" }, finalize: pending, gate: pending },
  })
}

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

test("start, status, and schema validation enforce lean contracts and reject malformed payloads", async () => {
  const calls = respond(start)
  const request = { request: "Draft a short supply note.", agent: "draft", model: { providerID: "test", modelID: "scripted" } }
  const result = await startDraftingRun("matter-a", request)
  expect(result).toMatchObject({ profile: "lean", mode: "started", ownerSessionID: "ses_child1" })
  expect(JSON.parse(String(calls[0]?.init?.body))).toEqual(request)

  respond(start)
  await expect(startDraftingRun("matter-a", { ...request, model: { providerID: "test", modelID: "other" } })).rejects.toThrow()

  for (const [index, fixture] of [completed, blockedGate, blockedQuestion, failed, orphaned].entries()) {
    const statusCalls = respond(fixture)
    expect(await getDraftingStatus(`matter-${index}`)).toMatchObject({ profile: "lean", status: fixture.status })
    expect(statusCalls[0]?.url).toEndWith(`/matters/matter-${index}/drafting-status`)
  }
  respond(null)
  expect(await getDraftingStatus("matter-none")).toBeUndefined()
  respond({ ...completed, stages: { ...completed.stages, gate: undefined } })
  await expect(getDraftingStatus("matter-x")).rejects.toThrow()

  const northwind = { party: "Northwind Traders Ltd", basis: "stated" }
  expect(isLeanDraftingStatus({ ...completed, plan: { ...completed.plan, acting_for: { party: "Northwind", basis: "yes" } } })).toBe(false)
  expect(isLeanDraftingStatus({ ...completed, plan: { ...completed.plan, acting_for: "Northwind" } })).toBe(false)
  expect(isLeanDraftingStatus({ ...completed, plan: { ...completed.plan, acting_for: northwind, posture: "aggressive" } })).toBe(false)
  expect(isLeanDraftingStatus({ ...completed, plan: { ...completed.plan, acting_for: northwind, posture: "conservative", held_back: 2 } })).toBe(true)
  for (const badHeldBack of ["2", -1, 1.5]) {
    expect(isLeanDraftingStatus({ ...completed, plan: { ...completed.plan, acting_for: northwind, held_back: badHeldBack } })).toBe(false)
  }
  expect(isLeanDraftingStatus({ ...stoppedRun(), stopped: "yes" })).toBe(false)
})

test("connection_wait validation and headline tone distinguish active reconnects from settled runs", async () => {
  const now = Date.parse(start.status.updated_at)
  const connection_wait = { last_successful_contact: now - 10_000, next_retry: now + 5_000 }
  const body = { ...start.status, connection_wait }
  respond(body)
  const status = lean(await getDraftingStatus("matter-reconnecting"))
  expect(status).toMatchObject({ status: "running", in_flight: true, connection_wait })
  expect(runHeadline(status, now).tone).toBe("running")
  expect(runHeadline(status, now).title).toBe("Reconnecting…")
  for (const fixture of [completed, blockedGate, failed, orphaned, pausedRun()]) {
    expect(runHeadline(lean({ ...fixture, connection_wait }), now).title).not.toBe("Reconnecting…")
  }
  expect(runHeadline(lean({ ...body, connection_wait: null }), now).title).not.toBe("Reconnecting…")
  for (const invalid of ["waiting", {}, { ...connection_wait, next_retry: "soon" }, { ...connection_wait, last_successful_contact: -1 }]) {
    expect(isLeanDraftingStatus({ ...body, connection_wait: invalid })).toBe(false)
  }
})

test("review draft with unresolved checklist findings is distinguished from clean completion", async () => {
  const finding = {
    id: "checklist-check:checklist_unresolved:item-1",
    stage: "checklist-check",
    kind: "checklist_unresolved",
    detail: "Checklist requirement not satisfied: delivery period.",
    status: "disclosed",
    ref: { checklist_item: "item-1", checklist_class: "source_conflict", document: "d1" },
  }
  const body = { ...completed, open_items: [finding], gate: { ...completed.gate, delivery: "review_draft_with_findings", disclosed: [finding], items: [finding] } }
  expect(isLeanDraftingStatus(body)).toBe(true)
  respond(body)
  const status = (await getDraftingStatus("matter-findings")) as LeanDraftingStatus
  expect(status.gate?.delivery).toBe("review_draft_with_findings")
  expect(draftingStatusLabel(status).findings).toBe(true)
  expect(draftingStatusLabel(lean(completed)).findings).toBe(false)
  expect(draftingStatusLabel(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft" } })).findings).toBe(false)
  expect(draftingStatusLabel(lean(blockedGate)).findings).toBe(false)
})

test("resume, pause, stop, question answer/option/decide, and checklist refine routes send exact payloads", async () => {
  const model = { providerID: "test", modelID: "selected" }
  const resumed = respond(resume, 202)
  expect(await resumeLeanDraftingRun("matter-a", model)).toMatchObject({ profile: "lean", mode: "queued", status: { status: "running" } })
  expect(resumed[0]?.url).toEndWith("/matters/matter-a/lean/resume")
  expect(JSON.parse(String(resumed[0]?.init?.body))).toEqual({ model })

  expect(await continueModel(Promise.resolve({ model: "test/selected" }))).toEqual({ providerID: "test", modelID: "selected" })
  expect(await continueModel(Promise.reject(new Error("offline")))).toBeUndefined()
  expect(await continueModel(new Promise(() => {}), 10)).toBeUndefined()
  const resumedNoModel = respond(resume, 202)
  expect(await resumeLeanDraftingRun("matter-a")).toMatchObject({ profile: "lean", mode: "queued" })
  expect(JSON.parse(String(resumedNoModel[0]?.init?.body))).toEqual({})

  const paused = respond({ profile: "lean", mode: "paused", status: pausedRun() })
  expect(await pauseLeanDraftingRun("matter-a", "ses_child1")).toMatchObject({ mode: "paused" })
  expect(paused[0]?.url).toEndWith("/matters/matter-a/lean/pause")
  expect(JSON.parse(String(paused[0]?.init?.body))).toEqual({ ownerSessionID: "ses_child1" })
  respond({ profile: "lean", mode: "queued", request_id: "r1", status: pausedRun() })
  await expect(pauseLeanDraftingRun("matter-a", "ses_child1")).rejects.toThrow()

  const stopped = respond({ profile: "lean", mode: "stopped", status: stoppedRun() })
  expect(await stopLeanDraftingRun("matter-a", "ses_child1")).toMatchObject({ mode: "stopped" })
  expect(stopped[0]?.url).toEndWith("/matters/matter-a/lean/stop")
  expect(JSON.parse(String(stopped[0]?.init?.body))).toEqual({ ownerSessionID: "ses_child1" })
  respond({ profile: "lean", mode: "paused", status: pausedRun() })
  await expect(stopLeanDraftingRun("matter-a", "ses_child1")).rejects.toThrow()

  const route = `/matters/matter-a/lean/questions/${encodeURIComponent(advised.id)}`
  const answered = respond(answer, 202)
  await answerLeanQuestion("matter-a", advised.id, "No remedy.")
  expect(answered[0]?.url).toEndWith(route)
  expect(JSON.parse(String(answered[0]?.init?.body))).toEqual({ answer: "No remedy." })

  const chosen = respond(answer, 202)
  await chooseLeanOption("matter-a", advised.id, "o2")
  expect(JSON.parse(String(chosen[0]?.init?.body))).toEqual({ option_id: "o2" })

  const decided = respond(answer, 202)
  await decideLeanQuestion("matter-a", advised.id)
  expect(JSON.parse(String(decided[0]?.init?.body))).toEqual({ decide_for_me: true })

  const refined = respond({ profile: "lean", mode: "queued", request_id: "queued-refinement", status: { ...blockedGate, pending_requests: 1 } }, 202)
  const refineResult = await refineLeanChecklist("matter-check", "owner-conversation", "private-requirement-id")
  expect(refineResult.status.pending_requests).toBe(1)
  expect(refined[0]?.url).toEndWith("/matters/matter-check/lean/checklist/refine")
  expect(JSON.parse(String(refined[0]?.init?.body))).toEqual({
    ownerSessionID: "owner-conversation",
    parents: [{ parent_id: "private-requirement-id", scope: "target_body" }],
  })
  respond({ error: "The active drafting run changed.", code: "lean_run_changed" }, 409)
  await expect(refineLeanChecklist("matter-check", "old-owner", "private-requirement-id")).rejects.toThrow("The active drafting run changed.")

  respond(refusedOpen, 409)
  await expect(startDraftingRun("matter-a", { request: "Draft", agent: "draft", model })).rejects.toThrow()
  respond({ error: "A lean drafting run is already in progress.", code: "lean_run_in_progress" }, 409)
  await expect(startDraftingRun("matter-a", { request: "Draft", agent: "draft", model })).rejects.toThrow()
  respond({ error: "Already answered.", code: "already_answered" }, 409)
  await expect(decideLeanQuestion("matter-a", advised.id)).rejects.toThrow("already answered")
  respond({ error: "Invalid option.", code: "invalid_answer" }, 400)
  await expect(chooseLeanOption("matter-a", advised.id, "o9")).rejects.toThrow()
})

test("run state predicates govern Continue, updates, stopped flag, delivered docs, and headline tones", () => {
  expect([blockedGate, blockedQuestion, failed, orphaned].map((fixture) => leanContinueAvailable(lean(fixture)))).toEqual([true, true, true, true])
  expect([completed, start.status].map((fixture) => leanContinueAvailable(lean(fixture)))).toEqual([false, false])
  expect(leanDeliveredDocuments(lean(completed))).toEqual(["Supply note.docx", "Supply note - redline.docx", "Supply note - drafting notes.docx"])
  expect(leanDeliveredDocuments(lean(blockedGate))).toEqual([])

  const stopped = stoppedRun()
  expect(openQuestions(stopped)).toEqual([])
  expect(leanContinueAvailable(stopped)).toBe(false)
  expect([pausedRun(), lean(failed), lean(blockedQuestion), stopped, lean(completed)].map(leanRunTakesUpdates)).toEqual([true, true, true, false, false])

  // Only ingest's derived stopped flag marks a run stopped; stop_reason text alone does not.
  const worded = lean({ ...failed, stop_reason: "Stopped at your request." })
  expect([lean(failed), worded, stopped].map(leanRunStopped)).toEqual([false, false, true])
  expect(leanContinueAvailable(worded)).toBe(true)

  const now = Date.parse(completed.updated_at)
  expect(runHeadline(lean(completed), now).tone).toBe("done")
  expect(runHeadline(lean(blockedQuestion), now).tone).toBe("input")
  expect(runHeadline(lean(start.status), now).tone).toBe("running")
  expect([blockedGate, failed, orphaned, stopped].map((fixture) => runHeadline(lean(fixture), now).tone)).toEqual(["paused", "paused", "paused", "paused"])
  expect(draftingTurnRecoveryMode("token")).toBe("malformed-tool")
  expect(draftingTurnRecoveryMode(undefined)).toBe("none")
})

test("delegableQuestion enforces kind, status, mixed flag, and valid recommended option", () => {
  expect(delegableQuestion(advised)).toBe(true)
  expect(delegableQuestion({ ...advised, id: "package-scope-d1", kind: "scope" })).toBe(true)
  expect(delegableQuestion({ ...advised, kind: "instruction" })).toBe(true)
  for (const kind of ["fact", "conflict", "file_name"] as const) expect(delegableQuestion({ ...advised, kind })).toBe(false)
  expect(delegableQuestion({ ...advised, kind: "scope", mixed: true })).toBe(false)
  expect(delegableQuestion({ ...advised, advice: { unavailable: { reason: "No source." } } })).toBe(false)
  expect(delegableQuestion({ ...advised, kind: undefined })).toBe(false)
  expect(delegableQuestion({ ...advised, advice: undefined })).toBe(false)
  expect(delegableQuestion({ ...advised, status: "answered" })).toBe(false)
  const advice = advised.advice as Extract<NonNullable<LeanQuestion["advice"]>, { recommendation: unknown }>
  expect(delegableQuestion({ ...advised, advice: { ...advice, recommendation: { ...advice.recommendation, option: "o9" } } })).toBe(false)
})

test("open items, step groups, queued answers, and late drafting admission follow stage state", () => {
  const settled = leanSettledItems(lean(blockedGate))
  expect(settled?.map((item) => item.status)).toEqual(["resolved", "open"])
  expect(leanSettledItems(lean(completed))).toEqual([])
  expect(leanSettledItems(repairing())).toBeNull()

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
  expect(stepGroups(two).map((group) => [group.title, group.steps.length])).toEqual([
    ["Supply note", 4],
    ["Covering letter", 4],
    ["Finishing", 2],
  ])
  expect(runPosition(two)).toEqual({ text: "Document 2 of 2 · Review", document: "Covering letter" })
  expect(isLeanDraftingStatus({ ...two, stages: { "d1.draft": done } })).toBe(false)

  const queued = structuredClone(blockedQuestion) as LeanDraftingStatus
  const first = queued.questions[0]!
  queued.questions.push({ ...first, id: "another-question", question: "Who should receive the summary?" })
  expect(openQuestions(queued).map((question) => question.id)).toEqual([first.id, "another-question"])
  queued.queued_answer_ids = [first.id]
  expect(openQuestions(queued).map((question) => question.id)).toEqual(["another-question"])

  const checklistBlocked = lean(structuredClone(blockedGate))
  checklistBlocked.stages["checklist-check"] = { ...checklistBlocked.stages.gate!, status: "blocked" }
  expect(draftingUpdateAdmission(checklistBlocked)).toBe("clarification")
  expect(draftingUpdateAdmission({ ...checklistBlocked, in_flight: true })).toBe("closed")
  expect(draftingUpdateAdmission({ ...checklistBlocked, pending_requests: 1 })).toBe("closed")
  for (const state of ["running", "completed", "failed"] as const) {
    expect(draftingUpdateAdmission({ ...checklistBlocked, status: state })).toBe("closed")
  }
  expect(draftingUpdateAdmission({ ...pausedRun(), stages: { ...blockedGate.stages, "checklist-check": { status: "blocked" } } })).toBe("closed")
  expect(draftingUpdateAdmission({ ...pausedRun(), stages: { ...pausedRun().stages, "checklist-check": { status: "blocked" } } })).toBe("update")

  const doneStage = { status: "done" }
  const rerun = { ...checklistBlocked, status: "failed", stages: { spine: { status: "failed" }, "d1.draft": doneStage, "d1.apply": doneStage, finalize: doneStage, "checklist-check": { status: "pending" }, gate: { status: "pending" } } }
  expect(draftingUpdateAdmission(rerun)).toBe("update")
  expect(draftingRunDrafted(rerun)).toBe(false)
  const checking = { ...checklistBlocked, status: "running", in_flight: true, stages: { spine: doneStage, "d1.draft": doneStage, "d1.apply": doneStage, finalize: doneStage, "checklist-check": { status: "running" }, gate: { status: "pending" } } }
  expect(draftingUpdateAdmission(checking)).toBe("closed")
  expect(draftingRunDrafted(checking)).toBe(true)
  expect(draftingRunDrafted({ stages: { spine: doneStage, "d1.draft": { status: "running" } } })).toBe(false)
})
