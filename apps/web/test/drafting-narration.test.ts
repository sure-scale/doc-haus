import { expect, test } from "bun:test"
import type { SessionMessageInfo } from "@opencode/client"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import failed from "../../../dochaus/workflows/lean/fixtures/status-failed.json"
import { PROVIDER_STOP_MESSAGES } from "../../../dochaus/lib/provider-stop"
import { isLeanDraftingStatus, leanStageSessions, type LeanDraftingStatus } from "../src/api/ingest/lean"
import { narratorNotes, narratorTitle } from "../src/api/opencode/narrator"
import type { Part } from "../src/api/session-view"
import { stepLabel } from "../src/components/chat/step-labels"
import {
  activeStep,
  fileLabel,
  partActivity,
  reasoningHeadline,
  runTransition,
  stageMilestones,
  stoppedByTechnicalProblem,
  type ActivityLine,
} from "../src/components/narration/activity"
import { buildExcerpt, EXCERPT_LIMIT, stageSessionNotes } from "../src/components/narration/excerpt"
import { resultChecks, runOutcome } from "../src/components/narration/outcome"
import {
  checkpointKeys,
  NARRATION_DEBOUNCE_MS,
  NARRATION_FAILURE_LIMIT,
  NARRATION_MIN_GAP_MS,
  NARRATION_PROGRESS_MS,
  planNarration,
} from "../src/components/narration/schedule"
import { draftedIn, elapsed, narrationTimeline, rollUp, stepRows, VISIBLE_STEPS } from "../src/components/narration/timeline"

function lean(value: unknown): LeanDraftingStatus {
  if (!isLeanDraftingStatus(value)) throw new Error("fixture is not a lean drafting status")
  return value
}

const running = lean(start.status)
const done = lean(completed)
const question = lean(blockedQuestion)
const stopped = lean(failed)

function tool(name: string, state: Extract<Part, { type: "tool" }>["state"], id = `part-${name}`): Part {
  return { type: "tool", id, sessionID: "ses_child2", messageID: "msg", tool: name, state }
}

function activity(id: string, name: string, input: Record<string, unknown>, output = "FOUND") {
  const line = partActivity(tool(name, { status: "completed", input, output, title: "" }, id), 1, "matter")
  if (!line) throw new Error(`no activity for ${id}`)
  return line
}

function withFindings(count: number, others: LeanDraftingStatus["open_items"] = []) {
  const findings = Array.from({ length: count }, (_, i) => ({
    id: `checklist-check:checklist_unresolved:${i + 1}`,
    stage: "checklist-check",
    kind: "checklist_unresolved",
    detail: `Checklist requirement not satisfied: Requirement ${i + 1}.`,
    status: "disclosed" as const,
    ref: { checklist_item: `item-${i + 1}`, checklist_class: "omission" },
  }))
  const items = [...findings, ...others]
  return lean({ ...done, gate: { ...done.gate!, delivery: "review_draft_with_findings", disclosed: items, items } })
}

function narrationNote(id: string, ...keys: string[]) {
  return { id, at: Number(id.slice(1)), text: `Drafting is stopped pending your answers (${id}).`, keys }
}

test("tool and reasoning parts map to activity lines, clip long queries, and omit folder targets", () => {
  expect(partActivity(tool("read-document", { status: "running", input: { document: "sources/Supply Terms.docx" }, title: "" }), 5, "matter")).toMatchObject({
    id: "part-read-document",
    at: 5,
    kind: "tool",
    state: "running",
    target: "Supply Terms.docx",
  })
  expect(fileLabel("/m/.dochaus/runs/r1/plan.json")).not.toContain(".dochaus")
  expect(fileLabel(".dochaus/lean/coverage.json")).not.toContain(".dochaus")
  expect(fileLabel(".dochaus/runs/r1/other.txt")).not.toContain(".dochaus")

  for (const path of ["/m/matter", "/m/matter/", "C:\\m\\matter", "/m/matter/sources"]) {
    const folder = activity(`f:${path}`, "read", { path })
    expect(folder.group).toBe("read")
    expect(folder.target).toBeUndefined()
  }
  expect(activity("f-doc", "read", { filePath: "/m/matter/Order form.pdf" }).target).toBe("Order form.pdf")

  const readPart = tool("read", { status: "completed", input: { path: "/m/sources/Supply Terms.docx" }, output: "", title: "" }, "p3")
  if (readPart.type !== "tool") throw new Error("expected a tool part")
  expect(stepLabel(readPart, "/m")).toContain("Supply Terms.docx")

  const longQuery = "delivery obligations and the remedies available when goods arrive late or damaged in transit"
  const clipped = activity("s1", "search-document", { document: "a/terms.docx", query: longQuery })
  expect(clipped.text.length).toBeLessThan(longQuery.length + 30)
  expect(clipped.detail).toBe(`"${longQuery}"`)
  expect(activity("s2", "search-document", { document: "a/terms.docx", query: "delivery" }).detail).toBeUndefined()

  const missing = activity("q-miss", "verify-quote", { document: "a/terms.docx" }, "NOT FOUND: …")
  expect(missing).toMatchObject({ state: "done", flag: true })
  expect(activity("q-ok", "verify-quote", { document: "a/terms.docx" }, "FOUND").flag).toBeFalsy()

  const thought: Part = {
    type: "reasoning",
    id: "r1",
    sessionID: "s",
    messageID: "m",
    text: "**Mapping the sources**\nSome notes.\n\n**Weighing the remedies**\nMore notes.",
    time: { start: 7 },
  }
  expect(partActivity(thought, 9, "matter")).toMatchObject({ at: 7, kind: "thought", state: "running", text: "Weighing the remedies" })
  expect(reasoningHeadline("The draft needs a schedule. Then more.")).toBe("The draft needs a schedule.")
  expect(reasoningHeadline("   ")).toBeUndefined()
  expect(partActivity({ type: "text", id: "t", sessionID: "s", messageID: "m", text: "answer" }, 1, "matter")).toBeUndefined()
})

test("consecutive grouped tools fold with flags and errors break folding", () => {
  const lines = [
    activity("r1", "read-document", { document: "a/one.docx" }),
    activity("r2", "read-document", { document: "a/two.docx" }),
    activity("q1", "verify-quote", { document: "a/one.docx" }, "NOT FOUND: x"),
    activity("q2", "verify-quote", { document: "a/two.docx" }),
    activity("q3", "verify-quote", { document: "a/two.docx" }, "NOT FOUND: y"),
  ]
  const rows = stepRows(lines)
  expect(rows.map((row) => [row.id, row.count, row.target, row.flag, row.flagged])).toEqual([
    ["r1", 2, "two.docx", false, 0],
    ["q1", 3, "two.docx", true, 2],
  ])

  const failedRead = partActivity(tool("read-document", { status: "error", input: { document: "a/one.docx" }, error: "x", title: "" }, "e1"), 1, "matter")!
  expect(failedRead.group).toBeUndefined()
  expect(stepRows([lines[0]!, failedRead, lines[1]!])).toHaveLength(3)

  const withFolder = [
    activity("r1", "read", { path: "/m/matter/Supply Terms.docx" }),
    activity("r2", "read", { path: "/m/matter/Order form.pdf" }),
    activity("r3", "read", { path: "/m/matter" }),
  ]
  expect(stepRows(withFolder).map((row) => [row.count, row.target])).toEqual([[3, undefined]])
})

test("stage milestones track active step, stable stage IDs, retries, pauses, and technical stops", () => {
  expect(activeStep(running)).toBe("d1.draft")
  expect(stageMilestones(running)[0]).toMatchObject({ state: "running", id: `d1.draft:stage:${running.stages["d1.draft"]!.started_at}` })

  const finished = stageMilestones(done)
  expect(finished).toHaveLength(6)
  const draft = done.stages["d1.draft"]!
  expect(finished[0]).toMatchObject({
    id: `d1.draft:stage:${draft.started_at}`,
    step: "d1.draft",
    at: Date.parse(draft.started_at!),
    until: Date.parse(draft.finished_at!),
  })
  expect(stageMilestones(question).at(-1)?.state).toBe("paused")
  expect(stageMilestones(stopped).at(-1)?.state).toBe("error")
  expect(stageMilestones(lean({ ...running, in_flight: false }))[0]?.state).toBe("paused")
  expect(draftedIn(running)).toBeUndefined()
  expect(draftedIn(done)).toMatch(/^(under a minute|\d+m|\d+h \d+m)$/)

  const fault = lean({ ...stopped, stop_reason: "d1.draft: stopped by a runtime error: provider unavailable" })
  expect(stoppedByTechnicalProblem(fault)).toBe(true)
  expect(stoppedByTechnicalProblem(stopped)).toBe(false)
  expect(stageMilestones(fault).at(-1)?.state).toBe("paused")
  expect(runTransition(running, fault)).toBeUndefined()
  expect(runTransition(fault, running)?.state).toBe("done")
  expect(runTransition(running, lean({ ...running, in_flight: false }))?.state).toBe("paused")
  expect(runTransition(question, done)).toBeUndefined()
  expect(runTransition(running, running)).toBeUndefined()
})

test("run outcomes and result checks reflect delegability, provider stops, post-check edits, and unresolved findings", () => {
  expect(runOutcome(running)).toBeUndefined()
  expect(runOutcome(done)?.tone).toBe("done")
  expect(runOutcome(question)?.tone).toBe("attention")

  const asked = question.questions[0]!
  const advice = { options: [{ id: "o1", text: "No remedy" }, { id: "o2", text: "A remedy" }], recommendation: { option: "o2", rationale: "Asked for.", basis: [{ source_id: "request", quote: "a remedy" }] } }
  const withQuestions = (...kinds: Array<"judgment" | "scope" | "fact">) =>
    lean({ ...question, questions: kinds.map((kind, index) => ({ ...asked, id: `q-${index}`, kind, advice })) })
  expect(runOutcome(withQuestions("judgment"))?.text).toContain("Decide for me")
  expect(runOutcome(withQuestions("fact"))?.text).not.toContain("Decide for me")

  const expired = lean({ ...stopped, stop_reason: `d1.draft: ${PROVIDER_STOP_MESSAGES.auth}`, provider_stop: "auth" })
  expect(runOutcome(expired)).toEqual({ tone: "attention", text: PROVIDER_STOP_MESSAGES.auth })
  const spent = lean({ ...stopped, stop_reason: `d1.draft: ${PROVIDER_STOP_MESSAGES.credit}`, provider_stop: "credit" })
  expect(runOutcome(spent)).toEqual({ tone: "attention", text: PROVIDER_STOP_MESSAGES.credit })
  expect(buildExcerpt({ status: spent, plan: { keys: [`run:failed:${spent.updated_at}`], focus: `run:failed:${spent.updated_at}`, reason: "checkpoint" }, activity: [], now: 0 })).not.toContain("Settings")

  expect(resultChecks(done).map((c) => c.tone)).toEqual(["pass", "pass"])
  expect(resultChecks(lean({ ...done, gate: null }))).toEqual([])
  const item = (id: string, kind: string, status: "open" | "resolved" | "disclosed") => ({ id, stage: "d1.review", kind, detail: id, status })
  const disclosed = lean({
    ...done,
    gate: { ...done.gate!, items: [item("q1", "quote_miss", "disclosed"), item("q2", "quote_miss", "resolved"), item("p1", "placeholder", "disclosed"), item("c1", "checklist_open", "disclosed")] },
  })
  expect(resultChecks(disclosed).map((c) => c.tone)).toEqual(["pass", "warn", "neutral", "warn"])

  const changed = lean({ ...done, changed_after_check: { since: "2026-10-11T09:05:28.417Z", documents: ["Supply note.docx"], missing: [] } })
  expect(resultChecks(changed)[0]?.tone).toBe("warn")

  const two = withFindings(2, [item("p1", "placeholder", "disclosed")])
  expect(resultChecks(two)[0]?.tone).toBe("warn")
  expect(stageMilestones(two).at(-1)).toMatchObject({ step: "gate", state: "done", flag: true })
  expect(stageMilestones(done).at(-1)?.flag).toBeUndefined()
  const settledExcerpt = buildExcerpt({ status: two, plan: { keys: [`run:completed:${done.updated_at}`], focus: `run:completed:${done.updated_at}`, reason: "checkpoint" }, activity: [], now: Date.parse(done.updated_at) })
  expect(settledExcerpt).toContain("<unresolved_checklist_findings>")
  expect(settledExcerpt).not.toContain("Drafting is complete.")

  const queued = lean({ ...blockedQuestion, queued_answer_ids: blockedQuestion.questions.map((q) => q.id) })
  expect(checkpointKeys(queued).some((key) => key.startsWith("question:"))).toBe(false)
})

test("checkpoint scheduling debounces, respects min gap, emits progress, and caps excerpt size", () => {
  expect(checkpointKeys(running)).toEqual([])
  const keys = checkpointKeys(question)
  expect(keys[0]).toBe("plan:d1")
  expect(keys).toContain("question:q-4cc92c7c73")
  expect(keys.at(-1)).toBe(`run:blocked:${question.updated_at}`)
  expect(checkpointKeys(lean({ ...running, in_flight: false })).at(-1)).toBe(`run:paused:${running.updated_at}`)

  const now = Date.parse(question.updated_at) + 60_000
  const seenAt = new Map(keys.map((key) => [key, now - 1_000]))
  const base = { status: question, narrated: new Set<string>(), seenAt, now, inFlight: false, failures: 0, lastCallAt: 0, activitySince: 0 }
  expect(planNarration(base)).toBeUndefined()
  const later = { ...base, now: now + NARRATION_DEBOUNCE_MS }
  const plan = planNarration(later)
  expect(plan).toEqual({ keys, focus: `run:blocked:${question.updated_at}`, reason: "checkpoint" })
  expect(planNarration({ ...later, inFlight: true })).toBeUndefined()
  expect(planNarration({ ...later, failures: NARRATION_FAILURE_LIMIT })).toBeUndefined()
  expect(planNarration({ ...later, narrated: new Set(plan!.keys) })).toBeUndefined()

  const started = Date.parse(running.stages["d1.draft"]!.started_at!)
  const planned = lean({ ...running, plan: { documents: [{ id: "d1", title: "Supply note", purpose: "" }] } })
  const gapBase = { status: planned, narrated: new Set<string>(), seenAt: new Map([["plan:d1", started]]), now: started + 20_000, inFlight: false, failures: 0, lastCallAt: started + 5_000, activitySince: 3 }
  expect(planNarration(gapBase)).toBeUndefined()
  expect(planNarration({ ...gapBase, now: started + 5_000 + NARRATION_MIN_GAP_MS })?.keys).toEqual(["plan:d1"])

  const quiet = { ...gapBase, status: running, lastCallAt: 0 }
  expect(planNarration({ ...quiet, now: started + NARRATION_PROGRESS_MS - 1 })).toBeUndefined()
  expect(planNarration({ ...quiet, now: started + NARRATION_PROGRESS_MS })).toEqual({ keys: [`d1.draft:progress:${started}:1`], focus: "d1.draft", reason: "progress" })
  expect(planNarration({ ...quiet, now: started + NARRATION_PROGRESS_MS, activitySince: 0 })).toBeUndefined()

  const long = buildExcerpt({
    status: running,
    plan: { keys: ["d1.draft:progress:0:1"], focus: "d1.draft", reason: "progress" },
    activity: Array.from({ length: 200 }, (_, i): ActivityLine => ({ id: `l${i}`, at: i, kind: "tool", state: "done", text: "x".repeat(400) })),
    notes: { text: "draft ".repeat(5_000), reasoning: "thinking ".repeat(5_000) },
    now: Date.now(),
  })
  expect(long.length).toBeLessThanOrEqual(EXCERPT_LIMIT)
})

test("timeline groups stage activity, deduplicates notes, marks past pause notes, and rolls up research/reads", () => {
  const line = (id: string, at: number, extra: Partial<ActivityLine> = {}): ActivityLine => ({ id, at, kind: "tool", state: "done", text: id, ...extra })
  const stages = [
    line("d1.draft:stage:a", 10, { kind: "milestone", step: "d1.draft", until: 20, text: "Draft written" }),
    line("d1.review:stage:b", 20, { kind: "milestone", step: "d1.review", state: "running", text: "Reviewing" }),
  ]
  const events = [
    line("x", 5),
    line("r1", 11, { step: "d1.draft", group: "read", target: "a.docx" }),
    line("r2", 12, { step: "d1.draft", group: "read", target: "b.docx" }),
    line("q1", 15, { step: "d1.draft", group: "quote", flag: true }),
    line("q2", 16, { step: "d1.draft", group: "quote" }),
    line("t0", 19, { step: "d1.review", kind: "thought" }),
    ...[21, 22, 23, 24].map((at) => line(`v${at}`, at, { step: "d1.review" })),
  ]
  const notes = [
    { id: "n1", at: 20.5, text: "The draft is written.", keys: [] },
    { id: "n2", at: 25, text: " the draft is  written. ", keys: [] },
  ]
  const blocks = narrationTimeline([...stages, ...events], notes, true)
  expect(blocks.map((block) => block.kind)).toEqual(["steps", "stage", "stage", "note"])
  const draft = blocks[1]
  const review = blocks[2]
  if (draft?.kind !== "stage" || review?.kind !== "stage") throw new Error("expected two stage rows")
  expect(draft.lines.map((row) => row.id)).toEqual(["r1", "r2", "q1", "q2"])
  expect(draft.shown).toEqual([])
  expect(draft.rollUp.some((entry) => entry.flagged)).toBe(true)
  expect(review.shown.map((row) => row.id)).toEqual(["v22", "v23", "v24"].slice(-VISIBLE_STEPS))
  expect([elapsed(45_000), elapsed(252_000), elapsed(3_780_000)]).toEqual(["45s", "4m 12s", "1h 3m"])

  const paused = narrationNote("n1", "question:q-4cc92c7c73", `run:blocked:${question.updated_at}`)
  const progress = narrationNote("n2", "d1.draft:progress:1:0")
  const pastFlags = (status: LeanDraftingStatus) =>
    narrationTimeline([], [paused, progress], false, [], status).flatMap((block) => (block.kind === "note" ? [!!block.past] : []))
  expect(pastFlags(question)).toEqual([false, false])
  expect(pastFlags(running)).toEqual([true, false])
  expect(pastFlags(done)).toEqual([true, false])

  const blocked = partActivity(tool("official-source", { status: "completed", input: { url: "https://example.gov/law" }, output: "", title: "", metadata: { blocked: true } }), 1, "matter")!
  const partial = partActivity(tool("official-source", { status: "completed", input: { url: "https://example.gov/law" }, output: "", title: "", metadata: { partial: true } }), 1, "matter")!
  const empty = partActivity(tool("web-search", { status: "completed", input: { query: "law" }, output: "", title: "", metadata: { results: false } }), 1, "matter")!
  expect(rollUp([blocked, partial, empty]).at(-1)?.flagged).toContain("3")
  const error = partActivity(tool("webfetch", { status: "error", input: { url: "https://example.gov/law" }, error: "timeout", title: "" }), 1, "matter")!
  expect(rollUp([error])[0]?.flagged).toBeDefined()
})

test("narrator history and parallel stage sessions retain keys, workers, and bounded notes", () => {
  expect(narratorTitle("ses_child1")).toBe("doc.haus narrator · ses_child1")
  const messages = [
    { type: "user", id: "u1", metadata: { dochaus_narration: { keys: ["plan:d1"] } } },
    { type: "assistant", id: "a1", time: { created: 10, completed: 12 }, content: [{ type: "text", text: " The plan is ready. " }] },
    { type: "user", id: "u2", metadata: {} },
    { type: "assistant", id: "a2", time: { created: 20 }, content: [{ type: "reasoning", text: "only thinking" }] },
  ] as unknown as SessionMessageInfo[]
  expect(narratorNotes(messages)).toEqual([{ id: "a1", at: 12, text: "The plan is ready.", keys: ["plan:d1"] }])

  const row = { status: "running" as const, attempts: 1, started_at: null, finished_at: null }
  expect(leanStageSessions({ ...row, session_id: "legacy" })).toEqual(["legacy"])
  expect(leanStageSessions({ ...row, session_id: "worker-a", session_ids: ["worker-a", "worker-b"] })).toEqual(["worker-a", "worker-b"])
  expect(leanStageSessions({ ...row, session_id: null, session_ids: ["worker-a", "worker-b"] })).toEqual(["worker-a", "worker-b"])
  expect(leanStageSessions(undefined)).toEqual([])

  const notes = stageSessionNotes([
    { text: `First result ${"a".repeat(4000)}`, reasoning: `${"a".repeat(4000)} First reasoning` },
    { text: `Second result ${"b".repeat(4000)}`, reasoning: `${"b".repeat(4000)} Second reasoning` },
  ])!
  expect(notes.text).toContain("First result")
  expect(notes.text).toContain("Second result")
  expect(notes.text.length).toBeLessThan(1800)
  expect(notes.reasoning.length).toBeLessThan(1000)
})
