import { expect, test } from "bun:test"
import type { SessionMessageInfo } from "@opencode/client"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { MemoryRouter } from "react-router-dom"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import failed from "../../../dochaus/workflows/lean/fixtures/status-failed.json"
import { isLeanDraftingStatus, leanStageSessions, type LeanDraftingStatus } from "../src/api/ingest/lean"
import { narratorNotes, narratorTitle } from "../src/api/opencode/narrator"
import type { Part } from "../src/api/session-view"
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
import DraftingNarration from "../src/components/narration/DraftingNarration"
import { buildExcerpt, stageSessionNotes, EXCERPT_LIMIT } from "../src/components/narration/excerpt"
import { resultChecks, runOutcome } from "../src/components/narration/outcome"
import {
  checkpointKeys,
  NARRATION_DEBOUNCE_MS,
  NARRATION_FAILURE_LIMIT,
  NARRATION_MIN_GAP_MS,
  NARRATION_PROGRESS_MS,
  planNarration,
} from "../src/components/narration/schedule"
import { elapsed, narrationTimeline, rollUp, VISIBLE_STEPS } from "../src/components/narration/timeline"

// Lean status bodies recorded from the real ingest routes with a scripted model
// (dochaus/workflows/lean/fixtures); the narration reads the same shapes.

function lean(value: unknown): LeanDraftingStatus {
  if (!isLeanDraftingStatus(value)) throw new Error("fixture is not a lean drafting status")
  return value
}

const running = lean(start.status)
const done = lean(completed)
const question = lean(blockedQuestion)
const stopped = lean(failed)

function tool(tool: string, state: Extract<Part, { type: "tool" }>["state"]): Part {
  return { type: "tool", id: `part-${tool}`, sessionID: "ses_child2", messageID: "msg", tool, state }
}

test("tool parts read as plain activity and name the run's working files by role", () => {
  const read = partActivity(
    tool("read-document", { status: "running", input: { document: "sources/Supply Terms.docx" }, title: "" }),
    5,
    "matter",
  )
  expect(read).toMatchObject({
    id: "part-read-document",
    at: 5,
    kind: "tool",
    state: "running",
    text: "Reading Supply Terms.docx",
  })
  expect(fileLabel("/m/.dochaus/runs/r1/plan.json")).toBe("the drafting plan")
  expect(fileLabel(".dochaus/runs/r1/other.txt")).toBe("the run's working files")

  const search = partActivity(
    tool("search-document", { status: "running", input: { document: "a/terms.docx", query: "delivery" }, title: "" }),
    1,
    "matter",
  )
  expect(search?.text).toBe('Searching terms.docx for "delivery"')

  const missing = partActivity(
    tool("verify-quote", {
      status: "completed",
      input: { document: "a/terms.docx" },
      output: "NOT FOUND: …",
      title: "",
    }),
    1,
    "matter",
  )
  expect(missing).toMatchObject({ state: "done", text: "A quotation was not found word for word in terms.docx" })
  const confirmed = partActivity(
    tool("verify-quote", { status: "completed", input: { document: "a/terms.docx" }, output: "FOUND", title: "" }),
    1,
    "matter",
  )
  expect(confirmed?.text).toBe("Confirmed a quotation in terms.docx")

  const unreadable = partActivity(
    tool("read", { status: "error", input: { filePath: "/m/.dochaus/runs/r1/draft.md" }, error: "x", title: "" }),
    1,
    "matter",
  )
  expect(unreadable).toMatchObject({ state: "error", text: "Could not read the working draft" })
})

test("thinking parts surface their latest heading and answer text is not activity", () => {
  const thought: Part = {
    type: "reasoning",
    id: "r1",
    sessionID: "s",
    messageID: "m",
    text: "**Mapping the sources**\nSome notes.\n\n**Weighing the remedies**\nMore notes.",
    time: { start: 7 },
  }
  expect(partActivity(thought, 9, "matter")).toMatchObject({
    at: 7,
    kind: "thought",
    state: "running",
    text: "Weighing the remedies",
  })
  expect(reasoningHeadline("The draft needs a schedule. Then more.")).toBe("The draft needs a schedule.")
  expect(reasoningHeadline("   ")).toBeUndefined()
  expect(
    partActivity({ type: "text", id: "t", sessionID: "s", messageID: "m", text: "answer" }, 1, "matter"),
  ).toBeUndefined()
})

test("each stage run is one row that updates in place, in plain words, with its own duration", () => {
  expect(stageMilestones(running).map((line) => [line.state, line.text])).toEqual([["running", "Drafting"]])
  expect(activeStep(running)).toBe("d1.draft")

  const finished = stageMilestones(done)
  expect(finished.map((line) => line.text)).toEqual([
    "Draft written",
    "Reviewed the draft",
    "Proposed amendments reconciled",
    "Changes applied",
    "Drafting notes written",
    "Final workflow check passed",
  ])
  const draft = done.stages["d1.draft"]!
  expect(finished[0]).toMatchObject({
    id: `d1.draft:stage:${draft.started_at}`,
    step: "d1.draft",
    at: Date.parse(draft.started_at!),
    until: Date.parse(draft.finished_at!),
  })
  // The running row and its finished form share an id, so the row updates in place.
  expect(stageMilestones(running)[0]?.id).toBe(`d1.draft:stage:${running.stages["d1.draft"]!.started_at}`)
  expect(stageMilestones(question).at(-1)).toMatchObject({
    state: "paused",
    text: "Final check found 1 open item to resolve",
  })
  expect(stageMilestones(stopped).at(-1)).toMatchObject({ state: "error", text: "Stopped while drafting" })
  // Nothing is working on the run: its stage reads as paused rather than spinning.
  expect(stageMilestones(lean({ ...running, in_flight: false }))[0]).toMatchObject({
    state: "paused",
    text: "Paused while drafting",
  })

  const retry = lean({
    ...running,
    stages: { ...running.stages, "d1.draft": { ...running.stages["d1.draft"]!, attempts: 2 } },
  })
  expect(stageMilestones(retry)[0]?.text).toBe("Drafting (trying again)")
})

test("a settled or stalled run says in one line where it stands and how the lawyer moves it on", () => {
  expect(runOutcome(running)).toBeUndefined()
  expect(runOutcome(done)).toEqual({ tone: "done", text: "Drafting is complete." })
  expect(runOutcome(question)).toEqual({
    tone: "attention",
    text: "Drafting is waiting on a question only you can answer.",
  })
  const answered = lean({
    ...question,
    questions: question.questions.map((row) => ({ ...row, status: "answered" })),
  })
  expect(runOutcome(answered)).toEqual({
    tone: "attention",
    text: "Paused at the final check with 1 open item. Continue drafting (above) runs a repair round on them.",
  })
  expect(runOutcome(stopped)).toMatchObject({ tone: "attention" })
  const idle = lean({ ...running, in_flight: false })
  expect(runOutcome(idle)?.text).toContain("Continue drafting (above)")
  const texts = [running, done, question, answered, stopped, idle].map((status) => runOutcome(status)?.text ?? "")
  expect(texts.filter((text) => text.includes("Drafting panel"))).toEqual([])
})

test("a completed run claims only the checks its status shows", () => {
  expect(resultChecks(done)).toEqual([
    { tone: "pass", text: "Final workflow check passed" },
    { tone: "pass", text: "Quotes checked against sources" },
  ])
  const item = (id: string, kind: string, status: "open" | "resolved" | "disclosed") => ({
    id,
    stage: "d1.review",
    kind,
    detail: id,
    status,
  })
  const disclosed = lean({
    ...done,
    gate: {
      ...done.gate!,
      items: [
        item("q1", "quote_miss", "disclosed"),
        item("q2", "quote_miss", "disclosed"),
        item("q3", "quote_miss", "resolved"),
        item("p1", "placeholder", "disclosed"),
        item("l1", "left_out", "disclosed"),
        item("a1", "question", "disclosed"),
      ],
    },
  })
  expect(resultChecks(disclosed)).toEqual([
    { tone: "pass", text: "Final workflow check passed" },
    { tone: "warn", text: "2 quotes not found word for word" },
    { tone: "neutral", text: "1 placeholder to fill" },
    { tone: "neutral", text: "1 other item in the drafting notes" },
  ])
  expect(resultChecks(lean({ ...done, gate: null }))).toEqual([])
  const changed = lean({ ...done, changed_after_check: { since: "2026-10-11T09:05:28.417Z", documents: ["Supply note.docx"] } })
  expect(resultChecks(changed)).toEqual([
    { tone: "warn", text: "Changed after final check: review the draft" },
    { tone: "pass", text: "Quotes checked against sources" },
  ])
  expect(resultChecks(lean({ ...done, changed_after_check: null }))[0]).toEqual({ tone: "pass", text: "Final workflow check passed" })
})

test("a technical stop reads as a calm pause, and Continue reads as a resume", () => {
  const fault = lean({ ...stopped, stop_reason: "d1.draft: stopped by a runtime error: provider unavailable" })
  expect(stoppedByTechnicalProblem(fault)).toBe(true)
  expect(stoppedByTechnicalProblem(stopped)).toBe(false)
  expect(runOutcome(fault)).toMatchObject({ tone: "attention" })
  expect(runOutcome(fault)?.text).toStartWith("Paused after a technical problem.")
  expect(stageMilestones(fault).at(-1)).toMatchObject({
    state: "paused",
    text: "Paused while drafting after a technical problem",
  })

  expect(runTransition(running, fault)).toBeUndefined()
  expect(runTransition(fault, running)).toEqual({ state: "done", text: "Resumed, picking up with drafting" })
  expect(runTransition(running, lean({ ...running, in_flight: false }))).toMatchObject({ state: "paused" })
  expect(runTransition(question, done)).toBeUndefined()
  expect(runTransition(running, running)).toBeUndefined()
})

test("checkpoints cover finished stages, the plan, open questions and the run settling", () => {
  expect(checkpointKeys(running)).toEqual([])
  const keys = checkpointKeys(question)
  expect(keys[0]).toBe("plan:d1")
  expect(keys).toContain("question:q-4cc92c7c73")
  expect(keys).toContain(`gate:blocked:${question.stages.gate!.finished_at}`)
  expect(keys.at(-1)).toBe(`run:blocked:${question.updated_at}`)
  expect(checkpointKeys(lean({ ...running, in_flight: false })).at(-1)).toBe(`run:paused:${running.updated_at}`)
})

test("checkpoint updates wait out the debounce, then coalesce into one call", () => {
  const now = Date.parse(question.updated_at) + 60_000
  const seenAt = new Map(checkpointKeys(question).map((key) => [key, now - 1_000]))
  const base = {
    status: question,
    narrated: new Set<string>(),
    seenAt,
    now,
    inFlight: false,
    failures: 0,
    lastCallAt: 0,
    activitySince: 0,
  }
  expect(planNarration(base)).toBeUndefined()

  const later = { ...base, now: now + NARRATION_DEBOUNCE_MS }
  const plan = planNarration(later)
  expect(plan).toEqual({
    keys: checkpointKeys(question),
    focus: `run:blocked:${question.updated_at}`,
    reason: "checkpoint",
  })
  expect(planNarration({ ...later, inFlight: true })).toBeUndefined()
  expect(planNarration({ ...later, failures: NARRATION_FAILURE_LIMIT })).toBeUndefined()
  expect(planNarration({ ...later, narrated: new Set(plan!.keys) })).toBeUndefined()
})

test("a working run keeps a gap between updates and gives progress during a long stage", () => {
  const started = Date.parse(running.stages["d1.draft"]!.started_at!)
  const plan = lean({ ...running, plan: { documents: [{ id: "d1", title: "Supply note", purpose: "" }] } })
  const seenAt = new Map([["plan:d1", started]])
  const base = {
    status: plan,
    narrated: new Set<string>(),
    seenAt,
    now: started + 20_000,
    inFlight: false,
    failures: 0,
    lastCallAt: started + 5_000,
    activitySince: 3,
  }
  expect(planNarration(base)).toBeUndefined()
  expect(planNarration({ ...base, now: started + 5_000 + NARRATION_MIN_GAP_MS })?.keys).toEqual(["plan:d1"])

  const quiet = { ...base, status: running, lastCallAt: 0 }
  expect(planNarration({ ...quiet, now: started + NARRATION_PROGRESS_MS - 1 })).toBeUndefined()
  expect(planNarration({ ...quiet, now: started + NARRATION_PROGRESS_MS })).toEqual({
    keys: [`d1.draft:progress:${started}:1`],
    focus: "d1.draft",
    reason: "progress",
  })
  expect(planNarration({ ...quiet, now: started + NARRATION_PROGRESS_MS, activitySince: 0 })).toBeUndefined()
})

test("the excerpt carries the run's state in tagged sections and stays under its cap", () => {
  const excerpt = buildExcerpt({
    status: question,
    plan: { keys: checkpointKeys(question), focus: "question:q-4cc92c7c73", reason: "checkpoint" },
    activity: [{ id: "a", at: 1, kind: "tool", state: "done", text: "Read Supply Terms.docx" }],
    now: Date.parse(question.created_at) + 5 * 60_000,
  })
  expect(excerpt).toContain("<run>\nRun state: stopped, waiting on the lawyer.")
  expect(excerpt).toContain("Planned documents: Supply note.")
  expect(excerpt).toContain("Time since the run started: 5 minutes.")
  expect(excerpt).toContain("- Read Supply Terms.docx")
  expect(excerpt).toContain("<open_questions_for_the_lawyer>\n- Should late delivery carry a remedy?")
  expect(excerpt).not.toContain("<stage_output_so_far>")
  const settled = buildExcerpt({
    status: done,
    plan: {
      keys: [`run:completed:${done.updated_at}`],
      focus: `run:completed:${done.updated_at}`,
      reason: "checkpoint",
    },
    activity: [],
    now: Date.parse(done.updated_at),
  })
  expect(settled).toContain("Delivered documents: Supply note.docx; Supply note - redline.docx;")
  expect(settled).toContain("- The run settled: Drafting is complete.")

  const long = buildExcerpt({
    status: running,
    plan: { keys: ["d1.draft:progress:0:1"], focus: "d1.draft", reason: "progress" },
    activity: Array.from(
      { length: 200 },
      (_, i): ActivityLine => ({ id: `l${i}`, at: i, kind: "tool", state: "done", text: "x".repeat(400) }),
    ),
    notes: { text: "draft ".repeat(5_000), reasoning: "thinking ".repeat(5_000) },
    now: Date.now(),
  })
  expect(long.length).toBeLessThanOrEqual(EXCERPT_LIMIT)
  expect(long).toContain("Working on: Drafting.")
})

test("the timeline hangs each stage's activity from its row, rolls up finished stages and keeps the running one live", () => {
  const line = (id: string, at: number, extra: Partial<ActivityLine> = {}): ActivityLine => ({
    id,
    at,
    kind: "tool",
    state: "done",
    text: id,
    ...extra,
  })
  const stages = [
    line("d1.draft:stage:a", 10, { kind: "milestone", step: "d1.draft", until: 20, text: "Draft written" }),
    line("d1.review:stage:b", 20, { kind: "milestone", step: "d1.review", state: "running", text: "Reviewing the draft" }),
  ]
  const activity = [
    line("x", 5),
    line("r1", 11, { step: "d1.draft", group: "read", target: "a.docx" }),
    line("r2", 12, { step: "d1.draft", group: "read", target: "a.docx" }),
    line("r3", 13, { step: "d1.draft", group: "read", target: "b.docx" }),
    line("s1", 14, { step: "d1.draft", group: "search" }),
    line("q1", 15, { step: "d1.draft", group: "quote", flag: true }),
    line("q2", 16, { step: "d1.draft", group: "quote" }),
    // Stamped by a browser clock a moment behind the run's: still its own stage's.
    line("t0", 19, { step: "d1.review", kind: "thought" }),
    ...[21, 22, 23, 24].map((at) => line(`v${at}`, at, { step: "d1.review" })),
  ]
  const notes = [
    { id: "n1", at: 20.5, text: "The draft is written.", keys: [] },
    { id: "n2", at: 25, text: " the draft is  written. ", keys: [] },
  ]
  const blocks = narrationTimeline([...stages, ...activity], notes, true)
  expect(blocks.map((block) => block.kind)).toEqual(["steps", "stage", "stage", "note"])
  const draft = blocks[1]
  const review = blocks[2]
  if (draft?.kind !== "stage" || review?.kind !== "stage") throw new Error("expected two stage rows")
  expect(draft.lines.map((row) => row.id)).toEqual(["r1", "r2", "r3", "s1", "q1", "q2"])
  expect(draft.shown).toEqual([])
  expect(draft.rollUp).toEqual([
    { text: "3 reads of source material" },
    { text: "1 search of matter documents" },
    { text: "Checked 2 quotes", flagged: "1 not found word for word" },
  ])
  expect(review.lines.map((row) => row.id)).toEqual(["t0", "v21", "v22", "v23", "v24"])
  expect(review.shown.map((row) => row.id)).toEqual(["v22", "v23", "v24"].slice(-VISIBLE_STEPS))
  expect(review.rollUp).toEqual([])

  const two = [
    line("d1.apply:stage:a", 1, { kind: "milestone", step: "d1.apply", until: 2 }),
    line("d2.draft:stage:b", 3, { kind: "milestone", step: "d2.draft", until: 4 }),
    line("d2.review:stage:c", 5, { kind: "milestone", step: "d2.review", until: 6 }),
    line("gate:stage:d", 7, { kind: "milestone", step: "gate", until: 8 }),
  ]
  const planned = [
    { id: "d1", title: "`Supply` note" },
    { id: "d2", title: "cover letter" },
  ]
  expect(
    narrationTimeline(two, [], false, planned).map((block) => (block.kind === "heading" ? block.text : block.kind)),
  ).toEqual(["Supply note", "stage", "Cover letter", "stage", "stage", "Finishing", "stage"])
  expect(narrationTimeline(two, [], false, planned.slice(0, 1)).some((block) => block.kind === "heading")).toBe(false)
  expect([elapsed(45_000), elapsed(252_000), elapsed(3_780_000)]).toEqual(["45s", "4m 12s", "1h 3m"])
})

test("narrator history pairs each update with the checkpoints its prompt recorded", () => {
  expect(narratorTitle("ses_child1")).toBe("doc.haus narrator · ses_child1")
  const messages = [
    { type: "user", id: "u1", metadata: { dochaus_narration: { keys: ["plan:d1"] } } },
    {
      type: "assistant",
      id: "a1",
      time: { created: 10, completed: 12 },
      content: [{ type: "text", text: " The plan is ready. " }],
    },
    { type: "user", id: "u2", metadata: {} },
    { type: "assistant", id: "a2", time: { created: 20 }, content: [{ type: "reasoning", text: "only thinking" }] },
  ] as unknown as SessionMessageInfo[]
  expect(narratorNotes(messages)).toEqual([{ id: "a1", at: 12, text: "The plan is ready.", keys: ["plan:d1"] }])
})

test("after a reload the run's own conversation shows the request, where the run stands and, once done, its result", () => {
  const narration = {
    status: lean({ ...question, request: "Draft a short supply note from the supply terms." }),
    visible: true,
    lines: stageMilestones(question),
    notes: [],
    narrating: false,
    failed: false,
    onForeignPart: () => {},
  }
  const render = (value: typeof narration, showRequest: boolean) =>
    renderToStaticMarkup(
      createElement(
        MemoryRouter,
        null,
        createElement(DraftingNarration, { narration: value, matterId: "m1", showRequest, onViewDocument: () => {} }),
      ),
    )
  const html = render(narration, true)
  expect(html).toContain('<div class="msg user">Draft a short supply note from the supply terms.</div>')
  // The question dock above the composer says a question waits, lists it and takes the answer; the narration does none of that.
  expect(html).not.toContain("narration-outcome-note")
  expect(html).not.toContain("Should late delivery carry a remedy?")
  expect(html).not.toContain("<textarea")
  expect(html).not.toContain("Drafting panel")
  expect(html).not.toContain('id="run-result"')

  const anchored = render(
    { ...narration, status: lean({ ...done, request: "Draft it." }), lines: stageMilestones(done) },
    false,
  )
  expect(anchored).not.toContain("Draft it.")
  expect(anchored).toContain('<section id="run-result" class="narration-result"')
  expect(anchored).toContain("Ready for review")
  expect(anchored).toContain("Final workflow check passed")
  expect(anchored).toContain("Quotes checked against sources")
  // The draft with its redline and notes beneath it, each to open or download.
  expect(anchored.match(/class="drafted-group"/g)).toHaveLength(1)
  expect(anchored).toContain('title="Open Supply note - redline.docx"')
  expect(anchored.match(/class="icon-btn drafted-download"/g)).toHaveLength(3)
  expect(anchored).not.toContain("narration-outcome-note")
  expect(anchored).not.toContain("in Documents")
})

test("source reads count calls without claiming fragments or repeated reads are separate documents", () => {
  const inputs = [
    { document: "sources/Terms.docx" },
    { filePath: "/matter/.dochaus/source-chunks/terms-c001.txt" },
    { filePath: "/matter/.dochaus/source-chunks/terms-c002.txt" },
    { filePath: "/matter/.dochaus/source-chunks/terms-c001.txt" },
    { filePath: "/matter/.dochaus/source-index.json" },
  ]
  const lines = inputs.map((input, index) => partActivity(
    { ...tool(index === 0 ? "read-document" : "read", {
      status: "completed", input, output: "source text", title: "",
    }), id: `read-${index}` }, index, "matter",
  )!)
  expect(rollUp(lines)).toEqual([{ text: "5 reads of source material" }])
  expect(rollUp(lines.slice(0, 1))).toEqual([{ text: "1 read of source material" }])
  const searches = [1, 2].map((index) => partActivity(
    { ...tool("search-document", {
      status: "completed", input: { query: "operative terms" }, output: "passages", title: "",
    }), id: `search-${index}` }, index, "matter",
  )!)
  expect(searches[0]?.text).toBe('Searched the matter documents for "operative terms"')
  expect(rollUp(searches)).toEqual([{ text: "2 searches of matter documents" }])
})


test("queued answers stop narration asking after durable acceptance", () => {
  const status = lean({ ...blockedQuestion, queued_answer_ids: blockedQuestion.questions.map(question => question.id) })
  expect(runOutcome(status)?.text).not.toContain("only you can answer")
  expect(checkpointKeys(status).some(key => key.startsWith("question:"))).toBe(false)
})

test("evidence stages narrate actual work", () => {
  const row = { status: "running" as const, attempts: 1, started_at: "2026-10-12T10:00:00.000Z", finished_at: null }
  const status = lean({ ...running, stages: { discrepancy: row, spine: row, "d1.review-independent": row, "package-review": row, finalize: { ...row, status: "pending", started_at: null }, gate: { ...row, status: "pending", started_at: null } } })
  expect(stageMilestones(status).map(line => line.text)).toEqual([
    "Assessing discrepancies in the source materials", "Building the source-backed drafting spine",
    "Independently checking function and fidelity", "Reviewing final semantics and the document package",
  ])
})


test("research activity names discovery and retrieval without claiming verified law", () => {
  const activity = (name: string, input: Record<string, unknown>, metadata?: Record<string, unknown>) => partActivity(tool(name, { status: "completed", input, output: "retrieved text", title: name, metadata }), 1, "matter")!
  const search = activity("web-search", { query: "Applicable primary law" })
  const source = activity("official-source", { url: "https://example.gov/law" })
  const page = activity("webfetch", { url: "https://example.gov/law" })
  expect(search.text).toBe('Searched the web for "Applicable primary law"')
  expect(source.text).toBe("Retrieved source passages from example.gov")
  expect(page.text).toBe("Retrieved a web page from example.gov")
  expect(rollUp([search, source, page])).toEqual([{ text: "1 web search" }, { text: "1 web page retrieval call" }, { text: "1 source passage retrieval call" }])
  const blocked = activity("official-source", { url: "https://example.gov/law" }, { blocked: true })
  const partial = activity("official-source", { url: "https://example.gov/law" }, { partial: true })
  const empty = activity("web-search", { query: "Applicable primary law" }, { results: false })
  expect(blocked.text).toStartWith("Could not retrieve")
  expect(partial.text).toContain("partial retrieval")
  expect(empty.text).toBe("A web search returned no results")
  expect(activity("web-search", { query: "law" }, { results: false, authenticationFailed: true }).text).toBe("A web search could not authenticate")
  expect(rollUp([blocked, partial, empty]).at(-1)?.flagged).toBe("3 retrieval limitations")
  const error = partActivity(tool("webfetch", { status: "error", input: { url: "https://example.gov/law" }, error: "timeout", title: "" }), 1, "matter")!
  expect(error).toMatchObject({ state: "error", text: "Could not retrieve a web page from example.gov" })
  expect(rollUp([error])).toEqual([{ text: "", flagged: "1 step did not complete" }])
})


test("parallel specialist reviews narrate their own work, completion and paused state", () => {
  const row = { status: "running" as const, attempts: 1, started_at: "2026-10-12T10:00:00.000Z", finished_at: null }
  const stages = { "d1.review-wording": row, "d1.review-client-disadvantage": row, finalize: { ...row, status: "pending" as const, started_at: null }, gate: { ...row, status: "pending" as const, started_at: null } }
  const status = lean({ ...running, stages })
  expect(stageMilestones(status).map(line => [line.step, line.text])).toEqual([
    ["d1.review-wording", "Reviewing problematic wording"],
    ["d1.review-client-disadvantage", "Reviewing client disadvantage"],
  ])
  const finished = Object.fromEntries(Object.entries(stages).map(([key, value]) => [key, { ...value, status: "done", finished_at: "2026-10-12T10:01:00.000Z" }]))
  expect(stageMilestones(lean({ ...status, stages: finished })).map(line => line.text)).toEqual([
    "Problematic wording review completed", "Client disadvantage review completed",
  ])
  expect(stageMilestones(lean({ ...status, status: "failed", in_flight: false })).map(line => line.text)).toEqual([
    "Paused while reviewing problematic wording", "Paused while reviewing client disadvantage",
  ])
  expect(fileLabel(".dochaus/lean/d1/review-wording.md")).toBe("the wording review notes")
  expect(fileLabel(".dochaus/lean/d1/review-client-disadvantage.md")).toBe("the client disadvantage review notes")
  expect(fileLabel(".dochaus/lean/d1/review-independent.md")).toBe("the independent review notes")
})


test("amendment stages narrate preparation and verification before application", () => {
  const row = { status: "running" as const, attempts: 1, calls: 2, session_id: null, started_at: "2026-10-13T10:00:00.000Z", finished_at: null }
  const pending = { ...row, status: "pending" as const, started_at: null }
  const stages = { "d1.change-plan": row, "d1.change-prepare": row, "d1.changes": row, "d1.change-review": row, "d1.apply": pending, finalize: pending, gate: pending }
  const status = lean({ ...running, stages })
  expect(stageMilestones(status).map(line => line.text)).toEqual([
    "Assessing review findings for amendments", "Preparing amendment groups",
    "Reconciling the proposed amendments", "Independently checking the amended document",
  ])
  const finished = Object.fromEntries(Object.entries(stages).map(([key, value]) => [key, value.started_at ? { ...value, status: "done", finished_at: "2026-10-13T10:01:00.000Z" } : value]))
  expect(stageMilestones(lean({ ...status, stages: finished })).map(line => line.text)).toEqual([
    "Amendment plan prepared", "Amendment groups prepared", "Proposed amendments reconciled", "Amended document checked",
  ])
  expect(stageMilestones(lean({ ...status, status: "blocked", in_flight: false })).map(line => line.state)).toEqual(["paused", "paused", "paused", "paused"])
  expect(fileLabel(".dochaus/lean/d1/change-plan.json")).toBe("the amendment plan")
  expect(fileLabel(".dochaus/lean/d1/change-preparation.json")).toBe("the prepared amendments")
  expect(fileLabel(".dochaus/lean/d1/change-review.json")).toBe("the proposed amendment review")
})


test("parallel stage sessions retain every worker with singular status compatibility", () => {
  const row = { status: "running" as const, attempts: 1, started_at: null, finished_at: null }
  expect(leanStageSessions({ ...row, session_id: "legacy" })).toEqual(["legacy"])
  expect(leanStageSessions({ ...row, session_id: "worker-a", session_ids: ["worker-a", "worker-b"] })).toEqual(["worker-a", "worker-b"])
  expect(leanStageSessions({ ...row, session_id: null, session_ids: ["worker-a", "worker-b"] })).toEqual(["worker-a", "worker-b"])
  expect(leanStageSessions(undefined)).toEqual([])
  const notes = stageSessionNotes([{ text: `First result ${"a".repeat(4000)}`, reasoning: `${"a".repeat(4000)} First reasoning` }, { text: `Second result ${"b".repeat(4000)}`, reasoning: `${"b".repeat(4000)} Second reasoning` }])!
  expect(notes.text).toContain("First result")
  expect(notes.text).toContain("Second result")
  expect(notes.reasoning).toContain("First reasoning")
  expect(notes.reasoning).toContain("Second reasoning")
  expect(notes.text.length).toBeLessThan(1800)
  expect(notes.reasoning.length).toBeLessThan(1000)
})
