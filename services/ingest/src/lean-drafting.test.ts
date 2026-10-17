import { Database } from "bun:sqlite"
import { afterEach, describe, expect, test } from "bun:test"
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { drainLeanRequests, enqueueLeanRequest, pendingLeanRequests } from "../../../dochaus/workflows/lean/requests"
import { runLean } from "../../../dochaus/workflows/lean/runner"
import { LEAN_STAGE_AGENTS, type LeanDeps, type StageSessionPort } from "../../../dochaus/workflows/lean/stages"
import {
  initLeanState,
  leanDir,
  readLeanState,
  readLeanText,
  readQuestions,
  writeLeanState,
  writeLeanText,
  writeQuestions,
} from "../../../dochaus/workflows/lean/state"
import { readLeanEvents } from "../../../dochaus/workflows/lean/events"
import { FIXTURE_SOURCE_TEXT, happyPath, leanFixtureMatter, scriptedModel, noChanges } from "../../../dochaus/workflows/lean/test-support"
import {
  draftingStartProfile,
  leanDraftingIdle,
  leanDraftingRoutes,
  leanDraftingStatus,
  startLeanDraftingRun,
} from "./lean-drafting"

// The coordinator, state files and routes are real; only the model behind the
// engine session port is scripted.

const directories: string[] = []

afterEach(async () => {
  await Promise.all(directories.map((directory) => leanDraftingIdle(directory)))
  directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true }))
})

function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "lean-ingest-"))
  directories.push(directory)
  return directory
}

/** A model whose every reply waits for `release`, then fails as an engine error. */
function heldModel() {
  const created: Parameters<StageSessionPort["createSession"]>[0][] = []
  const prompts: string[] = []
  const gate = Promise.withResolvers<void>()
  const sessions: StageSessionPort = {
    createSession: async (input) => {
      created.push(input)
      return `ses_${created.length}`
    },
    prompt: async (input) => {
      prompts.push(input.agent)
      await gate.promise
      throw new Error("scripted engine outage")
    },
  }
  return { deps: { sessions } satisfies LeanDeps, created, prompts, release: () => gate.resolve() }
}

const MODEL = { providerID: "test", modelID: "scripted" }

describe("drafting start profile", () => {
  test("defaults to lean and rejects every other profile", () => {
    expect(draftingStartProfile(undefined)).toBe("lean")
    expect(draftingStartProfile(null)).toBe("lean")
    expect(draftingStartProfile("legacy")).toBeUndefined()
    expect(draftingStartProfile("lean")).toBe("lean")
    expect(draftingStartProfile("fast")).toBeUndefined()
    expect(draftingStartProfile(1)).toBeUndefined()
  })
})

const DRAFT_AGENT = LEAN_STAGE_AGENTS.draft
const doneRequests = (directory: string) => readdirSync(path.join(leanDir(directory), "requests", "done"))

test("full draft comparison projects structured adverse originals and queues explicit target scope", async () => {
  const directory = matter(), model = heldModel()
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })
  const state = initLeanState(directory, { request: "Use supplied terms.", ownerSessionID: "ses_owner", model: MODEL })
  writeLeanState(directory, { ...state, status: "blocked", stages: { ...state.stages, "checklist-check": { ...state.stages.discrepancy!, status: "blocked" } } })
  const item = { check: "Compare the terms.", basis: "directive", source_refs: [{ document: "request", locator: "request", quote: "Use supplied terms." }], qualifications: [], target: { document: "d1", role: "Terms", treatment: "operative" as const } }
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", input_sha: "source", status: "succeeded", issues: [], items: [{ id: "original", item }] }] }))
  const finding = { item_id: "original", status: "indeterminate", reason: "Complete current comparison is missing.", evidence: [], source_refs: item.source_refs, inspection_scope: "Initial sample", limitation: "Remaining paragraphs." }
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["original"], checks: [{ item_id: "original", input_sha: "check", status: "succeeded", issues: [], finding }] }))
  expect(leanDraftingStatus(directory)?.checklist_refinement).toEqual({ available: true, parents: [{ parent_id: "original", check: item.check, target: item.target, finding: { status: "indeterminate", reason: finding.reason }, scopes: ["target_body"] }] })
  const body = { ownerSessionID: "ses_owner", parents: [{ parent_id: "original", scope: "target_body" }] }
  const post = () => routes.request("/m1/lean/checklist/refine", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
  writeLeanText(directory, "checklist/checks.json", "{broken")
  expect(leanDraftingStatus(directory)?.checklist_refinement.parents).toEqual([])
  // Malformed saved evidence is unavailable; it cannot authorize a new scope.
  expect((await post()).status).toBe(400)
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["original"], checks: [{ item_id: "original", input_sha: "check", status: "succeeded", issues: [], finding }] }))
  try {
    expect((await post()).status).toBe(202)
    expect(leanDraftingStatus(directory)?.checklist_refinement.available).toBe(false)
  } finally { model.release(); await leanDraftingIdle(directory) }
  expect(readLeanText(directory, "request.md")?.trim()).toBe("Use supplied terms.")
})

test("checklist refinement API validates the settled owner and adverse original parent before queueing", async () => {
  const directory = matter()
  const model = heldModel()
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })
  const state = initLeanState(directory, { request: "Draft from supplied terms.", ownerSessionID: "ses_owner", model: MODEL })
  const post = (body: unknown) => routes.request("/m1/lean/checklist/refine", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
  const selection = { ownerSessionID: "ses_owner", parents: [{ parent_id: "requirement-a", scope: "source_collection" }] }
  expect((await post({ ...selection, parents: [] })).status).toBe(400)
  expect((await post({ ...selection, ownerSessionID: "ses_previous" })).status).toBe(409)
  expect((await post(selection)).status).toBe(409)
  writeLeanState(directory, { ...state, status: "blocked", stages: { ...state.stages, "checklist-check": { ...state.stages.discrepancy!, status: "blocked" } } })
  const item = { check: "Compare the terms.", basis: "directive", source_refs: [{ document: "request", locator: "request", quote: "Draft from supplied terms." }], qualifications: [], target: { document: "d1", role: "Terms", treatment: "operative" as const } }
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", input_sha: "source", status: "succeeded", issues: [], items: [{ id: "requirement-a", item }] }] }))
  const finding = { item_id: "requirement-a", reason: "Current inspection result.", evidence: [], source_refs: item.source_refs, inspection_scope: "Current paragraphs", limitation: "None" }
  const checks = (status: string) => JSON.stringify({ extraction_complete: true, expected_items: ["requirement-a"], checks: [{ item_id: "requirement-a", input_sha: "check", status: "succeeded", issues: [], finding: { ...finding, status } }] })
  writeLeanText(directory, "checklist/checks.json", checks("satisfied"))
  expect((await post(selection)).status).toBe(400)
  writeLeanText(directory, "checklist/checks.json", checks("indeterminate"))
  expect((await post({ ...selection, parents: [{ parent_id: "missing", scope: "source_collection" }] })).status).toBe(400)
  expect(model.created).toEqual([])
  try {
    const response = await post(selection)
    expect(response.status).toBe(202)
    expect(await response.json()).toMatchObject({ profile: "lean", mode: "queued" })
  } finally {
    model.release()
    await leanDraftingIdle(directory)
  }
  expect(readLeanText(directory, "updated-notes.md")).toBeUndefined()
})

describe("lean drafting start and Continue", () => {
  test("start runs the coordinator in the background; Continue is queued (202) and drained by the running coordinator", async () => {
    const directory = matter()
    const model = heldModel()
    const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })

    expect(leanDraftingStatus(directory)).toBeUndefined()
    expect((await routes.request("/m1/lean/resume", { method: "POST" })).status).toBe(404)

    const started = await startLeanDraftingRun(directory, { request: "Draft the agreement.", agent: "draft", model: MODEL }, model.deps)
    expect(started).toMatchObject({ profile: "lean", mode: "started", ownerSessionID: "ses_1", status: { profile: "lean", in_flight: true, request: "Draft the agreement." } })
    expect(model.created[0]).toMatchObject({ agent: "draft", model: MODEL, metadata: { dochaus_role: "lean-drafting-owner" } })
    expect(model.created[0].parentID).toBeUndefined()

    // One coordinator per matter: a second start is refused; Continues queue.
    await expect(startLeanDraftingRun(directory, { request: "Again.", agent: "draft", model: MODEL }, model.deps))
      .rejects.toMatchObject({ code: "lean_run_in_progress" })
    const [first, second] = await Promise.all([1, 2].map(() => routes.request("/m1/lean/resume", { method: "POST" })))
    expect([first.status, second.status]).toEqual([202, 202])
    const bodies = [await first.json(), await second.json()]
    expect(bodies[0]).toMatchObject({ profile: "lean", mode: "queued", status: { in_flight: true } })
    expect(bodies[0].request_id).not.toBe(bodies[1].request_id)
    expect(pendingLeanRequests(directory)).toHaveLength(2)

    model.release()
    await leanDraftingIdle(directory)
    // The stopped stage is re-run once for the queued Continues, by the same coordinator.
    expect(model.prompts).toEqual([LEAN_STAGE_AGENTS.discrepancy, LEAN_STAGE_AGENTS.discrepancy])
    expect(pendingLeanRequests(directory)).toEqual([])
    expect(doneRequests(directory)).toHaveLength(2)
    const failed = await (await routes.request("/m1/lean/state")).json()
    expect(failed).toMatchObject({
      profile: "lean", status: "failed", in_flight: false, pending_requests: 0,
      stop_reason: "discrepancy: stopped by a runtime error: scripted engine outage",
      stages: { discrepancy: { status: "failed" } },
      plan: null,
    })
    expect(model.created[1]).toMatchObject({ parentID: "ses_1", agent: LEAN_STAGE_AGENTS.discrepancy })

    // A Continue after the run stopped starts it again.
    const resumed = await routes.request("/m1/lean/resume", { method: "POST" })
    expect(resumed.status).toBe(202)
    expect(await resumed.json()).toMatchObject({ mode: "queued", status: { in_flight: true } })
    await leanDraftingIdle(directory)
    expect(model.prompts).toEqual([LEAN_STAGE_AGENTS.discrepancy, LEAN_STAGE_AGENTS.discrepancy, LEAN_STAGE_AGENTS.discrepancy])
    expect(readLeanState(directory)).toMatchObject({ status: "failed", owner_session_id: "ses_1" })
    expect(doneRequests(directory)).toHaveLength(3)
  })

  test("an answer is validated at once, queued while the run works, and recorded by the coordinator; never by the run itself", async () => {
    const directory = matter()
    const model = heldModel()
    const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })
    const answer = (qid: string, body: unknown) =>
      routes.request(`/m1/lean/questions/${qid}`, { method: "POST", body: JSON.stringify(body), headers: { "content-type": "application/json" } })

    expect((await answer("q1", { answer: "Yes." })).status).toBe(404)
    await startLeanDraftingRun(directory, { request: "Draft the agreement.", agent: "draft", model: MODEL }, model.deps)
    writeQuestions(directory, [{ id: "q1", stage: "changes", question: "Which party bears the cost?", context: "", status: "open" }])

    expect((await answer("q9", { answer: "The buyer." })).status).toBe(404)
    const empty = await answer("q1", { answer: "  " })
    expect(empty.status).toBe(400)
    expect(await empty.json()).toMatchObject({ code: "invalid_answer" })
    const queued = await answer("q1", { answer: "The buyer." })
    expect(queued.status).toBe(202)
    expect(await queued.json()).toMatchObject({ mode: "queued", status: { in_flight: true, pending_requests: 1, queued_answer_ids: ["q1"] } })
    expect(readQuestions(directory)[0].status).toBe("open")

    model.release()
    await leanDraftingIdle(directory)
    expect(readQuestions(directory)[0]).toMatchObject({ status: "answered", answer: "The buyer." })
    expect(leanDraftingStatus(directory)).toMatchObject({ questions: [{ id: "q1", status: "answered" }], pending_requests: 0, queued_answer_ids: [] })
  })
})

describe("a new drafting request on a matter that already has a lean run", () => {
  function previousRun(directory: string, status: "completed" | "failed" | "blocked" | "running") {
    const state = initLeanState(directory, { request: "Draft the first agreement.", ownerSessionID: "ses_previous", model: MODEL })
    return writeLeanState(directory, { ...state, status, stop_reason: status === "blocked" || status === "failed" ? `${status} in an earlier pass` : null })
  }
  const archiveRoot = (directory: string) => path.join(directory, ".dochaus", "lean-archive")

  test("starts a fresh run once the previous run completed or failed; the previous run is archived and its delivered files stay", async () => {
    for (const status of ["completed", "failed"] as const) {
      const directory = matter()
      const previous = previousRun(directory, status)
      writeFileSync(path.join(directory, "Supply note.docx"), "delivered bytes")
      const model = heldModel()

      const started = await startLeanDraftingRun(directory, { request: "Draft the second agreement.", agent: "draft", model: MODEL }, model.deps)
      expect(started).toMatchObject({ mode: "started", ownerSessionID: "ses_1", status: { status: "running", owner_session_id: "ses_1", in_flight: true } })
      expect(readLeanText(directory, "request.md")).toBe("Draft the second agreement.\n")
      const archived = readdirSync(archiveRoot(directory))
      expect(archived).toHaveLength(1)
      expect(archived[0]).toContain("ses_previous")
      const archive = path.join(archiveRoot(directory), archived[0]!)
      expect(JSON.parse(readFileSync(path.join(archive, "state.json"), "utf8"))).toMatchObject({ status, owner_session_id: "ses_previous", revision: previous.revision })
      expect(readFileSync(path.join(archive, "request.md"), "utf8")).toBe("Draft the first agreement.\n")
      expect(readFileSync(path.join(directory, "Supply note.docx"), "utf8")).toBe("delivered bytes")

      model.release()
      await leanDraftingIdle(directory)
      expect(readLeanState(directory)).toMatchObject({ owner_session_id: "ses_1" })
    }
  })

  test("refuses while the run is in flight (in progress) or blocked or paused (continue it), leaving that run in place", async () => {
    const inFlight = matter()
    const model = heldModel()
    await startLeanDraftingRun(inFlight, { request: "Draft the agreement.", agent: "draft", model: MODEL }, model.deps)
    await expect(startLeanDraftingRun(inFlight, { request: "Draft another.", agent: "draft", model: MODEL }, model.deps))
      .rejects.toMatchObject({ code: "lean_run_in_progress" })
    expect(readLeanText(inFlight, "request.md")).toBe("Draft the agreement.\n")
    model.release()
    await leanDraftingIdle(inFlight)

    // "running" with no coordinator is a run paused by an ingest restart.
    for (const status of ["blocked", "running"] as const) {
      const directory = matter()
      previousRun(directory, status)
      const idle = heldModel()
      await expect(startLeanDraftingRun(directory, { request: "Draft another.", agent: "draft", model: MODEL }, idle.deps))
        .rejects.toMatchObject({ code: "lean_run_awaiting_continue" })
      expect(idle.created).toEqual([])
      expect(readLeanState(directory)).toMatchObject({ status, owner_session_id: "ses_previous" })
      expect(existsSync(archiveRoot(directory))).toBe(false)
    }
  })
})

test("chat updates are queued for the matching run and late updates retain an explicit refusal", async () => {
  const directory = matter()
  const model = heldModel()
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })
  const start = await startLeanDraftingRun(directory, { request: "Draft a short note.", agent: "draft", model: MODEL }, model.deps)
  const send = (text: unknown, ownerSessionID = start.ownerSessionID) => routes.request("/matter/lean/updates", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text, ownerSessionID }),
  })
  try {
    expect((await send("Be concise.", "different-run")).status).toBe(409)
    expect((await send(" ")).status).toBe(400)
    expect(pendingLeanRequests(directory)).toHaveLength(0)
    const response = await send("Use concise headings.")
    expect(response.status).toBe(202)
    expect(await response.json()).toMatchObject({ mode: "queued", status: { updates: [{ text: "Use concise headings.", status: "queued", progress: "received" }] } })
    const state = readLeanState(directory)!
    writeLeanState(directory, { ...state, stages: { ...state.stages, "d1.apply": { ...state.stages["d1.apply"]!, status: "running" } } })
    const late = await send("Change the scope.")
    expect(late.status).toBe(409)
    expect(await late.json()).toMatchObject({ code: "lean_update_too_late" })
    expect(pendingLeanRequests(directory)).toHaveLength(1)
  } finally { model.release() }
  await leanDraftingIdle(directory)
  expect(readLeanText(directory, "updated-notes.md")).toContain("Use concise headings.")
  expect(readLeanText(directory, "updated-notes.md")).not.toContain("Change the scope.")
})


test("the HTTP answer flow holds a run until every intake question is answered", async () => {
  const directory = matter()
  let intake = true
  const model = scriptedModel({ ...happyPath,
    [LEAN_STAGE_AGENTS.intake]: (prompt, index) => {
      if (intake) {
        intake = false
        return JSON.stringify({ questions: [
          { question: "Which governing law applies?", context: "The legal approach depends on it." },
          { question: "Whose position should this protect?", context: "The drafting approach depends on it." },
        ] })
      }
      return JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" })
    },
    [LEAN_STAGE_AGENTS.review]: () => "No corrections needed.",
    [LEAN_STAGE_AGENTS.changes]: noChanges,
  })
  const deps = { sessions: model.port }
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => deps })
  await startLeanDraftingRun(directory, { request: "Draft a supply note.", agent: "draft", model: MODEL }, deps)
  await leanDraftingIdle(directory)
  const waiting = leanDraftingStatus(directory)!
  expect(waiting.status).toBe("blocked")
  expect(waiting.questions).toHaveLength(2)
  expect(waiting.stages["d1.review"]?.status).toBe("pending")
  const post = (id: string, answer: string) => routes.request(`/m1/lean/questions/${id}`, {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ answer }),
  })
  expect((await post(waiting.questions[0]!.id, "Singapore law.")).status).toBe(202)
  await leanDraftingIdle(directory)
  expect(leanDraftingStatus(directory)!.status).toBe("blocked")
  expect(model.prompts).toHaveLength(2)
  expect((await routes.request("/m1/lean/resume", { method: "POST" })).status).toBe(202)
  await leanDraftingIdle(directory)
  expect(model.prompts).toHaveLength(2)
  expect((await post(waiting.questions[1]!.id, "The supplier.")).status).toBe(202)
  await leanDraftingIdle(directory)
  expect(leanDraftingStatus(directory)!.status).toBe("completed")
  expect(model.prompts.findLast(prompt => prompt.agent === LEAN_STAGE_AGENTS.intake)!.text).toContain("Singapore law.")
  expect(model.prompts.findLast(prompt => prompt.agent === LEAN_STAGE_AGENTS.intake)!.text).toContain("The supplier.")
  expect(leanDraftingStatus(directory)!.questions.every(question => question.status === "answered")).toBe(true)
}, 120_000)

test("updates are only marked applied after current package review and become stale on later instructions", async () => {
  const directory = leanFixtureMatter("lean-update-progress-")
  directories.push(directory)
  const model = scriptedModel(happyPath)
  enqueueLeanRequest(directory, { kind: "update", text: "Keep headings concise." })
  expect(leanDraftingStatus(directory)?.updates[0]).toMatchObject({ status: "queued", progress: "received" })
  const completed = await runLean(directory, { sessions: model.port })
  expect(completed.status).toBe("completed")
  expect(leanDraftingStatus(directory)?.updates[0]).toMatchObject({ status: "included", progress: "applied" })
  enqueueLeanRequest(directory, { kind: "update", text: "Use a descriptive title." })
  expect(leanDraftingStatus(directory)?.updates[1]).toMatchObject({ progress: "received" })
  drainLeanRequests(directory)
  expect(leanDraftingStatus(directory)?.updates.every(update => "progress" in update && update.progress === "requires-rework")).toBe(true)
}, 120_000)


test("creating an empty source index during registration preserves applied update progress", async () => {
  const directory = leanFixtureMatter("lean-empty-index-update-")
  directories.push(directory)
  rmSync(path.join(directory, ".dochaus/legal.db"))
  enqueueLeanRequest(directory, { kind: "update", text: "Keep headings concise." })
  const model = scriptedModel({ ...happyPath, [LEAN_STAGE_AGENTS.review]: () => "No corrections required.", [LEAN_STAGE_AGENTS.changes]: noChanges })
  const registerDeliverable: NonNullable<LeanDeps["registerDeliverable"]> = async (matter, filename, bytes) => {
    const db = new Database(path.join(matter, ".dochaus/legal.db"))
    db.run("CREATE TABLE IF NOT EXISTS chunks (doc_name TEXT, text TEXT)")
    db.close()
    writeFileSync(path.join(matter, filename), bytes)
    return filename
  }
  expect((await runLean(directory, { sessions: model.port, registerDeliverable })).status).toBe("completed")
  expect(leanDraftingStatus(directory)?.updates[0]).toMatchObject({ status: "included", progress: "applied" })
  const resumed = scriptedModel({})
  expect((await runLean(directory, { sessions: resumed.port, registerDeliverable })).status).toBe("completed")
  expect(resumed.prompts).toHaveLength(0)
})


test("amendment status rows preserve stage order, worker-only sessions and the update cutoff", async () => {
  const directory = matter()
  const model = heldModel()
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })
  const start = await startLeanDraftingRun(directory, { request: "Draft a short note.", agent: "draft", model: MODEL }, model.deps)
  const order = ["d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply"] as const
  const initial = readLeanState(directory)!
  expect(Object.keys(initial.stages).filter(step => order.includes(step as typeof order[number]))).toEqual([...order])
  try {
    for (const step of order.slice(0, -1)) {
      const state = readLeanState(directory)!
      writeLeanState(directory, { ...state, stages: { ...state.stages, [step]: { ...state.stages[step]!, status: "running", session_id: null, session_ids: ["ses_worker_a", "ses_worker_b"], calls: 2, outputs: ["d1/change-preparation.json"] } } })
      const response = await routes.request("/matter/lean/updates", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: `Instruction during ${step}.`, ownerSessionID: start.ownerSessionID }),
      })
      expect(response.status).toBe(202)
      expect(await response.json()).toMatchObject({ status: { stages: { [step]: { status: "running", session_id: null, session_ids: ["ses_worker_a", "ses_worker_b"], calls: 2, outputs: ["d1/change-preparation.json"] } } } })
    }
    writeQuestions(directory, [{ id: "amendment-choice", stage: "change-plan", document: "d1", question: "Which allocation should the amendment use?", context: "The source leaves this decision open.", status: "open" }])
    const state = readLeanState(directory)!
    writeLeanState(directory, { ...state, status: "blocked", stages: { ...state.stages, "d1.change-plan": { ...state.stages["d1.change-plan"]!, status: "blocked" }, "d1.apply": { ...state.stages["d1.apply"]!, status: "running" } } })
    const late = await routes.request("/matter/lean/updates", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: "Change the allocation.", ownerSessionID: start.ownerSessionID }),
    })
    expect(late.status).toBe(409)
    expect(await late.json()).toMatchObject({ code: "lean_update_too_late" })
    const answered = await routes.request("/matter/lean/questions/amendment-choice", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ answer: "Use the buyer allocation." }),
    })
    expect(answered.status).toBe(202)
    expect(await answered.json()).toMatchObject({ status: { questions: [{ stage: "change-plan", status: "open" }], queued_answer_ids: ["amendment-choice"] } })
    expect(readQuestions(directory)[0]!.status).toBe("open")
  } finally { model.release() }
  await leanDraftingIdle(directory)
})

test("idle blocked checklist accepts a clarification while busy, changed-owner and completed late requests conflict", async () => {
  const directory = leanFixtureMatter("lean-held-clarification-")
  directories.push(directory)
  const inspect = scriptedModel({ ...happyPath,
    "lean-checklist-extract": text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!)
      return JSON.stringify({ items: input.source.document === "request" ? [{ check: "Check the supplied delivery condition.", basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" as const } }] : [] })
    },
    "lean-checklist-check": text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!)
      return JSON.stringify({ item_id: input.item_id, status: "indeterminate", reason: "Inspect applicability.", evidence: input.current_evidence, source_refs: input.item.source_refs, inspection_scope: "Current target.", limitation: "Applicability needs inspection." })
    },
  })
  expect((await runLean(directory, { sessions: inspect.port })).status).toBe("blocked")
  const ownerSessionID = readLeanState(directory)!.owner_session_id
  const model = heldModel()
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => model.deps })
  const send = (owner = ownerSessionID) => routes.request("/matter/lean/updates", {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ ownerSessionID: owner, text: "Inspect the existing delivery condition against the source." }),
  })
  expect((await send("stale-owner")).status).toBe(409)
  expect(pendingLeanRequests(directory)).toHaveLength(0)
  try {
    const admitted = await send()
    expect(admitted.status).toBe(202)
    expect(await admitted.json()).toMatchObject({ status: { updates: [{ kind: "clarification" }] } })
    expect((await send()).status).toBe(409)
  } finally { model.release() }
  await leanDraftingIdle(directory)
  expect(leanDraftingStatus(directory)).toMatchObject({ status: "blocked", updates: [{ kind: "clarification", status: "included" }] })
  expect(readLeanText(directory, "updated-notes.md")).toBeUndefined()
  // Failed inspection keeps the original finding, never a synthetic pass.
  expect(readLeanText(directory, "checklist/checks.json")).toContain('"original_finding"')
  expect(readLeanText(directory, "checklist/checks.json")).toContain('"status": "indeterminate"')
  const settled = readLeanState(directory)!
  writeLeanState(directory, { ...settled, status: "completed" })
  const completed = await send()
  expect(completed.status).toBe(409)
  expect(await completed.json()).toMatchObject({ code: "lean_update_too_late" })
  expect(pendingLeanRequests(directory)).toHaveLength(0)
}, 120_000)


test("checker events time the model wait and the check as numbers, never content; status reports the last progress", async () => {
  const SLEEP_MS = 30
  const directory = leanFixtureMatter("lean-timing-")
  directories.push(directory)
  const scripted = scriptedModel(happyPath)
  // Every reply is held for a known time first, like an engine call in flight.
  const deps: LeanDeps = { sessions: { ...scripted.port, prompt: async (input) => {
    await Bun.sleep(SLEEP_MS)
    return scripted.port.prompt(input)
  } } }
  const routes = leanDraftingRoutes({ directory: () => directory, deps: () => deps })
  enqueueLeanRequest(directory, { kind: "update", text: "Keep the synthetic headings short." })
  expect((await routes.request("/m1/lean/resume", { method: "POST" })).status).toBe(202)
  await leanDraftingIdle(directory)
  const status = await (await routes.request("/m1/lean/state")).json() as { status: string; updated_at: string; last_progress_at: string }
  expect(status.status).toBe("completed")
  const events = readLeanEvents(directory)
  const checkers = events.filter(event => event.event === "checker")
  expect(checkers.length).toBeGreaterThan(5)
  // A review prefetched beside its sibling counts only the wait left when its own step began.
  const prefetched = new Set(["review-independent", "review-wording", "review-client-disadvantage"])
  const checkerKeys = new Set(["at", "stage", "document", "event", "step", "attempt", "issues", "wait_ms", "check_ms"])
  for (const event of checkers) {
    expect(Object.keys(event).filter(key => !checkerKeys.has(key))).toEqual([])
    expect(event.issues).toEqual([])
    expect(event.check_ms).toBeGreaterThanOrEqual(0)
    expect(event.wait_ms).toBeGreaterThanOrEqual(prefetched.has(event.stage) ? 0 : SLEEP_MS)
  }
  const boundaries = events.filter(event => event.event === "started" || event.event === "skipped")
  expect(boundaries.length).toBeGreaterThan(0)
  for (const event of boundaries) expect(event.check_ms).toBeGreaterThanOrEqual(0)
  // Timing is metadata only: no prompt, request, update or source wording reaches the log.
  const log = readFileSync(path.join(leanDir(directory), "events.jsonl"), "utf8")
  const prompts = scripted.prompts.map(prompt => prompt.text.slice(0, 120)).filter(text => text.trim().length >= 20)
  expect(prompts.length).toBeGreaterThan(0)
  for (const text of [FIXTURE_SOURCE_TEXT, "Keep the synthetic headings short.", "Draft a short supply note", ...prompts]) expect(log).not.toContain(text)
  expect(status.last_progress_at >= events.at(-1)!.at).toBe(true)
  expect(status.last_progress_at >= status.updated_at).toBe(true)
}, 120_000)
