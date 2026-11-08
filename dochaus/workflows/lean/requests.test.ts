import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { focusedNotesInputs } from "./focused-notes"
import { drainLeanRequests, enqueueLeanRequest, pendingLeanRequests, leanUserUpdates, checklistRefinementSelections, isChecklistRefinement } from "./requests"
import { runLean } from "./runner"
import { quoteFound } from "./quotes"
import { hashStageInputs, leanDir, readLeanJson, readLeanState, readLeanText, readQuestions, writeLeanBytes, writeLeanText, writeQuestions, type LeanStageState, type Question } from "./state"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { withFactOptions } from "./fact-answers"
import { writeLeanPlan } from "./plan"
import { AGENT, changes, cleanupMatters, draftingNotes, happyPath, matter, sameSession, scriptedModel, type Reply } from "./test-support"

// The lawyer request queue: POSTs append request files, the lean runner is the
// only consumer. Real files, real runner; only the model is scripted.

afterEach(cleanupMatters)

const requestsDir = (directory: string) => path.join(leanDir(directory), "requests")
const stageKey = (directory: string, step: "gate" | "checklist-check") => hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))

function blockedGateFixture() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Storage notice", purpose: "Describe storage handling.", parts: [{ heading: "Storage", scope: "Container identification." }] }])
  const offered = withFactOptions({ id: "q-label", stage: "intake", kind: "fact", status: "open", question: "What is the label design?", context: "The design is not supplied." })
  const question: Question = { ...offered, status: "answered", answer_kind: "option", option_id: "o1", answer: offered.options![0]!.text, fact_status: "provisional" }
  writeQuestions(directory, [question])
  writeLeanBytes(directory, "d1/body-applied.docx", Buffer.from("synthetic Word version one"))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p-label", text: "The label design remains unknown.", sha256: "synthetic" }]))
  writeLeanText(directory, "gate.json", JSON.stringify({ status: "blocked", blocking: [{ id: "gate:label", stage: "gate", kind: "question", status: "open", detail: "The unknown design needs a visible marker.", ref: { document: "d1", pid: "p-label", question_id: question.id } }] }))
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 5, substantive_attempts: 5, requested_at: "saved", items: [] }))
  const original = readLeanState(directory)!
  const checklist: LeanStageState = { ...original.stages["checklist-check"]!, status: "done", input_sha: stageKey(directory, "checklist-check") }
  const gate: LeanStageState = { ...original.stages.gate!, status: "blocked", input_sha: stageKey(directory, "gate") }
  const stopped = { ...original, status: "blocked" as const, stages: { ...original.stages,
    "checklist-check": checklist,
    gate,
  } }
  return { directory, stopped, question }
}

describe("lean request queue", () => {
  test("technical checklist refinement waits for settlement and does not become client instructions", () => {
    const directory = matter()
    const original = readLeanState(directory)!
    const request = enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: "requirement-a", scope: "source_collection" }] })
    drainLeanRequests(directory)
    expect(pendingLeanRequests(directory)).toEqual([`${request.id}.json`])
    expect(checklistRefinementSelections(directory)).toEqual([])
    const stopped = { ...original, status: "blocked" as const, stages: { ...original.stages, "checklist-check": { ...original.stages.discrepancy!, status: "blocked" as const } } }
    expect(drainLeanRequests(directory, stopped)).toEqual({ answered: 0, updated: 0, repairRound: false })
    expect(checklistRefinementSelections(directory)).toEqual([{ parent_id: "requirement-a", scope: "source_collection" }])
    expect(leanUserUpdates(directory)).toEqual([])
    expect(readLeanText(directory, "updated-notes.md")).toBeUndefined()
    expect(readLeanState(directory)).toEqual(original)
    expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ outcome: { status: "refinement_included" } })
  })

  test("a Continue after a pause only resumes: an earlier gate or checklist verdict starts no repair or refinement", () => {
    const directory = matter()
    const original = readLeanState(directory)!
    const blocking = [{ id: "gate:omission:earlier", stage: "gate", kind: "omission", detail: "An earlier final check item.", ref: {}, status: "open" }]
    writeLeanText(directory, "gate.json", `${JSON.stringify({ status: "blocked", delivery: null, blocking }, null, 2)}\n`)
    writeLeanText(directory, "repairs.json", `${JSON.stringify({ round: 2, requested_at: "2026-10-01T00:00:00.000Z", items: [], substantive_attempts: 2 }, null, 2)}\n`)
    const repairs = readLeanText(directory, "repairs.json")
    const paused = { ...original, status: "blocked" as const, stop_reason: "Paused at your request.", paused_at: "2026-10-17T00:00:00.000Z" }
    const resume = enqueueLeanRequest(directory, { kind: "continue" })
    expect(drainLeanRequests(directory, paused)).toEqual({ answered: 0, updated: 0, repairRound: false })
    expect(readLeanJson(directory, `requests/done/${resume.id}.json`)).toMatchObject({ outcome: { status: "resume" } })
    expect(readLeanText(directory, "repairs.json")).toBe(repairs)
    // A checklist row still blocked from an earlier pass is not a settled hold.
    const refine = enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: "requirement-a", scope: "target_body" }] })
    drainLeanRequests(directory, { ...paused, stages: { ...original.stages, "checklist-check": { ...original.stages.discrepancy!, status: "blocked" as const } } })
    expect(readLeanJson(directory, `requests/done/${refine.id}.json`)).toMatchObject({ outcome: { status: "rejected" } })
    expect(checklistRefinementSelections(directory)).toEqual([])
    // The same gate verdict, settled rather than paused, is a request for a repair round.
    const { paused_at: _paused, ...unkeyed } = paused
    const settled = { ...unkeyed, stages: { ...unkeyed.stages,
      "checklist-check": { ...unkeyed.stages["checklist-check"]!, status: "done" as const, input_sha: stageKey(directory, "checklist-check") },
      gate: { ...unkeyed.stages.gate!, status: "blocked" as const, input_sha: stageKey(directory, "gate") },
    } }
    enqueueLeanRequest(directory, { kind: "continue" })
    expect(drainLeanRequests(directory, settled)).toMatchObject({ repairRound: true })
    expect(readLeanText(directory, "repairs.json")).not.toBe(repairs)
  })

  for (const stale of ["word", "answer", "verification-contract", "missing-gate-key", "pending-gate", "pending-checklist", "missing-checklist-key"] as const)
    test(`a blocked gate Continue resumes current checks before repair when ${stale} is obsolete or unverifiable`, () => {
      const { directory, stopped, question } = blockedGateFixture()
      const repairs = readLeanText(directory, "repairs.json")
      if (stale === "word") {
        const gateBefore = stopped.stages.gate.input_sha
        const checklistBefore = stopped.stages["checklist-check"].input_sha
        if (typeof gateBefore !== "string" || typeof checklistBefore !== "string") throw new Error("The fixture needs current gate and checklist keys.")
        writeLeanBytes(directory, "d1/body-applied.docx", Buffer.from("synthetic Word version two"))
        // The gate's own input list reads paragraph records; its upstream
        // checklist key is what notices changed accepted Word bytes.
        expect(stageKey(directory, "gate")).toBe(gateBefore)
        expect(stageKey(directory, "checklist-check")).not.toBe(checklistBefore)
      }
      if (stale === "answer") writeQuestions(directory, [{ ...question, answer: "No additional information is available; disclose the unresolved design." }])
      if (stale === "verification-contract") {
        for (const step of ["gate", "checklist-check"] as const) {
          const inputs = stageInputs(directory, step)
          expect(inputs.some(input => input.name === "provisional-fact-marker-contract")).toBe(true)
          stopped.stages[step].input_sha = hashStageInputs(step, LEAN_PROMPT_VERSION, inputs.map(input => input.name === "provisional-fact-marker-contract" ? { ...input, bytes: "previous marker verification contract" } : input))
        }
      }
      if (stale === "missing-gate-key") stopped.stages.gate.input_sha = null
      if (stale === "pending-gate") stopped.stages.gate = { ...stopped.stages.gate, status: "pending" }
      if (stale === "pending-checklist") stopped.stages["checklist-check"] = { ...stopped.stages["checklist-check"], status: "pending" }
      if (stale === "missing-checklist-key") stopped.stages["checklist-check"].input_sha = null
      const resume = enqueueLeanRequest(directory, { kind: "continue" })
      expect(drainLeanRequests(directory, stopped)).toEqual({ answered: 0, updated: 0, repairRound: false })
      expect(readLeanText(directory, "repairs.json")).toBe(repairs)
      expect(readLeanJson(directory, `requests/done/${resume.id}.json`)).toMatchObject({ outcome: { status: "resume" } })
    })

  test("a genuinely current blocked gate still admits the existing manual repair under its allowance", () => {
    const { directory, stopped } = blockedGateFixture()
    writeLeanText(directory, "repairs.json", JSON.stringify({ round: 2, substantive_attempts: 2, requested_at: "saved", items: [] }))
    stopped.stages.gate.input_sha = stageKey(directory, "gate")
    const resume = enqueueLeanRequest(directory, { kind: "continue" })
    expect(drainLeanRequests(directory, stopped)).toMatchObject({ repairRound: true })
    expect(readLeanJson(directory, "repairs.json")).toMatchObject({ round: 3, substantive_attempts: 2 })
    expect(readLeanJson(directory, `requests/done/${resume.id}.json`)).toMatchObject({ outcome: { status: "repair_round" } })
  })

  test("refinement rejects other stop points and malformed selections", () => {
    const directory = matter()
    const request = enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: "requirement-a", scope: "cited_sources" }] })
    drainLeanRequests(directory, readLeanState(directory))
    expect(checklistRefinementSelections(directory)).toEqual([])
    expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ outcome: { status: "rejected" } })
    for (const parents of [[], [{ parent_id: "a", scope: "unknown" }], [{ parent_id: "a", scope: "cited_sources" }, { parent_id: "a", scope: "source_collection" }]])
      expect(isChecklistRefinement({ kind: "refine_checklist", parents })).toBe(false)
  })

  test("requests enqueued at the same time from several processes all land under unique names", async () => {
    const directory = matter()
    const script = `import { enqueueLeanRequest } from ${JSON.stringify(path.join(import.meta.dir, "requests.ts"))}
for (let i = 0; i < 20; i++) enqueueLeanRequest(process.argv[1], { kind: "continue" })`
    const children = Array.from({ length: 4 }, () => Bun.spawn([process.execPath, "-e", script, directory], { stdout: "ignore", stderr: "pipe" }))
    expect(await Promise.all(children.map(child => child.exited))).toEqual([0, 0, 0, 0])
    const pending = pendingLeanRequests(directory)
    expect(pending).toHaveLength(80)
    expect(new Set(pending.map(name => JSON.parse(readFileSync(path.join(requestsDir(directory), name), "utf8")).id)).size).toBe(80)
    expect(readdirSync(path.join(requestsDir(directory), ".incoming"))).toEqual([])
  }, 60_000)

  test("drain processes answers in arrival order and holds a Continue until the run has settled", async () => {
    const directory = matter()
    const first = enqueueLeanRequest(directory, { kind: "answer", question_id: "q_missing", answer: "Yes." })
    const resume = enqueueLeanRequest(directory, { kind: "continue" })
    expect(pendingLeanRequests(directory)).toEqual([`${first.id}.json`, `${resume.id}.json`])

    expect(drainLeanRequests(directory)).toEqual({ answered: 0, updated: 0, repairRound: false })
    expect(pendingLeanRequests(directory)).toEqual([`${resume.id}.json`])
    expect(readLeanJson(directory, `requests/done/${first.id}.json`)).toMatchObject({ kind: "answer", outcome: { status: "rejected" } })

    expect(drainLeanRequests(directory, readLeanState(directory))).toEqual({ answered: 0, updated: 0, repairRound: false })
    expect(pendingLeanRequests(directory)).toEqual([])
    expect(readLeanJson(directory, `requests/done/${resume.id}.json`)).toMatchObject({ kind: "continue", outcome: { status: "resume" } })
  })

  test("a pending answer waits for the other answers; Continue cannot bypass them", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.intake]: () => JSON.stringify({ questions: [
      { question: "Which law should govern?", kind: "instruction", context: "Determines the legal approach." },
      { question: "Which party do you represent?", kind: "fact", context: "Determines whose position to protect." },
    ] }) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
    const questions = readQuestions(directory)
    expect(questions).toHaveLength(2)
    expect(model.prompts.map(row => row.agent)).toEqual([AGENT.instructions, AGENT.discrepancy, AGENT.discrepancy, AGENT.intake, AGENT.advise, AGENT.advise])
    expect(readLeanText(directory, "d1/draft.md")).toBeUndefined()
    expect(readLeanJson(directory, "plan.json")).toBeUndefined()
    enqueueLeanRequest(directory, { kind: "answer", question_id: questions[0]!.id, answer: "Singapore law." })
    enqueueLeanRequest(directory, { kind: "continue" })
    const idle = scriptedModel({})
    expect((await runLean(directory, { sessions: idle.port })).status).toBe("blocked")
    expect(idle.prompts).toHaveLength(0)
    expect(idle.sessions).toHaveLength(0)
    expect(readLeanState(directory)?.stages.research?.status).toBe("pending")
    expect(readQuestions(directory).map(question => question.status)).toEqual(["answered", "open"])
    enqueueLeanRequest(directory, { kind: "answer", question_id: questions[1]!.id, answer: "The supplier." })
    const resumed = scriptedModel(happyPath)
    expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
    for (const agent of [AGENT.intake, AGENT.research, AGENT["research-review"], AGENT.draft, AGENT.review, AGENT.changes, AGENT.finalize]) {
      const prompt = resumed.prompts.find(row => row.agent === agent)!.text
      expect(prompt).toContain("Singapore law.")
      expect(prompt).toContain("The supplier.")
    }
    expect(pendingLeanRequests(directory)).toEqual([])
  }, 120_000)

  test("a Continue that arrives during a run is processed before the run returns", async () => {
    const directory = matter()
    const reviewing: Reply = (prompt, index) => {
      if (index === 0) enqueueLeanRequest(directory, { kind: "continue" })
      return happyPath[AGENT.review]!(prompt, index)
    }
    const model = scriptedModel({ ...happyPath, [AGENT.review]: reviewing })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.counts().slice(0, 3)).toEqual([3, 1, 1])
    // Continue does not repeat any accepted draft/review or focused notes assignment.
    const notes = model.prompts.filter(prompt => prompt.agent === AGENT.finalize)
    expect(notes.map(prompt => prompt.text).toSorted()).toEqual(focusedNotesInputs(directory).map(input => input.text).toSorted())
    expect(new Set(notes.map(prompt => prompt.sessionID)).size).toBe(notes.length)
    expect(pendingLeanRequests(directory)).toEqual([])
    const done = readdirSync(path.join(requestsDir(directory), "done"))
    expect(done).toHaveLength(1)
    expect(readLeanJson(directory, `requests/done/${done[0]}`)).toMatchObject({ kind: "continue", outcome: { status: "resume" } })
  }, 120_000)
})

test("a chat update during review persists and reopens the review, not the accepted draft, before changes are applied", async () => {
  const directory = matter()
  const text = "Use concise headings throughout the document."
  let queued = false
  const model = scriptedModel({ ...happyPath, [AGENT.review]: (prompt, index) => {
    if (!queued) {
      queued = true
      enqueueLeanRequest(directory, { kind: "update", text })
      expect(leanUserUpdates(directory)).toMatchObject([{ text, status: "queued" }])
    }
    return happyPath[AGENT.review]!(prompt, index)
  } })
  const result = await runLean(directory, { sessions: model.port })
  expect(result.status).toBe("completed")
  expect(model.count(AGENT.review)).toBe(2)
  const outlines = model.prompts.filter(p => p.agent === AGENT.draft && p.text.includes("Reply now with the outline"))
  expect(outlines).toHaveLength(1)
  expect(outlines[0]!.text).not.toContain(text)
  expect(model.prompts.filter(p => p.agent === AGENT.review).at(-1)!.text).toContain(text)
  for (const agent of [AGENT.changes, AGENT.finalize])
    expect(model.prompts.find(p => p.agent === agent)!.text).toContain(text)
  expect(readLeanText(directory, "updated-notes.md")).toContain(text)
  expect(readLeanText(directory, "request.md")).not.toContain(text)
  expect(leanUserUpdates(directory)).toMatchObject([{ text, status: "included" }])
  expect(pendingLeanRequests(directory)).toEqual([])
}, 120_000)

test("updates retain arrival order and support exact source quotes without combining messages", () => {
  const directory = matter()
  enqueueLeanRequest(directory, { kind: "update", text: "Use short headings." })
  enqueueLeanRequest(directory, { kind: "update", text: "Keep the definitions in full." })
  expect(drainLeanRequests(directory).updated).toBe(2)
  expect(leanUserUpdates(directory).map(update => update.text)).toEqual(["Use short headings.", "Keep the definitions in full."])
  expect(quoteFound(directory, { source_id: "request", quote: "Keep the definitions in full." })).toBe(true)
  expect(quoteFound(directory, { source_id: "request", quote: "Use short headings. Keep the definitions in full." })).toBe(false)
})
