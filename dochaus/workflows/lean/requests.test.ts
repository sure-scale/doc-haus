import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { drainLeanRequests, enqueueLeanRequest, pendingLeanRequests, leanUserUpdates, checklistRefinementSelections, isChecklistRefinement } from "./requests"
import { runLean } from "./runner"
import { quoteFound } from "./quotes"
import { leanDir, readLeanJson, readLeanState, readLeanText, readQuestions } from "./state"
import { AGENT, changes, cleanupMatters, draftingNotes, happyPath, matter, sameSession, scriptedModel, type Reply } from "./test-support"

// The lawyer request queue: POSTs append request files, the lean runner is the
// only consumer. Real files, real runner; only the model is scripted.

afterEach(cleanupMatters)

const requestsDir = (directory: string) => path.join(leanDir(directory), "requests")

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
      { question: "Which law should govern?", context: "Determines the legal approach." },
      { question: "Which party do you represent?", context: "Determines whose position to protect." },
    ] }) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
    const questions = readQuestions(directory)
    expect(questions).toHaveLength(2)
    expect(model.prompts).toHaveLength(2)
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
    expect(model.counts()).toEqual([3, 1, 1, 1])
    expect(pendingLeanRequests(directory)).toEqual([])
    const done = readdirSync(path.join(requestsDir(directory), "done"))
    expect(done).toHaveLength(1)
    expect(readLeanJson(directory, `requests/done/${done[0]}`)).toMatchObject({ kind: "continue", outcome: { status: "resume" } })
  }, 120_000)
})

test("a chat update during review persists and reopens earlier work before changes are applied", async () => {
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
  expect(outlines).toHaveLength(2)
  expect(outlines[0]!.text).not.toContain(text)
  expect(outlines[1]!.text).toContain(text)
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
