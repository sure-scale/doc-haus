import { Database } from "bun:sqlite"
import { afterEach, expect, spyOn, test } from "bun:test"
import path from "node:path"
import { leanSteps } from "./plan"
import { matterSourceBasis, matterSourceIDs, snapshotRead, withSourceSnapshot } from "./quotes"
import { checklistRefinementSelections, enqueueLeanRequest, leanUserUpdates } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { readLeanState, writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

// One synchronous operation reads the matter's sources and processed requests
// once. The values, and so every step key, are those of fresh reads.

// Statements that read chunk text in full, as in the latency harness.
const TEXT_SCAN = /^\s*SELECT\b[^;]*\btext\b[^;]*\bFROM chunks\b/i

afterEach(cleanupMatters)

function countScans<T>(work: () => T) {
  const queries = spyOn(Database.prototype, "query")
  try {
    const value = work()
    return { value, scans: queries.mock.calls.filter(([sql]) => TEXT_SCAN.test(String(sql))).length }
  } finally {
    queries.mockRestore()
  }
}

test("every step key inside one snapshot equals a fresh read and the key its finished stage recorded", async () => {
  const directory = matter()
  enqueueLeanRequest(directory, { kind: "update", text: "Keep the headings short." })
  expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
  expect(leanUserUpdates(directory)).toMatchObject([{ status: "included" }])
  const steps = leanSteps(directory)
  expect(steps.length).toBeGreaterThan(5)
  const fresh = steps.map(step => stageInputSha(directory, step))
  const scoped = countScans(() => withSourceSnapshot(() => steps.map(step => stageInputSha(directory, step))))
  expect(scoped.value).toEqual(fresh)
  expect(scoped.scans).toBe(1)
  const state = readLeanState(directory)!
  for (const [index, step] of steps.entries()) if (state.stages[step]?.status === "done") expect(state.stages[step]!.input_sha).toBe(fresh[index]!)
  // A resumed run finds every finished stage current and redoes none of them.
  const resumed = scriptedModel({})
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.prompts).toHaveLength(0)
}, 120_000)

test("a source revised between step boundaries re-keys every later step and asks stale prefetched reviews again", async () => {
  const directory = matter()
  const scripted = scriptedModel(happyPath)
  const revised = { done: false }
  // The revision lands while the first review is asked, after the remaining
  // reviews were keyed and prompted on the original text.
  const port = { ...scripted.port, prompt: (input: Parameters<typeof scripted.port.prompt>[0]) => {
    if (input.agent === AGENT.review && !revised.done) {
      revised.done = true
      const db = new Database(path.join(directory, ".dochaus", "legal.db"))
      try { db.run("UPDATE chunks SET text = text || ' Revised timing applies.'") } finally { db.close() }
    }
    return scripted.port.prompt(input)
  } }
  // Checklist checking finds the extraction taken at drafting stale.
  const first = await runLean(directory, { sessions: port })
  expect(revised.done).toBe(true)
  expect(first.status).toBe("blocked")
  expect(first.stop_reason).toContain("Source extraction is incomplete or stale")
  const steps = leanSteps(directory)
  const revisedAt = steps.indexOf("d1.review")
  for (const [index, step] of steps.slice(0, steps.indexOf("checklist-check")).entries()) {
    expect(first.stages[step]?.status).toBe("done")
    if (index <= revisedAt) expect(first.stages[step]!.input_sha).not.toBe(stageInputSha(directory, step))
    else expect(first.stages[step]!.input_sha).toBe(stageInputSha(directory, step))
  }
  // Each prefetched review keyed on the original text was asked again, not accepted.
  for (const review of ["review-independent", "review-wording", "review-client-disadvantage"] as const)
    expect(scripted.prompts.filter(row => row.agent === AGENT[review])).toHaveLength(2)
  // Resume redoes the stale steps; once current, a further resume asks nothing.
  const resumed = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.prompts.length).toBeGreaterThan(0)
  for (const step of steps) expect(readLeanState(directory)!.stages[step]!.input_sha).toBe(stageInputSha(directory, step))
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
}, 120_000)

test("outside a snapshot every read is fresh; a nested snapshot joins the outer one", () => {
  const directory = matter()
  expect(countScans(() => [matterSourceBasis(directory), matterSourceBasis(directory)]).scans).toBe(2)
  const nested = countScans(() => withSourceSnapshot(() => [matterSourceBasis(directory), withSourceSnapshot(() => matterSourceBasis(directory))]))
  expect(nested.scans).toBe(1)
  expect(nested.value[0]).toBe(nested.value[1]!)
  // Callers get their own copy of the ids, so one cannot alter another's.
  withSourceSnapshot(() => {
    matterSourceIDs(directory).push("injected.docx")
    expect(matterSourceIDs(directory)).not.toContain("injected.docx")
  })
})

test("an async operation is refused and its snapshot ends with the synchronous call", async () => {
  const directory = matter()
  const later = Promise.withResolvers<number>()
  expect(() => withSourceSnapshot(async () => {
    matterSourceBasis(directory)
    await Bun.sleep(1)
    later.resolve(countScans(() => matterSourceBasis(directory)).scans)
  })).toThrow("withSourceSnapshot needs a synchronous operation")
  // After its await the operation reads fresh source text again.
  expect(await later.promise).toBe(1)
  // A failed operation also ends its snapshot.
  expect(() => withSourceSnapshot(() => { throw new Error("stopped") })).toThrow("stopped")
  expect(countScans(() => [matterSourceBasis(directory), matterSourceBasis(directory)]).scans).toBe(2)
})

test("a caller cannot alter a snapshot value another caller reads", () => {
  const directory = matter()
  const selection = { parent_id: "p1", scope: "target_body" as const }
  writeLeanText(directory, "requests/done/000000000000001-000000-aaaaaaaa.json", `${JSON.stringify({ kind: "refine_checklist", parents: [selection], id: "000000000000001-000000-aaaaaaaa", outcome: { status: "refinement_included" } })}\n`)
  writeLeanText(directory, "requests/done/000000000000002-000000-bbbbbbbb.json", `${JSON.stringify({ kind: "update", text: "Keep the headings short.", id: "000000000000002-000000-bbbbbbbb", outcome: { status: "included" } })}\n`)
  withSourceSnapshot(() => {
    const processed = checklistRefinementSelections(directory)
    expect(() => { (processed[0] as { scope: string }).scope = "cited_sources" }).toThrow(TypeError)
    expect(checklistRefinementSelections(directory)).toEqual([selection])
    leanUserUpdates(directory)[0]!.text = "Changed by one caller."
    expect(leanUserUpdates(directory)).toMatchObject([{ text: "Keep the headings short." }])
    // Checklist source inputs share this path: nested values are frozen too.
    const inputs = snapshotRead("nested", () => ({ passages: [{ quote: "within thirty days" }] }))
    expect(() => { inputs.passages[0]!.quote = "within sixty days" }).toThrow(TypeError)
    expect(() => { inputs.passages.push({ quote: "added" }) }).toThrow(TypeError)
    expect(snapshotRead("nested", () => ({ passages: [{ quote: "unread" }] }))).toEqual({ passages: [{ quote: "within thirty days" }] })
  })
})
