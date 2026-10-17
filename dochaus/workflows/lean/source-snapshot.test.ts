import { Database } from "bun:sqlite"
import { afterEach, expect, spyOn, test } from "bun:test"
import { leanSteps } from "./plan"
import { matterSourceBasis, matterSourceIDs, withSourceSnapshot } from "./quotes"
import { enqueueLeanRequest, leanUserUpdates } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { readLeanState } from "./state"
import { cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

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
