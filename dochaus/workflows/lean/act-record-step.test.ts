import { afterEach, expect, test } from "bun:test"
import { ACT_RECORD_SAMPLES, readActRecord } from "./act-record"
import { readLeanEvents } from "./events"
import { runLean } from "./runner"
import { writeQuestions, type Question } from "./state"
import { AGENT, actRecordJobs, cleanupMatters, happyPath, listedActs, matter, scriptedModel } from "./test-support"

// The act-record step in the lean walk, driven through the real coordinator
// with a scripted model: where it runs, what it records and when a resume
// runs it again.

afterEach(cleanupMatters)

/** WHAT samples plus the DELIVERABLES job; WHO and HOW are asked only of listed acts. */
const UNLISTED_JOBS = ACT_RECORD_SAMPLES + 1

const answered = (answer: string): Question => ({
  id: "q-scope", stage: "intake", question: "Which goods does the note cover?", context: "The terms do not list the goods.", status: "answered", answer, answered_at: "2026-10-31T09:00:00.000Z",
})

test("the acts step runs once, after research review and before the spine, and records its counts", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  const state = await runLean(directory, { sessions: model.port })

  expect(state.status).toBe("completed")
  const agents = model.prompts.map(prompt => prompt.agent)
  const acts = agents.flatMap((agent, at) => agent === AGENT.acts ? [at] : [])
  expect(acts).toHaveLength(UNLISTED_JOBS)
  expect(Math.min(...acts)).toBeGreaterThan(agents.indexOf(AGENT["research-review"]))
  expect(Math.max(...acts)).toBeLessThan(agents.indexOf(AGENT.spine))
  expect(state.stages.acts).toMatchObject({ status: "done", outputs: ["acts.json"], calls: UNLISTED_JOBS })
  expect(readActRecord(directory)?.status).toBe("none")
  expect(readLeanEvents(directory).filter(event => event.event === "act_record")).toEqual([expect.objectContaining({
    stage: "acts", outcome: "none",
    counts: { acts: 0, deliverables: 0, what_samples: ACT_RECORD_SAMPLES, what_passed: ACT_RECORD_SAMPLES },
    jobs: { what: "passed", deliverables: "passed", who: "not_asked", how: "not_asked" },
  })])
})

test("a listed act is described by WHO and HOW before the spine", async () => {
  const directory = matter()
  const act = { act: "prepare the supply note", source: "request", quote: "Draft a short supply note" }
  const model = scriptedModel({
    ...happyPath, [AGENT.acts]: actRecordJobs({ what: () => JSON.stringify({ acts: [act] }) }),
    // The outline's placement job, now shown the listed act, places it in the one planned document.
    [AGENT.coverage]: prompt => JSON.stringify({ placements: listedActs(prompt).map(({ index }) => ({ act: index, document: "d1" })) }),
  })

  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const record = readActRecord(directory)
  expect(record?.status).toBe("listed")
  expect(record?.acts).toEqual([expect.objectContaining({ ...act, who: expect.objectContaining({ authorised_by: [] }), how: expect.objectContaining({ instrument: null }) })])
  expect(model.count(AGENT.acts)).toBe(UNLISTED_JOBS + 2)
})

test("a resumed unchanged run does not run the acts step again; a changed answer runs it once more", async () => {
  const directory = matter()
  writeQuestions(directory, [answered("Only the goods in the first order.")])
  expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")

  const unchanged = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: unchanged.port })).status).toBe("completed")
  expect(unchanged.count(AGENT.acts)).toBe(0)

  writeQuestions(directory, [answered("Every order placed under the terms.")])
  const changed = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: changed.port })).status).toBe("completed")
  const asked = changed.prompts.filter(prompt => prompt.agent === AGENT.acts)
  expect(asked).toHaveLength(UNLISTED_JOBS)
  // Each job is shown the lawyer's current answer.
  for (const prompt of asked) expect(prompt.text).toContain("Every order placed under the terms.")
})
