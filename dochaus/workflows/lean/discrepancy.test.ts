import { afterEach, expect, test } from "bun:test"
import { runDiscrepancyStage } from "./discrepancy"
import { evidenceOutputValid } from "./evidence"
import { writeInstructionsFile } from "./instructions"
import { readLeanJson, readLeanState, writeLeanText } from "./state"
import { AGENT, cleanupMatters, matter, scriptedModel, SOURCE, SOURCE_TEXT, type Reply } from "./test-support"

afterEach(cleanupMatters)

// The discrepancy stage's two jobs, told apart by the gaps job's <differences>
// block; a retry in a session answers as the job that session asked.
const ENTRY = { detail: "Delivery term appears in the source", evidence: [{ source_id: SOURCE, quote: SOURCE_TEXT }], limitations: [], material: false, status: "benign", resolution: "Delivery and payment are distinct events." }
const GAP = { detail: "The supply terms state no governing law; the note needs one.", requirement: { source_id: "request", quote: "Draft a short supply note" },
  evidence: [{ source_id: SOURCE, quote: "Payment is due on delivery." }], limitations: [], material: true, lawyer: true }
const RESTATED = { ...GAP, evidence: ENTRY.evidence }

function jobs(gaps: (attempt: number) => object) {
  const asked = new Map<string | undefined, "differences" | "gaps">()
  const attempts = { differences: 0, gaps: 0 }
  const reply: Reply = (prompt, _index, sessionID) => {
    if (!asked.has(sessionID)) asked.set(sessionID, prompt.includes("<differences>\n") ? "gaps" : "differences")
    const job = asked.get(sessionID)!
    const attempt = attempts[job]++
    return JSON.stringify(job === "differences" ? { entries: [ENTRY], limitations: ["Differences limit."] } : gaps(attempt))
  }
  return { reply, attempts, sessions: () => [...asked.values()] }
}

async function run(directory: string, reply: Reply) {
  const model = scriptedModel({ [AGENT.discrepancy]: reply })
  const result = await runDiscrepancyStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {})
  return { result, model }
}

test("the differences job runs first; the gaps job is shown its entries and the stage writes discrepancy.json once", async () => {
  const directory = matter()
  const scripted = jobs(() => ({ gaps: [GAP], limitations: ["Gaps limit."] }))
  const { result, model } = await run(directory, scripted.reply)
  expect(result).toEqual({ status: "done", outputs: ["discrepancy.json"], openItems: [] })
  expect(scripted.sessions()).toEqual(["differences", "gaps"])
  expect(model.prompts[1]!.text).toContain(ENTRY.detail)
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [ENTRY], gaps: [GAP], limitations: ["Differences limit.", "Gaps limit."] })
  expect(evidenceOutputValid(directory, "discrepancy")).toBe(true)
})

test("a gaps reply that fails its check retries in its own session without asking the differences job again", async () => {
  const directory = matter()
  const scripted = jobs(attempt => ({ gaps: [attempt ? GAP : RESTATED], limitations: [] }))
  const { result, model } = await run(directory, scripted.reply)
  expect(result.status).toBe("done")
  expect(scripted.attempts).toEqual({ differences: 1, gaps: 2 })
  expect(model.sessions).toHaveLength(2)
  const retry = model.prompts[2]!
  expect(retry.sessionID).toBe(model.prompts[1]!.sessionID)
  expect(retry.text).toContain("gaps[0] quotes the same evidence as differences[0]")
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [ENTRY], gaps: [GAP], limitations: ["Differences limit."] })
})

test("a run resumed after the differences job ran asks only the gaps job; a new instruction list does too", async () => {
  const directory = matter()
  const failing = await run(directory, jobs(() => ({ gaps: [RESTATED], limitations: [] })).reply)
  expect(failing.result).toEqual({ status: "failed", reason: expect.stringMatching(/^discrepancy: gaps checker failed after 3 retries: gaps\[0\] quotes the same evidence/) })
  expect(readLeanJson(directory, "discrepancy.json")).toBeUndefined()

  const resumed = jobs(() => ({ gaps: [GAP], limitations: [] }))
  expect((await run(directory, resumed.reply)).result.status).toBe("done")
  expect(resumed.attempts).toEqual({ differences: 0, gaps: 1 })
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [ENTRY], gaps: [GAP], limitations: ["Differences limit."] })

  writeLeanText(directory, "instructions.json", JSON.stringify({ items: [{ id: "i1", text: "Keep the note to one page.", source_id: "request", quote: "Draft a short supply note" }], documents: [] }))
  writeInstructionsFile(directory)
  const instructed = jobs(() => ({ gaps: [], limitations: [] }))
  const again = await run(directory, instructed.reply)
  expect(again.result.status).toBe("done")
  expect(instructed.attempts).toEqual({ differences: 0, gaps: 1 })
  expect(again.model.prompts[0]!.text).toContain("Keep the note to one page.")

  // A later instruction re-keys the differences job.
  writeLeanText(directory, "updated-notes.md", "Use the earlier delivery term.")
  const updated = jobs(() => ({ gaps: [], limitations: [] }))
  expect((await run(directory, updated.reply)).result.status).toBe("done")
  expect(updated.attempts).toEqual({ differences: 1, gaps: 1 })
})
