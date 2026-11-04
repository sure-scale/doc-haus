import { afterEach, expect, test } from "bun:test"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { cleanupMatters, matter } from "./test-support"

// Caution attaches to the points the evidence leaves open: every document stage
// reads the qualify-by-point rule, and the research handoff marks an unverified
// point rather than leaving it out.

afterEach(cleanupMatters)

test("every document stage qualifies by point and states an unverified point with a verify marker", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
  for (const step of ["d1.draft", "d1.review", "d1.changes", "finalize"] as const) {
    const prompt = stagePrompt(directory, step)
    expect([step, prompt]).toEqual([step, expect.stringContaining("Qualify by point: qualifying language, placeholders and reservations attach to the specific points the sources leave open or contested.")])
    expect(prompt).toContain("An instruction or answer to use qualifying language applies to those open points, not to the document as a whole.")
    expect(prompt).toContain("A limitation describes the state of the evidence; it never directs omission.")
    expect(prompt).toContain("with its effective date and a visible [open item: verify ...] marker when its official text was not retrieved; it is left out only when nothing was found.")
  }
})
