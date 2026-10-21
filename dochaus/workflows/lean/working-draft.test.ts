import { afterEach, expect, test } from "bun:test"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { lawyerWords, quoteFound } from "./quotes"
import { outlineBasis } from "./stages"
import { initLeanState, readLeanText, WORKING_DRAFT } from "./state"
import { cleanupMatters, matter } from "./test-support"

// An upgraded quick draft (#282) reaches the full run as an unverified working
// draft: the outline and drafting prompts show it, labelled as neither source nor
// instruction, and it never becomes the lawyer's quotable words.

afterEach(cleanupMatters)

const REQUEST = "Draft a short supply note from the supply terms."
const QUICK = "Supply note\n\nThe supplier delivers within a fortnight of the order."

function upgraded() {
  const directory = matter()
  initLeanState(directory, { request: REQUEST, ownerSessionID: "ses_owner", model: { providerID: "test", modelID: "scripted" }, workingDraft: `${QUICK}\r\n` })
  return directory
}

test("the outline and each drafting prompt carry the working draft as unverified, never as the request", () => {
  const directory = upgraded()
  expect(readLeanText(directory, WORKING_DRAFT)).toBe(`${QUICK}\n`)
  const outline = stagePrompt(directory, "d1.draft", { outline: true })
  const block = /<working_draft>\n([\s\S]*?)\n<\/working_draft>/.exec(outline)?.[1]
  expect(block).toContain(QUICK)
  expect(block).toContain("never quote or cite it")
  expect(/<request>\n([\s\S]*?)\n<\/request>/.exec(outline)?.[1]).toBe(REQUEST)
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Summarise delivery.", parts: [{ heading: "Delivery", scope: "Delivery period." }] }])
  expect(stagePrompt(directory, "d1.draft")).toContain(`<working_draft>`)
})

test("the working draft is never the lawyer's words, and only a run with one changes the outline key", () => {
  const plain = matter()
  const directory = matter()
  const before = outlineBasis(directory)
  expect(outlineBasis(plain)).toBe(before)
  initLeanState(directory, { request: REQUEST, ownerSessionID: "ses_owner", model: { providerID: "test", modelID: "scripted" }, workingDraft: QUICK })
  expect(lawyerWords(directory).join("\n")).not.toContain("fortnight")
  expect(quoteFound(directory, { source_id: "request", quote: "The supplier delivers within a fortnight of the order." })).toBe(false)
  expect(quoteFound(directory, { source_id: "request", quote: "Draft a short supply note from the supply terms." })).toBe(true)
  expect(stagePrompt(plain, "d1.draft", { outline: true })).not.toContain("<working_draft>")
  expect(outlineBasis(directory)).not.toBe(before)
})
