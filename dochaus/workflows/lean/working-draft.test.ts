import { afterEach, expect, test } from "bun:test"
import { evidencePrompt, gapsPrompt } from "./evidence"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { lawyerWords, quoteFound } from "./quotes"
import { acceptResearchReply, researchPrompt } from "./research"
import { outlineBasis, stageInputs } from "./stages"
import { initLeanState, readLeanText, writeLeanText, WORKING_DRAFT } from "./state"
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

const PRESENCE = /<working_draft_input>\n([\s\S]*?)\n<\/working_draft_input>/

test("the differences, gaps and intake prompts report a saved working draft as an input without its text or authority", () => {
  const directory = upgraded()
  for (const prompt of [evidencePrompt(directory, "discrepancy"), gapsPrompt(directory), researchPrompt(directory, "intake", "")]) {
    const block = PRESENCE.exec(prompt)?.[1]
    expect(block).toContain("A working draft of the requested document is saved for this run.")
    expect(block).toContain("It is a draft input only: not a source, not a record of facts and not approved wording.")
    expect(block).toContain("Do not quote or cite it as authority")
    expect(block).toContain("do not assume any other document has been supplied")
    expect(block).toContain("The outline and drafting steps receive its text separately from the matter documents.")
    expect(prompt).not.toContain(QUICK)
  }
})

test("the report is absent without a nonblank working draft and stays out of the other pre-draft prompts", () => {
  const plain = matter()
  const blank = matter()
  writeLeanText(blank, WORKING_DRAFT, " \n")
  for (const directory of [plain, blank])
    for (const prompt of [evidencePrompt(directory, "discrepancy"), gapsPrompt(directory), researchPrompt(directory, "intake", "")])
      expect(prompt).not.toContain("<working_draft_input>")
  const directory = upgraded()
  for (const prompt of [researchPrompt(directory, "research", ""), researchPrompt(directory, "research-review", ""), evidencePrompt(directory, "spine")])
    expect(prompt).not.toContain("<working_draft_input>")
  const outline = stagePrompt(directory, "d1.draft", { outline: true })
  expect(outline).toContain("<working_draft>")
  expect(outline).not.toContain("<working_draft_input>")
})

test("only the discrepancy and intake keys carry the working-draft presence context", () => {
  const plain = matter()
  const directory = upgraded()
  const names = (at: string, step: Parameters<typeof stageInputs>[1]) => stageInputs(at, step).map(part => part.name)
  for (const step of ["discrepancy", "intake"] as const) {
    expect(names(directory, step)).toContain("working-draft-context")
    expect(names(plain, step)).not.toContain("working-draft-context")
  }
  for (const step of ["research", "spine"] as const) expect(names(directory, step)).not.toContain("working-draft-context")
})

test("a saved working draft leaves an indispensable missing comparison original open", () => {
  const directory = matter()
  initLeanState(directory, {
    request: "Use the working draft and compare it against the signed supply agreement.",
    ownerSessionID: "ses_owner", model: null, workingDraft: QUICK,
  })
  const prompt = researchPrompt(directory, "intake", "")
  expect(prompt).toContain("compare it against the signed supply agreement")
  expect(PRESENCE.exec(prompt)?.[1]).toContain("do not assume any other document has been supplied")
  const question = "Please supply the signed supply agreement for comparison."
  const accepted = acceptResearchReply(directory, "intake", JSON.stringify({ questions: [{
    question, kind: "fact", context: "The requested comparison depends on the signed agreement.",
  }] }))
  expect(accepted.issues).toEqual([])
  expect(accepted.questions).toMatchObject([{ question, status: "open", kind: "fact" }])
})
