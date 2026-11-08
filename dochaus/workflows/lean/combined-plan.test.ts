import { afterEach, describe, expect, test } from "bun:test"
import { operativeDocument, parsePlan, readLeanPlan, writeLeanPlan, type PlannedDocument } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { readLeanText, writeLeanText } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, happyPath, matter, registering, scriptedModel, type Reply } from "./test-support"

// "Deliverable: combined": the outline says one document and the requested memo
// are delivered together in one named file. The parser holds it to a Memo line,
// to one such document and to a document other than the memo; a combined file
// that cannot be built stands as the body on the last attempt. Naming and
// export are outside this contract.

afterEach(cleanupMatters)

const REQUEST = "Draft a supply note and a cover memo, and send them together as supply-pack.docx."
const NEEDS_MEMO = '"Deliverable: combined" puts a document and the requested memo in one file, so it needs a "Memo:" line'
const ONE_DOCUMENT = "only one document can share the file with the memo"
const document = (title: string, ...lines: string[]) => [`# ${title}`, "Form: other", ...lines, "## Delivery", "What is delivered."].join("\n")
const outline = (...lines: string[]) => ["Acting for: not stated", ...lines].join("\n")
const MEMO = "# Cover memo\nForm: advice"

describe("a document delivered with the requested memo in one file", () => {
  test("parses with a Memo line, also for a revision compared against its original", () => {
    const parsed = parsePlan(outline("Memo: cover memo", document("Supply note", "Mode: new", "File: supply-pack.docx", "Deliverable: combined"), MEMO), REQUEST)
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.memo).toBe("cover memo")
    expect(parsed.value?.documents[0]).toMatchObject({ title: "Supply note", mode: "new", file: "supply-pack.docx", deliverable: "combined" })

    const revised = parsePlan(outline("Memo: cover memo", document("Supply note", "Mode: revise", "Compare against: original.docx", "**Deliverable:** Combined"), MEMO), REQUEST)
    expect(revised.issues).toEqual([])
    expect(revised.value?.documents[0]).toMatchObject({ mode: "revise", comparison: "original.docx", deliverable: "combined" })
  })

  test("without a memo it is retry feedback, and the document stands as the body", () => {
    const parsed = parsePlan(outline(document("Supply note", "Deliverable: combined")), REQUEST)
    expect(parsed.issues).toEqual([expect.stringContaining(NEEDS_MEMO)])
    expect(parsed.issues[0]).toContain('otherwise write "Deliverable: body"')
    expect(parsed.value?.documents[0]?.deliverable).toBe("body")
    expect(parsed.value).not.toHaveProperty("memo")

    // A Memo line the request does not support leaves no memo to combine with.
    const ungrounded = parsePlan(outline("Memo: board paper", document("Supply note", "Deliverable: combined")), REQUEST)
    expect(ungrounded.issues).toEqual([expect.stringContaining('names "board paper"'), expect.stringContaining(NEEDS_MEMO)])
    expect(ungrounded.value?.documents[0]?.deliverable).toBe("body")

    // Without combined, a missing memo is not an issue.
    expect(parsePlan(outline(document("Supply note", "Deliverable: body")), REQUEST).issues).toEqual([])
  })

  test("more than one combined document is retry feedback, and each stands as the body", () => {
    const parsed = parsePlan(outline("Memo: cover memo", document("Supply note", "Deliverable: combined"), document("Delivery schedule", "Deliverable: combined"), MEMO), REQUEST)
    expect(parsed.issues).toEqual([expect.stringContaining(ONE_DOCUMENT)])
    expect(parsed.issues[0]).toContain('"Supply note", "Delivery schedule"')
    expect(parsed.value?.documents.map(row => row.deliverable)).toEqual(["body", "body", undefined])
    expect(parsed.value?.memo).toBe("cover memo")

    const one = parsePlan(outline("Memo: cover memo", document("Supply note", "Deliverable: combined"), document("Delivery schedule", "Deliverable: redline"), MEMO), REQUEST)
    expect(one.issues).toEqual([])
    expect(one.value?.documents.map(row => row.deliverable)).toEqual(["combined", "redline", undefined])
  })

  test("under the memo itself it is retry feedback, and the memo keeps no deliverable", () => {
    const parsed = parsePlan(outline("Memo: cover memo", document("Supply note", "File: supply-pack.docx"), `${MEMO}\nDeliverable: combined`), REQUEST)
    expect(parsed.issues).toEqual(['"Deliverable: combined" is under the memo "Cover memo" itself. Write it under the document the request asks to deliver with the memo, and leave it out under the memo.'])
    expect(parsed.value?.documents.map(row => [row.title, row.deliverable])).toEqual([["Supply note", undefined], ["Cover memo", undefined]])
  })

  test("a saved plan accepts combined and rejects an unknown deliverable", () => {
    const directory = matter()
    const [planned] = parsePlan(outline("Memo: cover memo", document("Supply note", "Deliverable: combined"), MEMO), REQUEST).value!.documents
    const plan = writeLeanPlan(directory, [planned!], undefined, undefined, { name: "cover memo" })
    expect(readLeanPlan(directory)?.documents[0]?.deliverable).toBe("combined")
    writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, documents: [{ ...plan.documents[0], deliverable: "bundle" }] }))
    expect(readLeanPlan(directory)).toBeUndefined()
  })

  test("operative drafting inputs do not carry the deliverable", () => {
    const body: PlannedDocument = { id: "d1", title: "Supply note", purpose: "What is delivered.", parts: [{ heading: "Delivery", scope: "What is delivered." }], mode: "new" }
    const combined = operativeDocument({ ...body, deliverable: "combined", file: "supply-pack.docx" })
    expect(combined).not.toHaveProperty("deliverable")
    expect(combined).toEqual(operativeDocument({ ...body, deliverable: "body" }))
  })

  test("the outline prompt states when to write combined", () => {
    expect(stagePrompt(matter(), "d1.draft", { outline: true }))
      .toContain('"Deliverable: combined" when the request asks for this document and its memo in one named file, otherwise "Deliverable: body"; combined needs the "Memo:" line and goes under that document, not under the memo.')
  })

  test("an outline that keeps combined without a memo through the last retry runs with the document as the body", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", REQUEST)
    const deliverables = registering()
    const draft = drafter(DRAFT)
    const combined: Reply = (prompt, index) => {
      const reply = draft(prompt, index)
      return reply.startsWith("Acting for:") ? reply.replace("# Supply note", "# Supply note\nDeliverable: combined") : reply
    }
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: combined })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes(NEEDS_MEMO))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanPlan(directory)?.documents[0]?.deliverable).toBe("body")
    expect(readLeanPlan(directory)).not.toHaveProperty("memo")
    expect(readLeanText(directory, "d1/draft.md")?.split("\n")[0]).toBe("# Supply note")
  }, 120_000)
})
