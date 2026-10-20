import { afterEach, describe, expect, test } from "bun:test"
import { parsePlan, readLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { writeLeanText } from "./state"
import { AGENT, DRAFT, LETTER, changesFirst, cleanupMatters, drafter, happyPath, matter, registering, reviewFirst, scriptedModel, type Reply } from "./test-support"

// The "Memo:" line names a memo the run writes as its notes, so a planned
// document repeating that memo would be drafted, reviewed and redlined twice.
// The parser compares content words; the repeat is retry feedback, and on the
// last attempt the requested memo stays with every planned document. A memo
// that merely names the document it covers is not a repeat of that document.

afterEach(cleanupMatters)

const document = (title: string, ...lines: string[]) => [`# ${title}`, "What it is for.", "Form: other", ...lines, "## Terms", "What it covers."].join("\n")
const outline = (...lines: string[]) => ["Acting for: not stated", ...lines].join("\n")
const REMEDY = 'Remove that document, or, if the request asks for a separate memo with its own audience or purpose, leave out the "Memo:" line.'
const repeats = (title: string, memo: string) => `The document "${title}" repeats the requested memo "${memo}", which the run writes as its notes. ${REMEDY}`

describe("a planned document that repeats the Memo line", () => {
  test("is retry feedback when its title's words are within the memo's, or the memo's within its title's", () => {
    const shorter = "Draft a licence for the workshop unit, with a brief covering note to the tenant."
    const parsed = parsePlan(outline("Memo: brief covering note to the tenant", document("Workshop licence"), document("Covering Note to Tenant")), shorter)
    expect(parsed.issues).toEqual([repeats("Covering Note to Tenant", "brief covering note to the tenant")])

    const same = "Draft a licence for the workshop unit and a covering memo to the tenant on the points still open."
    const titled = parsePlan(outline("Memo: covering memo to the tenant on the points still open", document("Workshop licence"), document("Covering Memo to the Tenant: Points Still Open")), same)
    expect(titled.issues).toEqual([repeats("Covering Memo to the Tenant: Points Still Open", "covering memo to the tenant on the points still open")])

    const longer = "Draft a licence for the workshop unit, with a covering note."
    expect(parsePlan(outline("Memo: covering note", document("Workshop licence"), document("Covering Note for the Tenant")), longer).issues)
      .toEqual([repeats("Covering Note for the Tenant", "covering note")])
  })

  test("a distinct memo document, and outlines without a Memo line, are accepted as before", () => {
    const request = "Draft advice on the tax treatment of the unit sale, with a cover note."
    const distinct = parsePlan(outline("Memo: cover note", document("Advice on Tax Treatment")), request)
    expect(distinct.issues).toEqual([])
    expect(distinct.value).toMatchObject({ memo: "cover note", documents: [{ title: "Advice on Tax Treatment" }] })

    const plain = "Draft a licence for the workshop unit, with a brief covering note to the tenant."
    for (const memo of [[], ["Memo: none"]]) {
      const parsed = parsePlan(outline(...memo, document("Workshop licence"), document("Covering Note to Tenant")), plain)
      expect(parsed.issues).toEqual([])
      expect(parsed.value).not.toHaveProperty("memo")
      expect(parsed.value?.documents.map(row => row.title)).toEqual(["Workshop licence", "Covering Note to Tenant"])
    }
  })

  test("on the last attempt the requested memo is kept, every document stays, and a combined document stays combined", () => {
    const request = "Draft a licence for the workshop unit and a brief covering note to the tenant, sent together as licence-pack.docx."
    const parsed = parsePlan(outline("Memo: brief covering note to the tenant", document("Workshop licence", "File: licence-pack.docx", "Deliverable: combined"), document("Covering Note to Tenant")), request)
    expect(parsed.issues).toEqual([repeats("Covering Note to Tenant", "brief covering note to the tenant")])
    expect(parsed.value?.memo).toBe("brief covering note to the tenant")
    expect(parsed.value?.documents.map(row => [row.title, row.deliverable])).toEqual([["Workshop licence", "combined"], ["Covering Note to Tenant", undefined]])
  })

  test("a memo that names the document it covers is not a repeat of that document", () => {
    const contract = parsePlan(outline("Memo: explanatory note on the Employment Contract", document("Employment Contract")),
      "Draft the Employment Contract for the new CFO and an explanatory note on the Employment Contract for the board.")
    expect(contract.issues).toEqual([])
    expect(contract.value?.memo).toBe("explanatory note on the Employment Contract")

    const lease = parsePlan(outline("Memo: cover memo on the lease amendment", document("Lease Amendment", "File: amendment-pack.docx", "Deliverable: combined")),
      "Draft the lease amendment and a cover memo on the lease amendment, sent together as amendment-pack.docx.")
    expect(lease.issues).toEqual([])
    expect(lease.value?.memo).toBe("cover memo on the lease amendment")
    expect(lease.value?.documents[0]?.deliverable).toBe("combined")

    const note = parsePlan(outline("Memo: note", document("Note Purchase Agreement")), "Draft a Note Purchase Agreement and a note to the client.")
    expect(note.issues).toEqual([])
    expect(note.value?.memo).toBe("note")
  })

  test("a document that is the memo is still a repeat when the memo also names its subject", () => {
    const request = "Draft a supply note, with a cover memo on the supply note."
    const parsed = parsePlan(outline("Memo: cover memo on the supply note", document("Cover memo"), document("Supply note")), request)
    expect(parsed.issues).toEqual([repeats("Cover memo", "cover memo on the supply note")])
    expect(parsed.value?.memo).toBe("cover memo on the supply note")
  })

  test("the outline prompt says the memo is not also planned as a document", () => {
    expect(stagePrompt(matter(), "d1.draft", { outline: true })).toContain('otherwise leave it out. Do not also plan the memo it names as a "# " document.')
  })
})

describe("a repeated memo through the run", () => {
  test("is retried, then the requested memo is kept with both documents and the memo as the run's notes", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, with a covering letter.")
    const draft = drafter(DRAFT, LETTER)
    const memoDrafter: Reply = (prompt, index) => draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: covering letter")
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter, [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes(repeats("Covering letter", "covering letter")))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanPlan(directory)).toHaveProperty("memo", "covering letter")
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title])).toEqual([["d1", "Supply note"], ["d2", "Covering letter"]])
  }, 120_000)

  test("a memo naming the document it covers is not retried and is kept", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, with a covering letter on the supply note.")
    const draft = drafter(DRAFT)
    const memoDrafter: Reply = (prompt, index) => draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: covering letter on the supply note")
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter, [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes("repeats the requested memo"))).toHaveLength(0)
    expect(readLeanPlan(directory)).toHaveProperty("memo", "covering letter on the supply note")
  }, 120_000)
})
