import { afterEach, describe, expect, test } from "bun:test"
import { parsePlan, readLeanPlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { readLeanText, writeLeanText } from "./state"
import { AGENT, COVER_MEMO, DRAFT, changesFirst, cleanupMatters, drafter, happyPath, matter, registering, reviewFirst, scriptedModel, type Reply } from "./test-support"

// The "Memo:" line names the memo the request asks for, and the memo is a
// planned document: the line must name exactly one. That document is the
// workflow's memo, planned as advice that no one executes, drafted last, in its
// own parts, or in the parts of a cover memo when it lists none. A Memo line
// naming no planned document is retry feedback, and on the last attempt the
// memo is added in those parts.

afterEach(cleanupMatters)

const document = (title: string, ...lines: string[]) => [`# ${title}`, "What it is for.", "Form: other", ...lines, "## Terms", "What it covers."].join("\n")
const outline = (...lines: string[]) => ["Acting for: not stated", ...lines].join("\n")
const MEMO_PARTS = ["What this is and what you need to do", "Decisions taken and why", "Discrepancies and how handled", "What the documents do", "Research relied on and its limits"]
const unplanned = (memo: string, titles: string[]) =>
  `The "Memo:" line names "${memo}", but no planned title is that memo; the planned titles are ${titles.map(title => `"${title}"`).join(", ")}. Plan the memo as one "# " document whose title has the memo's words, with "Form: advice", and plan its parts as for any document.`

describe("the Memo line names one planned document", () => {
  test("that document is the memo: last, advice, executed by no one, keeping its title, purpose, file and parts", () => {
    const request = "Draft a licence for the workshop unit and a brief covering note to the tenant, delivered as tenant-note.docx."
    const parsed = parsePlan(outline("Memo: brief covering note to the tenant",
      document("Covering Note to Tenant", "File: tenant-note.docx", "Form: letter", "Executed by: Landlord", "Mode: revise"), document("Workshop licence")), request)
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.memo).toBe("brief covering note to the tenant")
    expect(parsed.value?.documents.map(row => row.title)).toEqual(["Workshop licence", "Covering Note to Tenant"])
    const memo = parsed.value!.documents[1]!
    expect(memo).toMatchObject({ title: "Covering Note to Tenant", purpose: "What it is for.", file: "tenant-note.docx", form: "advice", executed_by: [], mode: "new" })
    expect(memo.parts).toEqual([{ heading: "Terms", scope: "What it covers.", display_heading: "" }])
  })

  test("a memo planned with parts keeps them in order; one planned without parts gets the five", () => {
    const request = "Draft a supply note, with a memo on the pricing issues."
    const issues = ["Pricing mechanism", "Indexation", "Termination exposure", "Recommendation"]
    const parted = parsePlan(outline("Memo: memo on the pricing issues", document("Supply note"),
      "# Memo on the Pricing Issues", "Form: advice", ...issues.flatMap(heading => [`## ${heading}`, `What ${heading.toLowerCase()} covers.`])), request)
    expect(parted.issues).toEqual([])
    expect(parted.value?.documents.at(-1)?.parts.map(part => part.heading)).toEqual(issues)

    const bare = parsePlan(outline("Memo: cover memo", "# Cover memo", "Form: advice", document("Supply note")), "Draft a supply note, with a cover memo.")
    expect(bare.issues).toEqual([])
    const memo = bare.value!.documents.at(-1)!
    expect(memo.parts.map(part => part.heading)).toEqual(MEMO_PARTS)
    expect(memo.parts.every(part => /^At most \d+ words: /.test(part.scope) && part.display_heading === undefined)).toBe(true)
  })

  test("a memorandum titled for its reader is the memo the line names", () => {
    const parsed = parsePlan(outline("Memo: cover memo", document("Supply note"), document("Cover Memorandum to Partner")), "Draft a supply note, with a cover memo to the partner.")
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.documents.map(row => [row.title, row.form])).toEqual([["Supply note", "other"], ["Cover Memorandum to Partner", "advice"]])
  })

  test("a memo naming the document it covers is not that document", () => {
    const parsed = parsePlan(outline("Memo: explanatory note on the Employment Contract", document("Employment Contract"), document("Explanatory note")),
      "Draft the Employment Contract for the new CFO and an explanatory note on the Employment Contract for the board.")
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.documents.map(row => [row.title, row.form])).toEqual([["Employment Contract", "other"], ["Explanatory note", "advice"]])
  })

  test("a Memo line naming no planned document is retry feedback; on the last attempt the memo is added in its parts", () => {
    const parsed = parsePlan(outline("Memo: cover note", document("Advice on Tax Treatment")), "Draft advice on the tax treatment of the unit sale, with a cover note.")
    expect(parsed.issues).toEqual([unplanned("cover note", ["Advice on Tax Treatment"])])
    expect(parsed.value?.documents.map(row => row.title)).toEqual(["Advice on Tax Treatment", "Cover note"])
    expect(parsed.value?.documents[1]).toMatchObject({ form: "advice", executed_by: [], mode: "new" })
    expect(parsed.value?.documents[1]?.purpose).toContain("cover memo")
    expect(parsed.value?.documents[1]?.parts.map(part => part.heading)).toEqual(MEMO_PARTS)
  })

  test("a one-word Memo line names a memo titled with more words", () => {
    const parsed = parsePlan(outline("Memo: memo", document("Supply note"), "# Cover Memo", "Form: advice"), "Draft a supply note, with a memo.")
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.documents.map(row => [row.title, row.form])).toEqual([["Supply note", "other"], ["Cover Memo", "advice"]])
  })

  test("with no title naming the memo, the last attempt adopts the one document planned as advice or titled a memo", () => {
    const request = "Draft a supply note, with a cover note."
    const advice = parsePlan(outline("Memo: cover note", document("Supply note"), "# Advice to the Buyer", "Form: advice"), request)
    expect(advice.issues).toEqual([unplanned("cover note", ["Supply note", "Advice to the Buyer"])])
    expect(advice.value?.documents.map(row => [row.title, row.form])).toEqual([["Supply note", "other"], ["Advice to the Buyer", "advice"]])
    expect(advice.value?.documents[1]?.parts.map(part => part.heading)).toEqual(MEMO_PARTS)

    const titled = parsePlan(outline("Memo: cover note", document("Board Memorandum"), document("Supply note")), request)
    expect(titled.value?.documents.map(row => [row.title, row.form])).toEqual([["Supply note", "other"], ["Board Memorandum", "advice"]])

    // Of two such documents, neither is adopted: the memo is added.
    const two = parsePlan(outline("Memo: cover note", document("Board Memorandum"), "# Advice to the Buyer", "Form: advice", "## Terms"), request)
    expect(two.value?.documents.map(row => row.title)).toEqual(["Board Memorandum", "Advice to the Buyer", "Cover note"])
  })

  test("several documents repeating the memo are retry feedback; on the last attempt the one planned as advice is the memo", () => {
    const request = "Draft a licence for the workshop unit, with a covering note to the tenant."
    const parsed = parsePlan(outline("Memo: covering note to the tenant", document("Covering Note"), "# Covering Note to the Tenant", "Form: advice", document("Workshop licence")), request)
    expect(parsed.issues).toEqual([`The "Memo:" line names "covering note to the tenant", which "Covering Note", "Covering Note to the Tenant" all repeat. Plan the memo as one "# " document, and give the others titles that do not repeat it.`])
    expect(parsed.value?.documents.map(row => [row.title, row.form])).toEqual([["Covering Note", "other"], ["Workshop licence", "other"], ["Covering Note to the Tenant", "advice"]])

    const none = parsePlan(outline("Memo: covering note to the tenant", document("Covering Note to the Tenant"), document("Covering Note"), document("Workshop licence")), request)
    expect(none.value?.documents.map(row => row.title)).toEqual(["Covering Note to the Tenant", "Workshop licence", "Covering Note"])
  })

  test("without a Memo line, or with a non-answer, the plan is as outlined", () => {
    const plain = "Draft a licence for the workshop unit, with a brief covering note to the tenant."
    for (const memo of [[], ["Memo: none"]]) {
      const parsed = parsePlan(outline(...memo, document("Covering Note to Tenant"), document("Workshop licence")), plain)
      expect(parsed.issues).toEqual([])
      expect(parsed.value).not.toHaveProperty("memo")
      expect(parsed.value?.documents.map(row => [row.title, row.form, row.parts.map(part => part.heading)])).toEqual([["Covering Note to Tenant", "other", ["Terms"]], ["Workshop licence", "other", ["Terms"]]])
    }
  })

  test("a memo the request names only by its file name is in the request's words", () => {
    const request = "Please prepare the supply note and cover-memo-to-client.docx."
    const memo = ["# Cover Memo to Client", "File: cover-memo-to-client.docx", "Form: advice", "## Terms", "What it covers."]
    const parsed = parsePlan(outline("Memo: Cover Memo to Client", ...memo, document("Supply note")), request)
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.memo).toBe("Cover Memo to Client")
    expect(parsed.value?.documents.map(row => [row.title, row.form, row.file])).toEqual([["Supply note", "other", undefined], ["Cover Memo to Client", "advice", "cover-memo-to-client.docx"]])
    expect(parsePlan(outline("Memo: Advice Memo", ...memo, document("Supply note")), request).issues[0]).toContain('The "Memo:" line names "Advice Memo", but the request does not use those words.')
  })

  test("without a Memo line, the one document planned as advice, when no other is and another document is planned, is the memo whatever its title, planned last", () => {
    const request = "Please prepare the supply note and cover-memo-to-client.docx."
    const memo = ["# Cover Memo to Client", "Form: advice", "## Terms", "What it covers."]
    const parsed = parsePlan(outline(...memo, document("Supply note")), request)
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.memo).toBe("Cover Memo to Client")
    expect(parsed.value?.documents.map(row => row.title)).toEqual(["Supply note", "Cover Memo to Client"])
    expect(parsed.value?.documents[1]).toMatchObject({ form: "advice", executed_by: [], mode: "new" })

    // Its title need not call it a memo, and another document titled one does
    // not stop it.
    const adopted = (...lines: string[]) => {
      const plan = parsePlan(outline(...lines), request).value
      return [plan?.memo, plan?.documents.map(row => [row.title, row.form])]
    }
    expect(adopted("# Advice on Pricing", "Form: advice", "## Terms", document("Supply note")))
      .toEqual(["Advice on Pricing", [["Supply note", "other"], ["Advice on Pricing", "advice"]]])
    expect(adopted(...memo, document("Board Memorandum"), document("Supply note")))
      .toEqual(["Cover Memo to Client", [["Board Memorandum", "other"], ["Supply note", "other"], ["Cover Memo to Client", "advice"]]])

    // Not when it is planned in another form, an advice letter included,
    // another document is advice too, or it is the only document planned,
    // memo-titled or not: the plan is as outlined.
    const asOutlined = (...lines: string[]) => {
      const plan = parsePlan(outline(...lines), request).value
      expect(plan).not.toHaveProperty("memo")
      return plan?.documents.map(row => row.title)
    }
    expect(asOutlined("# Memorandum of Understanding", "Form: agreement", "## Terms", document("Supply note"))).toEqual(["Memorandum of Understanding", "Supply note"])
    expect(asOutlined("# Cover Memo to Client", "Form: advice letter", "## Terms", document("Supply note"))).toEqual(["Cover Memo to Client", "Supply note"])
    expect(asOutlined(...memo, "# Advice on Pricing", "Form: advice", "## Terms", document("Supply note"))).toEqual(["Cover Memo to Client", "Advice on Pricing", "Supply note"])
    expect(asOutlined("# Advice on Pricing", "Form: advice", "## Terms")).toEqual(["Advice on Pricing"])
    expect(asOutlined("# Memo on Pricing", "Form: advice", "## Terms")).toEqual(["Memo on Pricing"])
  })

  test("a Memo line the request does not word is retry feedback to rename it, and the one advice document stays the memo", () => {
    const name = "Cover Memo on Supply Note Pricing"
    const parsed = parsePlan(outline(`Memo: ${name}`, document("Supply note"), "# Cover Memo on Pricing", "Form: advice", "## Terms", "What it covers."),
      "Draft a supply note, with a cover memo on pricing.")
    expect(parsed.issues).toEqual([`The "Memo:" line names "${name}", but the request does not use those words. Name the memo in a few of the request's own words, as it writes them, and keep the memo planned as its "# " document. Leave the line and that document out only when the request does not ask for a memo or note.`])
    expect(parsed.value?.memo).toBe("Cover Memo on Pricing")
    expect(parsed.value?.documents.map(row => [row.title, row.form])).toEqual([["Supply note", "other"], ["Cover Memo on Pricing", "advice"]])
  })

  test("plan.json records the memo document by its id, and an id outside the plan is not read", () => {
    const directory = matter()
    const plan = writeLeanPlan(directory, [{ title: "Supply note", purpose: "p", parts: [{ heading: "Terms", scope: "s" }] }, { title: "Cover memo", purpose: "p", parts: [{ heading: "Terms", scope: "s" }] }],
      undefined, undefined, { name: "cover memo", document: "d2" })
    expect(JSON.parse(readLeanText(directory, "plan.json")!)).toMatchObject({ memo: "cover memo", memo_document: "d2" })
    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", memo_document: "d2" })
    writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, memo_document: "d3" }))
    expect(readLeanPlan(directory)).toBeUndefined()
  })

  test("the outline prompt names any memo the request asks for in the Memo line and plans its parts as for any document", () => {
    const prompt = stagePrompt(matter(), "d1.draft", { outline: true })
    expect(prompt).toContain('When the request asks for a memo or note alongside the documents, whatever its subject, write one line "Memo:')
    expect(prompt).toContain('plan that memo as a "# " document with that title and "Form: advice"; plan its parts as for any document.')
    expect(prompt).not.toContain("its parts are set by the workflow")
    expect(prompt).not.toContain("own subject, audience or analysis")
    expect(prompt).toContain("A different delivery format does not create another substantive document.")
  })
})


describe("a requested memo through the run", () => {
  test("is planned last as its own document, by id", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, with a cover memo.")
    const draft = drafter(COVER_MEMO, DRAFT)
    const memoDrafter: Reply = (prompt, index) => draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo")
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter, [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
    const run = await runLean(directory, { sessions: model.port, registerDeliverable: registering().registerDeliverable })
    expect([run.status, run.stop_reason]).toEqual(["completed", null])

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes("no planned document is that memo"))).toHaveLength(0)
    const plan = readLeanPlan(directory)
    expect(plan).toMatchObject({ memo: "cover memo", memo_document: "d2" })
    expect(plan?.documents.map(row => [row.id, row.title, row.form])).toEqual([["d1", "Supply note", "other"], ["d2", "Cover memo", "advice"]])
    expect(plan?.documents[1]?.parts.map(part => part.heading)).toEqual(MEMO_PARTS)
  }, 120_000)

  test("a memo the outline never plans is retried, then added as the last document", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, with a cover memo.")
    const draft = drafter(DRAFT, COVER_MEMO)
    const memoDrafter: Reply = (prompt, index) => /^Draft part /m.test(prompt) ? draft(prompt, index)
      : draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo").split("\n# Cover memo")[0]!
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter, [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
    const run = await runLean(directory, { sessions: model.port, registerDeliverable: registering().registerDeliverable })
    expect([run.status, run.stop_reason]).toEqual(["completed", null])

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes(unplanned("cover memo", ["Supply note"])))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title, row.form])).toEqual([["d1", "Supply note", "other"], ["d2", "Cover memo", "advice"]])
    expect(readLeanPlan(directory)?.memo_document).toBe("d2")
  }, 120_000)
})
