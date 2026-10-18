import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { readFileSync } from "node:fs"
import path from "node:path"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { checkDraftingNotes, readChangeList } from "./checks"
import { memoNotesFile } from "./deliverables"
import { runItemIDs, withoutRefMarkers } from "./delivered-memo"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { openItem } from "./quotes"
import { runLean } from "./runner"
import { acceptStageReply, disclosableItems, LEAN_MAX_RETRIES, stageOutputValid } from "./stages"
import { leanFile, readLeanState, readLeanText, writeLeanText, writeOpenItems } from "./state"
import { AGENT, DRAFT, DRAFTED, cleanupMatters, drafter, happyPath, matter, pidOf, sameSession, scriptedModel, type Reply } from "./test-support"

// In a requested memo an internal id may appear only inside a "[ref: …]"
// marker: delivery removes those markers and nothing else, so an id anywhere
// else would reach the reader. The drafting-notes check holds an item as
// disclosed only by its id in a marker and retries a run id written outside
// one. On the last attempt an internal id alone does not fail the notes: they
// are saved as the model wrote them. Without a requested memo the check is
// unchanged.

afterEach(cleanupMatters)

const DOCUMENT = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const PID = "p0123456789ab"
const CHECKLIST_ID = "source-0123456789ab-3"
const GAP = openItem("research", "research_gap", "No authority was found for the delivery period.", { document: "d1" })
const CONDITION = openItem("apply", "placeholder", "Confirm the condition standard chosen.", { document: "d1", amendment_id: "a1" })
const MISSING = (ids: string[]) =>
  `The drafting notes do not mention these open items: ${ids.join(", ")}. Say what each is and what the lawyer needs to do about it, ending that statement with its exact id as "[ref: <id>]".`
const STRAY = (ids: string[]) =>
  `These internal ids appear outside a "[ref: …]" marker: ${ids.join(", ")}. The memo is for the reader: describe each point in words and keep each id only in the marker that ends its statement.`

/** A matter whose run records a paragraph pointer and checklist ids beside the
 * open items, as runItemIDs reads them. */
function memoMatter(plan: "memo" | "none" | "notes file only" = "memo") {
  const directory = matter()
  if (plan === "memo") writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
  else writeLeanPlan(directory, [DOCUMENT])
  // A delivery selection's rename names a notes file without a memo.
  if (plan === "notes file only") writeLeanText(directory, "plan.json", JSON.stringify({ ...readLeanPlan(directory)!, notes_file: "supply-advice.docx" }))
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "sha", text: "The supplier shall deliver the goods." }]))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: [CHECKLIST_ID, "governing-law", "item-1"], checks: [] }))
  return directory
}

const memo = (...lines: string[]) => ["# Cover memo", "", ...lines, ""].join("\n")
const docxText = (bytes: Uint8Array) => readDocxCommentParagraphs(bytes).map(row => row.paragraph_text).join("\n")
/** A matter source named `name`, so a citation of it by name is supported. */
const addSource = (directory: string, name: string, text: string) => {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(name, "1", text, 0, text.length)
  db.close()
}

/** A supply note run with a requested cover memo whose finalize model writes
 * the same memo on every attempt: its title, then `lines` for the open item
 * ids its prompt lists. A retry in the same session does not list them again.
 * `written` is the memo the model last wrote; `model` resumes the run. */
async function stubbornMemoRun(lines: (ids: string[]) => string[], replies: Record<string, Reply> = {}) {
  const directory = matter()
  writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, with a cover memo on the open points.")
  const draft = drafter(DRAFT)
  const memoDrafter: Reply = (prompt, index) => draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo")
  let ids: string[] = [], written = ""
  const stubborn: Reply = prompt => {
    const blocks = [...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)]
    if (blocks.length) ids = blocks.flatMap(match => [...match[1]!.matchAll(/^- (\S+):/gm)].map(row => row[1]!))
    return written = memo(...lines(ids))
  }
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter, [AGENT.finalize]: stubborn, ...replies })
  const run = await runLean(directory, { sessions: model.port })
  return { directory, run, written, model }
}

describe("a requested memo", () => {
  test("an item id written only in prose is not disclosed, and is also a stray id", () => {
    const directory = memoMatter()
    const notes = memo(`Confirm the condition standard chosen, item ${CONDITION.id}.`, `No authority was found for the delivery period. [ref: ${GAP.id}]`)
    expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([MISSING([CONDITION.id]), STRAY([CONDITION.id])])
  })

  test("an item id inside a ref marker is disclosed", () => {
    const directory = memoMatter()
    const notes = memo(`No authority was found for the delivery period; confirm it before use. [ref: ${GAP.id}]`,
      `- The condition standard is for the lawyer to confirm. [ref: ${CONDITION.id}]`)
    expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([])
  })

  test("one refs marker discloses several ids, in each marker form delivery removes", () => {
    const directory = memoMatter()
    for (const marker of [`[refs: ${GAP.id}; ${CONDITION.id}]`, `[refs: ${GAP.id}, ${CONDITION.id}]`, `[Ref: ${GAP.id}] [ ref :${CONDITION.id} ]`]) {
      const notes = memo(`The delivery period and the condition standard are for the lawyer to confirm. ${marker}`)
      expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([])
      // Check and delivery agree on what a marker is: the delivered line has none.
      expect(withoutRefMarkers(notes)).toBe(memo("The delivery period and the condition standard are for the lawyer to confirm."))
    }
    // An id is matched whole: a longer id in the marker does not disclose it.
    expect(checkDraftingNotes(directory, memo(`Both points are open. [refs: ${GAP.id}0; ${CONDITION.id}]`), [GAP, CONDITION]).issues).toEqual([MISSING([GAP.id])])
  })

  test("a structured run id outside a marker is retried by name, even beside its marker", () => {
    const directory = memoMatter()
    const notes = memo(`Paragraph ${PID} omits the delivery period, against requirement \`${CHECKLIST_ID}\`. [ref: ${GAP.id}]`,
      `- ${CONDITION.id}: the condition standard is for the lawyer to confirm. [ref: ${CONDITION.id}]`)
    expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([STRAY([CONDITION.id, PID, CHECKLIST_ID])])
  })

  test("a short label or word-like id in prose that the run does not record is no stray id", () => {
    const directory = memoMatter()
    const notes = memo(`The supply note carries change b1; the governing-law clause and item-1 of the schedule stand. [refs: ${GAP.id}; ${CONDITION.id}]`)
    expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([])
  })

  test("on the last attempt a structured run id in prose does not fail the notes, which are saved as written", async () => {
    const directory = memoMatter()
    writeOpenItems(directory, [GAP, CONDITION])
    const reply = memo(`Paragraph ${PID} omits the delivery period, against requirement \`${CHECKLIST_ID}\`. [ref: ${GAP.id}]`,
      `- ${CONDITION.id}: The condition standard is for the lawyer to confirm. [ref: ${CONDITION.id}]`)
    // Before the last attempt the issue is retried and nothing is saved.
    expect((await acceptStageReply(directory, "finalize", reply)).issues).toEqual([STRAY([CONDITION.id, PID, CHECKLIST_ID])])
    expect(readLeanText(directory, "drafting-notes.md")).toBeUndefined()
    // On the last attempt the notes are accepted and saved byte for byte.
    expect((await acceptStageReply(directory, "finalize", reply, { final: true })).issues).toEqual([])
    expect(readLeanText(directory, "drafting-notes.md")).toBe(reply)
    // A resume keeps the saved notes, a later prose id included; an item they
    // no longer disclose re-runs them.
    expect(stageOutputValid(directory, "finalize")).toBe(true)
    writeLeanText(directory, "drafting-notes.md", `${reply}\nSee paragraph ${PID}.\n`)
    expect(stageOutputValid(directory, "finalize")).toBe(true)
    writeLeanText(directory, "drafting-notes.md", reply.replace(` [ref: ${GAP.id}]`, ""))
    expect(stageOutputValid(directory, "finalize")).toBe(false)
  })

  test("through the run, a memo that keeps naming open item ids in its prose completes on its last attempt as the model wrote it", async () => {
    // The applied change leaves an open item for the memo to disclose.
    const { directory, run, written } = await stubbornMemoRun(ids => ["The draft follows the supply terms.", "",
      ...ids.map(id => `- ${id}: The lawyer must resolve this before the draft is used. [ref: ${id}]`)], {
      [AGENT.changes]: sameSession(prompt => JSON.stringify({ amendments: [{ pid: pidOf(prompt, DRAFTED),
        updated: "The supplier shall deliver the goods in good condition [open item: confirm the condition standard chosen].", reason: "Drafting judgment.", sources: [] }], questions: [] })),
    })
    expect(run.status).toBe("completed")
    expect(readLeanState(directory)?.stages.finalize?.attempts).toBe(LEAN_MAX_RETRIES + 1)
    const ids = disclosableItems(directory).map(item => item.id)
    expect(ids.length).toBeGreaterThan(0)
    expect(readLeanText(directory, "drafting-notes.md")).toBe(written)
    // Delivery removes only the markers.
    const delivered = readLeanText(directory, "delivered-notes.md")!
    for (const id of ids) expect(delivered).toContain(`- ${id}: The lawyer must resolve this before the draft is used.\n`)
    expect(stageOutputValid(directory, "finalize")).toBe(true)
  }, 120_000)
})

describe("a recorded change id in a requested memo", () => {
  const CHANGE_STRAY = (ids: string[]) =>
    `These change ids appear outside a "[ref: …]" marker: ${ids.join(", ")}. The memo is for the reader: describe each change by what it does, not its id.`
  const amendment = (id: string) => ({ id, pid: PID, existing_sha256: "sha", updated: "The supplier shall deliver the goods promptly.", reason: "Drafting judgment.", sources: [] })
  /** A memo run whose change list records amendments a1 to a3, a withheld
   * amendment w1 and an explanation i1, as the change checks number them. */
  function changeMatter(withMemo = true) {
    const directory = memoMatter(withMemo ? "memo" : "none")
    writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: ["a1", "a2", "a3"].map(amendment), questions: [], shared_paragraphs: [],
      withheld: [{ amendment: amendment("w1"), issues: ["Quote not found."] }], explanations: [{ id: "i1", pid: PID, reason: "Initial revision.", sources: [] }] }))
    return directory
  }
  const disclosed = `[refs: ${GAP.id}; ${CONDITION.id}]`

  test("outside a marker it is retried, alone, in a list, in a range or as an open item's amendment", () => {
    const directory = changeMatter()
    expect(checkDraftingNotes(directory, memo(`- **Cover memo inserted first (a1).** The memo now leads the pack. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY(["a1"])])
    expect(checkDraftingNotes(directory, memo(`- Compelled-disclosure hierarchy (a2, a3 and w1); the explanation i1 stands. ${disclosed}`), [GAP, CONDITION]).issues)
      .toEqual([CHANGE_STRAY(["a2", "a3", "w1", "i1"])])
    expect(checkDraftingNotes(directory, memo(`Three review changes were applied (a1 to a3). ${disclosed}`), [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY(["a1", "a3"])])
    expect(checkDraftingNotes(directory, memo(`As amended by a1 and a2, the price rises; under a3's terms it is fixed. ${disclosed}`), [GAP, CONDITION]).issues)
      .toEqual([CHANGE_STRAY(["a1", "a2", "a3"])])
    // The amendment an open item names is a change id even with no change list.
    expect(checkDraftingNotes(memoMatter(), memo(`The condition standard (a1) is open. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY(["a1"])])
  })

  test("joined to another id by a slash, hyphen or en dash it is still retried", () => {
    const directory = changeMatter()
    expect(checkDraftingNotes(directory, memo(`Three review changes were applied (a1-a3). ${disclosed}`), [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY(["a1", "a3"])])
    expect(checkDraftingNotes(directory, memo(`Changes a2/w1 stand. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY(["a2", "w1"])])
    expect(checkDraftingNotes(directory, memo(`Changes a1–a2 stand. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY(["a1", "a2"])])
    // Joined to a word or inside a longer token it is not.
    expect(checkDraftingNotes(directory, memo(`Rows a1-x and xa1-a2 stand. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([])
  })

  test("inside a marker it is removed at delivery and is no stray id", () => {
    const directory = changeMatter()
    expect(checkDraftingNotes(directory, memo(`The memo now leads the pack. [refs: ${GAP.id}; ${CONDITION.id}; a1]`), [GAP, CONDITION]).issues).toEqual([])
  })

  test("text of the same shape that is not a recorded change id is no stray: a list label, a quotation, a citation, a longer id or another case", () => {
    const directory = changeMatter()
    for (const line of ["Paragraph (a) and item (b) stand.", 'The term sheet says "see note a1 below".', "Rule 3(a1) applies.", "Rows a12 and a1x and Schedule A1 stand.", "Footnote[^a1].",
      "Section 3(1)(a1) of the 1985 Act applies.", "Regulation 4(2)(a1) applies.", "Sub-paragraph (a1) of Schedule 2 applies.",
      "See s. 3(1)(a1), reg. 4(2)(b), clause 12.1(d) and Art. 6(1)(f).", "Schedule 2, para. 3 and Section 21A apply."])
      expect(checkDraftingNotes(directory, memo(`${line} ${disclosed}`), [GAP, CONDITION]).issues).toEqual([])
  })

  // A locator's number starts with a digit; a sub-division's may be a letter
  // in parentheses. Any other token after a locator word is prose.
  test("after a locator word whose number does not start with a digit it is retried before the last attempt", () => {
    const directory = changeMatter()
    for (const [line, id] of Object.entries({ "The new clause (a1) states the period.": "a1", "The clause a1 adds sets the period.": "a1",
      "The section a2 inserts is new.": "a2", "Under clause a1, the period is fixed.": "a1", "The paragraph i1 explains stands.": "i1",
      "The rule w1 withheld is open.": "w1", "**Clause a1.** The period is fixed.": "a1" })) {
      const notes = memo(`${line} ${disclosed}`)
      expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([CHANGE_STRAY([id])])
      expect(checkDraftingNotes(directory, notes, [GAP, CONDITION], { final: true }).issues).toEqual([])
    }
    expect(checkDraftingNotes(directory, memo(`Clauses a2 and part a3 stand; see s. a1, reg. a2, para. (a3), cl. a1, sch. a2 and art. a3. ${disclosed}`), [GAP, CONDITION]).issues)
      .toEqual([CHANGE_STRAY(["a1", "a2", "a3"])])
  })

  test("without a requested memo the notes may name it", () => {
    const directory = changeMatter(false)
    const notes = ["# Drafting notes", "", `- Change a1 was applied and w1 withheld; ${GAP.id} and ${CONDITION.id} are open.`, ""].join("\n")
    expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([])
  })

  test("the memo prompt asks for changes by what they do; the notes prompt is unchanged", () => {
    const rule = "describe each change by what it does, never by its change id"
    expect(stagePrompt(changeMatter(), "finalize")).toContain(rule)
    expect(stagePrompt(changeMatter(false), "finalize")).not.toContain(rule)
  })

  test("the ids matched are the change-list ids the run records, in runItemIDs' order, then each open item's amendment", () => {
    const directory = changeMatter()
    writeLeanPlan(directory, [DOCUMENT, { title: "Supply agreement", purpose: "Full terms", parts: [{ heading: "Price", scope: "Price terms" }] }], undefined, undefined,
      { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    writeLeanText(directory, "d2/changes.json", JSON.stringify({ amendments: ["a1", "a4"].map(amendment), questions: [], shared_paragraphs: [],
      withheld: [{ amendment: amendment("w2"), issues: ["Quote not found."] }], explanations: [{ id: "i2", pid: PID, reason: "Initial revision.", sources: [] }] }))
    const pending = openItem("apply", "unapplied_amendment", "Amendment a9 was not applied.", { document: "d2", amendment_id: "a9" })
    const listed = ["a1", "a2", "a3", "w1", "i1", "a4", "w2", "i2"]
    const named = memo(`Changes a9, ${[...listed].reverse().join(", ")} and b1 are listed. [refs: ${GAP.id}; ${CONDITION.id}; ${pending.id}]`)
    expect(checkDraftingNotes(directory, named, [GAP, CONDITION, pending]).issues).toEqual([CHANGE_STRAY([...listed, "a9"])])
    // The same change-list ids, in the same order, as the run's own id list.
    expect(runItemIDs(directory).filter(id => /^[awi]\d+$/.test(id))).toEqual(listed)
  })

  const OTHER_GAP = openItem("research", "research_gap", "No authority was found for the payment date.", { document: "d1" })
  // The model keeps a change id where the check reads prose: in a heading's
  // parenthesis, in a list of ids, after a clause number, inside a verified
  // quotation opened curly and closed straight, and as the name of a source.
  const STUBBORN = [
    `- **Cover memo inserted first (a1).** The memo now leads the pack. [refs: ${GAP.id}; ${CONDITION.id}]`,
    "- Compelled-disclosure hierarchy (a2, a3, w1); the explanation i1 stands. [ref: a1]",
    "Clause 4 (a2) follows the source, which says “Schedule a1 sets the price\" [source: a1] (source: a1).",
  ]

  test("on the last attempt it does not fail the notes, which are saved as written", async () => {
    const directory = changeMatter()
    addSource(directory, "a1", "Schedule a1 sets the price.")
    writeOpenItems(directory, [GAP, CONDITION])
    const reply = memo(...STUBBORN)
    // Before the last attempt the issue is retried and nothing is saved.
    expect((await acceptStageReply(directory, "finalize", reply)).issues).toEqual([CHANGE_STRAY(["a1", "a2", "a3", "w1", "i1"])])
    expect(readLeanText(directory, "drafting-notes.md")).toBeUndefined()
    // On the last attempt the notes are accepted and saved and rendered as written.
    expect((await acceptStageReply(directory, "finalize", reply, { final: true })).issues).toEqual([])
    expect(readLeanText(directory, "drafting-notes.md")).toBe(reply)
    expect(docxText(readFileSync(leanFile(directory, "drafting-notes.docx")))).toContain("Compelled-disclosure hierarchy (a2, a3, w1); the explanation i1 stands.")
    // A resume keeps the saved notes, so a completed finalize is not re-run.
    expect(stageOutputValid(directory, "finalize")).toBe(true)
  })

  test("on the last attempt other issues still fail, and without a requested memo the notes are saved as written", async () => {
    const memoRun = changeMatter()
    writeOpenItems(memoRun, [GAP, CONDITION, OTHER_GAP])
    expect((await acceptStageReply(memoRun, "finalize", memo(STUBBORN[0]!), { final: true })).issues).toEqual([MISSING([OTHER_GAP.id])])
    expect(readLeanText(memoRun, "drafting-notes.md")).toBeUndefined()
    const plain = changeMatter(false)
    writeOpenItems(plain, [GAP, CONDITION])
    const notes = ["# Drafting notes", "", `- Change a1 was applied and w1 withheld (a2, a3); ${GAP.id} and ${CONDITION.id} are open.`, ""].join("\n")
    expect((await acceptStageReply(plain, "finalize", notes, { final: true })).issues).toEqual([])
    expect(readLeanText(plain, "drafting-notes.md")).toBe(notes)
  })

  test("through the run, a memo that keeps naming a change id completes on its last attempt as the model wrote it", async () => {
    const { directory, run, written } = await stubbornMemoRun(ids => ["The delivery clause now states the period (a1).", "",
      ...ids.map(id => `- The lawyer must resolve this before the draft is used. [ref: ${id}]`)])
    expect(run.status).toBe("completed")
    expect(readChangeList(directory, "d1")?.amendments.map(row => row.id)).toEqual(["a1"])
    expect(readLeanState(directory)?.stages.finalize?.attempts).toBe(LEAN_MAX_RETRIES + 1)
    expect(readLeanText(directory, "drafting-notes.md")).toBe(written)
    const delivered = readLeanText(directory, "delivered-notes.md")!
    expect(delivered).toContain("The delivery clause now states the period (a1).\n")
    expect(delivered).not.toContain("[ref")
  }, 120_000)
})

describe("a planned document id in a requested memo", () => {
  const OTHER = { title: "Supply agreement", purpose: "Full terms", parts: [{ heading: "Price", scope: "Price terms" }] }
  const DOCUMENT_STRAY = (ids: string[]) =>
    `These document ids appear outside a "[ref: …]" marker: ${ids.join(", ")}. The memo is for the reader: name each document by its title or file name, not its id.`
  /** A memo run planning d1 and d2. */
  function twoDocuments(withMemo = true) {
    const directory = matter()
    writeLeanPlan(directory, [DOCUMENT, OTHER], undefined, undefined, withMemo ? { name: "cover memo", file: memoNotesFile([DOCUMENT, OTHER], "cover memo") } : undefined)
    return directory
  }
  const disclosed = `[refs: ${GAP.id}; ${CONDITION.id}]`

  test("outside a marker it is retried, asking for the document's title or file name", () => {
    const directory = twoDocuments()
    expect(checkDraftingNotes(directory, memo(`The agreement (d2) sets the price. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([DOCUMENT_STRAY(["d2"])])
    expect(checkDraftingNotes(directory, memo(`- d1 — Supply note.docx and Agreement (d2) are drafted. ${disclosed}`), [GAP, CONDITION]).issues)
      .toEqual([DOCUMENT_STRAY(["d1", "d2"])])
  })

  test("inside a marker it is removed at delivery and is no stray id", () => {
    const directory = twoDocuments()
    expect(checkDraftingNotes(directory, memo(`The supply agreement sets the price. [refs: ${GAP.id}; ${CONDITION.id}; d2]`), [GAP, CONDITION]).issues).toEqual([])
  })

  test("it is matched whole, not inside a word or a longer id", () => {
    const directory = twoDocuments()
    expect(checkDraftingNotes(directory, memo(`Schedule d12 and the add1 and d2x rows stand. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([])
    expect(checkDraftingNotes(directory, memo(`Rows add1-d2, d1-x and d12/d13 stand. ${disclosed}`), [GAP, CONDITION]).issues).toEqual([])
  })

  test("joined to another id by a slash, hyphen or en dash it is still retried", () => {
    const directory = twoDocuments()
    for (const line of ["d1/d2 are drafted.", "The d1-d2 cross-references agree.", "The d1–d2 terms agree."])
      expect(checkDraftingNotes(directory, memo(`${line} ${disclosed}`), [GAP, CONDITION]).issues).toEqual([DOCUMENT_STRAY(["d1", "d2"])])
  })

  test("after a locator word whose number does not start with a digit it is retried before the last attempt", () => {
    const directory = twoDocuments()
    for (const [line, id] of Object.entries({ "Part d2 sets the price.": "d2", "Schedule d2 sets the price.": "d2", "Article d1 is drafted.": "d1" })) {
      const notes = memo(`${line} ${disclosed}`)
      expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([DOCUMENT_STRAY([id])])
      expect(checkDraftingNotes(directory, notes, [GAP, CONDITION], { final: true }).issues).toEqual([])
    }
    expect(checkDraftingNotes(directory, memo(`Article d1 and rules d2 apply; see ss. d1, regs. d2, paras. d1 and arts. d2. ${disclosed}`), [GAP, CONDITION]).issues)
      .toEqual([DOCUMENT_STRAY(["d1", "d2"])])
  })

  test("without a requested memo the notes may name it", () => {
    const directory = twoDocuments(false)
    const notes = ["# Drafting notes", "", `- d1 and d2 are drafted; ${GAP.id} and ${CONDITION.id} are open.`, ""].join("\n")
    expect(checkDraftingNotes(directory, notes, [GAP, CONDITION]).issues).toEqual([])
  })

  test("the memo prompt asks for documents by title or file name; the notes prompt is unchanged", () => {
    const rule = "name each document by its title or file name, never by its internal id"
    expect(stagePrompt(twoDocuments(), "finalize")).toContain(rule)
    expect(stagePrompt(twoDocuments(false), "finalize")).not.toContain(rule)
  })

  // The model keeps a document id where the check reads prose: opening a list
  // item, in a parenthesis, after its own title, joined to another id and after
  // a determiner, beside a verified quotation and its source citations.
  const KEEPS = [
    `- d1 is drafted, and the agreement (d2) sets the price. ${disclosed}`,
    "The Supply note (d1) and the d1/d2 delivery terms agree; the executed d2 is filed.",
    "The source says “Schedule d1 sets the price\" [source: d1] (source: d1).",
  ]

  test("on the last attempt it does not fail the notes, which are saved as written", async () => {
    const directory = twoDocuments()
    addSource(directory, "d1", "Schedule d1 sets the price.")
    writeOpenItems(directory, [GAP, CONDITION])
    const reply = memo(...KEEPS)
    // Before the last attempt the issue is retried and nothing is saved.
    expect((await acceptStageReply(directory, "finalize", reply)).issues).toEqual([DOCUMENT_STRAY(["d1", "d2"])])
    expect(readLeanText(directory, "drafting-notes.md")).toBeUndefined()
    // On the last attempt the notes are accepted and saved and rendered as written.
    expect((await acceptStageReply(directory, "finalize", reply, { final: true })).issues).toEqual([])
    expect(readLeanText(directory, "drafting-notes.md")).toBe(reply)
    expect(docxText(readFileSync(leanFile(directory, "drafting-notes.docx")))).toContain("the agreement (d2) sets the price")
    // A resume keeps the saved notes, a later prose id included.
    expect(stageOutputValid(directory, "finalize")).toBe(true)
    writeLeanText(directory, "drafting-notes.md", `${reply}\nSee d2.\n`)
    expect(stageOutputValid(directory, "finalize")).toBe(true)
  })

  test("on the last attempt without a requested memo the notes are saved as written", async () => {
    const directory = twoDocuments(false)
    writeOpenItems(directory, [GAP, CONDITION])
    const notes = ["# Drafting notes", "", `- d1 and the agreement (d2) are drafted; ${GAP.id} and ${CONDITION.id} are open.`, ""].join("\n")
    expect((await acceptStageReply(directory, "finalize", notes, { final: true })).issues).toEqual([])
    expect(readLeanText(directory, "drafting-notes.md")).toBe(notes)
    expect(stageOutputValid(directory, "finalize")).toBe(true)
  })

  test("through the run, a memo that keeps naming a document id completes on its last attempt as the model wrote it", async () => {
    const { directory, run, written, model } = await stubbornMemoRun(ids => ["d1 now states the delivery period.", "",
      ...ids.map(id => `- The lawyer must resolve this before the draft is used. [ref: ${id}]`)])
    expect(run.status).toBe("completed")
    expect(readLeanState(directory)?.stages.finalize?.attempts).toBe(LEAN_MAX_RETRIES + 1)
    expect(readLeanText(directory, "drafting-notes.md")).toBe(written)
    const delivered = readLeanText(directory, "delivered-notes.md")!
    expect(delivered).toContain("\nd1 now states the delivery period.\n")
    expect(delivered).not.toContain("[ref")
    // A resume keeps the saved memo: finalize is not asked again.
    const calls = model.count(AGENT.finalize)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    expect(model.count(AGENT.finalize)).toBe(calls)
    expect(readLeanText(directory, "drafting-notes.md")).toBe(written)
  }, 120_000)
})

describe("without a requested memo the check is unchanged", () => {
  const OLD_MISSING = (ids: string[]) => `The drafting notes do not mention these open items: ${ids.join(", ")}. Name each id and say what the lawyer needs to do about it.`
  /** The check as it was: every item id anywhere in the notes. */
  const before = (content: string, ids: string[]) => {
    const missing = ids.filter(id => !content.includes(id))
    return missing.length ? [OLD_MISSING(missing)] : []
  }

  for (const plan of ["none", "notes file only"] as const)
    test(`plan with ${plan === "none" ? "no memo" : "a notes file but no memo"}`, () => {
      const directory = memoMatter(plan)
      const prose = ["# Drafting notes", "", `- ${CONDITION.id}: confirm the condition standard chosen in paragraph ${PID}.`, `- Requirement ${CHECKLIST_ID} is open.`, ""].join("\n")
      expect(checkDraftingNotes(directory, prose, [CONDITION]).issues).toEqual([])
      expect(checkDraftingNotes(directory, prose, [CONDITION]).issues).toEqual(before(prose, [CONDITION.id]))
      expect(checkDraftingNotes(directory, prose, [GAP, CONDITION]).issues).toEqual([OLD_MISSING([GAP.id])])
      expect(checkDraftingNotes(directory, prose, [GAP, CONDITION]).issues).toEqual(before(prose, [GAP.id, CONDITION.id]))
    })
})

describe("a planned document id inside a quotation, a citation, a footnote marker or a locator's number", () => {
  const check = (line: string) => {
    const directory = matter()
    writeLeanPlan(directory, [DOCUMENT, { title: "Supply agreement", purpose: "Full terms", parts: [{ heading: "Price", scope: "Price terms" }] }], undefined, undefined,
      { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    return checkDraftingNotes(directory, memo(`${line} [refs: ${GAP.id}; ${CONDITION.id}]`), [GAP, CONDITION]).issues
  }
  test("is not a stray", () => {
    for (const line of ['Rule 3(d1) applies.', 'The clause says "the d1 schedule governs".', "Footnote[^d1].", "The clause says “the d1 schedule governs”.",
      "Section 3(1)(d1) applies.", "Regulation 4(2)(d2) applies.", "Sub-paragraph (d1) of Schedule 2 applies.", "Article 6(1)(d1) and rule 3(d2) apply.",
      "See ss. 3(d1), regs. 4(d2), paras. 2(d1) and arts. 6(d2)."])
      expect(check(line)).toEqual([])
  })
  test("a standalone id in prose is still caught", () => {
    for (const line of ["The draft d1 is ready.", "Doc 1 (d1).", "Under d1's terms the price is fixed.", "The risk is the buyer's. d1 carries it."]) expect(check(line)).toHaveLength(1)
  })
})

// The notes are lawyer-facing text: a placeholder of their own is a gap no open
// item records, so it is retried. Quoting a body placeholder exactly as its open
// item records it is disclosure, not a gap; an id, in a "[ref: …]" marker or
// not, is not a placeholder.
describe("a placeholder in the drafting notes' own text", () => {
  const HOLE = openItem("apply", "placeholder", "The body contains the placeholder [insert date].", { document: "d1", placeholder: "[insert date]" })
  const NOTES_PLACEHOLDER = (placeholders: string[]) =>
    `The drafting notes leave these placeholders in their own text: ${placeholders.join(", ")}. State the point in words, or quote the body's placeholder exactly as its open item records it.`
  /** Notes that disclose GAP and HOLE as the mode asks, after `line`. A memo's
   * marker may also name an id the run no longer lists; delivery removes it. */
  const notes = (withMemo: boolean, line: string) => withMemo
    ? memo(line, `No authority was found for the delivery period. [refs: ${GAP.id}; apply:placeholder:0123456789ab]`,
      `The body keeps [INSERT DATE] for the start date; supply it. [ref: ${HOLE.id}]`)
    : ["# Drafting notes", "", line, `- ${GAP.id}: research the delivery period.`, `- [${HOLE.id}]: the body keeps [insert date]; supply the start date.`, ""].join("\n")

  for (const withMemo of [false, true])
    test(`is retried ${withMemo ? "in a requested memo" : "without a requested memo"}; a quoted open placeholder, an id and a ref marker are not`, () => {
      const directory = memoMatter(withMemo ? "memo" : "none")
      expect(checkDraftingNotes(directory, notes(withMemo, "The notice period is [insert period]."), [GAP, HOLE]).issues).toEqual([NOTES_PLACEHOLDER(["[insert period]"])])
      expect(checkDraftingNotes(directory, notes(withMemo, "The notice period follows the source."), [GAP, HOLE]).issues).toEqual([])
      // The same bracket is a gap of the notes when no open item records it.
      expect(checkDraftingNotes(directory, notes(withMemo, "The notice period follows the source."), [GAP, { ...HOLE, ref: { document: "d1" } }]).issues)
        .toEqual([NOTES_PLACEHOLDER([withMemo ? "[INSERT DATE]" : "[insert date]"])])
    })

  test("a line of underscores only is not a placeholder, as in the body", () => {
    const directory = memoMatter("none")
    expect(checkDraftingNotes(directory, notes(false, "Signed:\n\n______________"), [GAP, HOLE]).issues).toEqual([])
  })

  test("on the last attempt it does not fail the notes, which are saved as written", async () => {
    const directory = memoMatter()
    writeOpenItems(directory, [GAP, HOLE])
    const reply = notes(true, "The notice period is [insert period].")
    // Before the last attempt the issue is retried and nothing is saved.
    expect((await acceptStageReply(directory, "finalize", reply)).issues).toEqual([NOTES_PLACEHOLDER(["[insert period]"])])
    expect(readLeanText(directory, "drafting-notes.md")).toBeUndefined()
    // On the last attempt the notes are accepted and saved as written; a resume keeps them.
    expect((await acceptStageReply(directory, "finalize", reply, { final: true })).issues).toEqual([])
    expect(readLeanText(directory, "drafting-notes.md")).toBe(reply)
    expect(stageOutputValid(directory, "finalize")).toBe(true)
  })
})
