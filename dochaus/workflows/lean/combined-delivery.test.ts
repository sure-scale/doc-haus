import { afterEach, describe, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { renderMarkdownToDocx } from "../../lib/render"
import { combinedDelivery, combinedParts, memoNotesFile, recordLeanDocuments } from "./deliverables"
import { checkDeliverySelections, selectedDeliveryPlan } from "./delivery-selection"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan, type PlannedDocument } from "./plan"
import { writeLeanBytes, writeLeanText } from "./state"
import { cleanupMatters, matter, registering, scriptedModel } from "./test-support"

// A document delivered with the requested memo in one file ("Deliverable:
// combined"). The combined file is built at download time from the registered
// body and notes, so registration keeps the bare requested name free: the body
// is registered as "<stem> - clean" and the notes under the memo's name. Real
// registration path; only the registry writes are local.

afterEach(cleanupMatters)

type Document = Omit<PlannedDocument, "id">
const SUPPLY: Document = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const SCHEDULE: Document = { title: "Delivery schedule", purpose: "Delivery dates", parts: [{ heading: "Dates", scope: "Delivery dates" }] }
const COMBINED: Document = { ...SUPPLY, file: "supply-pack.docx", deliverable: "combined" }
const memo = (documents: Document[]) => ({ name: "cover memo", file: memoNotesFile(documents, "cover memo") })

/** Writes each planned document's body and redline and the notes, then
 * registers them as a passed gate does; returns the registered names. */
async function registered(directory: string, documents: Document[]) {
  const files = [...documents.flatMap((_, index) => [`d${index + 1}/body-applied.docx`, `d${index + 1}/redline.docx`]), "drafting-notes.docx", "drafting-notes.md"]
  for (const [index, file] of files.entries())
    writeLeanBytes(directory, file, file.endsWith(".md") ? new TextEncoder().encode(`# Notes ${index}\n`) : (await renderMarkdownToDocx(`# Record ${index}\n\nGenerated record ${index}.`)).bytes)
  const registration = registering()
  const summary = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: registration.registerDeliverable }, true)
  return { names: registration.names, rows: summary.documents.map(row => [row.kind, row.document, row.deliverable]) }
}

describe("registration of a combined document", () => {
  test("registers the body as '<stem> - clean', the notes under the memo's name, and nothing under the requested name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    const { names, rows } = await registered(directory, [COMBINED])
    expect(rows).toEqual([
      ["body", "d1", "supply-pack - clean.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["drafting_notes", null, "supply-pack - cover memo.docx"],
      ["drafting_notes_markdown", null, null],
    ])
    expect(names).not.toContain("supply-pack.docx")
    expect(existsSync(path.join(directory, "supply-pack.docx"))).toBe(false)
  }, 60_000)

  test("without combined, registered names are unchanged", async () => {
    const plain = matter()
    writeLeanPlan(plain, [SUPPLY])
    expect((await registered(plain, [SUPPLY])).rows).toEqual([
      ["body", "d1", "Supply note.docx"],
      ["redline", "d1", "Supply note - redline.docx"],
      ["drafting_notes", null, "Supply note - drafting notes.docx"],
      ["drafting_notes_markdown", null, null],
    ])

    const named = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx", deliverable: "body" }, { ...SCHEDULE, file: "schedule-review.docx", deliverable: "redline" }]
    writeLeanPlan(named, documents, undefined, undefined, memo(documents))
    expect((await registered(named, documents)).rows).toEqual([
      ["body", "d1", "supply-pack.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["body", "d2", "schedule-review - clean.docx"],
      ["redline", "d2", "schedule-review.docx"],
      ["drafting_notes", null, "supply-pack - cover memo.docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)
})

describe("delivery names that collide", () => {
  const named = (documents: Document[], name: string) => ({ name, file: memoNotesFile(documents, name) })

  test("another document named like the combined document does not take the combined file's name", async () => {
    const directory = matter()
    const documents: Document[] = [COMBINED, { ...SCHEDULE, title: "supply-pack" }]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    const { names, rows } = await registered(directory, documents)
    expect(rows).toContainEqual(["body", "d2", "supply-pack (2).docx"])
    expect(names).not.toContain("supply-pack.docx")
  }, 60_000)

  test("the memo does not take a later combined document's name", async () => {
    // A memo file not named after the combined document (an earlier plan's, or a
    // lawyer's edit) can repeat its name.
    const directory = matter()
    const documents: Document[] = [SCHEDULE, { ...SUPPLY, file: "Delivery schedule - cover memo.docx", deliverable: "combined" }]
    writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo", file: "Delivery schedule - cover memo.docx" })
    const { names, rows } = await registered(directory, documents)
    expect(rows).toContainEqual(["drafting_notes", null, "Delivery schedule - cover memo (2).docx"])
    expect(names).not.toContain("Delivery schedule - cover memo.docx")
  }, 60_000)

  test("a memo named like the redline yields: the redline keeps its name", async () => {
    const directory = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx" }]
    writeLeanPlan(directory, documents, undefined, undefined, named(documents, "redline"))
    expect((await registered(directory, documents)).rows).toEqual([
      ["body", "d1", "supply-pack.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["drafting_notes", null, "supply-pack - redline (2).docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)

  test("a memo named like the clean body yields: the combined or redline document's body keeps its name", async () => {
    const combined = matter()
    writeLeanPlan(combined, [COMBINED], undefined, undefined, named([COMBINED], "clean"))
    expect((await registered(combined, [COMBINED])).rows).toEqual([
      ["body", "d1", "supply-pack - clean.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["drafting_notes", null, "supply-pack - clean (2).docx"],
      ["drafting_notes_markdown", null, null],
    ])

    const redline = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx", deliverable: "redline" }]
    writeLeanPlan(redline, documents, undefined, undefined, named(documents, "clean"))
    expect((await registered(redline, documents)).rows).toEqual([
      ["body", "d1", "supply-pack - clean.docx"],
      ["redline", "d1", "supply-pack.docx"],
      ["drafting_notes", null, "supply-pack - clean (2).docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)

  test("without a memo or a combined document, colliding names resolve as before", async () => {
    // A requested name stays its document's; an untitled document with the same title yields.
    const requested = matter()
    const documents: Document[] = [{ ...SCHEDULE, title: "supply-pack" }, { ...SUPPLY, file: "supply-pack.docx" }]
    writeLeanPlan(requested, documents)
    expect((await registered(requested, documents)).rows).toEqual([
      ["body", "d1", "supply-pack (2).docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["body", "d2", "supply-pack.docx"],
      ["redline", "d2", "supply-pack - redline (2).docx"],
      ["drafting_notes", null, "supply-pack - drafting notes.docx"],
      ["drafting_notes_markdown", null, null],
    ])

    // A notes name the delivery selection requested keeps its name.
    const notes = matter()
    const plan = writeLeanPlan(notes, [SUPPLY, SCHEDULE])
    writeLeanText(notes, "plan.json", `${JSON.stringify({ ...plan, notes_file: "Delivery schedule.docx" }, null, 2)}\n`)
    expect((await registered(notes, [SUPPLY, SCHEDULE])).rows).toEqual([
      ["body", "d1", "Supply note.docx"],
      ["redline", "d1", "Supply note - redline.docx"],
      ["body", "d2", "Delivery schedule (2).docx"],
      ["redline", "d2", "Delivery schedule - redline.docx"],
      ["drafting_notes", null, "Delivery schedule.docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)
})

describe("combinedDelivery", () => {
  test("is the combined document's id and its requested delivery stem, given a memo", () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(combinedDelivery(directory)).toEqual({ document: "d1", name: "supply-pack" })
    // Without a requested file name the stem is the title, as the body's would be.
    const untitled = { ...SUPPLY, deliverable: "combined" as const }
    writeLeanPlan(directory, [SCHEDULE, untitled], undefined, undefined, memo([SCHEDULE, untitled]))
    expect(combinedDelivery(directory)).toEqual({ document: "d2", name: "Supply note" })
  })

  test("is undefined without a memo, without a combined document, or with more than one", () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED])
    expect(combinedDelivery(directory)).toBeUndefined()
    const body = { ...COMBINED, deliverable: "body" as const }
    writeLeanPlan(directory, [body], undefined, undefined, memo([body]))
    expect(combinedDelivery(directory)).toBeUndefined()
    const both = [COMBINED, { ...SCHEDULE, deliverable: "combined" as const }]
    writeLeanPlan(directory, both, undefined, undefined, memo(both))
    expect(combinedDelivery(directory)).toBeUndefined()
  })

  test("a name cut at the length limit keeps whole characters", () => {
    // 119 letters, then a character outside the Basic Multilingual Plane at the cut.
    const long = { ...COMBINED, file: `${"a".repeat(119)}\u{1F600}tail.docx` }
    const directory = matter()
    writeLeanPlan(directory, [long], undefined, undefined, memo([long]))
    const name = combinedDelivery(directory)!.name
    expect(name).toBe(`${"a".repeat(119)}\u{1F600}`)
    expect(encodeURIComponent(`${name}.docx`)).toEndWith("%F0%9F%98%80.docx")
    // The memo's file shortens the document's stem, at the same character here.
    const notes = memoNotesFile([{ title: "x", file: `${"a".repeat(106)}\u{1F600}tail.docx` }], "cover memo")
    expect(notes).toBe(`${"a".repeat(106)}\u{1F600} - cover memo.docx`)
  })
})

describe("the requested memo's notes file in a combined package", () => {
  test("is named after the combined document wherever it stands, before another document's requested file", () => {
    expect(memoNotesFile([SCHEDULE, COMBINED], "cover memo")).toBe("supply-pack - cover memo.docx")
    // An earlier document titled like the memo, without a file, does not name it.
    expect(memoNotesFile([{ ...SCHEDULE, title: "Cover memo" }, COMBINED], "cover memo")).toBe("supply-pack - cover memo.docx")
    expect(memoNotesFile([{ ...SCHEDULE, file: "schedule.docx" }, { ...SUPPLY, deliverable: "combined" }], "cover memo")).toBe("Supply note - cover memo.docx")
    // More than one combined document builds no combined file (combinedDelivery).
    expect(memoNotesFile([SCHEDULE, { ...SUPPLY, deliverable: "combined" }, { ...COMBINED, title: "Annex", file: "annex.docx" }], "cover memo")).toBe("annex - cover memo.docx")
  })

  test("registration and the combined file's parts bind the notes to that name", async () => {
    const directory = matter()
    const documents: Document[] = [SCHEDULE, COMBINED]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    const { names, rows } = await registered(directory, documents)
    expect(rows).toContainEqual(["drafting_notes", null, "supply-pack - cover memo.docx"])
    expect(names).not.toContain("supply-pack.docx")
    expect(combinedParts(directory)).toEqual({ document: "d2", name: "supply-pack", body: "supply-pack - clean.docx", notes: "supply-pack - cover memo.docx" })
  }, 60_000)
})

describe("package review's selection on a combined document", () => {
  const INSTRUCTION = "Send the note and the cover memo together as Supply Pack Final.docx"
  const selection = (role: "body" | "redline", file = "Supply Pack Final.docx") => ({ document: "d1" as const, role, file, instruction_quote: INSTRUCTION })

  test("a body selection names the combined file and keeps it combined; the same name is no change", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(checkDeliverySelections(directory, [selection("body")])).toEqual([])

    const selected = selectedDeliveryPlan(directory, [selection("body")])!
    expect(selected.documents[0]).toEqual({ ...COMBINED, id: "d1", file: "Supply Pack Final.docx", deliverable: "combined" })
    expect(selected).toMatchObject({ memo: "cover memo", notes_file: "Supply Pack Final - cover memo.docx" })
    writeLeanText(directory, "plan.json", `${JSON.stringify(selected, null, 2)}\n`)
    expect(combinedDelivery(directory)).toEqual({ document: "d1", name: "Supply Pack Final" })
    expect(selectedDeliveryPlan(directory, [selection("body")])).toBeUndefined()
  })

  test("a redline selection still makes it a redline deliverable", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(selectedDeliveryPlan(directory, [selection("redline")])?.documents[0]).toMatchObject({ file: "Supply Pack Final.docx", deliverable: "redline" })
    expect(selectedDeliveryPlan(directory, [selection("redline", COMBINED.file)])?.documents[0]).toMatchObject({ file: "supply-pack.docx", deliverable: "redline" })
  })

  test("a body selection on a combined document that is not first moves the memo's file with it; renaming another document leaves it", () => {
    const directory = matter()
    const documents: Document[] = [{ ...SCHEDULE, file: "schedule.docx" }, COMBINED]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    expect(selectedDeliveryPlan(directory, [{ ...selection("body"), document: "d2" }])).toMatchObject({ memo: "cover memo", notes_file: "Supply Pack Final - cover memo.docx" })
    expect(selectedDeliveryPlan(directory, [selection("body", "Schedule Final.docx")])).toMatchObject({ memo: "cover memo", notes_file: "supply-pack - cover memo.docx" })
    // A memo file not named this way is left alone.
    writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo", file: "schedule - cover memo.docx" })
    expect(selectedDeliveryPlan(directory, [{ ...selection("body"), document: "d2" }])).toMatchObject({ memo: "cover memo", notes_file: "schedule - cover memo.docx" })
  })

  const PACKAGE_FINDING = "Distinct named primary artifacts need distinct filenames. Report a conflicting naming instruction as a package finding."
  const ONE_ROW = '"Supply Pack Final.docx" is the one file that holds the memo and document d1. Select it by one row only: document "d1", role body, file "Supply Pack Final.docx". Code builds it from the memo and the clean body at download; give no drafting_notes row for it.'
  const notesRow = { document: null, role: "drafting_notes" as const, file: "Supply Pack Final.docx", instruction_quote: INSTRUCTION }

  test("a notes row beside the body row under the combined file's name is refused with the one-row rule, not as a package finding", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(checkDeliverySelections(directory, [notesRow, selection("body")])).toEqual([ONE_ROW])
  })

  test("a notes row under the plan's combined file name is refused with the one-row rule when no body row names the file", () => {
    // Registration would otherwise name the notes after the combined file and suffix them.
    const SCHEDULE_INSTRUCTION = "Deliver the schedule as Schedule Final.docx"
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}. ${SCHEDULE_INSTRUCTION}.`)
    const documents: Document[] = [{ ...COMBINED, file: "supply pack final.docx" }, SCHEDULE]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    expect(checkDeliverySelections(directory, [notesRow])).toEqual([ONE_ROW])
    // A row for another document does not name the combined file.
    expect(checkDeliverySelections(directory, [{ document: "d2", role: "body", file: "Schedule Final.docx", instruction_quote: SCHEDULE_INSTRUCTION }, notesRow])).toEqual([ONE_ROW])
  })

  test("a notes row under a name the reply leaves free, or without a combined file, keeps today's result", () => {
    // The reply renames the combined file, or delivers that document as a redline (selectedDeliveryPlan).
    const NOTE_INSTRUCTION = "Deliver the note itself as Supply Note Final.docx"
    const combined = matter()
    writeLeanText(combined, "request.md", `${INSTRUCTION}. ${NOTE_INSTRUCTION}.`)
    const documents: Document[] = [{ ...COMBINED, file: "Supply Pack Final.docx" }]
    writeLeanPlan(combined, documents, undefined, undefined, memo(documents))
    for (const role of ["body", "redline"] as const)
      expect(checkDeliverySelections(combined, [{ ...selection(role, "Supply Note Final.docx"), instruction_quote: NOTE_INSTRUCTION }, notesRow])).toEqual([])
    // Another notes name in a combined package.
    expect(checkDeliverySelections(combined, [{ ...notesRow, file: "Supply Note Final.docx", instruction_quote: NOTE_INSTRUCTION }])).toEqual([])
    // A body deliverable under that name with the memo, and a combined document without one.
    const plain = matter()
    writeLeanText(plain, "request.md", `${INSTRUCTION}.`)
    const body: Document[] = [{ ...SUPPLY, file: "Supply Pack Final.docx", deliverable: "body" }]
    writeLeanPlan(plain, body, undefined, undefined, memo(body))
    expect(checkDeliverySelections(plain, [notesRow])).toEqual([])
    writeLeanPlan(plain, documents)
    expect(checkDeliverySelections(plain, [notesRow])).toEqual([])
  })

  test("any other repeated file name keeps the package-finding refusal", () => {
    // Two documents under one name in a combined package.
    const combined = matter()
    writeLeanText(combined, "request.md", `${INSTRUCTION}.`)
    const documents: Document[] = [COMBINED, SCHEDULE]
    writeLeanPlan(combined, documents, undefined, undefined, memo(documents))
    expect(checkDeliverySelections(combined, [selection("body"), { ...selection("body"), document: "d2" }])).toEqual([PACKAGE_FINDING])
    // A notes row under a redline row's name.
    expect(checkDeliverySelections(combined, [notesRow, selection("redline")])).toEqual([PACKAGE_FINDING])
    // The same two rows without a combined document.
    const plain = matter()
    writeLeanText(plain, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(plain, [SUPPLY], undefined, undefined, memo([SUPPLY]))
    expect(checkDeliverySelections(plain, [notesRow, selection("body")])).toEqual([PACKAGE_FINDING])
  })

  test("the selection contract states the one-row rule only for a combined package", () => {
    const rule = 'The requested file that holds the memo and document d1 is selected by one row: document "d1", role body, the requested file name. Code builds that file from the memo and the clean body at download, so give no drafting_notes row for it.'
    const combined = matter()
    writeLeanPlan(combined, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(packageReviewPrompt(combined)).toContain(rule)
    const plain = matter()
    writeLeanPlan(plain, [SUPPLY], undefined, undefined, memo([SUPPLY]))
    expect(packageReviewPrompt(plain)).not.toContain("holds the memo and document")
    writeLeanPlan(plain, [COMBINED])
    expect(packageReviewPrompt(plain)).not.toContain("holds the memo and document")
  })
})
