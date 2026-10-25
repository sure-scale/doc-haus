// Run from dochaus/: bun test workflows/lean/apply-crossrefs.test.ts
import { afterEach, expect, test } from "bun:test"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { OPEN_ITEM_HIGHLIGHT_FILL, highlightOpenItemsInDocx } from "../../lib/open-items"
import { renderMarkdownToDocx } from "../../lib/render"
import { runApply } from "./apply-stage"
import { shownText, type ParagraphRow } from "./checks"
import { writeLeanPlan } from "./plan"
import { paragraphRows } from "./stages"
import { paragraphFields, unlink } from "../../lib/docx-fields"
import { leanFile, readLeanJson, readLeanState, writeLeanBytes, writeLeanText, type OpenItem } from "./state"
import { cleanupMatters, matter, SOURCE } from "./test-support"

afterEach(cleanupMatters)

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

// A neutral agreement citing clauses, sub-clauses and a schedule, and one
// clause it does not have.
const AGREEMENT = `# Services Agreement

## 1. Definitions

1.1 In this Agreement the terms in clause 2.1 and Schedule 1 apply.

1.2 Clause 3 governs payment; nothing in clause 9 limits it.

## 2. Services

2.1 The Supplier shall provide the Services.

2.2 The Customer shall pay as set out in clauses 3.1 and 3.2.

## 3. Payment

3.1 Invoices are payable within thirty days.

3.2 Late payment under clause 3.1 carries interest, subject to **clause 2.2**.

## Schedule 1 – Services

The services listed here.`

type Amendment = { pid: string; existing_sha256: string; operation?: "delete" | "split"; updated: string | string[] }

async function drafted(markdown = AGREEMENT) {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Services Agreement", form: "agreement", purpose: "", parts: [{ heading: "Services", scope: "" }] }])
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(markdown, undefined, undefined, { form: "agreement" })).bytes)
  const rows = await paragraphRows(leanFile(directory, "d1/body.docx"))
  writeLeanText(directory, "d1/paragraphs.json", `${JSON.stringify(rows, null, 2)}\n`)
  const row = (start: string) => rows.find(row => row.text.startsWith(start))!
  return { directory, row }
}

async function apply(directory: string, amendments: Amendment[]) {
  writeLeanText(directory, "d1/changes.json", JSON.stringify({
    amendments: amendments.map((row, at) => ({ id: `a${at + 1}`, reason: "A change.", sources: [{ source_id: "terms.txt", quote: "terms" }], ...row })),
    explanations: [], questions: [],
  }))
  const run = await runApply(directory, readLeanState(directory)!, "d1")
  expect(run).toMatchObject({ status: "done" })
  const rows = readLeanJson(directory, "d1/paragraphs-applied.json") as ParagraphRow[]
  return { rows, openItems: run.status === "done" ? run.openItems : [] as OpenItem[], fields: fields(readFileSync(leanFile(directory, "d1/body-applied.docx")), rows) }
}

// Each REF field of the applied body: the paragraph it is in, its cached
// number and the paragraph its bookmark is on, as a reader sees them.
function fields(bytes: Uint8Array, rows: ParagraphRow[]) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const shown = new Map(rows.map(row => [row.pid, shownText(row)]))
  const paragraphs = Array.from(doc.getElementsByTagNameNS(W, "p"))
  const pid = (paragraph: Element) => Array.from(paragraph.getElementsByTagNameNS(W, "bookmarkStart"))
    .map(node => node.getAttributeNS(W, "name") ?? "").find(name => name.startsWith("_dh_"))?.slice(4)
  const marks = new Map(paragraphs.flatMap(paragraph => Array.from(paragraph.getElementsByTagNameNS(W, "bookmarkStart"))
    .map(node => [node.getAttributeNS(W, "name")!, shown.get(pid(paragraph) ?? "")] as const)))
  return paragraphs.flatMap(paragraph => {
    const found: { from: string; cached: string; target?: string }[] = []
    let current: { instr: string; cached: string; result: boolean } | undefined
    for (const node of Array.from(paragraph.getElementsByTagNameNS(W, "*"))) {
      const type = node.localName === "fldChar" ? node.getAttributeNS(W, "fldCharType") : undefined
      if (type === "begin") current = { instr: "", cached: "", result: false }
      else if (type === "separate" && current) current.result = true
      else if (type === "end" && current) {
        const name = /REF (\S+)/.exec(current.instr)?.[1] ?? ""
        found.push({ from: shown.get(pid(paragraph) ?? "")!.slice(0, 4), cached: current.cached, target: marks.has(name) ? marks.get(name) ?? "(outside the body rows)" : undefined })
        current = undefined
      } else if (node.localName === "instrText" && current) current.instr += node.textContent
      else if (node.localName === "t" && current?.result) current.cached += node.textContent
    }
    return found
  })
}

// The tracked changes of a document's redline, in document order: what each
// inserts or deletes, by the opening words of its paragraph as accepted; how
// many field codes are tracked (a field the comparer could not match), and how
// many formatting changes.
function revisions(directory: string) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(readFileSync(leanFile(directory, "d1/redline.docx")))["word/document.xml"]!), "application/xml")
  const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
  const text = (node: Element, name: string) => all(node, name).map(t => t.textContent).join("")
  const changes = all(doc, "p").flatMap(paragraph => (["ins", "del"] as const).flatMap(kind => all(paragraph, kind)
    .map(node => [(text(paragraph, "t") || text(paragraph, "delText")).slice(0, 12), kind, text(node, kind === "ins" ? "t" : "delText")])
    .filter(([, , changed]) => changed)))
  const fields = ["ins", "del"].flatMap(kind => all(doc, kind)).flatMap(node => ["fldChar", "instrText", "delInstrText"].flatMap(name => all(node, name))).length
  return { changes, fields, formatting: all(doc, "rPrChange").length }
}

// Each paragraph's XML in a body, by its paragraph pointer.
function paragraphXml(directory: string, file: string) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(readFileSync(leanFile(directory, file)))["word/document.xml"]!), "application/xml")
  const pid = (paragraph: Element) => Array.from(paragraph.getElementsByTagNameNS(W, "bookmarkStart")).map(node => node.getAttributeNS(W, "name") ?? "").find(name => name.startsWith("_dh_"))?.slice(4)
  return new Map(Array.from(doc.getElementsByTagNameNS(W, "p")).map(paragraph => [pid(paragraph), new XMLSerializer().serializeToString(paragraph)]))
}

// The text of each run of a body shaded as an open item, in document order.
function shadedOpenItems(bytes: Uint8Array) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  return Array.from(doc.getElementsByTagNameNS(W, "r"))
    .filter(run => Array.from(run.getElementsByTagNameNS(W, "shd")).some(shd => shd.getAttributeNS(W, "fill") === OPEN_ITEM_HIGHLIGHT_FILL))
    .map(run => Array.from(run.getElementsByTagNameNS(W, "t")).map(text => text.textContent).join(""))
}

test("the redline marks only the words that changed where a change turns a cross-reference into an open item or renumbers it", async () => {
  const { directory, row } = await drafted()
  const removed = row("The Supplier shall provide")
  await apply(directory, [{ pid: removed.pid, existing_sha256: removed.sha256, operation: "delete", updated: "" }])
  expect(revisions(directory)).toEqual({
    changes: [
      ["In this Agre", "ins", "[open item: refers to "], ["In this Agre", "ins", ", which was removed]"],
      ["The Supplier", "del", "The Supplier shall provide the Services."],
      ["Late payment", "ins", "2.1"], ["Late payment", "del", "2.2"],
    ],
    fields: 0,
    formatting: 0,
  })
}, 60_000)

test("an amendment beside a cross-reference keeps it live; one that rewrites its number links the new number", async () => {
  const { directory, row } = await drafted()
  const beside = row("In this Agreement")
  const rewrite = row("The Customer shall pay")
  const { rows, openItems, fields: after } = await apply(directory, [
    { pid: beside.pid, existing_sha256: beside.sha256, updated: "In this Agreement the defined terms in clause 2.1 and Schedule 1 apply." },
    { pid: rewrite.pid, existing_sha256: rewrite.sha256, updated: "The Customer shall pay as set out in clauses 3.1 and 1.2." },
  ])
  expect(rows.map(shownText)).toContain("1.1 In this Agreement the defined terms in clause 2.1 and Schedule 1 apply.")
  expect(rows.map(shownText)).toContain("2.2 The Customer shall pay as set out in clauses 3.1 and 1.2.")
  expect(after.map(field => [field.from, field.cached, field.target?.slice(0, 4)])).toEqual([
    ["1.1 ", "2.1", "2.1 "], ["1.1 ", "1", "Sche"], ["1.2 ", "3", "3. P"], ["2.2 ", "3.1", "3.1 "], ["2.2 ", "1.2", "1.2 "], ["3.2 ", "3.1", "3.1 "], ["3.2 ", "2.2", "2.2 "],
  ])
  // The one to a clause the agreement does not have stays text and is reported.
  expect(openItems.filter(item => item.kind === "consistency").map(item => item.detail)).toEqual([
    "Clause 1.2 refers to clause 9, but no clause in the final text is numbered 9.",
  ])
}, 60_000)

test("a reference an amendment extends or writes points at the clause it names; the paragraphs no amendment changed keep their XML", async () => {
  const { directory, row } = await drafted()
  const extended = row("Clause 3 governs")
  const written = row("Invoices are payable")
  const stale = row("The Customer shall pay")
  const { rows, fields: after } = await apply(directory, [
    { pid: extended.pid, existing_sha256: extended.sha256, updated: "Clause 3.2 governs payment; nothing in clause 9 limits it." },
    { pid: written.pid, existing_sha256: written.sha256, updated: "Invoices are payable within thirty days of delivery under clause 2.1." },
    // Written against other text, it is not applied: its paragraph stays as drafted.
    { pid: stale.pid, existing_sha256: "0".repeat(64), updated: "The Customer shall pay as set out in clause 1.2." },
  ])
  expect(rows.map(shownText)).toEqual(expect.arrayContaining([
    "1.2 Clause 3.2 governs payment; nothing in clause 9 limits it.",
    "3.1 Invoices are payable within thirty days of delivery under clause 2.1.",
  ]))
  // "3" and the ".2" the amendment typed after it are one reference, to 3.2.
  expect(after.map(field => [field.from, field.cached, field.target?.slice(0, 4)])).toEqual([
    ["1.1 ", "2.1", "2.1 "], ["1.1 ", "1", "Sche"], ["1.2 ", "3.2", "3.2 "], ["2.2 ", "3.1", "3.1 "], ["2.2 ", "3.2", "3.2 "], ["3.1 ", "2.1", "2.1 "], ["3.2 ", "3.1", "3.1 "], ["3.2 ", "2.2", "2.2 "],
  ])
  const drafted_ = paragraphXml(directory, "d1/body.docx")
  const applied = paragraphXml(directory, "d1/body-applied.docx")
  const unchanged = [...drafted_.keys()].filter(pid => pid && pid !== extended.pid && pid !== written.pid)
  expect(unchanged).toContain(stale.pid)
  for (const pid of unchanged) expect(applied.get(pid)).toBe(drafted_.get(pid)!)
  // The redline shows the words the amendments changed, not their paragraphs replaced.
  expect(revisions(directory)).toEqual({
    changes: [
      ["Clause 3.2 g", "ins", "3"], ["Clause 3.2 g", "ins", ".2"], ["Clause 3.2 g", "del", "3"],
      ["Invoices are", "ins", " of delivery under clause "], ["Invoices are", "ins", "2.1"],
    ],
    fields: 0,
    formatting: 0,
  })
}, 60_000)

test("a reference an amendment writes stays text in a document kept to typed numbers and in advice", async () => {
  for (const kept of ["typed", "advice"] as const) {
    const directory = matter()
    // The original the document revises types its clause numbers.
    if (kept === "typed") writeFileSync(path.join(directory, SOURCE), (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { numbering: false, form: "agreement" })).bytes)
    writeLeanPlan(directory, [{ title: "Services Agreement", form: kept === "advice" ? "advice" : "agreement", purpose: "", parts: [{ heading: "Services", scope: "" }], ...(kept === "typed" ? { comparison: SOURCE } : {}) }])
    writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: kept === "advice" ? "advice" : undefined })).bytes)
    const drafted = await paragraphRows(leanFile(directory, "d1/body.docx"))
    writeLeanText(directory, "d1/paragraphs.json", `${JSON.stringify(drafted, null, 2)}\n`)
    const written = drafted.find(row => row.text.startsWith("Invoices are payable"))!
    const { rows, fields: after } = await apply(directory, [{ pid: written.pid, existing_sha256: written.sha256, updated: "Invoices are payable within thirty days of delivery under clause 2.1." }])
    expect(rows.map(shownText)).toContain("3.1 Invoices are payable within thirty days of delivery under clause 2.1.")
    expect([kept, after]).toEqual([kept, []])
  }
}, 60_000)

test("a changed paragraph citing a clause the changes removed keeps the open item, not a reference to the clause now under that number", async () => {
  const { directory, row } = await drafted()
  const removed = row("The Supplier shall provide")
  const citing = row("In this Agreement")
  const { rows, openItems, fields: after } = await apply(directory, [
    { pid: removed.pid, existing_sha256: removed.sha256, operation: "delete", updated: "" },
    { pid: citing.pid, existing_sha256: citing.sha256, updated: "In this Agreement the defined terms in clause 2.1 and Schedule 1 apply." },
  ])
  const marker = "[open item: refers to 2.1, which was removed]"
  expect(rows.map(shownText)).toContain(`1.1 In this Agreement the defined terms in clause ${marker} and Schedule 1 apply.`)
  expect(after.filter(field => field.from === "1.1 ").map(field => [field.cached, field.target?.slice(0, 4)])).toEqual([["1", "Sche"]])
  expect(openItems.filter(item => item.ref.placeholder === marker).map(item => item.ref.pid)).toEqual([citing.pid])
}, 60_000)

test("deleting a clause shows the new numbers of the clauses after it; a reference to the deleted clause becomes an open item", async () => {
  const { directory, row } = await drafted()
  const removed = row("The Supplier shall provide")
  const { rows, openItems, fields: after } = await apply(directory, [{ pid: removed.pid, existing_sha256: removed.sha256, operation: "delete", updated: "" }])
  expect(rows.map(shownText)).not.toContain("2.1 The Supplier shall provide the Services.")
  expect(rows.map(shownText)).toContain("3.2 Late payment under clause 3.1 carries interest, subject to clause 2.1.")
  expect(after.map(field => [field.from, field.cached, field.target?.slice(0, 4)])).toEqual([
    ["1.1 ", "1", "Sche"], ["1.2 ", "3", "3. P"], ["2.1 ", "3.1", "3.1 "], ["2.1 ", "3.2", "3.2 "], ["3.2 ", "3.1", "3.1 "], ["3.2 ", "2.1", "2.1 "],
  ])
  // Its old number now names another clause, so it names nothing and is reported.
  const marker = "[open item: refers to 2.1, which was removed]"
  expect(rows.map(shownText)).toContain(`1.1 In this Agreement the terms in clause ${marker} and Schedule 1 apply.`)
  expect(openItems.filter(item => item.ref.placeholder === marker).map(item => item.ref.pid)).toEqual([row("In this Agreement").pid])
}, 60_000)

test("deleting a sub-clause two clauses cite leaves no reference to the clause now under its number", async () => {
  const { directory, row } = await drafted()
  const removed = row("Invoices are payable")
  const { rows, openItems, fields: after } = await apply(directory, [{ pid: removed.pid, existing_sha256: removed.sha256, operation: "delete", updated: "" }])
  const marker = "[open item: refers to 3.1, which was removed]"
  expect(rows.map(shownText)).toEqual(expect.arrayContaining([
    `2.2 The Customer shall pay as set out in clauses ${marker} and 3.1.`,
    `3.1 Late payment under clause ${marker} carries interest, subject to clause 2.2.`,
  ]))
  expect(after.map(field => [field.from, field.cached, field.target?.slice(0, 4)])).toEqual([
    ["1.1 ", "2.1", "2.1 "], ["1.1 ", "1", "Sche"], ["1.2 ", "3", "3. P"], ["2.2 ", "3.1", "3.1 "], ["3.1 ", "2.2", "2.2 "],
  ])
  expect(openItems.filter(item => item.ref.placeholder === marker).map(item => item.ref.pid).toSorted())
    .toEqual([row("The Customer shall pay").pid, row("Late payment").pid].toSorted())
}, 60_000)

test("a changed paragraph keeps a live reference the change did not rewrite where a removed clause's open item breaks the list it was in", async () => {
  const { directory, row } = await drafted()
  const removed = row("Invoices are payable")
  const citing = row("The Customer shall pay")
  const { rows, fields: after } = await apply(directory, [
    { pid: removed.pid, existing_sha256: removed.sha256, operation: "delete", updated: "" },
    { pid: citing.pid, existing_sha256: citing.sha256, updated: "The Customer shall promptly pay as set out in clauses 3.1 and 3.2." },
  ])
  const marker = "[open item: refers to 3.1, which was removed]"
  expect(rows.map(shownText)).toContain(`2.2 The Customer shall promptly pay as set out in clauses ${marker} and 3.1.`)
  // The former 3.2, now 3.1, still follows its clause.
  expect(after.filter(field => field.from === "2.2 ").map(field => [field.cached, field.target?.slice(0, 4)])).toEqual([["3.1", "3.1 "]])
  expect(after.find(field => field.from === "2.2 ")?.target).toStartWith("3.1 Late payment")
}, 60_000)

test("a clause citing itself keeps the number typed, when drafted and when a change writes it", async () => {
  const own = "Invoices are payable within thirty days under this clause 3.1."
  const consistency = (items: OpenItem[]) => items.filter(item => item.kind === "consistency").map(item => item.detail)
  // Drafted: the body's other references are linked, this one is not.
  const draft = await drafted(AGREEMENT.replace("3.1 Invoices are payable within thirty days.", `3.1 ${own}`))
  const kept = draft.row("Invoices are payable")
  const unchanged = await apply(draft.directory, [{ pid: kept.pid, existing_sha256: "0".repeat(64), updated: own }])
  expect(unchanged.rows.map(shownText)).toContain(`3.1 ${own}`)
  expect(unchanged.fields.filter(field => field.from === "3.1 ")).toEqual([])
  expect(unchanged.fields.length).toBeGreaterThan(5)
  // Written by a change, it stays typed too, and it is no inconsistency.
  const { directory, row } = await drafted()
  const written = row("Invoices are payable")
  const { rows, openItems, fields: after } = await apply(directory, [{ pid: written.pid, existing_sha256: written.sha256, updated: own }])
  expect(rows.map(shownText)).toContain(`3.1 ${own}`)
  expect(after.filter(field => field.from === "3.1 ")).toEqual([])
  expect(consistency(openItems)).toEqual(consistency(unchanged.openItems))
}, 60_000)

test("a clause split in parts stays the target of its references, its bookmark on its first part only", async () => {
  const { directory, row } = await drafted()
  const split = row("Invoices are payable")
  const { rows, fields: after } = await apply(directory, [{ pid: split.pid, existing_sha256: split.sha256, operation: "split", updated: ["Invoices are payable within thirty days.", "Invoices are sent monthly."] }])
  expect(rows.map(shownText)).toContain("Invoices are sent monthly.")
  expect(after.filter(field => field.cached === "3.1").map(field => field.target)).toEqual(["3.1 Invoices are payable within thirty days.", "3.1 Invoices are payable within thirty days."])
  const xml = strFromU8(unzipSync(readFileSync(leanFile(directory, "d1/body-applied.docx")))["word/document.xml"]!)
  const paragraph = xml.split("</w:p>").find(part => part.includes("Invoices are payable"))!
  const id = /<w:bookmarkStart w:id="(\d+)" w:name="_Ref\d+"\/>/.exec(paragraph)?.[1]
  expect(paragraph).toContain(`<w:bookmarkEnd w:id="${id}"/>`)
}, 60_000)

test("an open item a change writes is shaded as drafting shades one; the paragraphs no change touched, and the body applied again, keep their bytes", async () => {
  const { directory, row } = await drafted(AGREEMENT.replace("2.1 The Supplier shall provide the Services.", "2.1 The Supplier shall provide the Services [open item: service levels]."))
  const written = row("Invoices are payable")
  const amendment = { pid: written.pid, existing_sha256: written.sha256, updated: "Invoices are payable within thirty days [open item: confirm the payment period]." }
  await apply(directory, [amendment])
  const applied = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  expect(shadedOpenItems(applied)).toEqual(["[open item: service levels]", "[open item: confirm the payment period]"])
  const drafted_ = paragraphXml(directory, "d1/body.docx")
  const after = paragraphXml(directory, "d1/body-applied.docx")
  const unchanged = [...drafted_.keys()].filter(pid => pid && pid !== written.pid)
  expect(unchanged.length).toBeGreaterThan(5)
  for (const pid of unchanged) expect(after.get(pid)).toBe(drafted_.get(pid)!)
  // The redline marks the words the change wrote, its markers shaded as in the body.
  expect(revisions(directory)).toEqual({ changes: [["Invoices are", "ins", " "], ["Invoices are", "ins", "[open item: confirm the payment period]"]], fields: 0, formatting: 0 })
  expect(shadedOpenItems(readFileSync(leanFile(directory, "d1/redline.docx")))).toEqual(["[open item: service levels]", "[open item: confirm the payment period]"])
  // Shaded again, the body does not change; applied again, it is the same bytes.
  expect(highlightOpenItemsInDocx(applied)).toBe(applied)
  await apply(directory, [amendment])
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx")).equals(applied)).toBe(true)
}, 60_000)

test("the redline shades the whole of an open item a change rewrote word by word, and nothing around it", async () => {
  const { directory, row } = await drafted(AGREEMENT.replace("3.1 Invoices are payable within thirty days.", "3.1 Invoices are payable within thirty days [open item: confirm the payment period with the client]."))
  const written = row("Invoices are payable")
  const marker = "[open item: confirm the revised payment period with the finance team]"
  await apply(directory, [{ pid: written.pid, existing_sha256: written.sha256, updated: `Invoices are payable within thirty days ${marker}.` }])
  expect(shadedOpenItems(readFileSync(leanFile(directory, "d1/body-applied.docx")))).toEqual([marker])
  // The comparison splits the marker at the words the change rewrote; every
  // piece the redline shows is shaded, the words before and after it are not,
  // and no formatting change is tracked.
  const redline = readFileSync(leanFile(directory, "d1/redline.docx"))
  const pieces = shadedOpenItems(redline)
  expect(pieces.length).toBeGreaterThan(1)
  expect(pieces.join("")).toBe(marker)
  expect(revisions(directory).formatting).toBe(0)
}, 60_000)

test("an open item citing a bracketed heading that a change rewrites is shaded whole, to its own closing bracket, and nothing around it", async () => {
  const { directory, row } = await drafted(AGREEMENT
    .replace("2.1 The Supplier shall provide the Services.", "2.1 The Supplier shall provide the Services in accordance with this agreement and the Schedule [open item: confirm the service scope — confirm all services are subject to this clause 2 and to [Payment and invoicing] is intended].")
    .replace("3.1 Invoices are payable within thirty days.", "3.1 Invoices are payable within thirty days [open item: confirm the payment period — a dispute is then addressed under [Disputes and mediation]]. Unless disputed, an invoice is final."))
  const scope = "[open item: confirm the service scope — confirm all services are subject to this clause 2 and clause 3 is intended]"
  const period = "[open item: confirm the payment period — a dispute is then addressed under Disputes and Mediation]"
  const pending = "[open item: confirm the suspension procedure]"
  const services = row("The Supplier shall provide")
  const invoices = row("Invoices are payable")
  await apply(directory, [
    { pid: services.pid, existing_sha256: services.sha256, updated: `The Supplier shall provide the Services in accordance with this clause 2, clause 3 and the Schedule ${scope}.` },
    { pid: invoices.pid, existing_sha256: invoices.sha256, updated: `While both parties trade, invoices are payable within thirty days ${period}. Unless disputed, an invoice is final, and while a dispute subsists payment is suspended pending ${pending}.` },
  ])
  // Only the markers' text is shaded, all of it, in the body and the redline,
  // with no formatting change tracked.
  expect(shadedOpenItems(readFileSync(leanFile(directory, "d1/body-applied.docx"))).join("")).toBe(scope + period + pending)
  expect(shadedOpenItems(readFileSync(leanFile(directory, "d1/redline.docx"))).join("")).toBe(scope + period + pending)
  expect(revisions(directory).formatting).toBe(0)
}, 60_000)

test("a drafting placeholder is shaded whole in the body and the redline, one drafting split and one a change writes, and nothing around them", async () => {
  // Drafting renders the bold word in a run of its own, so no single run holds the placeholder.
  const { directory, row } = await drafted(AGREEMENT.replace("2.1 The Supplier shall provide the Services.", "2.1 The Supplier shall provide the Services by [insert the **agreed** date]."))
  const invoices = row("Invoices are payable")
  await apply(directory, [{ pid: invoices.pid, existing_sha256: invoices.sha256, updated: "Invoices are payable within [TBC – finance team] days of receipt." }])
  const applied = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  const redline = readFileSync(leanFile(directory, "d1/redline.docx"))
  expect(shadedOpenItems(applied).join("")).toBe("[insert the agreed date][TBC – finance team]")
  expect(shadedOpenItems(redline).join("")).toBe("[insert the agreed date][TBC – finance team]")
  // The shading the body adds to the drafted placeholder is no tracked formatting change.
  expect(revisions(directory).formatting).toBe(0)
  expect(shownParagraphs(redline, "accept")).toEqual(shownParagraphs(applied))
  expect(shownParagraphs(redline, "reject")).toEqual(shownParagraphs(readFileSync(leanFile(directory, "d1/body.docx"))))
}, 60_000)

// A lawyer's original numbered by a Word list, as the drafted body is but
// without its paragraph pointers: its references typed, or its own fields.
function suppliedOriginal(drafted: Uint8Array, references: "typed" | "fields") {
  const parts = unzipSync(drafted)
  const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  if (references === "typed") for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) for (const field of paragraphFields(paragraph)) unlink(field)
  for (const mark of Array.from(doc.getElementsByTagNameNS(W, "bookmarkStart"))) if (mark.getAttributeNS(W, "name")?.startsWith("_dh_")) mark.setAttributeNS(W, "w:name", `_lawyer${mark.getAttributeNS(W, "id")}`)
  return zipSync({ ...parts, "word/document.xml": strToU8(new XMLSerializer().serializeToString(doc)) })
}

// The paragraphs of a body as Word shows them, its tracked changes accepted or rejected.
function shownParagraphs(bytes: Uint8Array, changes?: "accept" | "reject") {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
  for (const node of changes ? all(doc, changes === "accept" ? "del" : "ins") : []) if (node.parentNode && (node.parentNode as Element).localName !== "rPr") node.parentNode.removeChild(node)
  return all(doc, "p").map(paragraph => all(paragraph, "*").map(node => node.localName === "t" || node.localName === "delText" ? node.textContent : "").join("").trim()).filter(Boolean)
}

test("revising a supplied original, the redline marks only the words that changed in paragraphs that cite clauses; the body keeps its references live", async () => {
  const cases = [
    { references: "typed", start: "Invoices are payable", updated: "Invoices are payable within forty days.", changes: [["Invoices are", "ins", "forty"], ["Invoices are", "del", "thirty"]] },
    { references: "typed", start: "Late payment", updated: "Late payment under clause 3.1 carries default interest, subject to clause 2.2.", changes: [["Late payment", "ins", "default "]] },
    { references: "fields", start: "Invoices are payable", updated: "Invoices are payable within forty days.", changes: [["Invoices are", "ins", "forty"], ["Invoices are", "del", "thirty"]] },
  ] satisfies { references: "typed" | "fields"; start: string; updated: string; changes: string[][] }[]
  for (const { references, start, updated, changes } of cases) {
    const directory = matter()
    const body = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes
    const original = suppliedOriginal(body, references)
    writeFileSync(path.join(directory, SOURCE), original)
    writeLeanPlan(directory, [{ title: "Services Agreement", form: "agreement", purpose: "", parts: [{ heading: "Services", scope: "" }], comparison: SOURCE }])
    writeLeanBytes(directory, "d1/body.docx", body)
    const rows = await paragraphRows(leanFile(directory, "d1/body.docx"))
    writeLeanText(directory, "d1/paragraphs.json", `${JSON.stringify(rows, null, 2)}\n`)
    const row = rows.find(row => row.text.startsWith(start))!
    const { fields: after } = await apply(directory, [{ pid: row.pid, existing_sha256: row.sha256, updated }])
    expect(after.length).toBeGreaterThan(5)
    expect([references, start, revisions(directory)]).toEqual([references, start, { changes, fields: 0, formatting: 0 }])
    // Accepted, the redline reads as the body; rejected, as the original.
    const redline = readFileSync(leanFile(directory, "d1/redline.docx"))
    expect(shownParagraphs(redline, "accept")).toEqual(shownParagraphs(readFileSync(leanFile(directory, "d1/body-applied.docx"))))
    expect(shownParagraphs(redline, "reject")).toEqual(shownParagraphs(original))
  }
}, 60_000)
