// Run from dochaus/: bun test workflows/lean/apply-crossrefs.test.ts
import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../lib/render"
import { runApply } from "./apply-stage"
import { shownText, type ParagraphRow } from "./checks"
import { writeLeanPlan } from "./plan"
import { paragraphRows } from "./stages"
import { leanFile, readLeanJson, readLeanState, writeLeanBytes, writeLeanText, type OpenItem } from "./state"
import { cleanupMatters, matter } from "./test-support"

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

async function drafted() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Services Agreement", form: "agreement", purpose: "", parts: [{ heading: "Services", scope: "" }] }])
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes)
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
// inserts or deletes, by the opening words of its paragraph as accepted, and
// how many field codes are tracked (a field the comparer could not match).
function revisions(directory: string) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(readFileSync(leanFile(directory, "d1/redline.docx")))["word/document.xml"]!), "application/xml")
  const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
  const text = (node: Element, name: string) => all(node, name).map(t => t.textContent).join("")
  const changes = all(doc, "p").flatMap(paragraph => (["ins", "del"] as const).flatMap(kind => all(paragraph, kind)
    .map(node => [(text(paragraph, "t") || text(paragraph, "delText")).slice(0, 12), kind, text(node, kind === "ins" ? "t" : "delText")])
    .filter(([, , changed]) => changed)))
  const fields = ["ins", "del"].flatMap(kind => all(doc, kind)).flatMap(node => ["fldChar", "instrText", "delInstrText"].flatMap(name => all(node, name))).length
  return { changes, fields }
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
  })
}, 60_000)

test("an amendment beside a cross-reference keeps it live; one that rewrites its number leaves that number as text", async () => {
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
    ["1.1 ", "2.1", "2.1 "], ["1.1 ", "1", "Sche"], ["1.2 ", "3", "3. P"], ["2.2 ", "3.1", "3.1 "], ["3.2 ", "3.1", "3.1 "], ["3.2 ", "2.2", "2.2 "],
  ])
  // The reference left as text is checked like any other; the one to a clause
  // the agreement does not have is reported.
  expect(openItems.filter(item => item.kind === "consistency").map(item => item.detail)).toEqual([
    "Clause 1.2 refers to clause 9, but no clause in the final text is numbered 9.",
  ])
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
