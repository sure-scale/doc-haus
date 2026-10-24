// Run from dochaus/: bun test lib/docx-crossrefs.test.ts
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { openRevision } from "./docx-amendments"
import { linkCrossReferences, referenceFields } from "./docx-crossrefs"
import { ensureParagraphRows } from "./drafting-amendment"
import { shownText } from "./list-label"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const BODY = "word/document.xml"

// A neutral agreement citing clauses, sub-clauses and a schedule, with one
// reference to a clause it does not have.
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

// An agreement: a document whose references to its own clauses are linked.
const render = async (markdown: string) => (await renderMarkdownToDocx(markdown, undefined, undefined, { form: "agreement" })).bytes

let workspace: string
const saved = process.env.WORKSPACE_ROOT
beforeAll(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "docx-crossrefs-"))
  process.env.WORKSPACE_ROOT = workspace
})
afterAll(() => {
  if (saved === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = saved
  rmSync(workspace, { recursive: true, force: true })
})

const elements = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))

// Each complex field of a paragraph: its instruction and its cached result.
function fields(paragraph: Element) {
  const found: { instr: string; result: string; resultRuns: Element[] }[] = []
  let current: (typeof found)[number] | undefined
  let mode: "instr" | "result" | undefined
  for (const node of elements(paragraph, "*")) {
    if (node.localName === "fldChar") {
      const type = node.getAttributeNS(W, "fldCharType")
      if (type === "begin") { current = { instr: "", result: "", resultRuns: [] }; mode = "instr" }
      if (type === "separate") mode = "result"
      if (type === "end" && current) { found.push(current); current = undefined; mode = undefined }
    }
    if (node.localName === "instrText" && current && mode === "instr") current.instr += node.textContent ?? ""
    if (node.localName === "t" && current && mode === "result") {
      current.result += node.textContent ?? ""
      current.resultRuns.push(node.parentNode as Element)
    }
  }
  return found
}

// Every REF field of the body: the paragraph it is in and the one its bookmark
// is on, as a reader sees them, its cached number and what its bookmark covers.
async function references(bytes: Uint8Array) {
  const opened = await openRevision(bytes)
  try {
    const body = opened.paragraphs.filter(row => row.part === BODY)
    const shown = (row: (typeof body)[number]) => shownText({ text: row.match.text, label: row.metadata.list?.generatedLabel?.trim() || undefined })
    const marks = new Map<string, { paragraph: (typeof body)[number]; start: Element }>()
    for (const row of body) for (const start of elements(row.node, "bookmarkStart")) marks.set(start.getAttributeNS(W, "name")!, { paragraph: row, start })
    return body.flatMap(row => fields(row.node).map(field => {
      const name = /^ REF (\S+) /.exec(field.instr)?.[1] ?? ""
      const mark = marks.get(name)
      return { from: shown(row), instr: field.instr.replace(/_Ref\d+/, "_Ref#"), cached: field.result, name, target: mark ? shown(mark.paragraph) : undefined, covers: mark ? covered(mark.start) : undefined, runs: field.resultRuns }
    }))
  } finally {
    opened.session.close()
  }
}

// The text between a bookmark's start and its end in the same paragraph.
function covered(start: Element) {
  const id = start.getAttributeNS(W, "id")
  let text = ""
  let inside = false
  for (const node of elements(start.parentNode as Element, "*")) {
    if (node === start) inside = true
    else if (node.localName === "bookmarkEnd" && node.getAttributeNS(W, "id") === id) break
    else if (inside && node.localName === "t") text += node.textContent ?? ""
  }
  return text
}

function bodyXml(bytes: Uint8Array) {
  return new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[BODY]!), "application/xml")
}

function withBody(bytes: Uint8Array, edit: (doc: Document) => void) {
  const parts = unzipSync(bytes)
  const doc = bodyXml(bytes)
  edit(doc)
  parts[BODY] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

const paragraphText = (paragraph: Element) => elements(paragraph, "t").map(node => node.textContent).join("")
const paragraphStarting = (doc: Document, text: string) => elements(doc, "p").find(row => paragraphText(row).startsWith(text))!

// A Word edit: a new clause at the level of `before`, inserted above it.
function insertClauseBefore(bytes: Uint8Array, before: string, text: string) {
  return withBody(bytes, doc => {
    const target = paragraphStarting(doc, before)
    const copy = target.cloneNode(true) as Element
    for (const node of [...elements(copy, "bookmarkStart"), ...elements(copy, "bookmarkEnd")]) node.parentNode!.removeChild(node)
    const runs = elements(copy, "r")
    runs.slice(1).forEach(run => run.parentNode!.removeChild(run))
    elements(runs[0]!, "t")[0]!.textContent = text
    target.parentNode!.insertBefore(copy, target)
  })
}

describe("linkCrossReferences", () => {
  test("render links each resolved reference to a bookmark on its target, its cached text the number shown", async () => {
    const bytes = await render(AGREEMENT)
    const found = await references(bytes)
    expect(found.map(({ from, instr, cached, target, covers }) => ({ from: from.slice(0, 4), instr, cached, target, ...(instr.includes("\\r") ? {} : { covers }) }))).toEqual([
      { from: "1.1 ", instr: " REF _Ref# \\r \\h ", cached: "2.1", target: "2.1 The Supplier shall provide the Services." },
      { from: "1.1 ", instr: " REF _Ref# \\h ", cached: "1", target: "Schedule 1 – Services", covers: "1" },
      { from: "1.2 ", instr: " REF _Ref# \\r \\h ", cached: "3", target: "3. Payment" },
      { from: "2.2 ", instr: " REF _Ref# \\r \\h ", cached: "3.1", target: "3.1 Invoices are payable within thirty days." },
      { from: "2.2 ", instr: " REF _Ref# \\r \\h ", cached: "3.2", target: "3.2 Late payment under clause 3.1 carries interest, subject to clause 2.2." },
      { from: "3.2 ", instr: " REF _Ref# \\r \\h ", cached: "3.1", target: "3.1 Invoices are payable within thirty days." },
      { from: "3.2 ", instr: " REF _Ref# \\r \\h ", cached: "2.2", target: "2.2 The Customer shall pay as set out in clauses 3.1 and 3.2." },
    ])
    // A reader sees the same text; the reference word stays text, and a
    // reference to a clause the document does not have is not linked.
    const opened = await openRevision(bytes)
    try {
      expect(opened.paragraphs.map(row => row.match.text)).toContain("Clause 3 governs payment; nothing in clause 9 limits it.")
    } finally {
      opened.session.close()
    }
    // The cached number keeps the formatting of the text it replaced.
    const bold = found.find(row => row.cached === "2.2")!
    expect(bold.runs.every(run => elements(run, "b").length === 1)).toBe(true)
    // One hidden _Ref bookmark per target, apart from the paragraph pointers.
    const names = elements(bodyXml(bytes), "bookmarkStart").map(node => node.getAttributeNS(W, "name")!)
    const refs = names.filter(name => name.startsWith("_Ref"))
    expect(refs.every(name => /^_Ref\d{9}$/.test(name))).toBe(true)
    expect(new Set(refs).size).toBe(refs.length)
    expect(refs).toHaveLength(new Set(found.map(row => row.name)).size)
    expect(names.filter(name => name.startsWith("_dh_")).length).toBeGreaterThan(0)
    const ids = elements(bodyXml(bytes), "bookmarkStart").map(node => node.getAttributeNS(W, "id"))
    expect(new Set(ids).size).toBe(ids.length)
  }, 60_000)

  test("quoted text, a schedule's own paragraph numbers and lettered clause numbers are not linked", async () => {
    // A neutral deed of variation: its quotations are clauses of the agreement
    // it varies, one running on over two paragraphs.
    const deed = `# Deed of Variation

## 1. Definitions

1.1 "Original Agreement" means the agreement varied as set out in clause 2.1.

1.2 Words defined in the Original Agreement have the same meaning in this Deed.

## 2. Variations

2.1 The Original Agreement is varied as set out in this clause 2.

2.2 Clause 7.1 of the Original Agreement is replaced with: "7.1 The Supplier shall deliver under clause 1.2."

2.3 Clause 9 of the Original Agreement is replaced with:

“9.1 Either party may terminate under clause 2.1.

“9.2 Notice is given under clause 1.1.”

2.4 A new clause 2A and a new clause 2.1A follow clause 2.2.

## Schedule 1 – Procedure

## 1. Notices

1.1 Notice is given under clause 1.2 of this Schedule and clause 2.1.

1.2 Notices are in writing.`
    const found = await references(await render(deed))
    expect(found.map(row => [row.from.slice(0, 14), row.cached, row.target?.slice(0, 14)])).toEqual([
      ["1.1 \"Original ", "2.1", "2.1 The Origin"],
      ["2.1 The Origin", "2", "2. Variations"],
      ["2.4 A new clau", "2.2", "2.2 Clause 7.1"],
      ["1.1 Notice is ", "2.1", "2.1 The Origin"],
    ])
  }, 60_000)

  test("linking is idempotent and a re-render of the same text names the same bookmarks", async () => {
    const first = await render(AGREEMENT)
    expect(await linkCrossReferences(first)).toEqual(first)
    const again = await render(AGREEMENT)
    expect((await references(again)).map(row => row.name)).toEqual((await references(first)).map(row => row.name))
  }, 60_000)

  test("a document whose numbers are typed, advice, a letter and a document of no known form get no fields", async () => {
    const typed = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { numbering: false, form: "agreement" })).bytes
    expect(await linkCrossReferences(typed)).toBe(typed)
    expect((await references(await render(AGREEMENT))).length).toBeGreaterThan(0)
    // Advice and letters cite the clauses of the documents they discuss; a
    // document of no known form may be either, so its references stay typed.
    for (const form of ["advice", "letter", "advice letter", undefined] as const) {
      const rendered = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form })).bytes
      expect(elements(bodyXml(rendered), "numPr").length).toBeGreaterThan(0)
      expect((await references(rendered)).map(row => `${form}: ${row.from.slice(0, 4)} ${row.cached}`)).toEqual([])
    }
    const plain = withBody(await render(AGREEMENT), doc => unlinkAll(doc))
    expect(await references(await linkCrossReferences(plain, { form: "advice" }))).toEqual([])
  }, 60_000)

  test("after a clause is inserted, linking again refreshes each cached number from its target", async () => {
    const bytes = await render(AGREEMENT)
    const before = await references(bytes)
    const inserted = insertClauseBefore(bytes, "The Supplier shall provide", "The Supplier shall appoint a manager.")
    const refreshed = await references(await linkCrossReferences(inserted))
    expect(refreshed.map(row => row.name)).toEqual(before.map(row => row.name))
    expect(refreshed.map(row => [row.cached, row.target?.slice(0, 4)])).toEqual([
      ["2.2", "2.2 "], ["1", "Sche"], ["3", "3. P"], ["3.1", "3.1 "], ["3.2", "3.2 "], ["3.1", "3.1 "], ["2.3", "2.3 "],
    ])
    expect(refreshed[0]!.from).toBe("1.1 In this Agreement the terms in clause 2.2 and Schedule 1 apply.")
  }, 60_000)

  test("a reference whose target is gone becomes plain text and is not linked to the clause now under its number", async () => {
    const bytes = await render(AGREEMENT)
    const removed = withBody(bytes, doc => {
      const target = paragraphStarting(doc, "Invoices are payable")
      target.parentNode!.removeChild(target)
    })
    const after = await references(await linkCrossReferences(removed))
    expect(after.map(row => [row.from.slice(0, 4), row.cached, row.target?.slice(0, 4)])).toEqual([
      ["1.1 ", "2.1", "2.1 "], ["1.1 ", "1", "Sche"], ["1.2 ", "3", "3. P"], ["2.2 ", "3.1", "3.1 "], ["3.1 ", "2.2", "2.2 "],
    ])
    const opened = await openRevision(await linkCrossReferences(removed))
    try {
      expect(opened.paragraphs.map(row => row.match.text)).toContain("The Customer shall pay as set out in clauses 3.1 and 3.1.")
    } finally {
      opened.session.close()
    }
  }, 60_000)

  test("the reader gives each paragraph's fields by pointer: where the cached number is, and the paragraph or text it points at", async () => {
    const file = path.join(workspace, "read.docx")
    writeFileSync(file, await render(AGREEMENT))
    const rows = await ensureParagraphRows(file)
    const byPid = new Map(rows.map(row => [row.pid, row]))
    const read = [...referenceFields(readFileSync(file))].flatMap(([pid, links]) => links.map(link => ({
      cites: byPid.get(pid)!.text.slice(link.start, link.start + link.length), number: link.number,
      to: link.number ? shownText(byPid.get(link.target!)!) : link.covers,
    })))
    expect(read).toEqual([
      { cites: "2.1", number: true, to: "2.1 The Supplier shall provide the Services." },
      { cites: "1", number: false, to: "1" },
      { cites: "3", number: true, to: "3. Payment" },
      { cites: "3.1", number: true, to: "3.1 Invoices are payable within thirty days." },
      { cites: "3.2", number: true, to: "3.2 Late payment under clause 3.1 carries interest, subject to clause 2.2." },
      { cites: "3.1", number: true, to: "3.1 Invoices are payable within thirty days." },
      { cites: "2.2", number: true, to: "2.2 The Customer shall pay as set out in clauses 3.1 and 3.2." },
    ])
  }, 60_000)

  test.skipIf(!Bun.which("soffice"))("LibreOffice shows the renumbered references after a clause and a heading are inserted", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "docx-crossrefs-lo-"))
    try {
      const bytes = await render(AGREEMENT)
      const clause = insertClauseBefore(bytes, "The Supplier shall provide", "The Supplier shall appoint a manager.")
      const heading = insertClauseBefore(clause, "Payment", "Reporting")
      writeFileSync(path.join(dir, "edited.docx"), heading)
      const converted = Bun.spawnSync(["soffice", `-env:UserInstallation=file://${dir}/profile`, "--headless", "--convert-to", "pdf", "--outdir", dir, path.join(dir, "edited.docx")])
      expect(converted.exitCode).toBe(0)
      const { extractText, getDocumentProxy } = await import("unpdf")
      const { text } = await extractText(await getDocumentProxy(new Uint8Array(readFileSync(path.join(dir, "edited.pdf")))), { mergePages: true })
      const flat = text.replace(/\s+/g, " ")
      // The fields were not refreshed: LibreOffice computes them from their bookmarks.
      expect(flat).toContain("In this Agreement the terms in clause 2.2 and Schedule 1 apply.")
      expect(flat).toContain("Clause 4 governs payment; nothing in clause 9 limits it.")
      expect(flat).toContain("The Customer shall pay as set out in clauses 4.1 and 4.2.")
      expect(flat).toContain("Late payment under clause 4.1 carries interest, subject to clause 2.3.")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  }, 180_000)
})

// Every complex field of the body, unwrapped to its cached text.
function unlinkAll(doc: Document) {
  for (const node of [...elements(doc, "fldChar"), ...elements(doc, "instrText")]) {
    const run = node.parentNode as Element
    run.parentNode?.removeChild(run)
  }
}
