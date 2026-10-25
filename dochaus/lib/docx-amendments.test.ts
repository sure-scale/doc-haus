import { beforeAll, describe, expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { createDocxRevision, createDocxLocalization, applyDocxLocalization, applyDocxParagraphEdits, DocxTrackedRevisionsError, openRevision,
  pendingRevisionAnchors, pendingRevisionParagraphs, verifyPreservation } from "./docx-amendments"
import { paragraphTargets } from "./drafting-amendment"
import { renderMarkdownToDocx } from "./render"
import { docxodus } from "./docxodus"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../test/support/counterparty-markup-fixture"
import { revisionTablesOriginalDocx } from "../test/support/revision-tables-fixture"

describe("revision-bound DOCX amendments", () => {
  let source: Uint8Array
  const amendments = [
    { id: "permission", before_quote: "Approval is required.", replacement: "Written approval is required." },
    { id: "timing", before_quote: "when feasible", replacement: "promptly" },
  ]
  beforeAll(async () => {
    source = (await renderMarkdownToDocx([
      "# Working notes", "", "## Current procedure", "",
      "Current: Approval is required. Example: Approval is required.", "",
      "Keep this **separate paragraph** unchanged.", "", "## Contact procedure", "",
      "| Action | Direction |", "| --- | --- |",
      "| Notice | Notify **when feasible** after a *confirmed event*. |",
      "| Archive | Approval is required. |", "", "## Prior example", "",
      "Approval is required.", "",
    ].join("\n"))).bytes
  })

  test("rebuilds a complete, deterministic paragraph/table projection after reopening", async () => {
    const first = await createDocxRevision(source)
    const second = await createDocxRevision(new Uint8Array(source))
    expect(second).toEqual(JSON.parse(JSON.stringify(first)))
    expect(first.markdown).toContain("Notify **when feasible** after a *confirmed event*.")
    expect(first.markdown).toContain("Table 1, row 2, column 2")
    expect(first.paragraphs.find(p => p.text.includes("when feasible"))?.table?.row_context)
      .toEqual(["Notice", "Notify when feasible after a confirmed event."])
    const packet = await createDocxLocalization(source, amendments)
    expect(packet.candidates.filter(c => c.amendment_id === "permission")).toHaveLength(4)
    expect(JSON.stringify(packet)).not.toContain("Unid")
    expect(JSON.stringify(packet)).not.toContain("spanInElement")
  })

  test("handles survive subset, reordered amendments and renamed IDs without a retained session", async () => {
    const whole = await createDocxLocalization(source, amendments)
    const reversed = await createDocxLocalization(source, [...amendments].reverse())
    const subset = await createDocxLocalization(source, [{ ...amendments[0], id: "saved-row" }])
    const handles = whole.candidates.filter(c => c.amendment_id === "permission").map(c => c.target_handle)
    expect(subset.candidates.map(c => c.target_handle)).toEqual(handles)
    expect(reversed.candidates.filter(c => c.amendment_id === "permission").map(c => c.target_handle)).toEqual(handles)
    const target = subset.candidates.find(c => c.context_before.endsWith("Current: "))!
    const output = await applyDocxLocalization(new Uint8Array(source), JSON.parse(JSON.stringify(subset)), [
      { amendment_id: "saved-row", target_handle: target.target_handle },
    ])
    expect((await createDocxRevision(output.bytes)).markdown)
      .toContain("Current: Written approval is required. Example: Approval is required.")
  })

  test("applies only selected duplicate and table spans while preserving saved structure/styles", async () => {
    const packet = await createDocxLocalization(source, amendments)
    const permission = packet.candidates.find(c => c.amendment_id === "permission" && c.context_before.endsWith("Current: "))!
    const timing = packet.candidates.find(c => c.amendment_id === "timing")!
    const output = await applyDocxLocalization(source, JSON.parse(JSON.stringify(packet)), [
      { amendment_id: "permission", target_handle: permission.target_handle },
      { amendment_id: "timing", target_handle: timing.target_handle },
    ])
    const after = await createDocxRevision(output.bytes)
    expect(after.markdown).toContain("Current: Written approval is required. Example: Approval is required.")
    expect(after.markdown).toContain("Notify **promptly** after a *confirmed event*.")
    expect(after.paragraphs.filter(p => p.text === "Approval is required.")).toHaveLength(2)
    expect(after.markdown).toContain("Keep this **separate paragraph** unchanged.")
    const originalParts = unzipSync(source)
    const editedParts = unzipSync(output.bytes)
    expect(editedParts["word/styles.xml"]).toEqual(originalParts["word/styles.xml"])
    const xml = strFromU8(editedParts["word/document.xml"])
    expect(xml.match(/<w:tc\b/g)).toHaveLength(6)
    expect(xml).not.toContain("Unid")
    expect(xml).not.toMatch(/<w:(ins|del)\b/)
    expect(output.report.package_structure_verified).toBe(true)
  })

  test("rejects stale bytes, unresolved/missing selections, replacement injection, overlap and excessive competitors", async () => {
    const packet = await createDocxLocalization(source, amendments)
    const snapshot = new Uint8Array(source)
    const target = packet.candidates[0].target_handle
    const changed = (await renderMarkdownToDocx("Unrelated changed paragraph. Approval is required.")).bytes
    await expect(applyDocxLocalization(changed, packet, [])).rejects.toThrow("stale_docx_revision")
    await expect(applyDocxLocalization(source, packet, [])).rejects.toThrow("missing_or_duplicate_amendment_selection")
    await expect(applyDocxLocalization(source, packet, [
      { amendment_id: "permission", target_handle: "missing" },
      { amendment_id: "timing", target_handle: packet.candidates.find(c => c.amendment_id === "timing")!.target_handle },
    ])).rejects.toThrow("unknown_or_conflicting_target_handle")
    await expect(applyDocxLocalization(source, packet, [
      { amendment_id: "permission", unresolved_reason: "Cannot distinguish occurrences" },
      { amendment_id: "timing", target_handle: "missing" },
    ])).rejects.toThrow("unresolved_docx_localization")
    await expect(applyDocxLocalization(source, packet, [
      { amendment_id: "permission", target_handle: target, replacement: "invented wording" },
    ])).rejects.toThrow()
    await expect(createDocxLocalization(source, [amendments[0]], { maxCandidates: 2 }))
      .rejects.toThrow("docx_candidate_limit_exceeded:4")
    await expect(createDocxLocalization(source, [{ ...amendments[0], before_quote: "absent text" }]))
      .rejects.toThrow("before_quote_not_found")
    const duplicate = await createDocxLocalization(source, [amendments[0], { ...amendments[0], id: "same-span" }])
    await expect(applyDocxLocalization(source, duplicate, [
      { amendment_id: "permission", target_handle: duplicate.candidates[0].target_handle },
      { amendment_id: "same-span", target_handle: duplicate.candidates[0].target_handle },
    ])).rejects.toThrow("overlapping_docx_amendments")
    expect(source).toEqual(snapshot)
  })

  test("rebuilds persisted handles in a new Bun process", async () => {
    const packet = await createDocxLocalization(source, amendments)
    const child = Bun.spawn([process.execPath, "-e", `
      import {createDocxLocalization} from './docx-amendments';
      console.log(JSON.stringify(await createDocxLocalization(await Bun.stdin.bytes(), ${JSON.stringify(amendments)})));
    `], { cwd: import.meta.dirname, stdin: "pipe", stdout: "pipe", stderr: "pipe" })
    child.stdin.write(source)
    child.stdin.end()
    const rebuilt = await new Response(child.stdout).json()
    const error = await new Response(child.stderr).text()
    expect(await child.exited, error).toBe(0)
    expect(rebuilt).toEqual(JSON.parse(JSON.stringify(packet)))
  })

  test("supports descending nonoverlapping spans and reports the changed run formatting across split runs", async () => {
    const edits = [
      { id: "lead", before_quote: "Current:", replacement: "Live:" },
      { id: "comparison", before_quote: "Example:", replacement: "Prior:" },
      { id: "split-run", before_quote: "Notify when feasible", replacement: "Notify promptly" },
    ]
    const packet = await createDocxLocalization(source, edits)
    const output = await applyDocxLocalization(source, packet, packet.candidates.map(candidate => ({
      amendment_id: candidate.amendment_id, target_handle: candidate.target_handle,
    })))
    const after = await createDocxRevision(output.bytes)
    expect(after.markdown).toContain("Live: Approval is required. Prior: Approval is required.")
    expect(after.markdown).toContain("Notify **promptly** after a *confirmed event*.")
    expect(output.report.amendments.find(amendment => amendment.amendment_id === "split-run")?.replacement_formatting?.bold)
      .toBe(true)
  })

  test("refuses unsupported structural and tracked-revision inputs without publishing bytes", async () => {
    const structural = await createDocxLocalization(source, [{ ...amendments[0], replacement: "New paragraph.\nSecond paragraph." }])
    await expect(applyDocxLocalization(source, structural, [
      { amendment_id: "permission", target_handle: structural.candidates[0].target_handle },
    ])).rejects.toThrow("unsupported_structural_replacement")
    const empty = await createDocxLocalization(source, [{ id: "empty", before_quote: "Approval is required.", replacement: "" }])
    const whole = empty.candidates.find(candidate => candidate.paragraph.text === "Approval is required.")!
    await expect(applyDocxLocalization(source, empty, [{ amendment_id: "empty", target_handle: whole.target_handle }]))
      .rejects.toThrow("unsupported_empty_paragraph_replacement")
    const changed = (await renderMarkdownToDocx("A changed note.")).bytes
    const tracked = await (await docxodus()).compareDocuments(source, changed, { authorName: "Reviewer" })
    await expect(createDocxRevision(tracked)).rejects.toThrow("unsupported_existing_tracked_revisions")
  })

  test("a paragraph-addressed edit writes exactly the span its handle names, run formatting included", async () => {
    const quotes = [amendments[0], { id: "split-run", before_quote: "Notify when feasible", replacement: "Notify promptly" },
      { id: "three-runs", before_quote: "feasible after a confirmed", replacement: "feasible after a verified" }]
    for (const amendment of quotes) {
      const packet = await createDocxLocalization(source, [amendment])
      for (const candidate of packet.candidates) {
        const paragraph_handle = candidate.paragraph.paragraph_handle
        const occurrence = packet.candidates.filter(row => row.paragraph.paragraph_handle === paragraph_handle).indexOf(candidate)
        const byHandle = await applyDocxLocalization(source, packet, [{ amendment_id: amendment.id, target_handle: candidate.target_handle }])
        const byParagraph = await applyDocxParagraphEdits(source, [{ ...amendment, paragraph_handle, occurrence }])
        expect(byParagraph.report).toEqual(byHandle.report)
        expect(byParagraph.bytes).toEqual(byHandle.bytes)
      }
    }
  })

  test("paragraph-addressed edits apply together with handle selections and refuse overlap, absent occurrences and other revisions", async () => {
    const snapshot = new Uint8Array(source)
    const packet = await createDocxLocalization(source, [amendments[0]])
    const current = packet.candidates.find(row => row.context_before.endsWith("Current: "))!
    const edit = { ...amendments[0], paragraph_handle: current.paragraph.paragraph_handle, occurrence: 0 }
    const timing = await createDocxLocalization(source, [amendments[1]])
    const output = await applyDocxParagraphEdits(source, [edit, { ...edit, id: "example", occurrence: 1, replacement: "Prior approval is required." }],
      { packet: timing, selections: [{ amendment_id: "timing", target_handle: timing.candidates[0]!.target_handle }] })
    const after = await createDocxRevision(output.bytes)
    expect(after.markdown).toContain("Current: Written approval is required. Example: Prior approval is required.")
    expect(after.markdown).toContain("Notify **promptly** after a *confirmed event*.")
    expect(after.paragraphs.filter(p => p.text === "Approval is required.")).toHaveLength(2)
    expect(output.report.amendments.map(row => [row.amendment_id, row.paragraph_handle])).toEqual([
      ["timing", timing.candidates[0]!.paragraph.paragraph_handle], ["permission", edit.paragraph_handle], ["example", edit.paragraph_handle]])
    expect(output.report.package_structure_verified).toBe(true)
    await expect(applyDocxParagraphEdits(source, [{ ...edit, occurrence: 2 }])).rejects.toThrow("before_quote_not_found:permission")
    await expect(applyDocxParagraphEdits(source, [{ ...edit, paragraph_handle: "paragraph-missing" }])).rejects.toThrow("unknown_paragraph_handle")
    await expect(applyDocxParagraphEdits(output.bytes, [edit])).rejects.toThrow("unknown_paragraph_handle")
    await expect(applyDocxParagraphEdits(source, [edit, { ...edit, id: "overlap", before_quote: "Current: Approval", replacement: "Live: Approval" }]))
      .rejects.toThrow("overlapping_docx_amendments")
    await expect(applyDocxParagraphEdits(source, [edit, { ...edit, occurrence: 1 }])).rejects.toThrow("duplicate_amendment_id")
    await expect(applyDocxParagraphEdits(source, [{ ...edit, before_quote: "Approval\tis required." }])).rejects.toThrow("unsupported_cross_block_quote")
    await expect(applyDocxParagraphEdits(source, [{ ...edit, replacement: "One.\nTwo." }])).rejects.toThrow("unsupported_structural_replacement")
    await expect(applyDocxParagraphEdits(source, [])).rejects.toThrow("missing_docx_amendments")
    expect(source).toEqual(snapshot)
  })
})


for (const [name, middle] of [["bold", "**then**"], ["italic", "*then*"], ["hyperlink", "[then](https://example.invalid/)"]]) {
  test(`context localization preserves unchanged ${name} runs around duplicate placeholders`, async () => {
    const source = (await renderMarkdownToDocx(`# Working note\n\nPending [Insert label] ${middle} [Insert label].\n`)).bytes
    const before = await createDocxRevision(source)
    const paragraph = before.paragraphs.find(row => row.text.includes("Insert label"))!
    expect(paragraph.text).toBe("Pending [Insert label] then [Insert label].")
    const context = (row: typeof paragraph) => {
      const characters = row.formatting.flatMap(fragment => Array.from(fragment.text).map(text => ({ text, formatting: fragment.formatting })))
      const start = Array.from(row.text.slice(0, row.text.indexOf(" then "))).length
      return characters.slice(start, start + " then ".length)
    }
    const originalContext = context(paragraph)
    expect(originalContext[1]!.formatting[name === "hyperlink" ? "hyperlinkUrl" : name as "bold" | "italic"])
      .toBe(name === "hyperlink" ? "https://example.invalid/" : true)
    const packet = await createDocxLocalization(source, [
      { id: "first", before_quote: "[Insert label] then", replacement: "First open item then" },
      { id: "second", before_quote: "[Insert label].", replacement: "Second open item." },
    ])
    expect(packet.candidates.filter(row => row.amendment_id === "first")).toHaveLength(1)
    expect(packet.candidates.filter(row => row.amendment_id === "second")).toHaveLength(1)
    const output = await applyDocxLocalization(source, packet, packet.candidates.map(candidate => ({
      amendment_id: candidate.amendment_id, target_handle: candidate.target_handle,
    })))
    const after = (await createDocxRevision(output.bytes)).paragraphs.find(row => row.text.includes("open item"))!
    expect(after.text).toBe("Pending First open item then Second open item.")
    expect(context(after)).toEqual(originalContext)
    expect(output.report.package_structure_verified).toBe(true)
  })
}

test("context localization retains run formatting for insertion, deletion, no-op and changed Unicode text", async () => {
  const source = (await renderMarkdownToDocx("# Working note\n\nLead **value** tail.\n\nLead **😀** tail.\n")).bytes
  for (const [before_quote, replacement, expected] of [
    ["Lead value", "Lead new value", "Lead **new value** tail."],
    ["value tail.", "value.", "Lead **value**."],
    ["Lead value tail.", "Lead value tail.", "Lead **value** tail."],
    ["Lead 😀 tail.", "Lead 😁 tail.", "Lead **😁** tail."],
  ]) {
    const packet = await createDocxLocalization(source, [{ id: "change", before_quote, replacement }])
    expect(packet.candidates).toHaveLength(1)
    const output = await applyDocxLocalization(source, packet, [{ amendment_id: "change", target_handle: packet.candidates[0]!.target_handle }])
    expect((await createDocxRevision(output.bytes)).markdown).toContain(expected)
  }
})

test("a whole-paragraph candidate carries its neighbouring paragraphs, clipped at the edge nearest the quote", async () => {
  const before = `Earlier wording ${"a".repeat(200)} ends here.`
  const after = `Later wording begins ${"b".repeat(200)}.`
  const source = (await renderMarkdownToDocx(`${before}\n\nWhole paragraph.\n\n${after}\n`)).bytes
  const packet = await createDocxLocalization(source, [{ id: "whole", before_quote: "Whole paragraph.", replacement: "Changed paragraph." }])
  expect(packet.candidates.map(row => [row.context_before, row.context_after, row.previous_paragraph, row.next_paragraph]))
    .toEqual([["", "", before.slice(-160), after.slice(0, 160)]])
})

// Word writes paragraphs whose only content is a break, tab or field code. grep
// projects w:t text only, so these have no match; they must be skipped, not
// mapped onto the next text paragraph, and must survive an edit elsewhere.
for (const [name, content] of [
  ["page break", '<w:r><w:br w:type="page"/></w:r>'],
  ["text-wrapping break", "<w:r><w:br/></w:r>"],
  ["column break", '<w:r><w:br w:type="column"/></w:r>'],
  ["tab", "<w:r><w:tab/></w:r>"],
  ["field code without a result", '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'],
]) {
  test(`a ${name}-only paragraph has no projection text and survives an amendment of the next paragraph`, async () => {
    const rendered = unzipSync((await renderMarkdownToDocx("First paragraph.\n\nMiddle marker.\n\nThird paragraph.\n\nTail marker.\n")).bytes)
    const markers = /<w:r><w:t xml:space="preserve">(Middle|Tail) marker\.<\/w:t><\/w:r>/g
    const xml = strFromU8(rendered["word/document.xml"]!)
    expect(xml.match(markers)).toHaveLength(2)
    const source = zipSync({ ...rendered, "word/document.xml": strToU8(xml.replace(markers, content)) })
    const revision = await createDocxRevision(source)
    expect(revision.paragraphs.map(row => row.text)).toEqual(["First paragraph.", "Third paragraph."])
    const output = await applyDocxParagraphEdits(source, [{ id: "next", paragraph_handle: revision.paragraphs[1]!.paragraph_handle,
      occurrence: 0, before_quote: "Third", replacement: "Final" }])
    expect((await createDocxRevision(output.bytes)).paragraphs.map(row => row.text)).toEqual(["First paragraph.", "Final paragraph."])
    expect(output.report.package_structure_verified).toBe(true)
    const body = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]!)
      .match(/<w:p>[\s\S]*?<\/w:p>/g)!.map(paragraph => paragraph.replace(/\s+\/>/g, "/>"))
    const before = body(source)
    const after = body(output.bytes)
    expect(after).toHaveLength(4)
    expect([after[1], after[3]]).toEqual([before[1], before[3]])
    expect(after[1]).toContain(content)
  })
}

// A rendered deliverable's page furniture, as raw OOXML: a header with a legend
// and a draft line, and a footer whose page number is a PAGE field.
function furnishedDocx() {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  const part = (name: string) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name.replace(/\d+$/, "")}+xml"/>`
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${part("header1")}${part("footer1")}</Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${O}/footer" Target="footer1.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>`
      + paragraph("Notice must be given within five days.") + paragraph("Subject to contract, the parties will meet monthly.")
      + `<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>`),
    "word/header1.xml": xml(`<w:hdr xmlns:w="${W}">${paragraph("Subject to contract")}${paragraph("Doe &amp; Partners LLP draft: 20 October 2026")}</w:hdr>`),
    "word/footer1.xml": xml(`<w:ftr xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>`
      + `<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`),
  }, { mtime: new Date(1980, 0, 1) })
}

test("the revision view labels page header and footer text apart from the body; header text stays amendable", async () => {
  const source = furnishedDocx()
  const revision = await createDocxRevision(source)
  expect(revision.paragraphs.map(row => [row.scope, row.text])).toEqual([
    ["body", "Notice must be given within five days."], ["body", "Subject to contract, the parties will meet monthly."],
    ["hdr1", "Subject to contract"], ["hdr1", "Doe & Partners LLP draft: 20 October 2026"], ["ftr1", "Page 1"]])
  expect(revision.markdown).toBe([
    "Notice must be given within five days.", "Subject to contract, the parties will meet monthly.",
    "[Page header: page furniture, not body text]", "Subject to contract", "Doe & Partners LLP draft: 20 October 2026",
    "[Page footer: page furniture, not body text]", "Page 1",
  ].join("\n\n"))
  // A body-only view carries no label, and document text cannot forge one.
  const plain = (await renderMarkdownToDocx("Notice must be given within five days.\n\n[Page header: page furniture, not body text]\n")).bytes
  expect((await createDocxRevision(plain)).markdown)
    .toBe("Notice must be given within five days.\n\n\\[Page header: page furniture, not body text\\]")
  // The header remains a target when a change is about it; the footer's page-number paragraph is refused.
  const legend = await createDocxLocalization(source, [{ id: "legend", before_quote: "Subject to contract", replacement: "Subject to board approval" }])
  const header = legend.candidates.find(candidate => candidate.paragraph.scope === "hdr1")!
  // A candidate's neighbouring paragraphs stay inside its own part.
  expect(legend.candidates.map(row => [row.paragraph.scope, row.previous_paragraph, row.next_paragraph])).toEqual([
    ["body", "Notice must be given within five days.", null], ["hdr1", null, "Doe & Partners LLP draft: 20 October 2026"]])
  const output = await applyDocxLocalization(source, legend, [{ amendment_id: "legend", target_handle: header.target_handle }])
  expect((await createDocxRevision(output.bytes)).paragraphs.map(row => row.text)).toEqual(["Notice must be given within five days.",
    "Subject to contract, the parties will meet monthly.", "Subject to board approval", "Doe & Partners LLP draft: 20 October 2026", "Page 1"])
  const page = await createDocxLocalization(source, [{ id: "page", before_quote: "Page", replacement: "Sheet" }])
  await expect(applyDocxLocalization(source, page, [{ amendment_id: "page", target_handle: page.candidates[0]!.target_handle }]))
    .rejects.toThrow("unsupported_field_replacement")
})

test("the revision view shows each Word list number before its heading or paragraph, apart from the text quotes match", async () => {
  // render numbers the clauses with a Word list, so "2." and "2.1" are list labels, not run text.
  const bytes = (await renderMarkdownToDocx(["# Supply Agreement", "## 1. Delivery", "1.1 The Supplier shall deliver the cartons.",
    "## 2. Payment", "2.1 The Customer shall pay within thirty days."].join("\n\n"))).bytes
  const revision = await createDocxRevision(bytes)
  expect(revision.markdown).toContain("## 2. Payment\n\n2.1 The Customer shall pay within thirty days.")
  const last = revision.paragraphs.at(-1)!
  expect([last.label, last.text, last.headings.map(heading => heading.text)])
    .toEqual(["2.1", "The Customer shall pay within thirty days.", ["Supply Agreement", "2. Payment"]])
})

describe("live cross-references", () => {
  // A numbered agreement whose clause references render as REF fields.
  const AGREEMENT = ["# Services Agreement", "## 1. Definitions", "1.1 In this Agreement the terms in clause 2.1 and Schedule 1 apply.",
    "## 2. Services", "2.1 The Supplier shall provide the Services.", "2.2 The Customer shall pay as set out in clause 2.1.",
    "## Schedule 1 – Services", "The services listed here."].join("\n\n")
  const refs = (bytes: Uint8Array) => [...strFromU8(unzipSync(bytes)["word/document.xml"]!).matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>[\s\S]*?<w:t[^>]*>([^<]*)<\/w:t>/g)]
    .map(match => [match[1]!.replace(/_Ref\d+/, "_Ref#"), match[2]])

  test("an edit beside a REF field keeps it live; an edit of its number leaves that number as text, and the report says so", async () => {
    const source = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes
    expect(refs(source)).toEqual([[" REF _Ref# \\r \\h ", "2.1"], [" REF _Ref# \\h ", "1"], [" REF _Ref# \\r \\h ", "2.1"]])
    const revision = await createDocxRevision(source)
    const handle = (start: string) => revision.paragraphs.find(row => row.text.startsWith(start))!.paragraph_handle
    const beside = { id: "beside", paragraph_handle: handle("In this Agreement"), occurrence: 0, before_quote: "the terms", replacement: "the defined terms" }
    const output = await applyDocxParagraphEdits(source, [
      beside,
      { id: "number", paragraph_handle: handle("The Customer"), occurrence: 0, before_quote: "clause 2.1", replacement: "clause 2.2" },
    ])
    expect((await createDocxRevision(output.bytes)).paragraphs.map(row => row.text)).toContain("In this Agreement the defined terms in clause 2.1 and Schedule 1 apply.")
    expect((await createDocxRevision(output.bytes)).paragraphs.map(row => row.text)).toContain("The Customer shall pay as set out in clause 2.2.")
    expect(refs(output.bytes)).toEqual([[" REF _Ref# \\r \\h ", "2.1"], [" REF _Ref# \\h ", "1"]])
    expect(output.report.limitation).toEndWith("; 1 cross-reference field converted to text in paragraph 2.2")
    expect((await applyDocxParagraphEdits(source, [beside])).report.limitation).not.toContain("cross-reference")
  })

  test("a changed format of an existing run is refused, also in a batch that rewrites a cross-reference number", async () => {
    // Clause 2.3 sits in clause 2: numbers after the schedule heading stay typed.
    const source = (await renderMarkdownToDocx(AGREEMENT.replace("\n\n## Schedule 1", "\n\n2.3 Nothing here limits the **Supplier**.\n\n## Schedule 1"),
      undefined, undefined, { form: "agreement" })).bytes
    const revision = await createDocxRevision(source)
    const handle = (start: string) => revision.paragraphs.find(row => row.text.startsWith(start))!.paragraph_handle
    const plain = { id: "plain", paragraph_handle: handle("Nothing here"), occurrence: 0, before_quote: "Nothing here", replacement: "Nothing in this clause" }
    const number = { id: "number", paragraph_handle: handle("The Customer"), occurrence: 0, before_quote: "clause 2.1", replacement: "clause 2.2" }
    // A faulty engine edit: the untouched bold run "Supplier" gains capitals.
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    const dx = await docxodus()
    const probe = dx.openDocxSession(source, {})
    const prototype = Object.getPrototypeOf(probe)
    probe.close()
    const replaceMatch = prototype.replaceMatch
    prototype.replaceMatch = function (this: typeof probe, ...args: Parameters<typeof probe.replaceMatch>) {
      const result = replaceMatch.apply(this, args)
      const id = args[0].enclosingAnchor.id
      if (!this.raw.getXml(id).includes("Nothing in this clause")) return result
      const doc = new DOMParser().parseFromString(this.raw.getXml(id), "application/xml")
      const run = Array.from(doc.getElementsByTagNameNS(W, "r")).find(row => row.getElementsByTagNameNS(W, "t")[0]?.textContent === "Supplier")!
      run.getElementsByTagNameNS(W, "rPr")[0]!.appendChild(doc.createElementNS(W, "w:caps"))
      this.raw.replaceXml(id, new XMLSerializer().serializeToString(doc))
      return result
    }
    try {
      await expect(applyDocxParagraphEdits(source, [plain])).rejects.toThrow("docx_existing_run_properties_changed")
      await expect(applyDocxParagraphEdits(source, [plain, number])).rejects.toThrow("docx_existing_run_properties_changed")
    } finally { prototype.replaceMatch = replaceMatch }
    // Without the fault both batches apply.
    expect((await createDocxRevision((await applyDocxParagraphEdits(source, [plain, number])).bytes)).paragraphs.map(row => row.text))
      .toEqual(expect.arrayContaining(["Nothing in this clause limits the Supplier.", "The Customer shall pay as set out in clause 2.2."]))
  })
})

// Their markup stays pending while doc.haus edits other paragraphs: the layer
// reader projects the accepted text, and a paragraph with any pending revision
// is not edited.
describe("layered edits on a counterparty markup", () => {
  const source = counterpartyMarkupDocx()
  const [AVERY, MORGAN] = COUNTERPARTY_MARKUP.reviewers as [string, string]
  const layer = { pendingRevisions: "layer" } as const
  const revisions = async (bytes: Uint8Array) => {
    const session = (await docxodus()).openDocxSession(bytes, {})
    try {
      return session.listRevisions().map(row => ({ id: row.id, type: row.type, author: row.author, text: row.text, date: row.date, anchor: row.anchorId }))
    } finally { session.close() }
  }
  const handle = (revision: { paragraphs: { paragraph_handle: string; text: string }[] }, prefix: string) =>
    revision.paragraphs.find(row => row.text.startsWith(prefix))!.paragraph_handle
  const clause9 = { id: "law", occurrence: 0, before_quote: "the Province", replacement: "the State" }

  test("layer mode maps every paragraph, and names each paragraph with a pending change and its authors", async () => {
    const revision = await createDocxRevision(source, layer)
    expect(revision.paragraphs.map(row => [row.scope, row.text])).toEqual([
      ["body", "1. Supply. The Supplier shall deliver the Goods within forty-five days of each Order."],
      ["body", "2. Price. The Price is fixed for the first twelve months."],
      ["body", "4. Liability. Neither party is liable for indirect loss."],
      ["body", "3. Notices. Notices must be in writing."],
      ["body", "5. Term. This Agreement continues for two years."],
      ["body", "6. Audit. The Buyer may audit on ten Business Days' notice the Supplier's records."],
      ["body", "7. Assignment. Neither party may assign"],
      ["body", " without consent."],
      ["body", "8. Waiver. No waiver is effective unless in writing."],
      ["body", "9. Governing law. This Agreement is governed by the law of the Province."],
      ...COUNTERPARTY_MARKUP.comments.map(comment => ["cmt", comment.text]),
    ])
    expect((await paragraphTargets(source, layer)).map(row => row.paragraph_handle))
      .toEqual(revision.paragraphs.filter(row => row.scope === "body").map(row => row.paragraph_handle))
    // Both ends of the move are pending; the paragraph after a deleted paragraph mark is not.
    const pending = [["1", [AVERY]], ["2", [MORGAN]], ["4", [MORGAN]], ["3", [AVERY]], ["5", [AVERY]], ["6", [AVERY, MORGAN]], ["7", [MORGAN]], ["8", [MORGAN]]] as const
    const byText = (rows: Map<string, string[]>, text: (key: string) => string) => [...rows].map(([key, authors]) => [text(key).match(/^\d+/)?.[0] ?? text(key), authors])
    expect(byText(await pendingRevisionParagraphs(source), key => revision.paragraphs.find(row => row.paragraph_handle === key)!.text))
      .toEqual(pending.map(([text, authors]) => [text, [...authors]]))
    // The legacy tools' own session reaches the same paragraphs by anchor id.
    const session = (await docxodus()).openDocxSession(source, {})
    try {
      const texts = new Map(session.grep("^[\\s\\S]*$", { scope: 63, whitespace: 0 }).map(match => [match.enclosingAnchor.id, match.text]))
      expect(byText(pendingRevisionAnchors(session), key => texts.get(key)!)).toEqual(pending.map(([text, authors]) => [text, [...authors]]))
    } finally { session.close() }
  })

  test("an accept-mode edit to clause 9 is written untracked and keeps their ten revisions", async () => {
    const revision = await createDocxRevision(source, layer)
    const output = await applyDocxParagraphEdits(source, [{ ...clause9, paragraph_handle: handle(revision, "9.") }], undefined, layer)
    expect((await createDocxRevision(output.bytes, layer)).paragraphs[9]!.text)
      .toBe("9. Governing law. This Agreement is governed by the law of the State.")
    const before = await revisions(source)
    expect(before).toHaveLength(10)
    expect(await revisions(output.bytes)).toEqual(before)
    expect(output.report.package_structure_verified).toBe(true)
  })

  test("an edit inside or beside their pending change is refused, and nothing is written", async () => {
    const revision = await createDocxRevision(source, layer)
    const law = { ...clause9, paragraph_handle: handle(revision, "9.") }
    // Inside Avery's insertion an untracked edit would become her proposal.
    const inside = { id: "lead-time", paragraph_handle: handle(revision, "1."), occurrence: 0, before_quote: "forty-five", replacement: "sixty" }
    const beside = { id: "price-period", paragraph_handle: handle(revision, "2."), occurrence: 0, before_quote: "twelve months", replacement: "eighteen months" }
    await expect(applyDocxParagraphEdits(source, [inside], undefined, layer)).rejects.toThrow("docx_pending_revision:lead-time")
    await expect(applyDocxParagraphEdits(source, [law, beside], undefined, layer)).rejects.toThrow("docx_pending_revision:price-period")
    const packet = await createDocxLocalization(source, [{ id: "lead-time", before_quote: "forty-five", replacement: "sixty" }], {}, layer)
    const body = packet.candidates.find(row => row.paragraph.scope === "body")!
    await expect(applyDocxLocalization(source, packet, [{ amendment_id: "lead-time", target_handle: body.target_handle }], layer))
      .rejects.toThrow("docx_pending_revision:lead-time")
  })

  test("a tracked edit adds exactly two revisions under the author beside theirs and passes verification", async () => {
    const author = "Jordan Example"
    const tracked = { ...layer, revisionAuthor: author }
    // Prepared on the accepted view; applied as native tracked changes.
    expect(await createDocxRevision(source, tracked)).toEqual(await createDocxRevision(source, layer))
    const packet = await createDocxLocalization(source, [{ id: "law", before_quote: "the Province", replacement: "the State" }], {}, layer)
    const output = await applyDocxLocalization(source, packet, [{ amendment_id: "law", target_handle: packet.candidates[0]!.target_handle }], tracked)
    expect(output.report.package_structure_verified).toBe(true)
    const theirs = await revisions(source)
    const after = await revisions(output.bytes)
    expect(after.filter(row => row.author !== author)).toEqual(theirs)
    const clause = after.find(row => row.author === author)!.anchor
    // Whole words, as Word records an overtyped word, not the changed letters.
    expect(after.filter(row => row.author === author).map(row => [row.type, row.text, row.anchor])).toEqual([["delete", "Province", clause], ["insert", "State", clause]])
    expect((await createDocxRevision(output.bytes, layer)).paragraphs[9]!.text).toBe("9. Governing law. This Agreement is governed by the law of the State.")
    expect(strFromU8(unzipSync(output.bytes)["word/settings.xml"]!)).toContain("<w:trackRevisions")
    // Only Track Changes being switched on is tolerated in the settings part.
    const zoomed = unzipSync(output.bytes)
    zoomed["word/settings.xml"] = strToU8(strFromU8(zoomed["word/settings.xml"]!).replace("<w:trackRevisions", '<w:zoom w:percent="90"/><w:trackRevisions'))
    const [a, b, c] = [await openRevision(source, tracked), await openRevision(output.bytes, tracked), await openRevision(zipSync(zoomed), tracked)]
    try {
      expect(() => verifyPreservation(a, b, new Set([9]))).not.toThrow()
      expect(() => verifyPreservation(a, c, new Set([9]))).toThrow("docx_package_parts_changed")
    } finally { [a, b, c].forEach(opened => opened.session.close()) }
  })

  test("a paragraph in a table row they inserted is pending, though Docxodus lists only the row", async () => {
    const parts = unzipSync(revisionTablesOriginalDocx())
    const xml = strFromU8(parts["word/document.xml"]!)
    const row = xml.lastIndexOf("<w:tr>", xml.indexOf("Unit price")) + "<w:tr>".length
    parts["word/document.xml"] = strToU8(`${xml.slice(0, row)}<w:trPr><w:ins w:id="901" w:author="${AVERY}" w:date="2026-10-20T09:00:00Z"/></w:trPr>${xml.slice(row)}`)
    const inserted = zipSync(parts)
    const revision = await createDocxRevision(inserted, layer)
    expect([...(await pendingRevisionParagraphs(inserted))].map(([key, authors]) => [revision.paragraphs.find(row => row.paragraph_handle === key)!.text, authors]))
      .toEqual([["Unit price", [AVERY]], ["0.14 per carton", [AVERY]]])
    const session = (await docxodus()).openDocxSession(inserted, {})
    try { expect([...pendingRevisionAnchors(session).values()]).toEqual([[AVERY], [AVERY]]) } finally { session.close() }
    await expect(applyDocxParagraphEdits(inserted, [{ id: "price", paragraph_handle: handle(revision, "0.14"), occurrence: 0, before_quote: "0.14", replacement: "0.15" }],
      undefined, layer)).rejects.toThrow("docx_pending_revision:price")
  })

  test("refuse mode, the default, still refuses their markup", async () => {
    for (const options of [undefined, { pendingRevisions: "refuse" } as const]) {
      await expect(createDocxRevision(source, options)).rejects.toBeInstanceOf(DocxTrackedRevisionsError)
      await expect(createDocxLocalization(source, [{ id: "law", before_quote: "the Province", replacement: "the State" }], {}, options))
        .rejects.toBeInstanceOf(DocxTrackedRevisionsError)
      await expect(paragraphTargets(source, options)).rejects.toBeInstanceOf(DocxTrackedRevisionsError)
    }
  })
})
