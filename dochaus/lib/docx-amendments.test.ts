import { beforeAll, describe, expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { createDocxRevision, createDocxLocalization, applyDocxLocalization, applyDocxParagraphEdits, DocxTrackedRevisionsError, openRevision,
  pendingRevisionAnchors, pendingRevisionParagraphs, verifyPreservation } from "./docx-amendments"
import { paragraphTargets } from "./drafting-amendment"
import { renderMarkdownToDocx } from "./render"
import { wordChanges } from "./word-changes"
import { docxodus } from "./docxodus"
import { readDocxComments } from "./docx-comments"
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
  const source = (await renderMarkdownToDocx("# Working note\n\nLead **value** tail.\n\nLead **😀** tail.\n\nThe Supplier must notify the **Buyer** in writing.\n")).bytes
  for (const [before_quote, replacement, expected] of [
    // An inserted word is typed after the text before it, not into the bold word after it.
    ["Lead value", "Lead new value", "Lead new **value** tail."],
    ["value tail.", "value.", "Lead **value**."],
    ["Lead value tail.", "Lead value tail.", "Lead **value** tail."],
    ["Lead 😀 tail.", "Lead 😁 tail.", "Lead **😁** tail."],
    // After an earlier change, an inserted word is written with the text
    // before it, as the whole change is, and the bold word after it is left alone.
    ["Supplier must notify the Buyer", "Vendor must notify the new Buyer", "The Vendor must notify the new **Buyer** in writing."],
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

// A paragraph with its tabs and breaks as \t and \n, as the reader of the saved file sees it.
function laidOut(bytes: Uint8Array, start: string) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const shown = (paragraph: Element) => Array.from(paragraph.getElementsByTagNameNS(W, "*")).map(node => {
    for (let at = node.parentNode; at && at !== paragraph; at = at.parentNode) if ((at as Element).localName === "del") return ""
    return node.localName === "t" ? node.textContent : node.localName === "tab" ? "\t" : node.localName === "br" ? "\n" : ""
  }).join("")
  return Array.from(doc.getElementsByTagNameNS(W, "p")).map(shown).find(text => text.startsWith(start))
}

// A document whose paragraphs after the first are these runs.
async function runsDocx(paragraphs: string[]) {
  const markers = paragraphs.map((_, index) => `Marker ${index + 1}.`)
  const rendered = unzipSync((await renderMarkdownToDocx(`First paragraph.\n\n${markers.join("\n\n")}\n`)).bytes)
  const xml = paragraphs.reduce((text, runs, index) => {
    const placed = text.replace(`<w:r><w:t xml:space="preserve">${markers[index]}</w:t></w:r>`, runs)
    expect(placed).not.toBe(text)
    return placed
  }, strFromU8(rendered["word/document.xml"]!))
  return zipSync({ ...rendered, "word/document.xml": strToU8(xml) })
}

// Docxodus writes an edited run as its tabs and breaks, then all its text.
test("a tab or break stays where it was in a run the edit rewrites, and an insertion in that run is written, tracked or not", async () => {
  const source = await runsDocx(["<w:r><w:t>Acme Ltd</w:t><w:br/><w:t>1 Main Street, London</w:t></w:r>",
    '<w:r><w:t>(a)</w:t><w:tab/><w:t xml:space="preserve">the Supplier shall pay the Fee;</w:t></w:r>'])
  const revision = await createDocxRevision(source)
  for (const [start, before_quote, replacement, after] of [
    ["Acme", "London", "Leeds", "Acme Ltd\n1 Main Street, Leeds"],
    ["Acme", "Main Street", "Main High Street", "Acme Ltd\n1 Main High Street, London"],
    ["(a)", "shall pay", "shall promptly pay", "(a)\tthe Supplier shall promptly pay the Fee;"],
    ["(a)", "(a)", "(b)", "(b)\tthe Supplier shall pay the Fee;"],
  ] as const) {
    const paragraph_handle = revision.paragraphs.find(row => row.text.startsWith(start))!.paragraph_handle
    for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, {}] as const) {
      const output = await applyDocxParagraphEdits(source, [{ id: "edit", paragraph_handle, occurrence: 0, before_quote, replacement }], undefined, options)
      expect(laidOut(output.bytes, after.slice(0, 2))).toBe(after)
    }
  }
})

// The text reads through a tab or break: "1.1<tab>Definitions" is "1.1Definitions".
test("a change of a word read across a tab or break leaves the tab or break between the same characters, tracked or not", async () => {
  const bold = (text: string) => `<w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r>`
  const source = await runsDocx([`${bold("1.1")}<w:r><w:tab/></w:r>${bold("Definitions")}`,
    "<w:r><w:t>Acme Ltd</w:t><w:br/><w:t>1 Main Street</w:t></w:r>"])
  const revision = await createDocxRevision(source)
  for (const [start, before_quote, replacement, after] of [
    ["1.1", "1.1Definitions", "1.2Definitions", "1.2\tDefinitions"],
    ["1.1", "1.1Definitions", "1.1Definitions and Interpretation", "1.1\tDefinitions and Interpretation"],
    ["Acme", "Acme Ltd1 Main Street", "Acme Limited1 Main Street", "Acme Limited\n1 Main Street"],
  ] as const) {
    const paragraph_handle = revision.paragraphs.find(row => row.text.startsWith(start))!.paragraph_handle
    for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, {}] as const) {
      const output = await applyDocxParagraphEdits(source, [{ id: "edit", paragraph_handle, occurrence: 0, before_quote, replacement }], undefined, options)
      expect(laidOut(output.bytes, after.slice(0, 2))).toBe(after)
    }
  }
})

// Words that lead into the text after a tab or break are typed after it, in
// that text's formatting; words that trail the text before it stay before it.
test("an insertion at a tab or break lands on the side of the words it joins, tracked or not", async () => {
  const bold = (text: string) => `<w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r>`
  const source = await runsDocx(['<w:r><w:t>(a)</w:t><w:tab/><w:t xml:space="preserve">the Supplier shall pay the Fee;</w:t></w:r>',
    `${bold("1.1")}<w:r><w:tab/></w:r>${bold("Definitions")}`, "<w:r><w:t>Acme Ltd</w:t><w:br/><w:t>London</w:t></w:r>"])
  const revision = await createDocxRevision(source)
  for (const [start, before_quote, replacement, after] of [
    ["(a)", "(a)the Supplier", "(a)subject to clause 3, the Supplier", "(a)\tsubject to clause 3, the Supplier shall pay the Fee;"],
    ["1.1", "1.1Definitions", "1.1General Definitions", "1.1\tGeneral Definitions"],
    ["Acme", "Acme LtdLondon", "Acme LtdCentral London", "Acme Ltd\nCentral London"],
    ["Acme", "Acme LtdLondon", "Acme Ltd (UK)London", "Acme Ltd (UK)\nLondon"],
  ] as const) {
    const paragraph_handle = revision.paragraphs.find(row => row.text.startsWith(start))!.paragraph_handle
    for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, {}] as const) {
      const output = await applyDocxParagraphEdits(source, [{ id: "edit", paragraph_handle, occurrence: 0, before_quote, replacement }], undefined, options)
      expect(laidOut(output.bytes, after.slice(0, 2))).toBe(after)
      if (start !== "1.1") continue
      const shown = Array.from(new DOMParser().parseFromString(strFromU8(unzipSync(output.bytes)["word/document.xml"]!), "application/xml")
        .getElementsByTagName("w:t")).filter(node => node.textContent?.includes("General") && (node.parentNode!.parentNode as Element).localName !== "del")
      expect(shown.map(node => (node.parentNode as Element).getElementsByTagName("w:b").length > 0)).toEqual([true])
    }
  }
})

// LibreOffice saves each paragraph mark's run properties, an empty w:rPr too.
test("an edit keeps a paragraph mark's empty run properties, as LibreOffice saves them, tracked or not", async () => {
  const parts = unzipSync(await runsDocx(["<w:r><w:rPr/><w:t>Notices must be given in writing.</w:t></w:r>"]))
  const xml = strFromU8(parts["word/document.xml"]!)
  const marked = xml.replace('<w:p><w:bookmarkStart w:id="2"', '<w:p><w:pPr><w:rPr/></w:pPr><w:bookmarkStart w:id="2"')
  expect(marked).not.toBe(xml)
  const source = zipSync({ ...parts, "word/document.xml": strToU8(marked) })
  const paragraph_handle = (await createDocxRevision(source)).paragraphs.find(row => row.text.startsWith("Notices"))!.paragraph_handle
  for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, {}] as const) {
    const output = await applyDocxParagraphEdits(source, [{ id: "edit", paragraph_handle, occurrence: 0, before_quote: "in writing", replacement: "in writing or by email" }], undefined, options)
    expect(laidOut(output.bytes, "No")).toBe("Notices must be given in writing or by email.")
  }
})

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

  test("a sentence inserted on each side of a REF field, as Prepare joins two insertions, keeps it live, tracked or not", async () => {
    const source = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes
    const revision = await createDocxRevision(source, { pendingRevisions: "layer" })
    const quote = "The Customer shall pay as set out in clause 2.1."
    const joined = { id: "joined", paragraph_handle: revision.paragraphs.find(row => row.text === quote)!.paragraph_handle, occurrence: 0,
      before_quote: quote, replacement: `Interest accrues. ${quote} Late fees apply.` }
    for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, { pendingRevisions: "layer" }, {}] as const) {
      const output = await applyDocxParagraphEdits(source, [joined], undefined, options)
      expect(refs(output.bytes)).toEqual(refs(source))
      expect(output.report.limitation).not.toContain("cross-reference")
      expect(output.report.amendments[0]!.paragraph_after).toBe(joined.replacement)
    }
  })

  // Each changed word is its own change, so the words between them, a live
  // cross-reference included, are never deleted and typed again, and words
  // inserted just before one, after an earlier change, leave it alone.
  test("word changes around REF fields keep them live, tracked or not", async () => {
    const source = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes
    const revision = await createDocxRevision(source, { pendingRevisions: "layer" })
    for (const [start, before_quote, replacement, after] of [
      ["In this Agreement", "the terms in clause 2.1 and Schedule 1 apply.", "the words in clause 2.1 and Schedule 1 govern.",
        "In this Agreement the words in clause 2.1 and Schedule 1 govern."],
      ["The Customer", "The Customer shall pay as set out in clause 2.1.", "The Buyer shall pay as set out in clause 1.4 and 2.1.",
        "The Buyer shall pay as set out in clause 1.4 and 2.1."],
    ] as const) {
      const edit = { id: "words", paragraph_handle: revision.paragraphs.find(row => row.text.startsWith(start))!.paragraph_handle,
        occurrence: 0, before_quote, replacement }
      for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, { pendingRevisions: "layer" }, {}] as const) {
        const output = await applyDocxParagraphEdits(source, [edit], undefined, options)
        expect(refs(output.bytes)).toEqual(refs(source))
        expect(output.report.limitation).not.toContain("cross-reference")
        expect(output.report.amendments[0]!.paragraph_after).toBe(after)
      }
    }
  })

  // Inserted words are typed where they stand: they take the formatting of the
  // text before them, never that of a bold or italic term after them, and a
  // live cross-reference beside them stays live. Characters added at a word's
  // end take that word's formatting.
  test("inserted text takes the formatting before it and leaves a REF field beside it live, tracked or not", async () => {
    // The text with its bold and italic words marked, however Word splits them into runs.
    const marked = (row: Awaited<ReturnType<typeof createDocxRevision>>["paragraphs"][number]) => {
      let text = "", open = ""
      for (const fragment of row.formatting) {
        const mark = fragment.formatting.bold ? "**" : fragment.formatting.italic ? "*" : ""
        if (mark !== open) [text, open] = [text + open + mark, mark]
        text += fragment.text
      }
      return text + open
    }
    const source = (await renderMarkdownToDocx(AGREEMENT.replace("\n\n## Schedule 1",
      "\n\n2.3 The Supplier must notify the **Buyer** in writing under clause 2.1, and pay *the Fee* monthly.\n\n## Schedule 1"),
    undefined, undefined, { form: "agreement" })).bytes
    const revision = await createDocxRevision(source, { pendingRevisions: "layer" })
    const paragraph_handle = revision.paragraphs.find(row => row.text.startsWith("The Supplier must"))!.paragraph_handle
    for (const [before_quote, replacement, after] of [
      ["notify the Buyer", "notify the or Buyer", "The Supplier must notify the or **Buyer** in writing under clause 2.1, and pay *the Fee* monthly."],
      ["Supplier must notify the Buyer in writing under clause 2.1, and pay the Fee",
        "Vendor must notify the or Buyer in writing under clause 2.1, and pay Part A of the Fee",
        "The Vendor must notify the or **Buyer** in writing under clause 2.1, and pay Part A of *the Fee* monthly."],
      ["pay the Fee monthly", "pay the Fees monthly", "The Supplier must notify the **Buyer** in writing under clause 2.1, and pay *the Fees* monthly."],
      ["notify the Buyer in writing under clause 2.1, and pay the Fee monthly", "notify the Buyers in writing under clause 2.1, and pay the Fees monthly",
        "The Supplier must notify the **Buyers** in writing under clause 2.1, and pay *the Fees* monthly."],
      // Inserted just after the cross-reference, as the last change.
      ["Supplier must notify the Buyer in writing under clause 2.1,", "Vendor must notify the Buyer in writing under clause 2.1 in full,",
        "The Vendor must notify the **Buyer** in writing under clause 2.1 in full, and pay *the Fee* monthly."],
      // Two insertions around kept words, as Prepare joins them, the second just before the cross-reference.
      ["The Supplier must notify the Buyer in writing under clause 2.1,",
        "Notice is required. The Supplier must notify the Buyer in writing under clause 1.4 and 2.1,",
        "Notice is required. The Supplier must notify the **Buyer** in writing under clause 1.4 and 2.1, and pay *the Fee* monthly."],
    ] as const) {
      for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, {}] as const) {
        const output = await applyDocxParagraphEdits(source, [{ id: "words", paragraph_handle, occurrence: 0, before_quote, replacement }], undefined, options)
        expect(refs(output.bytes)).toEqual(refs(source))
        expect(output.report.limitation).not.toContain("cross-reference")
        expect(marked((await createDocxRevision(output.bytes, { pendingRevisions: "layer" })).paragraphs.find(row => row.label === "2.3")!)).toBe(after)
      }
    }
  })

  test("a cross-reference added beside a live one leaves it live, and the change is the whole numbers added, tracked or not", async () => {
    const source = (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes
    const paragraph_handle = (await createDocxRevision(source)).paragraphs.find(row => row.label === "2.2")!.paragraph_handle
    const [before_quote, replacement] = ["set out in clause 2.1.", "set out in clauses 3.2 and 2.1, as amended."]
    expect(wordChanges(before_quote, replacement)).toEqual([
      { offset: 11, replaced: "clause ", changed: "clauses 3.2 and " }, { offset: 21, replaced: "", changed: ", as amended" }])
    for (const options of [{ pendingRevisions: "layer", revisionAuthor: "Jordan Example" }, {}] as const) {
      const output = await applyDocxParagraphEdits(source, [{ id: "reference", paragraph_handle, occurrence: 0, before_quote, replacement }], undefined, options)
      expect(refs(output.bytes)).toEqual(refs(source))
      expect(output.report.limitation).not.toContain("cross-reference")
      expect((await createDocxRevision(output.bytes, { pendingRevisions: "layer" })).paragraphs.find(row => row.label === "2.2")!.text)
        .toBe("The Customer shall pay as set out in clauses 3.2 and 2.1, as amended.")
    }
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
      if (!this.raw.getXml(id).includes("in this clause")) return result
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

  // Avery's move puts clause 3 below clause 4, its typed number unchanged, so
  // the clauses read 2, 4, 3. Prepare's passes are told why.
  test("the paragraph their pending move puts below clause 4 is marked in the draft view and as a candidate's neighbour", async () => {
    const revision = await createDocxRevision(source, layer)
    expect(revision.paragraphs.filter(row => row.moved_here).map(row => row.text)).toEqual(["3. Notices. Notices must be in writing."])
    expect(revision.markdown).toContain("2. Price. The Price is fixed for the first twelve months.\n\n4. Liability. Neither party is liable for **indirect** loss.\n\n" +
      "[Moved here by a pending tracked change]\n\n3. Notices. Notices must be in writing.\n\n5. Term.")
    const packet = await createDocxLocalization(source, [{ id: "mitigation", before_quote: "Liability. Neither", replacement: "Liability. Each party shall mitigate its loss. Neither" }], {}, layer)
    expect(packet.candidates.map(row => [row.previous_paragraph, row.next_paragraph])).toEqual([
      ["2. Price. The Price is fixed for the first twelve months.", "[Moved here by a pending tracked change] 3. Notices. Notices must be in writing."]])
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

  // One requested change is one tracked change of whole words: an inserted
  // sentence is not written as the next word deleted and inserted again.
  test("a tracked insertion, deletion or one-word change is written as that change alone, beside theirs", async () => {
    const author = "Jordan Example"
    const tracked = { ...layer, revisionAuthor: author }
    const theirs = await revisions(source)
    for (const [before_quote, replacement, expected, text] of [
      ["This Agreement", "Each party submits to the courts of the Province. This Agreement", [["insert", "Each party submits to the courts of the Province. "]],
        "9. Governing law. Each party submits to the courts of the Province. This Agreement is governed by the law of the Province."],
      // A new sentence that opens with the word the next one does is still inserted before it.
      ["This Agreement is", "This clause binds the parties. This Agreement is", [["insert", "This clause binds the parties. "]],
        "9. Governing law. This clause binds the parties. This Agreement is governed by the law of the Province."],
      ["governed by the law of the Province.", "governed by the law.", [["delete", " of the Province"]],
        "9. Governing law. This Agreement is governed by the law."],
      ["the law of the Province", "the laws of the Province", [["delete", "law"], ["insert", "laws"]],
        "9. Governing law. This Agreement is governed by the laws of the Province."],
      // Two changed words, each its own change; the words between stay as they are.
      ["governed by the law of the Province", "governed by the laws of the State", [["delete", "law"], ["insert", "laws"], ["delete", "Province"], ["insert", "State"]],
        "9. Governing law. This Agreement is governed by the laws of the State."],
    ] as const) {
      const packet = await createDocxLocalization(source, [{ id: "law", before_quote, replacement }], {}, layer)
      const clause = packet.candidates.find(row => row.paragraph.text.startsWith("9."))!
      const output = await applyDocxLocalization(source, packet, [{ amendment_id: "law", target_handle: clause.target_handle }], tracked)
      expect(output.report.package_structure_verified).toBe(true)
      const after = await revisions(output.bytes)
      expect(after.filter(row => row.author !== author)).toEqual(theirs)
      expect(after.filter(row => row.author === author).map(row => [row.type, row.text])).toEqual(expected.map(row => [...row]))
      expect((await createDocxRevision(output.bytes, layer)).paragraphs[9]!.text).toBe(text)
    }
  })

  // A whole-paragraph quote ends sentences with the full stop the change sits
  // beside; the change takes the space after it, never the full stop.
  test("an inserted or deleted sentence keeps to the sentence boundary and the formatting of the text it joins", async () => {
    const author = "Jordan Example"
    const tracked = { ...layer, revisionAuthor: author }
    const CLAUSE = '<w:r><w:t xml:space="preserve">9. Governing law. This Agreement is governed by the law of the Province.</w:t></w:r>'
    const restyled = (runs: string) => {
      const parts = unzipSync(source)
      const xml = strFromU8(parts["word/document.xml"]!)
      expect(xml).toContain(CLAUSE)
      parts["word/document.xml"] = strToU8(xml.replace(CLAUSE, runs))
      return zipSync(parts)
    }
    const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
    const bold = (text: string) => `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`
    const heading = restyled(bold("9. Governing law.") + plain(" This Agreement is governed by the law of the Province."))
    const term = restyled(plain("9. Governing law. This Agreement is governed by the law of the ") + bold("Province") + plain("."))
    const notices = restyled(plain("9. Governing law. This Agreement is governed by the law of the Province. Notices are served in the Province."))
    const whole = "9. Governing law. This Agreement is governed by the law of the Province."
    const sentence = "Each party submits to the courts."
    // A sentence that ends with the words before it could also stand inside the sentence before.
    const repeat = "Each party submits to the courts of the Province."
    for (const [bytes, before_quote, replacement, expected] of [
      [source, whole, `9. Governing law. ${sentence} This Agreement is governed by the law of the Province.`, ["insert", `${sentence} `]],
      [heading, whole, `9. Governing law. ${sentence} This Agreement is governed by the law of the Province.`, ["insert", `${sentence} `]],
      [heading, "Governing law. This Agreement", `Governing law. ${sentence} This Agreement`, ["insert", `${sentence} `]],
      [source, "the law of the Province.", `the law of the Province. ${sentence}`, ["insert", ` ${sentence}`]],
      [term, "the law of the Province.", `the law of the Province. ${sentence}`, ["insert", ` ${sentence}`]],
      [source, whole, "9. Governing law.", ["delete", " This Agreement is governed by the law of the Province."]],
      [source, whole, `${whole} ${repeat}`, ["insert", ` ${repeat}`]],
      [source, "the Province.", `the Province. ${repeat}`, ["insert", ` ${repeat}`]],
      [notices, "the Province. Notices", `the Province. ${repeat} Notices`, ["insert", `${repeat} `]],
      [notices, "the law of the Province. Notices are served in the Province.", "the law of the Province.", ["delete", " Notices are served in the Province."]],
      [source, "the law of the Province", "the law of the Province and its courts", ["insert", " and its courts"]],
    ] as const) {
      const packet = await createDocxLocalization(bytes, [{ id: "law", before_quote, replacement }], {}, layer)
      const clause = packet.candidates.find(row => row.paragraph.text.startsWith("9."))!
      const output = await applyDocxLocalization(bytes, packet, [{ amendment_id: "law", target_handle: clause.target_handle }], tracked)
      expect((await revisions(output.bytes)).filter(row => row.author === author).map(row => [row.type, row.text])).toEqual([[...expected]])
      // The new sentence is body text, never part of the bold heading or term beside it.
      const xml = strFromU8(unzipSync(output.bytes)["word/document.xml"]!)
      for (const change of xml.match(/<w:ins [^>]*>[\s\S]*?<\/w:ins>/g)?.filter(change => change.includes(sentence)) ?? [])
        expect(change).not.toContain("<w:b")
    }
  })

  // Word's selection often takes a heading's trailing space with it. An
  // inserted word is typed where it stands, so it is not a defined term's.
  test("a tracked sentence inserted before text takes that text's formatting, and other insertions the formatting before them", async () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    const author = "Jordan Example"
    const CLAUSE = '<w:r><w:t xml:space="preserve">9. Governing law. This Agreement is governed by the law of the Province.</w:t></w:r>'
    const restyled = (runs: string) => {
      const parts = unzipSync(source)
      const xml = strFromU8(parts["word/document.xml"]!)
      expect(xml).toContain(CLAUSE)
      parts["word/document.xml"] = strToU8(xml.replace(CLAUSE, runs))
      return zipSync(parts)
    }
    const run = (style: string, text: string) => `<w:r>${style ? `<w:rPr>${style}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`
    const heading = restyled(run("<w:b/>", "9. Governing law. ") + run("", "This Agreement is governed by the law of the Province."))
    const term = restyled(run("", "9. Governing law. This Agreement is governed by the law of the ") + run("<w:b/>", "Province") + run("", "."))
    const act = restyled(run("", "9. Governing law. This Agreement is governed by the ") + run("<w:i/>", "Sale of Goods Act") + run("", " and the law of the Province."))
    const italic = restyled(run("<w:i/>", "9. Governing law. This Agreement is governed by the law of the Province."))
    const whole = "9. Governing law. This Agreement is governed by the law of the Province."
    for (const [bytes, before_quote, replacement, inserted, style] of [
      [heading, whole, "9. Governing law. Each party submits. This Agreement is governed by the law of the Province.", "Each party submits. ", ""],
      [heading, "This Agreement", "Each party submits. This Agreement", "Each party submits. ", ""],
      [term, "the law of the Province", "the law of the said Province", "said ", ""],
      [act, "governed by the Sale of Goods Act", "governed by the amended Sale of Goods Act", "amended ", ""],
      [italic, whole, `${whole} Each party submits.`, " Each party submits.", "<w:i/>"],
    ] as const) {
      const packet = await createDocxLocalization(bytes, [{ id: "law", before_quote, replacement }], {}, layer)
      const clause = packet.candidates.find(row => row.paragraph.text.startsWith("9."))!
      const output = await applyDocxLocalization(bytes, packet, [{ amendment_id: "law", target_handle: clause.target_handle }], { ...layer, revisionAuthor: author })
      expect((await revisions(output.bytes)).filter(row => row.author === author).map(row => [row.type, row.text])).toEqual([["insert", inserted]])
      const ours = Array.from(new DOMParser().parseFromString(strFromU8(unzipSync(output.bytes)["word/document.xml"]!), "application/xml")
        .getElementsByTagNameNS(W, "ins")).filter(node => node.getAttributeNS(W, "author") === author)
      expect(ours.flatMap(node => Array.from(node.getElementsByTagNameNS(W, "r"))).map(row => {
        const properties = row.getElementsByTagNameNS(W, "rPr")[0]
        return properties ? Array.from(properties.childNodes).filter(node => node.nodeType === 1).map(node => `<w:${(node as Element).localName}/>`).join("") : ""
      })).toEqual([style])
    }
  })

  test("a later tracked insertion stands beside doc.haus's own inserted sentence, never inside it", async () => {
    const author = "Jordan Example"
    const own = { ...layer, ownAuthors: [author], revisionAuthor: author }
    const write = async (bytes: Uint8Array, before_quote: string, replacement: string) => {
      const packet = await createDocxLocalization(bytes, [{ id: "law", before_quote, replacement }], {}, own)
      const clause = packet.candidates.find(row => row.paragraph.text.startsWith("9."))!
      return applyDocxLocalization(bytes, packet, [{ amendment_id: "law", target_handle: clause.target_handle }], own)
    }
    const first = (await write(source, "This Agreement", "Each party submits to the courts. This Agreement")).bytes
    await expect(write(first, "submits to the courts", "submits irrevocably to the courts")).rejects.toThrow("docx_own_revision:law")
    // Written as its own w:ins just after the first, which Docxodus lists with
    // it, whether or not the quote reaches back into the first.
    for (const [before_quote, replacement] of [["This Agreement", "Both parties agree. This Agreement"],
      ["the courts. This Agreement", "the courts. Both parties agree. This Agreement"]] as const) {
      const beside = await write(first, before_quote, replacement)
      expect((await revisions(beside.bytes)).filter(row => row.author === author).map(row => [row.type, row.text]))
        .toEqual([["insert", "Each party submits to the courts. Both parties agree. "]])
      expect((await createDocxRevision(beside.bytes, layer)).paragraphs[9]!.text)
        .toBe("9. Governing law. Each party submits to the courts. Both parties agree. This Agreement is governed by the law of the Province.")
    }
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

test("preserves annotated paragraphs and refuses edits inside controls", async () => {
  const original = (await renderMarkdownToDocx("Alpha Alpha.\n\nUnrelated Monday.")).bytes
  const fixture = (annotation: string) => {
    const parts = unzipSync(original), doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
    const text = Array.from(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "t")).find(node => node.textContent === "Alpha Alpha.")!
    const run = text.parentNode!, parent = run.parentNode!, wrapper = new DOMParser().parseFromString(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${annotation}</w:p>`, "application/xml")
    for (const child of Array.from(wrapper.documentElement.childNodes)) parent.insertBefore(doc.importNode(child, true), run)
    parent.removeChild(run); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
    return zipSync(parts)
  }
  const annotations = [
    '<w:r><w:t>Alpha </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="term"/><w:id w:val="27"/><w:lock w:val="contentLocked"/><w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000000}" w:xpath="/root/value"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Alpha</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>.</w:t></w:r>',
    '<w:bookmarkStart w:id="5" w:name="term"/><w:r><w:t>Alpha Alpha.</w:t></w:r><w:bookmarkEnd w:id="5"/>',
    '<w:commentRangeStart w:id="5"/><w:r><w:t>Alpha Alpha.</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r>',
  ]
  for (const [index, annotation] of annotations.entries()) {
    const bytes = fixture(annotation), revision = await createDocxRevision(bytes)
    const protectedParagraph = revision.paragraphs.find(p => p.text === "Alpha Alpha.")!
    const safeParagraph = revision.paragraphs.find(p => p.text === "Unrelated Monday.")!
    const edit = { id: "protected", paragraph_handle: protectedParagraph.paragraph_handle, occurrence: index === 0 ? 1 : 0, before_quote: "Alpha", replacement: "Beta" }
    if (index === 0) {
      for (const options of [{}, { pendingRevisions: "layer" as const, revisionAuthor: "doc.haus" }]) {
        await expect(applyDocxParagraphEdits(bytes, [edit], undefined, options)).rejects.toThrow("docx_protected_inline_content")
        const beside = { ...edit, occurrence: 0 }
        expect((await applyDocxParagraphEdits(bytes, [beside], undefined, options)).bytes.length).toBeGreaterThan(0)
      }
    }
    else expect((await applyDocxParagraphEdits(bytes, [edit])).bytes.length).toBeGreaterThan(0)
    const output = await applyDocxParagraphEdits(bytes, [{ id: "safe", paragraph_handle: safeParagraph.paragraph_handle, occurrence: 0, before_quote: "Monday", replacement: "Tuesday" }])
    expect((await createDocxRevision(output.bytes)).paragraphs.some(p => p.text === "Unrelated Tuesday.")).toBe(true)
  }
  const revisedControl = annotations[0]!.replace('<w:r><w:t>Alpha</w:t></w:r></w:sdtContent>', '<w:del w:id="42" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Old</w:delText></w:r></w:del><w:ins w:id="43" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Alpha</w:t></w:r></w:ins></w:sdtContent>')
  const prior = fixture(revisedControl), layer = { pendingRevisions: "layer" as const, revisionAuthor: "doc.haus", ownAuthors: ["doc.haus"] }
  const revision = await createDocxRevision(prior, layer), annotated = revision.paragraphs.find(p => p.text === "Alpha Alpha.")!
  const result = await applyDocxParagraphEdits(prior, [{ id: "beside", paragraph_handle: annotated.paragraph_handle, occurrence: 0, before_quote: "Alpha", replacement: "Beta" }], undefined, layer)
  const xml = strFromU8(unzipSync(result.bytes)["word/document.xml"]!)
  expect(xml).toContain('w:id="42"'); expect(xml).toContain('w:id="43"')

})

test("preservation refuses a control moved across identical text in an edited paragraph", async () => {
  const original = (await renderMarkdownToDocx("Alpha Alpha.")).bytes, parts = unzipSync(original)
  const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml"), W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const text = Array.from(doc.getElementsByTagNameNS(W, "t")).find(node => node.textContent === "Alpha Alpha.")!, run = text.parentNode!, parent = run.parentNode!
  const control = new DOMParser().parseFromString(`<w:p xmlns:w="${W}"><w:r><w:t>Alpha</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="term"/><w:id w:val="27"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Alpha</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>.</w:t></w:r></w:p>`, "application/xml")
  for (const child of Array.from(control.documentElement.childNodes)) parent.insertBefore(doc.importNode(child, true), run)
  parent.removeChild(run); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  const source = zipSync(parts), sdt = doc.getElementsByTagNameNS(W, "sdt")[0]!, space = sdt.previousSibling!, first = space.previousSibling!
  parent.insertBefore(sdt, first); parent.insertBefore(space, first)
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  const before = await openRevision(source), after = await openRevision(zipSync(parts))
  try {
    expect(before.revision.paragraphs[0]!.text).toBe(after.revision.paragraphs[0]!.text)
    expect(() => verifyPreservation(before, after, new Set([0]))).toThrow("docx_protected_inline_content_changed")
  } finally { before.session.close(); after.session.close() }
})

test("preservation checks controls hidden in prior deleted revisions", async () => {
  const original = (await renderMarkdownToDocx("Working Monday.")).bytes, parts = unzipSync(original), W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml"), paragraph = doc.getElementsByTagNameNS(W, "p")[0]!
  const revision = new DOMParser().parseFromString(`<w:del xmlns:w="${W}" w:id="42" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:sdt><w:sdtPr><w:tag w:val="protected"/><w:id w:val="27"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:delText>Old</w:delText></w:r></w:sdtContent></w:sdt></w:del>`, "application/xml")
  paragraph.appendChild(doc.importNode(revision.documentElement, true)); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  const source = zipSync(parts), tag = doc.getElementsByTagNameNS(W, "tag")[0]!
  tag.setAttributeNS(W, "w:val", "altered"); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  const options = { pendingRevisions: "layer" as const, ownAuthors: ["doc.haus"], revisionAuthor: "doc.haus" }
  const before = await openRevision(source, options), after = await openRevision(zipSync(parts), options)
  try {
    expect(before.revision.paragraphs[0]!.text).toBe(after.revision.paragraphs[0]!.text)
    expect(() => verifyPreservation(before, after, new Set([0]))).toThrow("docx_protected_inline_content_changed")
  } finally { before.session.close(); after.session.close() }
})

test("refuses edits inside an enclosing block content control", async () => {
  const original = (await renderMarkdownToDocx("Working Monday.")).bytes, parts = unzipSync(original), W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml"), paragraph = doc.getElementsByTagNameNS(W, "p")[0]!, parent = paragraph.parentNode!
  const wrapper = new DOMParser().parseFromString(`<w:sdt xmlns:w="${W}"><w:sdtPr><w:lock w:val="contentLocked"/><w:id w:val="27"/></w:sdtPr><w:sdtContent/></w:sdt>`, "application/xml")
  const imported = doc.importNode(wrapper.documentElement, true) as Element
  parent.replaceChild(imported, paragraph); imported.getElementsByTagNameNS(W, "sdtContent")[0]!.appendChild(paragraph)
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc)); const source = zipSync(parts)
  const revision = await createDocxRevision(source), selected = revision.paragraphs.find(p => p.text === "Working Monday.")!
  await expect(applyDocxParagraphEdits(source, [{ id: "inside", paragraph_handle: selected.paragraph_handle, occurrence: 0, before_quote: "Monday", replacement: "Tuesday" }])).rejects.toThrow("docx_protected_inline_content:inside")
})

test("preservation refuses moved comment and bookmark anchors with unchanged visible wording", async () => {
  const original = (await renderMarkdownToDocx("Alpha Alpha.")).bytes, W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  for (const tag of ["bookmarkStart", "commentRangeStart"]) {
    const parts = unzipSync(original), doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
    const text = Array.from(doc.getElementsByTagNameNS(W, "t")).find(node => node.textContent === "Alpha Alpha.")!, run = text.parentNode!, parent = run.parentNode!
    const fixture = new DOMParser().parseFromString(`<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:${tag} w:id="55"/><w:r><w:t>Alpha.</w:t></w:r></w:p>`, "application/xml")
    for (const child of Array.from(fixture.documentElement.childNodes)) parent.insertBefore(doc.importNode(child, true), run)
    parent.removeChild(run); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc)); const source = zipSync(parts)
    const marker = Array.from(doc.getElementsByTagNameNS(W, tag)).find(node => node.getAttributeNS(W, "id") === "55")!
    parent.insertBefore(marker, marker.previousSibling); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
    const before = await openRevision(source), after = await openRevision(zipSync(parts))
    try { expect(() => verifyPreservation(before, after, new Set([0]))).toThrow("docx_protected_inline_content_changed") }
    finally { before.session.close(); after.session.close() }
    // An independently added marker is not part of a paragraph edit's authority.
    parent.insertBefore(marker, marker.nextSibling?.nextSibling ?? null)
    const extra = doc.createElementNS(W, "w:commentRangeStart"); extra.setAttributeNS(W, "w:id", "56"); parent.appendChild(extra)
    parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
    const baseline = await openRevision(source), added = await openRevision(zipSync(parts))
    try { expect(() => verifyPreservation(baseline, added, new Set([0]))).toThrow("docx_protected_inline_content_changed") }
    finally { baseline.session.close(); added.session.close() }
  }
})

test("edits a simple unlocked rich-text control with tracked and untracked writes", async () => {
  const original = (await renderMarkdownToDocx("Before Alpha beta term after.\n\nOther paragraph.")).bytes, W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const fixture = (extra = "") => {
    const parts = unzipSync(original), doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
    const text = Array.from(doc.getElementsByTagNameNS(W, "t")).find(node => node.textContent === "Before Alpha beta term after.")!, run = text.parentNode!, parent = run.parentNode!
    const sample = new DOMParser().parseFromString(`<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="term"/><w:id w:val="27"/>${extra}</w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>Alpha beta term</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>`, "application/xml")
    for (const child of Array.from(sample.documentElement.childNodes)) parent.insertBefore(doc.importNode(child, true), run)
    parent.removeChild(run); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc)); return zipSync(parts)
  }
  for (const options of [{}, { pendingRevisions: "layer" as const, revisionAuthor: "doc.haus" }]) {
    const source = fixture(), revision = await createDocxRevision(source, options), p = revision.paragraphs.find(row => row.text.includes("Alpha beta"))!
    const edit = { id: "term", paragraph_handle: p.paragraph_handle, occurrence: 0, before_quote: "beta", replacement: "delta" }
    const output = await applyDocxParagraphEdits(source, [edit], undefined, options)
    const doc = new DOMParser().parseFromString(strFromU8(unzipSync(output.bytes)["word/document.xml"]!), "application/xml"), control = doc.getElementsByTagNameNS(W, "sdt")[0]!
    expect(control.getElementsByTagNameNS(W, "tag")[0]!.getAttributeNS(W, "val")).toBe("term")
    expect(control.getElementsByTagNameNS(W, "id")[0]!.getAttributeNS(W, "val")).toBe("27")
    expect(Array.from(control.getElementsByTagNameNS(W, "r")).every(run => run.getElementsByTagNameNS(W, "b").length === 1)).toBe(true)
    if (options.revisionAuthor) {
      const opened = await openRevision(source, { ...options, amendmentAuthors: { other: "Other" } }), ordinal = opened.revision.paragraphs.findIndex(row => row.text.includes("Alpha beta"))
      const span = new Map([[ordinal, [{ offset: p.text.indexOf("beta"), replaced: "beta", changed: "delta" }]]])
      const tamper = async (mutate: (control: Element) => void) => {
        const parts = unzipSync(output.bytes), altered = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
        mutate(altered.getElementsByTagNameNS(W, "sdt")[0]!); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(altered))
        const saved = await openRevision(zipSync(parts), options)
        try { expect(() => verifyPreservation(opened, saved, new Set([ordinal]), span)).toThrow("docx_protected_inline_content_changed") }
        finally { saved.session.close() }
      }
      try {
        await tamper(control => control.getElementsByTagNameNS(W, "tag")[0]!.setAttributeNS(W, "w:val", "wrong"))
        await tamper(control => control.getElementsByTagNameNS(W, "ins")[0]!.setAttributeNS(W, "w:author", "Other"))
        await tamper(control => control.getElementsByTagNameNS(W, "t")[0]!.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "default"))
        await tamper(control => { const text = control.getElementsByTagNameNS(W, "t")[0]!; text.appendChild(text.ownerDocument.createElementNS(W, "w:object")) })
        await tamper(control => { const bold = control.getElementsByTagNameNS(W, "b")[0]!; bold.parentNode!.replaceChild(bold.ownerDocument.createElementNS(W, "w:i"), bold) })
      } finally { opened.session.close() }
      const dx = await docxodus()
      for (const accept of [false, true]) {
        const session = dx.openDocxSession(output.bytes, { persistAnchorIds: false })
        try {
          for (const rev of session.listRevisions()) expect((accept ? session.acceptRevision(rev.id) : session.rejectRevision(rev.id)).success).toBe(true)
          const resolved = new DOMParser().parseFromString(strFromU8(unzipSync(session.save())["word/document.xml"]!), "application/xml")
          expect(resolved.getElementsByTagNameNS(W, "sdt")[0]!.textContent).toBe(accept ? "Alpha delta term" : "Alpha beta term")
        } finally { session.close() }
      }
    }
    const entire = { ...edit, before_quote: "Alpha beta term", replacement: "Gamma delta clause" }
    const entireOutput = await applyDocxParagraphEdits(source, [entire], undefined, options)
    expect((await createDocxRevision(entireOutput.bytes, { pendingRevisions: "layer" })).paragraphs.find(row => row.text.includes("Gamma delta clause"))).toBeDefined()
    if (options.revisionAuthor) {
      const dx = await docxodus()
      for (const accept of [false, true]) {
        const session = dx.openDocxSession(entireOutput.bytes, { persistAnchorIds: false })
        try {
          for (const rev of session.listRevisions()) expect((accept ? session.acceptRevision(rev.id) : session.rejectRevision(rev.id)).success).toBe(true)
          const resolved = new DOMParser().parseFromString(strFromU8(unzipSync(session.save())["word/document.xml"]!), "application/xml")
          expect(resolved.getElementsByTagNameNS(W, "sdt")[0]!.textContent).toBe(accept ? "Gamma delta clause" : "Alpha beta term")
        } finally { session.close() }
      }
    }
    // Shared words make this a partial edge edit, not an exact whole-control change.
    await expect(applyDocxParagraphEdits(source, [{ ...entire, replacement: "Alpha delta clause" }], undefined, options)).rejects.toThrow("docx_protected_inline_content")
    for (const [before_quote, replacement] of [["Alpha", "Gamma"], ["Before Alpha", "Before Gamma"], ["Alpha", "New Alpha"]])
      await expect(applyDocxParagraphEdits(source, [{ ...edit, before_quote, replacement }], undefined, options)).rejects.toThrow("docx_protected_inline_content")
    const nestedParts = unzipSync(source), nestedDoc = new DOMParser().parseFromString(strFromU8(nestedParts["word/document.xml"]!), "application/xml")
    const content = nestedDoc.getElementsByTagNameNS(W, "sdtContent")[0]!, child = content.firstChild!
    const nested = new DOMParser().parseFromString(`<w:sdt xmlns:w="${W}"><w:sdtPr><w:id w:val="28"/></w:sdtPr><w:sdtContent/></w:sdt>`, "application/xml")
    const wrapped = nestedDoc.importNode(nested.documentElement, true) as Element
    content.replaceChild(wrapped, child); wrapped.getElementsByTagNameNS(W, "sdtContent")[0]!.appendChild(child)
    nestedParts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(nestedDoc)); const nestedFile = zipSync(nestedParts)
    const nestedView = await createDocxRevision(nestedFile, options), nestedP = nestedView.paragraphs.find(row => row.text.includes("Alpha beta"))!
    await expect(applyDocxParagraphEdits(nestedFile, [{ ...edit, paragraph_handle: nestedP.paragraph_handle }], undefined, options)).rejects.toThrow("docx_protected_inline_content")
    for (const extra of ['<w:lock w:val="sdtLocked"/>', '<w:dataBinding w:storeItemID="id" w:xpath="/root"/>', '<w:text/>']) {
      const protectedFile = fixture(extra), view = await createDocxRevision(protectedFile, options), row = view.paragraphs.find(row => row.text.includes("Alpha beta"))!
      await expect(applyDocxParagraphEdits(protectedFile, [{ ...edit, paragraph_handle: row.paragraph_handle }], undefined, options)).rejects.toThrow("docx_protected_inline_content")
    }
  }
})

test("an edit inside a simple control preserves other reviewers' revisions and comments", async () => {
  const parts = unzipSync(counterpartyMarkupDocx()), W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", parser = new DOMParser()
  const doc = parser.parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  const text = Array.from(doc.getElementsByTagNameNS(W, "t")).find(node => node.textContent === "9. Governing law. This Agreement is governed by the law of the Province.")!
  const run = text.parentNode!, parent = run.parentNode!, block = parser.parseFromString(`<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">9. Governing law. </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="law"/><w:id w:val="27"/></w:sdtPr><w:sdtContent><w:r><w:t>This Agreement is governed by the law of the Province.</w:t></w:r></w:sdtContent></w:sdt></w:p>`, "application/xml")
  for (const child of Array.from(block.documentElement.childNodes)) parent.insertBefore(doc.importNode(child, true), run)
  parent.removeChild(run); parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc)); const source = zipSync(parts)
  const options = { pendingRevisions: "layer" as const, revisionAuthor: "doc.haus", ownAuthors: ["doc.haus"] }
  const view = await createDocxRevision(source, options), paragraph = view.paragraphs.find(row => row.text.includes("the Province."))!
  const output = await applyDocxParagraphEdits(source, [{ id: "law", paragraph_handle: paragraph.paragraph_handle, occurrence: 0, before_quote: "Province", replacement: "State" }], undefined, options)
  expect(readDocxComments(output.bytes)).toEqual(readDocxComments(source))
  const dx = await docxodus(), originalSession = dx.openDocxSession(source, { persistAnchorIds: false })
  const previous = originalSession.listRevisions(); originalSession.close()
  for (const accept of [false, true]) {
    const session = dx.openDocxSession(output.bytes, { persistAnchorIds: false })
    try {
      const known = new Set(previous.map(rev => rev.id))
      for (const rev of session.listRevisions().filter(rev => !known.has(rev.id))) expect((accept ? session.acceptRevision(rev.id) : session.rejectRevision(rev.id)).success).toBe(true)
      expect(session.listRevisions().map(({ id, author, type, text, date }) => ({ id, author, type, text, date }))).toEqual(previous.map(({ id, author, type, text, date }) => ({ id, author, type, text, date })))
      const result = parser.parseFromString(strFromU8(unzipSync(session.save())["word/document.xml"]!), "application/xml"), control = result.getElementsByTagNameNS(W, "sdt")[0]!
      expect(control.textContent).toBe(`This Agreement is governed by the law of the ${accept ? "State" : "Province"}.`)
      expect(control.getElementsByTagNameNS(W, "tag")[0]!.getAttributeNS(W, "val")).toBe("law")
    } finally { session.close() }
  }
})
