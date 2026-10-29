import { expect, test } from "bun:test"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { docxodus } from "./docxodus"
import { bake, binding, buildRedlined, redlineStaleReason, redlineState } from "./redline"
import { createDocxLocalization, createDocxRevision } from "../../../dochaus/lib/docx-amendments"
import { paragraphTargets, withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import type { RedlineRow } from "./db"

function row(before: string, after: string): RedlineRow {
  return { id: 1, doc_path: "example.docx", doc_name: "example.docx", scope: "clause", find_text: before, old_text: before,
    new_text: after, author: "Counsel", review_comment: "Explain the proposal.", anchor_id: null, created_at: 1,
    status: "pending", target_json: null, formatting_json: JSON.stringify({ spacingAfter: 6 }) }
}

test("bound wording proposals reject any stored formatting instead of silently dropping it", async () => {
  const source = await withParagraphPointers((await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes)
  const before = new Uint8Array(source)
  const wording = { ...row("Second paragraph.", "Updated paragraph."), formatting_json: null }
  const packet = await createDocxLocalization(source, [{ id: "1", before_quote: wording.find_text, replacement: wording.new_text }])
  const target = (await paragraphTargets(source))[1]!
  const bindings = [
    { version: 1, source_sha256: packet.revision.source_sha256, target_handle: packet.candidates[0]!.target_handle, amendment_id: "1" },
    { version: 2, pid: target.pid!, paragraph_sha256: target.sha256, occurrence: 0, base_sha256: packet.revision.source_sha256, amendment_id: "1" },
  ]
  for (const value of bindings) {
    const valid = { ...wording, target_json: JSON.stringify(value) }
    for (const options of [{}, { pendingRevisions: "layer" as const }]) {
      const state = await redlineState(source, [valid], options)
      expect(redlineStaleReason(valid, state, false)).toBeUndefined()
      expect((await buildRedlined(source, [valid], state, options)).length).toBeGreaterThan(0)
      expect((await createDocxRevision(await bake(source, [valid], state, options))).paragraphs.some(p => p.text === wording.new_text)).toBe(true)
      // Only the additional stored column changes: neither its syntax nor an
      // empty value can turn a wording binding into a formatting writer.
      for (const formatting_json of [JSON.stringify({ spacingAfter: 6 }), "", "null", "{invalid"]) {
        const invalid = { ...valid, formatting_json }
        expect(() => binding(invalid)).toThrow("Formatting cannot be combined")
        expect(redlineStaleReason(invalid, state, false)).toContain("invalid Word revision binding")
        await expect(buildRedlined(source, [invalid], state, options)).rejects.toThrow()
        await expect(bake(source, [invalid], state, options)).rejects.toThrow("Formatting cannot be combined")
      }
      const sibling = { ...row("First paragraph.", "First revised paragraph."), id: 2 }
      const invalid = { ...valid, formatting_json: JSON.stringify({ spacingAfter: 6 }) }
      await expect(buildRedlined(source, [sibling, invalid], state, options)).rejects.toThrow()
      await expect(bake(source, [sibling, invalid], state, options)).rejects.toThrow("Formatting cannot be combined")
      expect(source).toEqual(before)
    }
  }
})

test("main redline previews and accepts mixed wording and formatting atomically", async () => {
  const source = (await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes
  const proposal = row("Second paragraph.", "Updated paragraph.")
  const preview = await buildRedlined(source, [proposal])
  const dx = await docxodus(), session = dx.openDocxSession(preview, {})
  try {
    const changes = session.listRevisions()
    expect(changes.some(r => r.type === "format" && r.author === "Counsel")).toBe(true)
    expect(changes.some(r => r.type === "insert" && r.author === "Counsel")).toBe(true)
  } finally { session.close() }
  expect(strFromU8(unzipSync(preview)["word/comments.xml"]!)).toContain("Explain the proposal.")
  const accepted = await bake(source, [proposal])
  const xml = strFromU8(unzipSync(accepted)["word/document.xml"]!)
  expect((await (await import("../../../dochaus/lib/docx-amendments")).createDocxRevision(accepted)).paragraphs.some(p => p.text === "Updated paragraph.")).toBe(true)
  expect(xml).toContain('w:after="120"')
})

test("format only gets a review comment; ambiguous and invalid requests fail closed", async () => {
  const source = (await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes
  const same = row("Second paragraph.", "Second paragraph.")
  const preview = await buildRedlined(source, [same])
  expect(strFromU8(unzipSync(preview)["word/comments.xml"]!)).toContain("Explain the proposal.")
  expect(strFromU8(unzipSync(await bake(source, [same]))["word/document.xml"]!)).toContain('w:after="120"')
  const duplicate = (await renderMarkdownToDocx("Second paragraph.\n\nSecond paragraph.")).bytes
  await expect(bake(duplicate, [same])).rejects.toThrow()
  await expect(buildRedlined(duplicate, [same])).rejects.toThrow()
  await expect(bake(source, [{ ...same, formatting_json: '{"spacingAfter":999}' }])).rejects.toThrow()
})

test("a batch preserves each author and rejects the entire batch when a later target is stale", async () => {
  const source = (await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes
  const first = { ...row("First paragraph.", "First revised paragraph."), author: "Alex" }
  const second = { ...row("Second paragraph.", "Second revised paragraph."), id: 2, author: "Jordan" }
  const preview = await buildRedlined(source, [first, second])
  const session = (await docxodus()).openDocxSession(preview, {})
  try {
    const revisions = session.listRevisions()
    expect(revisions.some(r => r.type === "format" && r.author === "Alex")).toBe(true)
    expect(revisions.some(r => r.type === "format" && r.author === "Jordan")).toBe(true)
  } finally { session.close() }
  const before = new Uint8Array(source)
  await expect(bake(source, [first, { ...second, find_text: "Missing paragraph." }])).rejects.toThrow()
  expect(source).toEqual(before)
})

test("formatting on an uploaded markup keeps prior revisions and comments", async () => {
  const { counterpartyMarkupDocx } = await import("../../../dochaus/test/support/counterparty-markup-fixture")
  const { openRevision } = await import("../../../dochaus/lib/docx-amendments")
  const source = counterpartyMarkupDocx(), opened = await openRevision(source, { pendingRevisions: "layer" })
  let text: string
  try { text = opened.revision.paragraphs.find(p => p.text.includes("Governing law"))!.text } finally { opened.session.close() }
  const proposal = row(text!, text!)
  const prior = (await docxodus()).openDocxSession(source, {})
  const original = prior.listRevisions().map(r => [r.id, r.author, r.type, r.text]); prior.close()
  const output = await bake(source, [proposal], undefined, { pendingRevisions: "layer" })
  const session = (await docxodus()).openDocxSession(output, {})
  try {
    const revisions = session.listRevisions()
    expect(revisions.filter(r => r.author !== "Counsel").map(r => [r.id, r.author, r.type, r.text])).toEqual(original)
    expect(revisions.some(r => r.author === "Counsel" && r.type === "format")).toBe(true)
  } finally { session.close() }
  expect(unzipSync(output)["word/comments.xml"]).toEqual(unzipSync(source)["word/comments.xml"])
})

test("mixed wording can retain already-correct paragraph formatting", async () => {
  const source = (await renderMarkdownToDocx("First paragraph.\n\nSecond **important** paragraph.")).bytes
  const { applyDocxParagraphFormatting } = await import("../../../dochaus/lib/docx-paragraph-formatting")
  const formatted = await applyDocxParagraphFormatting(source, { paragraphText: "Second important paragraph.", format: { spacingAfter: 6 }, author: "Counsel", tracked: false })
  const proposal = row("Second important paragraph.", "Updated important paragraph.")
  const result = await buildRedlined(formatted, [proposal])
  const document = strFromU8(unzipSync(result)["word/document.xml"]!)
  expect(document).toContain("<w:b")
  expect(document).toContain('w:after="120"')
})
