import { beforeAll, describe, expect, test } from "bun:test"
import { strFromU8, unzipSync } from "fflate"
import { createDocxRevision, createDocxLocalization, applyDocxLocalization } from "./docx-amendments"
import { renderMarkdownToDocx } from "./render"
import { docxodus } from "./docxodus"

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
