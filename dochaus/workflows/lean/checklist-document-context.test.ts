import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { LEAN_STATE_SCHEMA, hashStageInputs, readLeanState, writeLeanText } from "./state"
import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { exportDocxComments, readDocxCommentParagraphs, readDocxComments, readDocxRevisionText } from "../../lib/docx-comments"
import { docxodus } from "../../lib/docxodus"
import { renderMarkdownToDocx } from "../../lib/render"
import { buildChecklistDocumentContext } from "./checklist-document-context"
import { assignChecklistTargetContext, checklistTargetAssignments } from "./checklist-target-context"
import { checklistCheckerInputSchema, checklistTextMeasuresMeaning, type ChecklistItem } from "./checklist-contract"
import { writeLeanPlan } from "./plan"
import { leanFile, writeLeanBytes } from "./state"

const item = (document = "record.docx", check = "Verify the violet inventory receipt."): ChecklistItem => ({
  check, basis: "directive", source_refs: [{ document: "request.txt", locator: "line 1", quote: "Verify the violet inventory receipt." }],
  qualifications: ["Subject to receipt confirmation."], target: { document, role: "Requested record", treatment: "operative" },
})
function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "checklist-context-"))
  writeLeanPlan(directory, [
    { title: "Record", purpose: "Requested record", file: "record.docx", parts: [{ heading: "Record", scope: "Record contents" }] },
    { title: "Explanation", purpose: "Requested explanation", file: "explanation.docx", parts: [{ heading: "Explanation", scope: "Explanation contents" }] },
  ])
  return directory
}
function acceptApplyForTest(directory: string, document = "d1") {
 const step = `${document}.apply` as `d${number}.apply`
 const prior = readLeanState(directory)
 writeLeanText(directory, "state.json", JSON.stringify({ schema: LEAN_STATE_SCHEMA, profile: "lean", stages: { ...prior?.stages,
 [step]: { status: "done", input_sha: hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)) } } }))
}
async function save(directory: string, file: string, markdown: string) {
  const bytes = (await renderMarkdownToDocx(markdown)).bytes
  writeLeanBytes(directory, file, bytes)
  if (file.endsWith("/body-applied.docx")) acceptApplyForTest(directory, file.split("/")[0])
  return bytes
}

test("applied bytes take priority, item stays unchanged, only target hints go inline", async () => {
  const directory = matter()
  try {
    await save(directory, "d1/body.docx", "Old violet inventory receipt.")
    const current = await save(directory, "d1/body-applied.docx", "# Record\n\nAdjacent before.\n\nThe violet inventory receipt is pending confirmation.\n\nAdjacent after.\n\nDistant unrelated sentence.")
    await save(directory, "d2/body.docx", "The other violet inventory receipt is explained here.")
    const unchanged = item(), before = JSON.stringify(unchanged)
    const instructions = [{ document: "later.txt", locator: "line 2", quote: "Receipt is not yet confirmed." }]
    const result = buildChecklistDocumentContext(directory, { item_id: "one", item: unchanged, settled_instructions: instructions })
    expect(result.input.item).toBe(unchanged)
    expect(JSON.stringify(unchanged)).toBe(before)
    expect(result.input.settled_instructions).toBe(instructions)
    expect(result.input.current_evidence.every(row => row.document === "record.docx")).toBe(true)
    expect(result.input.current_evidence.map(row => row.text)).toContain("The violet inventory receipt is pending confirmation.")
    expect(result.input.current_evidence.map(row => row.text)).not.toContain("The other violet inventory receipt is explained here.")
    expect(result.artifacts.find(row => row.value.id === "d1")!.value.paragraphs).toEqual(readDocxCommentParagraphs(current))
    expect(result.artifacts.find(row => row.value.id === "d1")!.value.artifact).toEndWith("body-applied.docx")
    expect(result.input.retrieval).toContain("NONEXHAUSTIVE")
    expect(result.input.retrieval).toContain("view=source")
    expect(result.input.retrieval).toContain("view=current_body")
    expect(result.input.retrieval).toContain("An unavailable view means no search occurred")
    expect(result.input.retrieval).toContain("next_text_offset")
    expect(result.input.retrieval).toContain("Native comments and inserted/deleted revision text")
    expect(checklistCheckerInputSchema.safeParse(result.input).success).toBe(true)
    expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(Buffer.from(current))
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("zero lexical hits gives no default passages while exact plan ID resolves", async () => {
  const directory = matter()
  try {
    await save(directory, "d1/body.docx", "# Record\n\nAmber cobalt jasper.")
    const requirement = item("d1", "Ultraviolet zirconium."); requirement.source_refs[0]!.quote = "Ultraviolet zirconium."; requirement.qualifications = []
    const result = buildChecklistDocumentContext(directory, { item_id: "empty", item: requirement })
    expect(result.input.current_evidence).toEqual([])
    expect(result.artifacts.find(row => row.value.id === "d1")!.value.paragraphs).toHaveLength(2)
    expect(result.input.retrieval).toContain("No matching hint does not establish absence")
    expect(result.warnings.some(warning => warning.includes("exact plan identity"))).toBe(false)
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("all records across three documents retain table cells and distinct duplicate-text PIDs", async () => {
  const directory = matter()
  try {
    const plan = JSON.parse(readFileSync(leanFile(directory, "plan.json"), "utf8"))
    plan.documents.push({ id: "d3", title: "Appendix", purpose: "Appendix", file: "appendix.docx", parts: [{ heading: "Appendix", scope: "Appendix" }] })
    writeFileSync(leanFile(directory, "plan.json"), JSON.stringify(plan))
    for (const id of ["d1", "d2", "d3"]) await save(directory, `${id}/body.docx`, `# Record ${id}\n\n` + Array.from({ length: 100 }, (_, i) => i === 25 || i === 75 ? "Repeated violet receipt." : `Entry ${i} for ${id}.`).join("\n\n") + "\n\n| Key | Value |\n| --- | --- |\n| Color | violet |")
    const result = buildChecklistDocumentContext(directory, { item_id: "scale", item: item() })
    expect(result.artifacts).toHaveLength(3)
    expect(result.artifacts.reduce((sum, a) => sum + a.value.paragraphs.length, 0)).toBe(315)
    expect(result.input.current_evidence.length).toBeLessThanOrEqual(6)
    const target = result.artifacts.find(row => row.value.id === "d1")!
    const repeated = target.value.paragraphs.filter(p => p.paragraph_text === "Repeated violet receipt.")
    expect(repeated).toHaveLength(2)
    expect(repeated[0]!.pids).not.toEqual(repeated[1]!.pids)
    expect(target.value.paragraphs.map(p => p.paragraph_text)).toContain("violet")
    for (const artifact of result.artifacts) {
      const file = path.join(directory, artifact.path); mkdirSync(path.dirname(file), { recursive: true }); writeFileSync(file, JSON.stringify(artifact.value))
    }
    for (const hint of result.input.current_evidence) {
      const index = Number(hint.locator.match(/\/paragraphs\/(\d+)$/)![1])
      expect(JSON.parse(readFileSync(path.join(directory, target.path), "utf8")).paragraphs[index].paragraph_text).toBe(hint.text)
    }
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("whitespace-only accepted paragraphs stay in records but are never emitted as checker evidence", async () => {
  const directory = matter()
  try {
    // A lone no-break space (signature spacer) and a spaces-plus-tab paragraph flank the lexical seed.
    const rendered = await renderMarkdownToDocx("# Record\n\nAdjacent before.\n\nZqspacerone\n\nThe violet inventory receipt is pending confirmation.\n\nZqspacertwo\n\nDistant unrelated sentence.")
    const zip = unzipSync(rendered.bytes)
    zip["word/document.xml"] = strToU8(strFromU8(zip["word/document.xml"]!)
      .replace("<w:t xml:space=\"preserve\">Zqspacerone</w:t>", "<w:t xml:space=\"preserve\">\u00a0</w:t>")
      .replace("<w:t xml:space=\"preserve\">Zqspacertwo</w:t>", "<w:t xml:space=\"preserve\">  </w:t><w:tab/><w:t xml:space=\"preserve\"> </w:t>"))
    writeLeanBytes(directory, "d1/body.docx", zipSync(zip))
    const result = buildChecklistDocumentContext(directory, { item_id: "blank", item: item() })
    const record = result.artifacts.find(row => row.value.id === "d1")!
    expect(record.value.paragraphs.map(row => row.paragraph_text)).toEqual(["Record", "Adjacent before.", "\u00a0", "The violet inventory receipt is pending confirmation.", "  \t ", "Distant unrelated sentence."])
    expect(checklistCheckerInputSchema.safeParse(result.input).error?.issues.map(issue => `${issue.path.join(".")}: ${issue.message}`) ?? []).toEqual([])
    // Blank neighbours are dropped, not replaced: the window stays the seed's immediate neighbours.
    expect(result.input.current_evidence.map(row => row.text)).toEqual(["The violet inventory receipt is pending confirmation."])
    expect(result.input.current_evidence[0]!.locator).toEndWith(`read-record ${record.path} /paragraphs/3`)
    // Full-body comparison keeps every assigned range, including blank paragraphs, but emits no blank evidence.
    const body = checklistTargetAssignments(directory, item()).assignments.find(unit => unit.assignment.kind === "target_body")!.assignment
    expect(body.ranges.map(range => range.paragraph_index)).toEqual([0, 1, 2, 3, 4, 5])
    const assigned = assignChecklistTargetContext(result, body)
    expect(checklistCheckerInputSchema.safeParse(assigned.input).error?.issues.map(issue => `${issue.path.join(".")}: ${issue.message}`) ?? []).toEqual([])
    expect(assigned.input.current_evidence.map(row => row.text)).toEqual(["Record", "Adjacent before.", "The violet inventory receipt is pending confirmation.", "Distant unrelated sentence."])
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("actual notes filename resolves; unknown and duplicate targets remain unresolved", async () => {
  const directory = matter()
  try {
    const plan = JSON.parse(readFileSync(leanFile(directory, "plan.json"), "utf8")); plan.notes_file = "review-notes.docx"; writeFileSync(leanFile(directory, "plan.json"), JSON.stringify(plan))
    await save(directory, "drafting-notes.docx", "The violet inventory receipt remains open.")
    const notes = buildChecklistDocumentContext(directory, { item_id: "notes", item: item("review-notes.docx") })
    expect(notes.input.current_evidence).toHaveLength(1)
    expect(notes.input.current_evidence[0]!.document).toBe("review-notes.docx")
    const missing = buildChecklistDocumentContext(directory, { item_id: "unknown", item: item("invented.docx") })
    expect(missing.input.current_evidence).toEqual([])
    expect(missing.warnings.join(" ")).toContain("0 exact plan identity matches")
    plan.documents[1].file = "record.docx"; writeFileSync(leanFile(directory, "plan.json"), JSON.stringify(plan))
    const duplicate = buildChecklistDocumentContext(directory, { item_id: "duplicate", item: item() })
    expect(duplicate.input.current_evidence).toEqual([])
    expect(duplicate.warnings.join(" ")).toContain("2 exact plan identity matches")
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("fresh calls observe changed, deleted and regenerated paragraphs without stale PID reuse", async () => {
  const directory = matter()
  try {
    const first = await save(directory, "d1/body.docx", "The violet inventory receipt is pending.\n\nRemove this paragraph.")
    const before = buildChecklistDocumentContext(directory, { item_id: "first", item: item() })
    const next = await save(directory, "d1/body.docx", "The violet inventory receipt is now confirmed.")
    const after = buildChecklistDocumentContext(directory, { item_id: "second", item: item() })
    expect(after.artifacts[0]!.value.sha256).not.toBe(before.artifacts[0]!.value.sha256)
    expect(after.artifacts[0]!.value.paragraphs).toEqual(readDocxCommentParagraphs(next))
    expect(after.artifacts[0]!.value.paragraphs).toHaveLength(1)
    expect(after.artifacts[0]!.value.paragraphs[0]!.pids).not.toEqual(readDocxCommentParagraphs(first)[0]!.pids)
    expect(after.input.current_evidence[0]!.text).toBe("The violet inventory receipt is now confirmed.")
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("missing PIDs and unreadable applied files remain explicit without mutation or old-body fallback", async () => {
  const directory = matter()
  try {
    const bytes = await save(directory, "d1/body.docx", "The violet inventory receipt is pending.")
    const zip = unzipSync(bytes)
    zip["word/document.xml"] = strToU8(strFromU8(zip["word/document.xml"]!).replace(/<w:bookmarkStart\b[^>]*\/>|<w:bookmarkEnd\b[^>]*\/>/g, ""))
    const noPointers = zipSync(zip); writeLeanBytes(directory, "d1/body.docx", noPointers)
    const missing = buildChecklistDocumentContext(directory, { item_id: "missing-pid", item: item() })
    expect(missing.warnings.join(" ")).toContain("lack native PIDs")
    expect(missing.input.current_evidence[0]!.locator).toContain("PIDs unavailable")
    expect(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx")))).toEqual(noPointers)
    writeLeanBytes(directory, "d1/body-applied.docx", new TextEncoder().encode("not a DOCX"))
    acceptApplyForTest(directory)
    const corrupt = buildChecklistDocumentContext(directory, { item_id: "corrupt", item: item() })
    expect(corrupt.input.current_evidence).toEqual([])
    expect(corrupt.warnings.join(" ")).toContain("could not be read")
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("redline native comments and revision text stay artifact-specific beside newer accepted body", async () => {
  const directory = matter()
  try {
    const prior = (await renderMarkdownToDocx("Prior violet inventory receipt.")).bytes
    const compared = (await renderMarkdownToDocx("Revised violet inventory receipt.")).bytes
    const redline = await (await docxodus()).compareDocuments(prior, compared, { authorName: "Reviewer" })
    const anchor = readDocxCommentParagraphs(redline)[0]
    const annotated = exportDocxComments(redline, [{ id: "explanation", author: "Reviewer", text: "Explains this comparison, not a later body.", anchor: { paragraph_text: anchor.paragraph_text, paragraph_index: anchor.paragraph_index } }])
    expect(annotated.unresolved).toEqual([])
    writeLeanBytes(directory, "d1/redline.docx", annotated.bytes)
    const current = await save(directory, "d1/body-applied.docx", "Latest violet inventory receipt.")
    const context = buildChecklistDocumentContext(directory, { item_id: "native", item: item() })
    const record = context.artifacts.find(artifact => artifact.value.id === "d1")!.value
    expect(record.paragraphs).toEqual(readDocxCommentParagraphs(current))
    expect(record.paragraphs[0].paragraph_text).toBe("Latest violet inventory receipt.")
    const comparison = record.annotations!.find(row => row.artifact.endsWith("/redline.docx"))!
    expect(comparison.comments).toEqual(readDocxComments(annotated.bytes))
    expect(comparison.revisions).toEqual(readDocxRevisionText(annotated.bytes))
    const properties = JSON.parse(comparison.properties)
    expect(properties.artifact).toBe(comparison.artifact)
    expect(properties.sha256).toBe(comparison.sha256)
    expect(properties.part).toBe("word/document.xml")
    expect(properties.revision_elements.ins).toBeGreaterThan(0)
    expect(properties.revision_elements.del).toBeGreaterThan(0)
    expect(properties.accepted_body_matches_current).toBe(false)
    const cleanProperties = JSON.parse(record.annotations!.find(row => row.artifact.endsWith("/body-applied.docx"))!.properties)
    expect(Object.values(cleanProperties.revision_elements).every(count => count === 0)).toBe(true)
    expect(cleanProperties.accepted_body_matches_current).toBe(true)
    expect(context.input.retrieval).toContain("/properties")
    expect(context.input.retrieval).toContain("never that every required change was made")
    writeLeanBytes(directory, "d1/body-applied.docx", compared)
    const aligned = buildChecklistDocumentContext(directory, { item_id: "native", item: item() })
    expect(JSON.parse(aligned.artifacts[0]!.value.annotations!.find(row => row.artifact.endsWith("/redline.docx"))!.properties).accepted_body_matches_current).toBe(true)
    expect(comparison.comments[0].markers[0].accepted_paragraph_text).toBe("Revised violet inventory receipt.")
    expect(comparison.comments[0].markers[0].pids).not.toEqual(record.paragraphs[0].pids)
    expect(record.annotations!.find(row => row.artifact.endsWith("/body-applied.docx"))!.comments).toEqual([])
    expect(context.input.current_evidence.every(row => !row.text.includes("Explains this comparison"))).toBe(true)
    expect(context.input.retrieval).toContain("/annotations")
    expect(context.input.retrieval).toContain("must not be remapped onto the current body by position")
    expect(context.input.retrieval).not.toContain("comments, deleted text, revision properties and rendered layout are NOT inspected")
    expect(readFileSync(leanFile(directory, "d1/redline.docx"))).toEqual(Buffer.from(annotated.bytes))
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("large multi-source matter exposes all exact source identities with bounded headings and no source text", async () => {
  const directory = matter()
  try {
    const originals = Array.from({ length: 4 }, (_, index) => ({
      document: `original-source-${index}.docx`, characters: 120000 + index,
      chunks: Array.from({ length: 90 }, (_, heading) => ({ heading_path: [`Part ${heading}`, `Source ${index}`],
        preview: `Private source body ${index}/${heading} must not be sent as inventory.` })),
    }))
    writeFileSync(path.join(directory, "source-map.json"), JSON.stringify({ documents: originals }))
    await save(directory, "d1/body.docx", "# Record\n\nThe violet inventory receipt is pending confirmation.")
    await save(directory, "d2/body.docx", "# Explanation\n\nThe proposed action awaits approval.")
    const unchanged = item(), before = JSON.stringify(unchanged)
    const result = buildChecklistDocumentContext(directory, { item_id: "original-discovery", item: unchanged })
    const marker = "Original source inventory (bounded navigation metadata, no source text): "
    const inventory = JSON.parse(result.input.retrieval.split(marker)[1]!.split("\n\n")[0]!)
    expect(inventory.map((row: { document: string }) => row.document)).toEqual(originals.map(row => row.document))
    for (const [index, row] of inventory.entries()) {
      expect(row.characters).toBe(120000 + index)
      expect(row.heading_paths).toHaveLength(6)
      expect(row.heading_path_count).toBe(90)
      expect(row.map_record).toBe(`read-record source-map.json /documents/${index}`)
    }
    expect(JSON.stringify(inventory).length).toBeLessThan(1800)
    expect(result.input.retrieval).not.toContain("Private source body")
    expect(result.input.retrieval).toContain("NOT a closed-world evidence set")
    expect(result.input.retrieval).toContain("targeted original-source lookup before claiming")
    expect(result.input.retrieval).toContain("proposed action separately from unknown approval and applicability")
    expect(result.input.current_evidence.every(row => row.document === unchanged.target.document)).toBe(true)
    expect(JSON.stringify(unchanged)).toBe(before)
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("annotation properties give exact accepted-body text measures for clean and redline artifacts, never a page count", async () => {
  const directory = matter()
  try {
    await save(directory, "d1/body-applied.docx", "# Record\n\nThe violet  receipt is pending.\n\n| Key | Value |\n| --- | --- |\n| Caf\u00e9 \u{1D538} | violet |")
    const prior = (await renderMarkdownToDocx("Prior violet inventory receipt.")).bytes
    const compared = (await renderMarkdownToDocx("Revised violet inventory receipt.")).bytes
    writeLeanBytes(directory, "d1/redline.docx", await (await docxodus()).compareDocuments(prior, compared, { authorName: "Reviewer" }))
    const context = buildChecklistDocumentContext(directory, { item_id: "length", item: item() })
    const annotations = context.artifacts.find(row => row.value.id === "d1")!.value.annotations!
    const properties = (suffix: string) => JSON.parse(annotations.find(row => row.artifact.endsWith(suffix))!.properties)
    // Accepted paragraphs: "Record", "The violet  receipt is pending.", "Key", "Value", "Café 𝔸", "violet".
    expect(context.artifacts.find(row => row.value.id === "d1")!.value.paragraphs).toHaveLength(6)
    expect(properties("/body-applied.docx").accepted_body_text_measures).toEqual({ paragraphs: 6, words: 11, characters: 57 })
    // The redline's own accepted text is the revised paragraph, not inserted plus deleted text.
    expect(properties("/redline.docx").accepted_body_text_measures).toEqual({ paragraphs: 1, words: 4, characters: 33 })
    expect(properties("/redline.docx").accepted_body_matches_current).toBe(false)
    for (const row of annotations) expect(row.properties).not.toMatch(/page/i)
    expect(context.input.retrieval).toContain("exact accepted-body text measures")
    expect(context.input.retrieval).toContain("Text measures are exact text counts, not rendered pagination")
    // Measures can disprove, never prove, a rendered-length directive.
    expect(context.input.retrieval).toContain(checklistTextMeasuresMeaning)
    expect(context.input.retrieval).toContain("Never report such a directive satisfied from text measures")
    expect(context.input.retrieval).toContain("a strict page or layout limit that the measures cannot decide remains indeterminate")
    expect(context.input.retrieval).not.toContain("may be judged from them")
    expect(context.input.retrieval).toContain("rendered layout are NOT covered")
    const again = buildChecklistDocumentContext(directory, { item_id: "length", item: item() })
    expect(again.artifacts.find(row => row.value.id === "d1")!.value.annotations!.map(row => row.properties)).toEqual(annotations.map(row => row.properties))
  } finally { rmSync(directory, { recursive: true, force: true }) }
})
