import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { LEAN_STATE_SCHEMA, hashStageInputs, readLeanState, writeLeanText } from "./state"
import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtempSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { combineMemoAndBody } from "../../lib/combined-docx"
import { exportDocxComments, readDocxCommentParagraphs, readDocxComments, readDocxRevisionText } from "../../lib/docx-comments"
import { docxodus } from "../../lib/docxodus"
import { renderMarkdownToDocx } from "../../lib/render"
import { ensureParagraphRows, paragraphSha256 } from "../../lib/drafting-amendment"
import { buildChecklistDocumentContext } from "./checklist-document-context"
import { assignChecklistTargetContext, checklistTargetAssignments } from "./checklist-target-context"
import { buildChecklistCheckerPrompt, checklistCheckerInputSchema, checklistTextMeasuresMeaning, type ChecklistItem } from "./checklist-contract"
import { combinedDelivery, deliveryArtifactEvidence, memoNotesFile } from "./deliverables"
import { withSourceSnapshot } from "./quotes"
import { writeLeanPlan } from "./plan"
import { leanFile, writeLeanBytes } from "./state"

const item = (document = "record.docx", check = "Verify the violet inventory receipt."): ChecklistItem => ({
  check, basis: "directive", source_refs: [{ document: "request.txt", locator: "line 1", quote: "Verify the violet inventory receipt." }],
  qualifications: ["Subject to receipt confirmation."], target: { document, role: "Requested record", treatment: "operative" },
})
// The delivery rows' fixed wording and the checker prompt's one sentence defining them.
const PLAIN_CODE_SETS = "Code sets this file's name (registration's naming logic on the current plan and matter files; registration may add \" (n)\" to keep an existing matter document) and its Word format."
const COMBINED_CODE_SETS = "Code sets this file's name (registration's naming logic on the current plan; registration adds \" (n)\" to keep another file that has the name), its Word format and its part order: the memo, a section break, then the body from a new page. Once the final gate passes, code builds it from the two registered parts and saves it in the matter; the download serves the saved file. The system replaces a file under the name only while it still holds, unchanged, the combined file the system last saved there, with no change pending review. When the build is refused, the file of that name holds the lawyer's accepted Word changes, or no numbered name is free, it is not saved, and a finalize item says so."
const NOT_COVERED = "Not covered: Word changes accepted after delivery, changes the final gate makes to the notes, and rendered layout."
const DELIVERY_SENTENCE = "A delivery_artifacts row in current_evidence records code-produced properties of a delivered file (its name, Word format and, for a combined file, its parts, their order and a build of it from the current parts). It is evidence only for a requirement about the delivered file itself (its name, format or composition), so cite it, as artifact evidence, only for such a requirement."
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
const sha = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
const texts = (bytes: Uint8Array) => readDocxCommentParagraphs(bytes).map(row => row.paragraph_text)
/** Every file in the matter with its bytes' hash: anything a context build could write. */
const files = (directory: string) => Object.fromEntries((readdirSync(directory, { recursive: true }) as string[]).toSorted()
  .filter(file => statSync(path.join(directory, file)).isFile()).map(file => [file, sha(readFileSync(path.join(directory, file)))]))
/** A target's delivery rows when none of them is a combined file. */
const plainRows = (rows: NonNullable<ReturnType<typeof buildChecklistDocumentContext>["delivery"]>) => rows.map(row => {
  if (row.kind === "combined") throw new Error("Unexpected combined delivery row.")
  return row
})
/** Passage hints only: every resolved target also receives its files' delivery rows. */
const hints = <T extends { locator: string }>(input: { current_evidence: T[] }) => input.current_evidence.filter(row => !row.locator.startsWith("delivery_artifacts"))
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
    expect(hints(result.input).every(row => row.document === "record.docx")).toBe(true)
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

test("a clause Word numbers is hinted with its list number beside, not inside, its text", async () => {
  const directory = matter()
  try {
    await save(directory, "d1/body.docx", "# Record\n\n## 1. Receipts\n\n1.1 The Supplier shall keep the violet inventory receipt.\n\n1.2 Unrelated wording.")
    const rows = await ensureParagraphRows(leanFile(directory, "d1/body.docx"))
    writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(rows.map(row => ({ ...row, sha256: paragraphSha256(row.text) }))))
    const result = buildChecklistDocumentContext(directory, { item_id: "one", item: item() })
    const hint = hints(result.input).find(row => row.text === "The Supplier shall keep the violet inventory receipt.")
    expect(hint?.locator).toContain("Word list number 1.1")
    expect(result.artifacts[0]!.value.paragraphs.find(row => row.paragraph_text === hint?.text)?.label).toBe("1.1")
    // A full-draft comparison unit shows the same number beside its assigned passage.
    const passage = assignChecklistTargetContext(result, checklistTargetAssignments(directory, item()).assignments[0]!.assignment).input.current_evidence
      .find(row => row.text === "The Supplier shall keep the violet inventory receipt.")
    expect(passage?.locator).toStartWith("Word list number 1.1 (printed before the text, not part of it); native PIDs ")
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("zero lexical hits gives no default passages while exact plan ID resolves", async () => {
  const directory = matter()
  try {
    await save(directory, "d1/body.docx", "# Record\n\nAmber cobalt jasper.")
    const requirement = item("d1", "Ultraviolet zirconium."); requirement.source_refs[0]!.quote = "Ultraviolet zirconium."; requirement.qualifications = []
    const result = buildChecklistDocumentContext(directory, { item_id: "empty", item: requirement })
    expect(hints(result.input)).toEqual([])
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
    expect(hints(result.input).length).toBeLessThanOrEqual(6)
    const target = result.artifacts.find(row => row.value.id === "d1")!
    const repeated = target.value.paragraphs.filter(p => p.paragraph_text === "Repeated violet receipt.")
    expect(repeated).toHaveLength(2)
    expect(repeated[0]!.pids).not.toEqual(repeated[1]!.pids)
    expect(target.value.paragraphs.map(p => p.paragraph_text)).toContain("violet")
    for (const artifact of result.artifacts) {
      const file = path.join(directory, artifact.path); mkdirSync(path.dirname(file), { recursive: true }); writeFileSync(file, JSON.stringify(artifact.value))
    }
    for (const hint of hints(result.input)) {
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
    expect(hints(result.input).map(row => row.text)).toEqual(["The violet inventory receipt is pending confirmation."])
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
    expect(hints(notes.input)).toHaveLength(1)
    expect(notes.input.current_evidence[0]!.document).toBe("review-notes.docx")
    expect(plainRows(notes.delivery!).map(row => [row.kind, row.planned_filename])).toEqual([["drafting_notes", "review-notes.docx"]])
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
    expect(hints(corrupt.input)).toEqual([])
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
    expect(hints(result.input).every(row => row.document === unchanged.target.document)).toBe(true)
    expect(JSON.stringify(unchanged)).toBe(before)
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("the original-source inventory carries a child's provenance as its own field", async () => {
  const directory = matter()
  try {
    const documents = ["Plain.docx", "Schedule.docx"].map(document => ({ document, characters: 10, chunks: [{ heading_path: ["Part 1"] }] }))
    writeFileSync(path.join(directory, "source-map.json"), JSON.stringify({ documents }))
    writeFileSync(path.join(directory, "source-index.json"), JSON.stringify({ sources: [{ document: "Thread.eml" }, { document: "Plain.docx" }, { document: "Schedule.docx", attachment_of: "Thread.eml" }] }))
    await save(directory, "d1/body.docx", "# Record\n\nThe violet inventory receipt is pending confirmation.")
    const result = buildChecklistDocumentContext(directory, { item_id: "original-discovery", item: item() })
    const marker = "Original source inventory (bounded navigation metadata, no source text): "
    const inventory = JSON.parse(result.input.retrieval.split(marker)[1]!.split("\n\n")[0]!)
    const row = (name: string) => inventory.find((entry: { document: string }) => entry.document === name)
    expect(row("Schedule.docx").provenance).toBe("attachment of Thread.eml")
    expect("provenance" in row("Plain.docx")).toBe(false)
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

test("a combined file's composition is code-built artifact evidence for its body and notes targets; nothing else changes", async () => {
  const planned = (deliverable?: "combined") => [
    { title: "Record", purpose: "Requested record", file: "record.docx", ...(deliverable ? { deliverable } : {}), parts: [{ heading: "Record", scope: "Record contents" }] },
    { title: "Explanation", purpose: "Requested explanation", file: "explanation.docx", parts: [{ heading: "Explanation", scope: "Explanation contents" }] },
  ]
  const memo = { name: "cover memo", file: memoNotesFile(planned("combined"), "cover memo") }
  const requirement = (document: string) => item(document, "Deliver one file with the cover memo first, then the record.")
  const targets = ["d1", "record.docx", "notes", "record - cover memo.docx", "d2"]
  const directory = mkdtempSync(path.join(tmpdir(), "checklist-context-"))
  try {
    expect(memo.file).toBe("record - cover memo.docx")
    // The same memo and notes file, with the record delivered on its own.
    writeLeanPlan(directory, planned(), undefined, undefined, memo)
    await save(directory, "d1/body.docx", "# Record\n\nThe violet inventory receipt is pending.")
    await save(directory, "d2/body.docx", "The violet inventory receipt is explained here.")
    await save(directory, "drafting-notes.docx", "# Cover memo\n\nThe violet inventory receipt is summarised.")
    const contexts = () => targets.map(target => buildChecklistDocumentContext(directory, { item_id: "package", item: requirement(target) }))
    const separate = contexts()
    writeLeanPlan(directory, planned("combined"), undefined, undefined, memo)
    const combined = contexts()
    for (const [index, target] of targets.entries()) {
      const before = separate[index]!, after = combined[index]!
      // Delivered on its own, a target has only its own files' rows.
      expect(before.delivery!.length).toBeGreaterThan(0)
      expect(before.delivery!.some(row => row.kind === "combined")).toBe(false)
      if (target === "d2") {
        // A document outside the combined file is unchanged, byte for byte.
        expect(after.delivery).toEqual(before.delivery!)
        expect(buildChecklistCheckerPrompt(after.input)).toBe(buildChecklistCheckerPrompt(before.input))
        continue
      }
      const supplied = after.input.current_evidence.at(-1)!
      expect(supplied).toMatchObject({ document: "record.docx", kind: "artifact" })
      expect(supplied.locator).toBe("delivery_artifacts row for the one delivered file that contains this target, from the current plan with registration's naming logic")
      expect(JSON.parse(supplied.text)).toEqual(after.delivery!.at(-1))
      expect(after.delivery!.at(-1)).toMatchObject({ kind: "combined", document: "d1", download_filename: "record.docx", contains: [
        { order: 1, kind: "drafting_notes", planned_filename: "record - cover memo.docx" },
        { order: 2, kind: "body", document: "d1", planned_filename: "record - clean.docx" },
      ], built_at_check: { not_built: "The body part is not yet written at check time." } })
      // The combined row is added after the target's own files; passages are byte-identical.
      expect(after.delivery!.slice(0, -1).map(row => [row.kind, row.document])).toEqual(before.delivery!.map(row => [row.kind, row.document]))
      expect(buildChecklistCheckerPrompt({ ...after.input, current_evidence: hints(after.input) })).toBe(buildChecklistCheckerPrompt({ ...before.input, current_evidence: hints(before.input) }))
      expect(checklistCheckerInputSchema.safeParse(after.input).success).toBe(true)
    }
    // The whole-target properties assignment receives the rows; a passage assignment does not.
    const target = combined[0]!
    const properties = assignChecklistTargetContext(target, { kind: "document_properties", ranges: [] })
    const rows = target.delivery!.length
    expect(properties.input.current_evidence.slice(-rows)).toEqual(target.input.current_evidence.slice(-rows))
    expect(properties.input.current_evidence.slice(0, -rows).every(row => row.kind === "artifact" && row.document.startsWith(".dochaus/lean/"))).toBe(true)
    expect(properties.delivery).toEqual(target.delivery!)
    const passages = assignChecklistTargetContext(target, checklistTargetAssignments(directory, requirement("d1")).assignments[0]!.assignment)
    expect(passages.input.current_evidence.every(row => row.kind === "body")).toBe(true)
    expect(passages).not.toHaveProperty("delivery")
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("a target delivered as a plain file receives the delivery rows of its own files, never another target's", async () => {
  const directory = matter()
  const named = (target: string) => item(target, "Save the record as record.docx.")
  const delivered = (target: string) => buildChecklistDocumentContext(directory, { item_id: "name", item: named(target) })
  const supplied = (context: ReturnType<typeof delivered>) => context.input.current_evidence.filter(row => row.locator.startsWith("delivery_artifacts"))
  try {
    await save(directory, "d1/body.docx", "# Record\n\nThe violet inventory receipt is pending.")
    await save(directory, "d2/body.docx", "The violet inventory receipt is explained here.")
    await save(directory, "drafting-notes.docx", "# Notes\n\nThe violet inventory receipt is summarised.")
    // A body target: its clean file under the requested name, and its redline.
    for (const target of ["record.docx", "d1"]) {
      const context = delivered(target)
      expect(plainRows(context.delivery!).map(row => [row.kind, row.document, row.planned_filename])).toEqual([["body", "d1", "record.docx"], ["redline", "d1", "record - redline.docx"]])
      expect(supplied(context).map(row => ({ ...row, text: JSON.parse(row.text) }))).toEqual(plainRows(context.delivery!).map(row => ({ document: row.planned_filename,
        locator: `delivery_artifacts row for a file that delivers this target (kind ${row.kind}), from the current plan with registration's naming logic`, text: row, kind: "artifact" })))
      // Passage hints stay first and unchanged; another document's files never appear.
      expect(context.input.current_evidence.slice(0, -2).every(row => row.kind === "body" && row.document === target)).toBe(true)
      expect(context.input.current_evidence.some(row => /explanation|"d2"/.test(row.document + row.text))).toBe(false)
      expect(checklistCheckerInputSchema.safeParse(context.input).success).toBe(true)
    }
    // Each row states what code sets and what it does not cover; a file not yet written says so.
    expect(delivered("d1").delivery![0]).toEqual({ kind: "body", document: "d1", path: ".dochaus/lean/d1/body-applied.docx", planned_filename: "record.docx",
      file_status: "not yet written", delivery_registration: "pending; code registers this file once the final gate passes", code_sets: PLAIN_CODE_SETS, not_covered: NOT_COVERED })
    // The checker prompt defines the rows once when they are supplied, and only then.
    const prompt = buildChecklistCheckerPrompt(delivered("d1").input)
    expect(prompt.split(DELIVERY_SENTENCE)).toHaveLength(2)
    expect(buildChecklistCheckerPrompt({ ...delivered("d1").input, current_evidence: hints(delivered("d1").input) })).not.toContain("delivery_artifacts")
    // The notes target: only the notes file.
    const notes = delivered("notes")
    expect(notes.delivery![0]).toMatchObject({ path: ".dochaus/lean/drafting-notes.docx", file_status: "present" })
    expect(plainRows(notes.delivery!).map(row => [row.kind, row.document, row.planned_filename])).toEqual([["drafting_notes", null, "record - drafting notes.docx"]])
    expect(supplied(notes).map(row => row.document)).toEqual(["record - drafting notes.docx"])
    // Another document: only its own files.
    expect(plainRows(delivered("explanation.docx").delivery!).map(row => row.planned_filename)).toEqual(["explanation.docx", "explanation - redline.docx"])
    // A redline deliverable carries the requested name; both of the target's files are rows.
    writeLeanPlan(directory, [
      { title: "Record", purpose: "Requested record", file: "record.docx", deliverable: "redline", parts: [{ heading: "Record", scope: "Record contents" }] },
      { title: "Explanation", purpose: "Requested explanation", file: "explanation.docx", parts: [{ heading: "Explanation", scope: "Explanation contents" }] },
    ])
    const redline = delivered("record.docx")
    expect(plainRows(redline.delivery!).map(row => [row.kind, row.planned_filename])).toEqual([["body", "record - clean.docx"], ["redline", "record.docx"]])
    // The whole-target properties unit receives the rows; a passage unit does not.
    const properties = assignChecklistTargetContext(redline, { kind: "document_properties", ranges: [] })
    expect(properties.delivery).toEqual(redline.delivery!)
    expect(supplied(properties)).toEqual(supplied(redline))
    expect(properties.input.current_evidence.slice(0, -2).every(row => row.kind === "artifact" && row.document.startsWith(".dochaus/lean/"))).toBe(true)
    const passages = assignChecklistTargetContext(redline, checklistTargetAssignments(directory, named("record.docx")).assignments[0]!.assignment)
    expect(passages.input.current_evidence.every(row => row.kind === "body")).toBe(true)
    expect(passages).not.toHaveProperty("delivery")
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

/** The record delivered in one file with its cover memo, beside a plain explanation. */
function combinedMatter() {
  const directory = mkdtempSync(path.join(tmpdir(), "checklist-context-"))
  const documents = [
    { title: "Record", purpose: "Requested record", file: "record.docx", deliverable: "combined" as const, parts: [{ heading: "Record", scope: "Record contents" }] },
    { title: "Explanation", purpose: "Requested explanation", file: "explanation.docx", parts: [{ heading: "Explanation", scope: "Explanation contents" }] },
  ]
  writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo", file: memoNotesFile(documents, "cover memo") })
  return directory
}
const packaged = (directory: string, target = "record.docx") => buildChecklistDocumentContext(directory, { item_id: "package", item: item(target, "Deliver one file named record.docx, the cover memo first, then the record.") })
/** The combined row exactly as the checker receives it. */
const combinedRow = (context: ReturnType<typeof packaged>) => JSON.parse(context.input.current_evidence.find(row => row.locator.startsWith("delivery_artifacts row for the one delivered file"))!.text)
const docxEdit = (bytes: Uint8Array, change: (xml: string) => string) => {
  const parts = unzipSync(bytes)
  parts["word/document.xml"] = strToU8(change(strFromU8(parts["word/document.xml"]!)))
  return zipSync(parts)
}

test("a combined file's row carries the exact properties of the file built from its current parts, built once per run and never written", async () => {
  const directory = combinedMatter()
  try {
    const memo = await save(directory, "drafting-notes.docx", "# Cover memo\n\nThe violet inventory receipt is summarised.\n\n- Confirm the receipt date.")
    const body = await save(directory, "d1/body-applied.docx", "# Record\n\n## Record\n\nThe violet inventory receipt is pending.")
    await save(directory, "d2/body.docx", "The violet inventory receipt is explained here.")
    const before = files(directory)
    const context = packaged(directory)
    expect(files(directory)).toEqual(before)
    const built = combineMemoAndBody({ memo, body })
    expect(texts(built)).toEqual([...texts(memo), ...texts(body)])
    expect(combinedRow(context)).toEqual({
      kind: "combined", document: "d1", download_filename: "record.docx", content_type: DOCX, files_delivered: 1,
      contains: [
        { order: 1, kind: "drafting_notes", role: "the requested memo", path: ".dochaus/lean/drafting-notes.docx", planned_filename: "record - cover memo.docx", file_status: "present", sha256: sha(memo), nonempty_paragraphs: texts(memo).length },
        { order: 2, kind: "body", document: "d1", role: "the clean body", path: ".dochaus/lean/d1/body-applied.docx", planned_filename: "record - clean.docx", file_status: "present", sha256: sha(body), nonempty_paragraphs: texts(body).length },
      ],
      built_at_check: { sha256: sha(built), nonempty_paragraphs: texts(built).length, paragraphs_equal_memo_then_body: true },
      code_sets: COMBINED_CODE_SETS, not_covered: NOT_COVERED,
    })
    // The name is the one the download gives the file.
    expect(combinedRow(context).download_filename).toBe(`${combinedDelivery(directory)!.name}.docx`)
    expect(context.input.current_evidence.at(-1)).toMatchObject({ document: "record.docx", kind: "artifact" })
    expect(buildChecklistCheckerPrompt(context.input)).toContain(DELIVERY_SENTENCE)
    expect(checklistCheckerInputSchema.safeParse(context.input).success).toBe(true)
    // The notes target receives the same row.
    expect(combinedRow(packaged(directory, "notes"))).toEqual(combinedRow(context))
    // A stage outside the checklist receives the row without a build.
    const unbuilt = deliveryArtifactEvidence(directory).at(-1)!
    expect(unbuilt).not.toHaveProperty("built_at_check")
    expect(JSON.stringify(unbuilt)).not.toContain(sha(memo))
    // One check run builds the file once; outside a run every build is fresh.
    const [first, second] = withSourceSnapshot(() => [packaged(directory), packaged(directory, "notes")])
    expect(first!.delivery!.at(-1)).toBe(second!.delivery!.at(-1))
    expect(packaged(directory).delivery!.at(-1)).not.toBe(packaged(directory).delivery!.at(-1))
    expect(files(directory)).toEqual(before)
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("a combined file the build refuses carries the refusal code and its parts' identities", async () => {
  const directory = combinedMatter()
  try {
    const memo = await save(directory, "drafting-notes.docx", "# Cover memo\n\nThe violet inventory receipt is summarised.")
    const body = await save(directory, "d1/body-applied.docx", "# Record\n\nThe violet inventory receipt is pending.")
    // A note reference: its text lives in a part the combined file does not carry.
    const noted = docxEdit(memo, xml => xml.replace("</w:t></w:r>", `</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`))
    expect(() => combineMemoAndBody({ memo: noted, body })).toThrow("unsupported_memo_dependency")
    writeLeanBytes(directory, "drafting-notes.docx", noted)
    const refused = combinedRow(packaged(directory))
    expect(refused.built_at_check).toEqual({ refused: "unsupported_memo_dependency" })
    expect(refused.contains.map((part: { sha256: string }) => part.sha256)).toEqual([sha(noted), sha(body)])
    // A memo that is not well-formed Word XML: the build's own code, and no paragraph count for it.
    const broken = docxEdit(memo, xml => xml.replace("<w:body>", "<w:body><w:p"))
    expect(() => combineMemoAndBody({ memo: broken, body })).toThrow("invalid_docx_xml:")
    writeLeanBytes(directory, "drafting-notes.docx", broken)
    const unreadable = combinedRow(packaged(directory))
    expect(unreadable.built_at_check).toEqual({ refused: "invalid_docx_xml" })
    expect(unreadable.contains[0]).toMatchObject({ file_status: "present", sha256: sha(broken) })
    expect(unreadable.contains[0]).not.toHaveProperty("nonempty_paragraphs")
    expect(unreadable.contains[1]).toMatchObject({ sha256: sha(body), nonempty_paragraphs: texts(body).length })
  } finally { rmSync(directory, { recursive: true, force: true }) }
})

test("a combined file with a part not yet written at check time says exactly that and is not built", async () => {
  const directory = combinedMatter()
  try {
    await save(directory, "d1/body.docx", "# Record\n\nThe violet inventory receipt is pending.")
    const neither = combinedRow(packaged(directory))
    expect(neither.built_at_check).toEqual({ not_built: "The memo and body parts are not yet written at check time." })
    expect(neither.contains.map((part: { file_status: string }) => part.file_status)).toEqual(["not yet written", "not yet written"])
    const memo = await save(directory, "drafting-notes.docx", "# Cover memo\n\nThe violet inventory receipt is summarised.")
    const body = combinedRow(packaged(directory))
    expect(body.built_at_check).toEqual({ not_built: "The body part is not yet written at check time." })
    expect(body.contains[0]).toMatchObject({ file_status: "present", sha256: sha(memo) })
    expect(body.contains[1]).toEqual({ order: 2, kind: "body", document: "d1", role: "the clean body", path: ".dochaus/lean/d1/body-applied.docx", planned_filename: "record - clean.docx", file_status: "not yet written" })
    rmSync(leanFile(directory, "drafting-notes.docx"))
    await save(directory, "d1/body-applied.docx", "# Record\n\nThe violet inventory receipt is pending.")
    expect(combinedRow(packaged(directory)).built_at_check).toEqual({ not_built: "The memo part is not yet written at check time." })
  } finally { rmSync(directory, { recursive: true, force: true }) }
})
