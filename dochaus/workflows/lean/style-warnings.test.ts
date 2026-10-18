import { afterAll, beforeEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { lintRenderedDraftMarkdown } from "../../lib/render"
import { changePlanPrompt } from "./change-pipeline"
import { finalGate } from "./gate"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { acceptStageReply, paragraphRows, stageOutputValid } from "./stages"
import { leanFile, readLeanText, writeLeanText } from "./state"
import { cleanupMatters, matter } from "./test-support"
import { draftingStyleWarnings } from "./style-warnings"

const workspace = mkdtempSync(path.join(tmpdir(), "lean-style-preferences-"))
const previous = process.env.WORKSPACE_ROOT
process.env.WORKSPACE_ROOT = workspace
const preferences = path.join(workspace, ".preferences", "preferences.json")
beforeEach(() => rmSync(path.dirname(preferences), { recursive: true, force: true }))
afterAll(() => {
  cleanupMatters()
  rmSync(workspace, { recursive: true, force: true })
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
})

function save(input: Record<string, unknown>) {
  mkdirSync(path.dirname(preferences), { recursive: true })
  writeFileSync(preferences, JSON.stringify(input))
}

test("render style lint uses the displayed canonical defaults and saved preferences", () => {
  expect(lintRenderedDraftMarkdown("# Note\n\nDelivery is due in 30 days.")).toEqual([])
  expect(lintRenderedDraftMarkdown("# Note\n\nDelivery is due in thirty (30) days.").join(" ")).toContain("numbers must be plain numerals")
  save({ dateFormat: "iso", numberStyle: "words-and-numerals" })
  const warnings = lintRenderedDraftMarkdown("# Note\n\nDelivery is due in 30 days from June 11, 2026.")
  expect(warnings.join(" ")).toContain("durations must be words then numerals")
  expect(warnings.join(" ")).toContain("dates must be written")
})

test("style advisories reach wording review, planning and final notes without blocking or changing quoted text", async () => {
  save({ dateFormat: "iso", numberStyle: "numerals" })
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
  const sourceWording = 'The source states "Delivery is due in thirty (30) days from June 11, 2026."'
  const body = `# Supply note\n\n## Delivery\n\n${sourceWording}`
  const accepted = await acceptStageReply(directory, "d1.draft", body)
  expect(accepted.issues).toEqual([])
  const rows = await paragraphRows(leanFile(directory, "d1/body.docx"))
  const warned = draftingStyleWarnings(rows)
  expect(warned).toHaveLength(1)
  expect(warned[0].text).toBe(sourceWording)
  for (const prompt of [stagePrompt(directory, "d1.review-wording"), changePlanPrompt(directory, "d1")]) {
    expect(prompt).toContain("<drafting_style_warnings>")
    expect(prompt).toContain(warned[0].pid)
    expect(prompt).toContain("numbers must be plain numerals")
    expect(prompt).toContain("Preserve exact source quotations")
    expect(prompt).toContain("Never add an open-item marker")
    expect(prompt).toContain("the final notes will disclose surviving warnings.")
  }
  expect(stagePrompt(directory, "d1.review-independent")).not.toContain("<drafting_style_warnings>")

  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows))
  const finalize = stagePrompt(directory, "finalize")
  expect(finalize).toContain("actual final body")
  expect(finalize).toContain("numbers must be plain numerals")
  const notes = await acceptStageReply(directory, "finalize", "# Drafting notes\n\nThe supplied wording has been preserved.")
  expect(notes.issues).toEqual([])
  expect(notes.openItems).toEqual([])
  const markdown = readLeanText(directory, "drafting-notes.md")!
  expect(markdown).toContain("## Drafting style advisories")
  expect(markdown).toContain("Supply note: numbers must be plain numerals")
  expect(markdown).toContain("may intentionally retain their format")
  expect(markdown).not.toContain(warned[0].pid)
  const delivered = await paragraphRows(leanFile(directory, "drafting-notes.docx"))
  expect(delivered.map(row => row.text).join("\n")).toContain("numbers must be plain numerals")
  expect(readLeanText(directory, "d1/paragraphs-applied.json")).toBe(JSON.stringify(rows))
  expect(stageOutputValid(directory, "finalize")).toBe(true)
  expect(finalGate(directory).status).toBe("pass")

  const corrected = rows.map(row => ({ ...row, text: row.text === sourceWording ? "Delivery is due in 30 days from 2026-06-11." : row.text }))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(corrected))
  expect(stagePrompt(directory, "finalize")).not.toContain("<drafting_style_warnings>")
  const retainedNote = `# Drafting notes\n\nThe source explanation says "thirty (30) days from June 11, 2026."`
  expect((await acceptStageReply(directory, "finalize", retainedNote)).issues).toEqual([])
  const notesWarning = readLeanText(directory, "drafting-notes.md")!
  expect(notesWarning.startsWith(retainedNote)).toBe(true)
  expect(notesWarning).toContain("Drafting notes: numbers must be plain numerals")
  expect(notesWarning).not.toContain("Supply note: numbers must be plain numerals")
  expect(notesWarning.match(/## Drafting style advisories/g)).toHaveLength(1)
  const retainedWordNote = (await paragraphRows(leanFile(directory, "drafting-notes.docx"))).map(row => row.text).join("\n")
  expect(retainedWordNote).toContain('The source explanation says "thirty (30) days from June 11, 2026."')
  expect(retainedWordNote).toContain("Drafting notes: numbers must be plain numerals")
  expect((await acceptStageReply(directory, "finalize", "# Drafting notes\n\nThe delivery terms use the selected formats.")).issues).toEqual([])
  expect(readLeanText(directory, "drafting-notes.md")).not.toContain("## Drafting style advisories")

  // Plain notes above lint their own quotation. A requested memo is delivered
  // as written: its quotations are not linted and no appendix is added, while
  // the final-body warnings still reach the memo writer.
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows))
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }], undefined, undefined, { name: "cover memo", file: "Supply note - cover memo.docx" })
  const memoPrompt = stagePrompt(directory, "finalize")
  expect(memoPrompt).toContain("numbers must be plain numerals")
  expect(memoPrompt).not.toContain("appends these retained style advisories")
  for (const prompt of [stagePrompt(directory, "d1.review-wording"), changePlanPrompt(directory, "d1")]) {
    expect(prompt).toContain("numbers must be plain numerals")
    expect(prompt).toContain("Retain a deviation when source fidelity or express instructions require it; surviving warnings are not listed in the requested memo.")
    expect(prompt).not.toContain("the final notes will disclose surviving warnings")
  }
  const memo = `# Cover memo\n\nThe source explanation says "thirty (30) days from June 11, 2026."`
  expect((await acceptStageReply(directory, "finalize", memo)).issues).toEqual([])
  expect(readLeanText(directory, "drafting-notes.md")).toBe(`${memo}\n`)
  const memoWord = (await paragraphRows(leanFile(directory, "drafting-notes.docx"))).map(row => row.text).join("\n")
  expect(memoWord).toContain('The source explanation says "thirty (30) days from June 11, 2026."')
  expect(memoWord).not.toContain("per firm preferences")
})
