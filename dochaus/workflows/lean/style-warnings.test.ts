import { Database } from "bun:sqlite"
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, utimesSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readFirmDraftingPreferences } from "../../lib/drafting-preferences"
import { matterDateFormat } from "../../lib/matter-client"
import { lintRenderedDraftMarkdown } from "../../lib/render"
import { changePlanPrompt } from "./change-pipeline"
import { finalGate } from "./gate"
import { leanSteps, readLeanPlan, runDateFormat, writeLeanPlan } from "./plan"
import { REVISION_PREFERENCE_SCOPE, stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { acceptStageReply, paragraphRows, stageInputs, stageOutputValid } from "./stages"
import { leanFile, readLeanText, writeLeanText } from "./state"
import { AGENT, SOURCE, SOURCE_TEXT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"
import { draftingStyleWarnings, styleScopedRows } from "./style-warnings"

let workspace = ""
let previous: string | undefined
let preferences = ""
beforeAll(() => {
  previous = process.env.WORKSPACE_ROOT
  workspace = mkdtempSync(path.join(tmpdir(), "lean-style-preferences-"))
  process.env.WORKSPACE_ROOT = workspace
  preferences = path.join(workspace, ".preferences", "preferences.json")
})
beforeEach(() => rmSync(path.dirname(preferences), { recursive: true, force: true }))
afterEach(cleanupMatters)
afterAll(() => {
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
  const warned = draftingStyleWarnings(directory, rows)
  expect(warned).toHaveLength(1)
  expect(warned[0].text).toBe(sourceWording)
  for (const prompt of [stagePrompt(directory, "d1.review-wording"), changePlanPrompt(directory, "d1")]) {
    expect(prompt).toContain("<drafting_style_warnings>")
    expect(prompt).toContain(warned[0].pid)
    expect(prompt).toContain("numbers must be plain numerals")
  }
  expect(stagePrompt(directory, "d1.review-independent")).not.toContain("<drafting_style_warnings>")

  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows))
  const finalize = stagePrompt(directory, "finalize")
  expect(finalize).toContain("<drafting_style_warnings>")
  expect(finalize).toContain("numbers must be plain numerals")
  const notes = await acceptStageReply(directory, "finalize", "# Drafting notes\n\nThe supplied wording has been preserved.")
  expect(notes.issues).toEqual([])
  expect(notes.openItems).toEqual([])
  const markdown = readLeanText(directory, "drafting-notes.md")!
  expect(markdown).toContain("## Drafting style advisories")
  expect(markdown).toContain("Supply note: numbers must be plain numerals")
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
})

test("in a revision, a date in the original's unchanged wording raises no style warning; the same date in changed wording still does", async () => {
  // The firm set month-day-year (every save from Settings confirms the preferences); the original writes day-month-year.
  save({ draftingConfigured: true, dateFormat: "month-day-year" })
  // The draft retypes the original's paragraph with plain quotes and single spacing.
  const directory = matter("Supply terms\n\nThe \u201cTerm\u201d starts  on 1 January 2027.\n\nPayment is due on 1 January 2027.")
  const plan = (mode: "new" | "revise") => writeLeanPlan(directory, [{ title: "Supply terms", purpose: "Revise the supply terms", parts: [{ heading: "Terms", scope: "Terms" }], mode, ...(mode === "revise" ? { comparison: SOURCE } : {}) }])
  plan("revise")
  const draft = '# Supply terms\n\n## Terms\n\nThe "Term" starts on 1 January 2027.\n\nPayment is due by 1 January 2027 at the latest.'
  expect((await acceptStageReply(directory, "d1.draft", draft)).issues).toEqual([])
  const rows = await paragraphRows(leanFile(directory, "d1/body.docx"))
  const kept = rows.find(row => row.text.includes("starts on"))!
  const changed = rows.find(row => row.text.startsWith("Payment"))!
  const warned = () => draftingStyleWarnings(directory, styleScopedRows(directory, "d1", rows)).map(row => row.pid)
  const block = (prompt: string) => /<drafting_style_warnings>[\s\S]*<\/drafting_style_warnings>/.exec(prompt)?.[0] ?? ""
  const prompts = () => [stagePrompt(directory, "d1.review-wording"), changePlanPrompt(directory, "d1")]
  // The final body keeps the original's date and corrects the changed one.
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows.map(row => row === changed ? { ...row, text: "Payment is due by January 1, 2027 at the latest." } : row)))
  const advisories = async () => {
    expect((await acceptStageReply(directory, "finalize", "# Drafting notes\n\nThe terms are revised.")).issues).toEqual([])
    return readLeanText(directory, "drafting-notes.md")!.includes("## Drafting style advisories")
  }

  expect(warned()).toEqual([changed.pid])
  for (const prompt of prompts()) {
    expect(block(prompt)).toContain(changed.pid)
    expect(block(prompt)).not.toContain(kept.pid)
  }
  for (const prompt of [...prompts(), stagePrompt(directory, "d1.changes")]) expect(prompt).toContain(REVISION_PREFERENCE_SCOPE)
  expect(await advisories()).toBe(false)

  // A new draft with the same words is checked in full, and is not told the revision's scope.
  plan("new")
  expect(warned()).toEqual([kept.pid, changed.pid])
  for (const prompt of prompts()) expect(block(prompt)).toContain(kept.pid)
  for (const prompt of [...prompts(), stagePrompt(directory, "d1.changes")]) expect(prompt).not.toContain(REVISION_PREFERENCE_SCOPE)
  expect(await advisories()).toBe(true)
})

test("until the firm sets its preferences, the date style the sources use governs the warnings, live before the outline and frozen after it; the firm's own setting governs once set", () => {
  // The firm has saved nothing; the sources write day-month-year dates.
  const directory = matter("Supply terms\n\nThe term starts on 1 January 2027 and ends on 31 December 2027.")
  const documents = [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }]
  const rows = [{ pid: "p1", text: "Delivery is due by January 1, 2027." }, { pid: "p2", text: "Payment is due by 1 February 2027." }]
  const warned = () => draftingStyleWarnings(directory, rows).map(row => [row.pid, row.warnings.join(" ")] as const)
  const flagged = (pid: string, example: string) => expect.arrayContaining([[pid, expect.stringContaining(`dates must be written ${example}`)]])
  expect(warned()).toHaveLength(1)
  expect(warned()).toEqual(flagged("p1", '"11 June 2026"'))
  // The outline froze the style; the plan governs the rest of the run.
  writeLeanPlan(directory, documents, undefined, undefined, undefined, undefined, undefined, "day-month-year")
  expect(warned()).toHaveLength(1)
  expect(warned()).toEqual(flagged("p1", '"11 June 2026"'))
  // The firm's own choice wins over the sources' style, as before.
  save({ draftingConfigured: true, dateFormat: "month-day-year" })
  expect(warned()).toHaveLength(1)
  expect(warned()).toEqual(flagged("p2", '"June 11, 2026"'))
})

describe("date format resolution", () => {
  const DAY_FIRST = `${SOURCE_TEXT} The order date is 30 June 2026 and the invoice date is 7 July 2026.`
  const rendered = (directory: string, step: Parameters<typeof stageInputs>[1]) =>
    stageInputs(directory, step).find(part => part.name === "drafting-preferences")?.bytes
  function addSource(directory: string, text: string) {
    const file = path.join(directory, ".dochaus", "legal.db")
    const db = new Database(file)
    db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("later.docx", "1", text, 0, text.length)
    db.close()
    utimesSync(file, 1_800_000_000, 1_800_000_000)
  }

  test("before the outline, the run takes the style the sources use; the firm's own setting wins once set", () => {
    const directory = matter(DAY_FIRST)
    expect(matterDateFormat(directory)).toBe("day-month-year")
    expect(runDateFormat(directory)).toBe("day-month-year")
    const before = rendered(directory, "intake")
    expect(before).toBe(readFirmDraftingPreferences("full", undefined, "day-month-year"))
    save({ draftingConfigured: true, dateFormat: "month-day-year" })
    expect(rendered(directory, "intake")).toBe(readFirmDraftingPreferences())
    expect(rendered(directory, "intake")).not.toBe(before)
  })

  test("a plan saved before the style was frozen into it takes the default, not the live sources", () => {
    const directory = matter(DAY_FIRST)
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
    expect(matterDateFormat(directory)).toBe("day-month-year")
    expect(runDateFormat(directory)).toBeUndefined()
    const defaultPrefs = new Set<string | Uint8Array>([readFirmDraftingPreferences("full"), readFirmDraftingPreferences("explanation")])
    for (const step of leanSteps(directory)) {
      const bytes = rendered(directory, step)
      if (bytes !== undefined) expect(defaultPrefs.has(bytes)).toBe(true)
    }
    expect(draftingStyleWarnings(directory, [{ pid: "p1", text: "Due 30 June 2026." }])).toHaveLength(1)
  })

  test("the outline freezes the sources' style into the plan; a source added to the matter moves the live reading but neither the run's style nor the preferences any stage reads", async () => {
    const directory = matter(DAY_FIRST)
    const model = scriptedModel(happyPath)
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("completed")
    expect(readLeanPlan(directory)?.date_format).toBe("day-month-year")
    expect(model.count(AGENT.review)).toBe(1)
    const before = new Map(leanSteps(directory).map(step => [step, rendered(directory, step)]))
    expect([...before.values()].filter(Boolean).length).toBeGreaterThan(0)
    const sourcePrefs = new Set<string | Uint8Array>([readFirmDraftingPreferences("full", undefined, "day-month-year"), readFirmDraftingPreferences("explanation", undefined, "day-month-year")])
    for (const bytes of before.values()) {
      if (bytes === undefined) continue
      expect(sourcePrefs.has(bytes)).toBe(true)
    }
    addSource(directory, "Paid June 1, 2026. Paid June 2, 2026. Paid June 3, 2026. Paid June 4, 2026.")
    expect(matterDateFormat(directory)).toBe("month-day-year")
    expect(runDateFormat(directory)).toBe("day-month-year")
    expect(rendered(directory, "d1.review")).toBe(readFirmDraftingPreferences("full", undefined, "day-month-year"))
    for (const step of leanSteps(directory)) expect(rendered(directory, step)).toBe(before.get(step))
  }, 120_000)
})
