import { Database } from "bun:sqlite"
import { afterEach, beforeEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, utimesSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readFirmDraftingPreferences } from "../../lib/drafting-preferences"
import { matterDateFormat } from "../../lib/matter-client"
import { leanSteps, readLeanPlan, runDateFormat, writeLeanPlan } from "./plan"
import { runLean } from "./runner"
import { stageInputs } from "./stages"
import { AGENT, SOURCE_TEXT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

// The date style a run writes until the firm sets its drafting preferences:
// the style the matter's sources use, frozen into the plan by the outline.
// Real coordinator; only the model is scripted.

const previous = process.env.WORKSPACE_ROOT
let workspace: string
beforeEach(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "lean-date-format-"))
  process.env.WORKSPACE_ROOT = workspace
})
afterEach(() => {
  cleanupMatters()
  rmSync(workspace, { recursive: true, force: true })
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
})

function save(input: Record<string, unknown>) {
  mkdirSync(path.join(workspace, ".preferences"), { recursive: true })
  writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify(input))
}
// The shared fixture's terms, dated the way the sources write dates.
const DAY_FIRST = `${SOURCE_TEXT} The order date is 30 June 2026 and the invoice date is 7 July 2026.`
const SOURCES = 'Write dates in the form "11 June 2026", the style the matter\'s sources use.'
const DEFAULT = 'Write dates in the form "June 11, 2026" (the default), and in one style throughout.'
const rendered = (directory: string, step: Parameters<typeof stageInputs>[1]) =>
  stageInputs(directory, step).find(part => part.name === "drafting-preferences")?.bytes
// A source added to the matter mid-run: the index ingest rewrites legal.db, dated apart.
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
  expect(rendered(directory, "intake")).toBe(readFirmDraftingPreferences("full", undefined, "day-month-year"))
  expect(rendered(directory, "intake")).toContain(SOURCES)
  save({ draftingConfigured: true, dateFormat: "month-day-year" })
  expect(rendered(directory, "intake")).toBe(readFirmDraftingPreferences())
  expect(rendered(directory, "intake")).toContain('Write dates in the form "June 11, 2026".')
  expect(rendered(directory, "intake")).not.toContain("sources use")
})

test("a plan saved before the style was frozen into it takes the default, not the live sources", () => {
  const directory = matter(DAY_FIRST)
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
  expect(matterDateFormat(directory)).toBe("day-month-year")
  expect(runDateFormat(directory)).toBeUndefined()
  for (const step of leanSteps(directory)) {
    const bytes = rendered(directory, step)
    if (bytes !== undefined) expect(bytes).toContain(DEFAULT)
  }
})

test("the outline freezes the sources' style into the plan; a source added to the matter moves the live reading but neither the run's style nor the preferences any stage reads", async () => {
  const directory = matter(DAY_FIRST)
  const model = scriptedModel(happyPath)
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("completed")
  expect(readLeanPlan(directory)?.date_format).toBe("day-month-year")
  expect(model.count(AGENT.review)).toBe(1)
  // Every stage that reads the preferences read the sources' style, not a firm's choice.
  const before = new Map(leanSteps(directory).map(step => [step, rendered(directory, step)]))
  expect([...before.values()].filter(Boolean).length).toBeGreaterThan(0)
  for (const bytes of before.values()) {
    if (bytes === undefined) continue
    expect(bytes).toContain(SOURCES)
    expect(bytes).not.toContain("has set these firm-wide preferences")
  }
  // Sources in another style arrive after the outline: the matter's live reading moves; the run's does not.
  addSource(directory, "Paid June 1, 2026. Paid June 2, 2026. Paid June 3, 2026. Paid June 4, 2026.")
  expect(matterDateFormat(directory)).toBe("month-day-year")
  expect(runDateFormat(directory)).toBe("day-month-year")
  expect(rendered(directory, "d1.review")).toBe(readFirmDraftingPreferences("full", undefined, "day-month-year"))
  for (const step of leanSteps(directory)) expect(rendered(directory, step)).toBe(before.get(step))
}, 120_000)
