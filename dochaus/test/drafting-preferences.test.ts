import { afterAll, beforeEach, expect, test } from "bun:test"
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LegalPlugin } from "../plugin/legal/policy"
import { recordSessionAgent } from "../plugin/legal/session"
import { DEFAULT_DRAFTING, LEGEND_MAX_LENGTH, normalizeLegend, normalizePreferences, renderInstructions } from "../lib/drafting-preferences"
import { writeLeanPlan } from "../workflows/lean/plan"

const previous = process.env.WORKSPACE_ROOT
const workspace = mkdtempSync(path.join(tmpdir(), "drafting-preferences-"))
const directory = path.join(workspace, "matter")
const preferences = path.join(workspace, ".preferences")
mkdirSync(directory)
process.env.WORKSPACE_ROOT = workspace

beforeEach(() => rmSync(preferences, { recursive: true, force: true }))
afterAll(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  rmSync(workspace, { recursive: true, force: true })
})

async function systemFor(agent = "lean-body") {
  const hooks = await LegalPlugin({ directory } as Parameters<typeof LegalPlugin>[0])
  const sessionID = `drafting-preferences-${agent}`
  recordSessionAgent(sessionID, agent)
  const output = { system: [] as string[] }
  await hooks["experimental.chat.system.transform"]!({ sessionID } as never, output)
  return output.system.join("\n")
}

function save(input: Record<string, unknown>) {
  mkdirSync(preferences, { recursive: true })
  writeFileSync(path.join(preferences, "preferences.json"), JSON.stringify(input))
}

test("fresh workspaces apply displayed defaults without creating preference files", async () => {
  const system = await systemFor()
  expect(system).toContain("<drafting_preferences>")
  expect(system).toContain("Use formal legal drafting style throughout.")
  expect(system).toContain("Keep explanations concise")
  expect(system).not.toContain("<plan_first_drafting>")
  expect(existsSync(preferences)).toBe(false)
})

test("a drafting run's stages take the Posture frozen in its plan; other work on the matter takes the matter's, else the firm's", async () => {
  const FAVOR = "Favor the represented person's", CONSERVATIVE = "Draft conservatively", BALANCED = "Use an even-handed, conventional approach"
  save({ posture: "client-favorable" })
  try {
    for (const agent of ["lean-review", "legal-review"]) expect(await systemFor(agent)).toContain(FAVOR)
    writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter", title: "Supply", posture: "conservative" }))
    for (const agent of ["lean-review", "legal-review"]) {
      const system = await systemFor(agent)
      expect(system).toContain(CONSERVATIVE)
      expect(system).not.toContain(FAVOR)
    }
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }], undefined, undefined, undefined, "balanced")
    const stage = await systemFor("lean-review")
    expect(stage).toContain(BALANCED)
    expect(stage).not.toContain(CONSERVATIVE)
    expect(await systemFor("legal-review")).toContain(CONSERVATIVE)
  } finally {
    rmSync(path.join(directory, "matter.json"), { force: true })
    rmSync(path.join(directory, ".dochaus"), { recursive: true, force: true })
  }
})

test("research defaults to the entire web and preserves explicit saved scope", () => {
  expect(DEFAULT_DRAFTING.webResearch).toBe("open")
  expect(normalizePreferences({}).webResearch).toBe("open")
  expect(normalizePreferences({ firm: "Existing firm" }).webResearch).toBe("open")
  for (const webResearch of ["official", "open"] as const) {
    expect(normalizePreferences({ webResearch }).webResearch).toBe(webResearch)
  }
  for (const webResearch of [null, false, 42, [], {}, "unsupported", undefined]) {
    expect(normalizePreferences({ webResearch }).webResearch).toBe("official")
  }
})

test("saved JSON is effective without generated markdown and is reread on the next turn", async () => {
  save({ detail: "detailed", houseStyle: "Use active voice.", process: "standard" })
  const first = await systemFor()
  expect(first).toContain("Explain your reasoning in detail")
  expect(first).toContain("Use active voice.")
  expect(first).not.toContain("<plan_first_drafting>")
  save({ detail: "concise", houseStyle: "Use short headings.", process: "plan-first" })
  const next = await systemFor()
  expect(next).toContain("Keep explanations concise")
  expect(next).toContain("Use short headings.")
  expect(next).not.toContain("Use active voice.")
  expect(next).not.toContain("<plan_first_drafting>")
})

test("stale markdown and operational preferences cannot enter model instructions", async () => {
  save({
    detail: "detailed",
    searchApiKey: "synthetic-search-secret",
    caseLawApiKey: "synthetic-case-law-secret",
    telemetry: { retention_days: 13, rate_overrides: { "synthetic-model": { input_per_mtok: 7 } } },
  })
  writeFileSync(path.join(preferences, "drafting.md"), "Synthetic stale instructions.")
  const system = await systemFor()
  expect(system).toContain("Explain your reasoning in detail")
  expect(system).toContain("Use formal legal drafting style throughout.")
  expect(system).not.toContain("Synthetic stale instructions.")
  expect(system).not.toContain("synthetic-search-secret")
  expect(system).not.toContain("synthetic-case-law-secret")
  expect(system).not.toContain("synthetic-model")
  expect(system).not.toContain("retention_days")
})


test("retired process values do not alter standing drafting instructions", async () => {
  save({ process: "plan-first" })
  const legacyPlan = await systemFor()
  save({ process: "standard" })
  expect(await systemFor()).toBe(legacyPlan)
  expect(legacyPlan).not.toContain("<plan_first_drafting>")
})

test("a retired stored setting is kept as saved and adds no drafting instruction", () => {
  expect(DEFAULT_DRAFTING).not.toHaveProperty("materialDecisionCheckpoints")
  for (const value of ["interactive", "off", "unsupported"]) {
    const saved = normalizePreferences({ materialDecisionCheckpoints: value })
    expect(saved).toHaveProperty("materialDecisionCheckpoints", value)
    expect(renderInstructions(saved)).toBe(renderInstructions(DEFAULT_DRAFTING))
  }
})

test("the quick-draft offer is a stored chat setting: Ask by default, Automatic and Off kept, never an instruction", () => {
  expect(DEFAULT_DRAFTING.quickDraftOffer).toBe("ask")
  for (const saved of ["ask", "auto", "off"]) {
    expect(normalizePreferences({ quickDraftOffer: saved }).quickDraftOffer).toBe(saved)
    expect(renderInstructions(normalizePreferences({ quickDraftOffer: saved }))).toBe(renderInstructions(DEFAULT_DRAFTING))
  }
  for (const invalid of ["always", "quick", "automatic", true, null, 1])
    expect(normalizePreferences({ quickDraftOffer: invalid }).quickDraftOffer).toBe("ask")
})

test("invalid drafting field types and choices fall back to the same valid defaults", () => {
  const baseline = renderInstructions(normalizePreferences(DEFAULT_DRAFTING))
  expect(baseline).toContain("Use formal legal drafting style throughout.")
  for (const invalid of [null, false, 42, [], {}, "unsupported", "toString", "__proto__"]) {
    const preferences = normalizePreferences({
      attorney: typeof invalid === "string" ? "" : invalid,
      firm: typeof invalid === "string" ? "" : invalid,
      houseStyle: typeof invalid === "string" ? "" : invalid,
      posture: invalid, formality: invalid, detail: invalid, dateFormat: invalid, numberStyle: invalid,
    })
    expect(renderInstructions(preferences)).toBe(baseline)
    expect(renderInstructions(preferences, "explanation")).not.toContain("undefined")
  }
  for (const invalid of [null, false, 42, [], "unsupported"]) {
    expect(normalizePreferences(invalid)).toEqual(DEFAULT_DRAFTING)
  }
})

test("normalization retains every supported drafting choice and exact valid text", () => {
  const choices = {
    posture: ["client-favorable", "balanced", "conservative"],
    formality: ["formal", "plain"], detail: ["concise", "detailed"],
    dateFormat: ["month-day-year", "day-month-year", "iso"],
    numberStyle: ["words-and-numerals", "numerals"],
  }
  for (const [field, values] of Object.entries(choices)) for (const value of values) {
    expect(normalizePreferences({ [field]: value })[field as keyof typeof choices]).toBe(value)
  }
  const text = { attorney: "A. Counsel", firm: "Counsel Practice", houseStyle: "  Keep active voice.\nPreserve meaningful qualifications.  " }
  expect(normalizePreferences(text)).toMatchObject(text)
})

test("an existing invalid saved drafting value cannot block the model context hook", async () => {
  save({ houseStyle: null, formality: "unsupported", detail: false, searchApiKey: "synthetic-secret" })
  const system = await systemFor()
  expect(system).toContain("Use formal legal drafting style throughout.")
  expect(system).toContain("Keep explanations concise")
  expect(system).not.toContain("undefined")
  expect(system).not.toContain("synthetic-secret")
})

test("page layout settings default to numbered pages and no legend, and stay out of model instructions", () => {
  const prefs = normalizePreferences({ pageNumbers: "roman", pageNumberPosition: "left", numberFirstPage: "no", draftLine: "no", draftLineFirm: 0, notesMarking: 0, agreementLegend: 7 })
  expect([prefs.pageNumbers, prefs.pageNumberPosition, prefs.numberFirstPage, prefs.draftLine, prefs.draftLineFirm, prefs.notesMarking, prefs.agreementLegend, prefs.adviceLegend]).toEqual(["plain", "center", true, true, true, true, "", ""])
  const set = normalizePreferences({ pageNumbers: "off", pageNumberPosition: "right", numberFirstPage: false, draftLine: false, draftLineFirm: false, notesMarking: false, agreementLegend: "Subject to contract" })
  expect([set.pageNumbers, set.pageNumberPosition, set.numberFirstPage, set.draftLine, set.draftLineFirm, set.notesMarking]).toEqual(["off", "right", false, false, false, false])
  expect(normalizePreferences({ pageNumbers: "dashed" }).pageNumbers).toBe("dashed")
  expect(renderInstructions(set)).toBe(renderInstructions(DEFAULT_DRAFTING))
})

test("a legend is one capped line of the firm's own text, its script's format characters kept", () => {
  expect(normalizeLegend("  Subject\nto\tcontract  ")).toBe("Subject to contract")
  expect(normalizeLegend("x".repeat(LEGEND_MAX_LENGTH + 10))).toHaveLength(LEGEND_MAX_LENGTH)
  expect(normalizeLegend("محرمانه‌است")).toBe("محرمانه‌است")
  expect(normalizeLegend(null)).toBe("")
})
