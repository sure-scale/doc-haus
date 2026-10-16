import { afterAll, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readFirmDraftingPreferences } from "../../../dochaus/lib/drafting-preferences"

const previous = process.env.WORKSPACE_ROOT
const root = mkdtempSync(path.join(tmpdir(), "dochaus-research-preferences-"))
process.env.WORKSPACE_ROOT = root
const { DEFAULT_DRAFTING, readDraftingPreferences, writeDraftingPreferences } = await import("./preferences")
afterAll(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  rmSync(root, { recursive: true, force: true })
})

test("new installations default to the entire web without creating preferences", () => {
  expect(readDraftingPreferences().webResearch).toBe("open")
  expect(readdirSync(root)).toEqual([])
})

test("CourtListener is off for new and existing preferences", () => {
  expect(readDraftingPreferences().courtlistenerMcpEnabled).toBe(false)
  writeDraftingPreferences({ ...DEFAULT_DRAFTING, caseLawApiKey: "synthetic-token" })
  const file = path.join(root, ".preferences", "preferences.json")
  const old = JSON.parse(readFileSync(file, "utf8"))
  delete old.courtlistenerMcpEnabled
  writeFileSync(file, JSON.stringify(old))
  expect(readDraftingPreferences().courtlistenerMcpEnabled).toBe(false)
  expect(readDraftingPreferences().caseLawApiKey).toBe("synthetic-token")
})

test("explicit opt-in round-trips without exposing credentials in instructions", () => {
  const saved = writeDraftingPreferences({ ...DEFAULT_DRAFTING, courtlistenerMcpEnabled: true, caseLawApiKey: "synthetic-token" })
  expect(saved.courtlistenerMcpEnabled).toBe(true)
  expect(readDraftingPreferences()).toEqual(saved)
  const instructions = readFileSync(path.join(root, ".preferences", "drafting.md"), "utf8")
  expect(instructions).not.toContain("synthetic-token")
  expect(instructions).not.toContain("courtlistenerMcpEnabled")
  expect(readFirmDraftingPreferences()).toBe(instructions)
  expect(readdirSync(path.join(root, ".preferences")).sort()).toEqual(["drafting.md", "preferences.json"])
})

test("writer normalizes invalid drafting values before saving canonical JSON and instructions", () => {
  const saved = writeDraftingPreferences({
    ...DEFAULT_DRAFTING,
    attorney: null, firm: 42, houseStyle: null,
    posture: "unsupported", formality: false, detail: {}, dateFormat: "unsupported", numberStyle: [],
    searchApiKey: "synthetic-search-secret", caseLawApiKey: "synthetic-case-secret",
    webResearch: "open", approvedSources: ["trusted.test"], courtlistenerMcpEnabled: true,
    process: "legacy-value", customOption: { retained: true },
    telemetry: { retention_days: 21, retention_days_source: "user", rate_overrides: { "synthetic-model": { input_per_mtok: 2 } } },
  })
  expect(saved).toMatchObject({
    attorney: "", firm: "", houseStyle: "", posture: "balanced", formality: "formal", detail: "concise", dateFormat: "month-day-year", numberStyle: "numerals",
    searchApiKey: "synthetic-search-secret", caseLawApiKey: "synthetic-case-secret",
    webResearch: "open", approvedSources: ["trusted.test"], courtlistenerMcpEnabled: true,
    process: "legacy-value", customOption: { retained: true },
    telemetry: { retention_days: 21, retention_days_source: "user", rate_overrides: { "synthetic-model": { input_per_mtok: 2 } } },
  })
  expect(readDraftingPreferences()).toEqual(saved)
  expect(JSON.parse(readFileSync(path.join(root, ".preferences/preferences.json"), "utf8"))).toEqual(saved)
  const instructions = readFileSync(path.join(root, ".preferences/drafting.md"), "utf8")
  expect(readFirmDraftingPreferences()).toBe(instructions)
  expect(instructions).toContain("Use formal legal drafting style throughout.")
  for (const secret of ["undefined", "synthetic-search-secret", "synthetic-case-secret", "synthetic-model", "customOption", "legacy-value"]) expect(instructions).not.toContain(secret)
})

test("an actual markdown export failure leaves the previous authoritative JSON and no temporary files", () => {
  const saved = writeDraftingPreferences({ ...DEFAULT_DRAFTING, houseStyle: "Keep the previous saved style." })
  const directory = path.join(root, ".preferences")
  const json = path.join(directory, "preferences.json")
  const markdown = path.join(directory, "drafting.md")
  const previousJson = readFileSync(json, "utf8")
  const previousMarkdown = readFileSync(markdown, "utf8")
  rmSync(markdown)
  mkdirSync(markdown)
  try {
    expect(() => writeDraftingPreferences({ ...saved, houseStyle: "A new style that cannot be exported." })).toThrow()
    expect(readFileSync(json, "utf8")).toBe(previousJson)
    expect(readDraftingPreferences()).toEqual(saved)
    expect(readFirmDraftingPreferences()).toBe(previousMarkdown)
    expect(readdirSync(directory).sort()).toEqual(["drafting.md", "preferences.json"])
  } finally {
    rmSync(markdown, { recursive: true })
    writeFileSync(markdown, previousMarkdown)
  }
})

test("research-only patches preserve unrelated drafting and operational preferences", () => {
  const baseline = writeDraftingPreferences({ ...DEFAULT_DRAFTING, attorney: "Test lawyer", houseStyle: "Preserve this style.", telemetry: { retention_days: 17, rate_overrides: {} } })
  const saved = writeDraftingPreferences({ webResearch: "open", approvedSources: ["Trusted.Example"], searchApiKey: " synthetic-key " })
  expect(saved).toMatchObject({ attorney: baseline.attorney, houseStyle: baseline.houseStyle, telemetry: baseline.telemetry,
    webResearch: "open", approvedSources: ["trusted.example"], searchApiKey: "synthetic-key" })
  const instructions = readFirmDraftingPreferences()
  expect(instructions).not.toContain("synthetic-key")
  expect(instructions).not.toContain("trusted.example")
})

test("invalid research saves reject without changing the saved consent or any preference", () => {
  const baseline = writeDraftingPreferences({ ...DEFAULT_DRAFTING, attorney: "Preserved lawyer" })
  for (const patch of [null, [], { webResearch: "all" }, { courtlistenerMcpEnabled: "true" }, { searchApiKey: "token\nheader" },
    { caseLawApiKey: {} }, { approvedSources: ["https://trusted.example/path"] }, { approvedSources: ["*.example"] }, { approvedSources: ["example.test:443"] }]) {
    expect(() => writeDraftingPreferences(patch)).toThrow()
    expect(readDraftingPreferences()).toEqual(baseline)
  }
})

test("malformed saved credentials are never echoed and cannot be overwritten by a partial save", () => {
  const file = path.join(root, ".preferences", "preferences.json")
  const previous = readFileSync(file, "utf8")
  try {
    for (const malformed of ['{"caseLawApiKey":"synthetic-secret", malformed', "null", "[]", "42", '"synthetic-secret"']) {
      writeFileSync(file, malformed)
      for (const read of [readDraftingPreferences, () => writeDraftingPreferences({ webResearch: "official" })]) {
        try { read(); throw new Error("Expected unreadable preferences to reject") } catch (error) {
          expect(String(error)).not.toContain("synthetic-secret")
          expect(String(error)).toContain("Saved preferences could not be read")
        }
      }
      expect(readFileSync(file, "utf8")).toBe(malformed)
    }
  } finally {
    writeFileSync(file, previous)
  }
})
