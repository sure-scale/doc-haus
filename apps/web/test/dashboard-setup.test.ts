import { expect, test } from "bun:test"
import type { DraftingPreferences } from "../src/api/ingest/preferences"
import type { SetupProgress } from "../src/api/ingest/setup"
import type { WordSetupStatus } from "../src/api/ingest/word"
import type { DashboardData } from "../src/components/dashboard/data"
import { setupTabProgress } from "../src/components/dashboard/SetupTab"

const setup: SetupProgress = { version: 1, completed: [], skipped: ["details", "research"], next: "done", ready: true }
const preferences = { attorney: "Alex Reed", firm: "Example Practice", practiceAreas: ["Commercial"], searchApiKey: "", houseStyle: "" } as DraftingPreferences
const data = (patch: Partial<DashboardData> = {}): DashboardData => ({ matters: [], ...patch })
const word = (status: WordSetupStatus["status"], available = false) => ({ status, available }) as WordSetupStatus

test("an unavailable or skipped step is not completed merely by onboarding bookkeeping", () => {
  expect(setupTabProgress(data(), null, []).done).toBe(0)
  expect(setupTabProgress(data(), setup, []).done).toBe(1)
  expect(setupTabProgress(data(), { ...setup, ready: false, completed: ["models", "details", "research"] }, []).done).toBe(0)
})

test("practice completion requires saved identity and law while practice areas are optional", () => {
  const practice = data({ preferences })
  expect(setupTabProgress(practice, setup, ["england-wales"])).toEqual({ done: 2, total: 5 })
  expect(setupTabProgress(practice, setup, []).done).toBe(1)
  expect(setupTabProgress(practice, setup, ["  "]).done).toBe(1)
  expect(setupTabProgress(data({ preferences: { ...preferences, practiceAreas: [] } }), setup, ["england-wales"]).done).toBe(2)
  expect(setupTabProgress(data({ preferences: { ...preferences, firm: " " } }), setup, ["england-wales"]).done).toBe(1)
})

test("research completion reflects a saved key, not a skipped research gate", () => {
  expect(setupTabProgress(data({ preferences }), null, []).done).toBe(0)
  expect(setupTabProgress(data({ preferences: { ...preferences, searchApiKey: "synthetic-key" } }), null, []).done).toBe(1)
  expect(setupTabProgress(data({ preferences: { ...preferences, searchApiKey: "  " } }), null, []).done).toBe(0)
})

test("Word requires ready setup or an available administrator-managed connection", () => {
  for (const status of ["unsupported", "not_configured", "needs_repair", "word_missing", "external"] as const) {
    expect(setupTabProgress(data({ word: word(status) }), null, []).done).toBe(0)
  }
  expect(setupTabProgress(data({ word: word("ready", true) }), null, []).done).toBe(1)
  expect(setupTabProgress(data({ word: word("external", true) }), null, []).done).toBe(1)
})

test("drafting completes after an explicit save without requiring house style or templates", () => {
  const styled = { ...preferences, houseStyle: "Use plain English." }
  expect(setupTabProgress(data({ preferences: styled }), null, []).done).toBe(0)
  expect(setupTabProgress(data({ preferences: styled, templates: 0 }), null, []).done).toBe(0)
  expect(setupTabProgress(data({ preferences, templates: 1 }), null, []).done).toBe(0)
  expect(setupTabProgress(data({ preferences: styled, templates: 1 }), null, []).done).toBe(0)
  expect(setupTabProgress(data({ preferences: { ...preferences, draftingConfigured: true } }), null, []).done).toBe(1)
})

test("complete saved configuration yields five, and deleting a prerequisite reduces progress", () => {
  const complete = data({ preferences: { ...preferences, draftingConfigured: true, searchApiKey: "synthetic-key" }, word: word("ready", true), templates: 0 })
  expect(setupTabProgress(complete, setup, ["england-wales"])).toEqual({ done: 5, total: 5 })
  expect(setupTabProgress(complete, setup, [])).toEqual({ done: 4, total: 5 })
})
