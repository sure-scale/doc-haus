import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"
import type { DraftingPreferences } from "../src/api/ingest/preferences"
import { defaultProvider, pickForProvider } from "../src/models"
import DraftingDetailsStep, { detailsPreferences } from "../src/components/onboarding/DraftingDetailsStep"
import { visibleFields } from "../src/components/settings/NativeProviderConnect"
import ModelsTab from "../src/components/settings/ModelsTab"
import ModelPickerStep from "../src/components/onboarding/ModelPickerStep"

type DateFormat = DraftingPreferences["dateFormat"]
const noop = () => {}
const done = async () => {}
// The Details step as setup shows it, and the date format its Dates control has pressed.
function datesShown(selectedJx: string[], saved: DateFormat = "month-day-year", dateFormatChoice: DateFormat | null = null) {
  const html = renderToStaticMarkup(createElement(DraftingDetailsStep, {
    drafting: { ...DEFAULT_DRAFTING, dateFormat: saved }, setDrafting: noop, jurisdictions: [], selectedJx, setSelectedJx: noop,
    dateFormatChoice, setDateFormatChoice: noop, setStep: noop, onClose: done,
  }))
  return html.match(/<button[^>]*aria-pressed="true"[^>]*>(?:<svg.*?<\/svg>)?(June 11, 2026|11 June 2026|2026-06-11)</)?.[1]
}
const saved = (selectedJx: string[], dateFormat: DateFormat = "month-day-year", choice: DateFormat | null = null) =>
  detailsPreferences({ ...DEFAULT_DRAFTING, dateFormat }, selectedJx, choice).dateFormat

test("model selection asks when several providers or models are available", () => {
  const a = { value: "one/a", label: "A", providerId: "one", modelId: "a" }
  const b = { value: "two/b", label: "B", providerId: "two", modelId: "b" }
  expect(defaultProvider([a, b])).toBe("")
  expect(defaultProvider([b])).toBe("two")
  expect(pickForProvider([a, b], "two")).toEqual({ primary: "two/b", fast: "two/b" })
  expect(pickForProvider([a, { ...a, value: "one/c", modelId: "c" }], "one")).toEqual({ primary: "", fast: "" })
})
test("removed saved models are guidance rather than inaccessible model options in Settings and setup", () => {
  const model = "opencode/retired"
  const settings = renderToStaticMarkup(createElement(ModelsTab, {
    models: [], switchTab: noop, provider: "", changeProvider: noop, readyProviders: [], model, setModel: noop,
    providerModels: [], smallModel: model, setSmallModelValue: noop, toast: noop, firstRun: false, onClose: noop,
  }))
  const setup = renderToStaticMarkup(createElement(ModelPickerStep, {
    connectedProviders: [], provider: "", changeProvider: noop, primary: model, setPrimary: noop, providerModels: [],
    fast: model, setFast: noop, setForceConnect: noop, setNotice: noop, onSaved: done,
  }))
  expect(settings).toContain("Choose another provider")
  for (const html of [settings, setup]) {
    expect(html).not.toContain('value="opencode/retired"')
    expect(html).not.toContain("opencode/retired · unavailable")
    expect(html).toMatch(/<button[^>]*disabled=""[^>]*>Save(?: and continue)?/)
  }
})
test("native auth fields retain conditional visibility", () => {
  const fields = [{ key: "kind", type: "string" as const }, { key: "account", type: "string" as const, when: [{ key: "kind", op: "eq" as const, value: "team" }] }, { key: "hidden", type: "string" as const, hidden: true }]
  expect(visibleFields(fields, {}).map(field => field.key)).toEqual(["kind"])
  expect(visibleFields(fields, { kind: "team" }).map(field => field.key)).toEqual(["kind", "account"])
  expect(visibleFields(fields, { kind: "personal" }).map(field => field.key)).toEqual(["kind"])
})

test("hidden native defaults are submitted only while their condition is active", () => {
  const fields = [{ key: "kind", type: "string" as const }, { key: "account", type: "string" as const, hidden: true, when: [{ key: "kind", op: "eq" as const, value: "team" }] }]
  expect(visibleFields(fields, { kind: "team" }, true).map(field => field.key)).toEqual(["kind", "account"])
  expect(visibleFields(fields, { kind: "personal" }, true).map(field => field.key)).toEqual(["kind"])
})

test("setup preselects the date format from the default jurisdictions and saves it", () => {
  // England & Wales: day-month-year, shown and saved over the month-day-year default.
  expect(datesShown(["EW"])).toBe("11 June 2026")
  expect(saved(["EW"])).toBe("day-month-year")
  // Delaware alone: month-day-year, even over a saved day-month-year.
  expect(datesShown(["US-DE"], "day-month-year")).toBe("June 11, 2026")
  expect(saved(["US-DE"], "day-month-year")).toBe("month-day-year")
  // A mix of US and other jurisdictions, or none, leaves the saved setting.
  for (const codes of [["EW", "US-NY"], []]) {
    expect(datesShown(codes, "iso")).toBe("2026-06-11")
    expect(saved(codes, "iso")).toBe("iso")
  }
  // The firm's own pick wins over the suggestion.
  expect(datesShown(["EW"], "month-day-year", "iso")).toBe("2026-06-11")
  expect(saved(["EW"], "month-day-year", "iso")).toBe("iso")
  expect(saved(["US-DE"], "month-day-year", "day-month-year")).toBe("day-month-year")
})

test("Save and finish keeps the other details the firm entered", () => {
  const drafting = { ...DEFAULT_DRAFTING, attorney: "Jane Doe", firm: "Fixture Partners LLP", agreementLegend: "Subject to contract", adviceLegend: "", houseStyle: "kept out" }
  expect(detailsPreferences(drafting, ["EW"], null)).toEqual({ attorney: "Jane Doe", firm: "Fixture Partners LLP", agreementLegend: "Subject to contract", adviceLegend: "", dateFormat: "day-month-year" })
})
