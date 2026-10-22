import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"
import type { DraftingPreferences } from "../src/api/ingest/preferences"
import { defaultProvider, pickForProvider } from "../src/models"
import DraftingDetailsStep, { detailsPreferences } from "../src/components/onboarding/DraftingDetailsStep"
import { visibleFields } from "../src/components/ai-setup/connection-fields"
import ModelFields from "../src/components/ai-setup/ModelFields"
import ProviderPicker from "../src/components/ai-setup/ProviderPicker"
import { connectionChoices, methodDefaults } from "../src/components/ai-setup/ConnectionEditor"
import type { Provider } from "../src/components/settings/types"
import type { IntegrationInfo } from "@opencode/client"

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
test("shared model fields retain unavailable choices and default fast to the primary model", () => {
  const html = renderToStaticMarkup(createElement(ModelFields, { models: [], primary: "provider/retired", fast: "", onPrimaryChange: noop, onFastChange: noop }))
  expect(html).toContain("provider/retired · unavailable")
  expect(html).toContain('<option value="" selected="">Same as primary</option>')
})
test("reusable sign-ins promote a provider into detected choices without selecting it", () => {
  const provider = (id: string): Provider => ({ id, name: id, activation: "auto", models: {}, nativeAvailable: false, credentialSources: [] })
  const html = renderToStaticMarkup(createElement(ProviderPicker, { providers: [provider("other"), { ...provider("alias"), integrationID: "sign-in" }], provider: "", onChange: noop, reusableProviderIDs: ["sign-in"] }))
  expect(html).toContain('<optgroup label="Detected"><option value="alias">alias</option></optgroup>')
  expect(html).toContain('<option value="" selected="">Select a provider</option>')
})
test("connection choices preserve native methods and hide conflicting host credentials", () => {
  const provider: Provider = { id: "google-vertex", name: "Vertex", activation: "auto", models: {}, nativeAvailable: false, credentialSources: [], host: { type: "google-adc", sources: ["ADC"], options: {} } }
  const integration: IntegrationInfo = { id: "google-vertex", name: "Vertex", methods: [{ type: "key", label: "New API key", form: [{ type: "string", key: "account", default: "personal" }] }], connections: [{ type: "credential", id: "saved", label: "Saved account", method: "key" }] }
  const choices = connectionChoices(provider, integration, [])
  expect(choices.map(choice => choice.value)).toEqual(["credential:saved", "method:0"])
  expect(methodDefaults(choices[1]?.method)).toEqual({ account: "personal" })
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
