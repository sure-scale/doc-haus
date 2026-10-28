import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { defaultProvider, pickForProvider } from "../src/models"
import { visibleFields } from "../src/components/ai-setup/connection-fields"
import ModelFields from "../src/components/ai-setup/ModelFields"
import ProviderPicker from "../src/components/ai-setup/ProviderPicker"
import { connectionChoices, methodDefaults } from "../src/components/ai-setup/ConnectionEditor"
import type { Provider } from "../src/components/settings/types"
import type { IntegrationInfo } from "@opencode/client"

const noop = () => {}
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
test("reusable sign-ins appear as detected choices without selecting them", () => {
  const provider = (id: string): Provider => ({ id, name: id, activation: "auto", models: {}, nativeAvailable: false, credentialSources: [] })
  const html = renderToStaticMarkup(createElement(ProviderPicker, { providers: [provider("other"), { ...provider("alias"), integrationID: "sign-in" }], provider: "", onChange: noop, reusableProviderIDs: ["sign-in"] }))
  expect(html).toContain('<option value="alias">alias</option>')
  expect(html.indexOf('value="__local_endpoint__"')).toBeLessThan(html.indexOf('value="alias"'))
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
