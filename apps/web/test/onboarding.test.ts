import { expect, test } from "bun:test"
import { defaultProvider, pickForProvider } from "../src/models"
import { visibleFields } from "../src/components/settings/NativeProviderConnect"

test("model selection asks when several providers or models are available", () => {
  const a = { value: "one/a", label: "A", providerId: "one", modelId: "a" }
  const b = { value: "two/b", label: "B", providerId: "two", modelId: "b" }
  expect(defaultProvider([a, b])).toBe("")
  expect(defaultProvider([b])).toBe("two")
  expect(pickForProvider([a, b], "two")).toEqual({ primary: "two/b", fast: "two/b" })
  expect(pickForProvider([a, { ...a, value: "one/c", modelId: "c" }], "one")).toEqual({ primary: "", fast: "" })
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
