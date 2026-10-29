import { expect, test } from "bun:test"
import { renderToStaticMarkup } from "react-dom/server"
import { checkDefinition, LiveComponent } from "../src/components/dashboard/visual/LiveComponents"
import { createDashboardSource } from "../src/components/dashboard/visual/source"

test("inherited source names produce a disconnected view rather than a runtime exception", () => {
  const html = renderToStaticMarkup(<LiveComponent definition={{ id: "missing", kind: "chart", props: { form: "bars", label: "Query" }, bindings: { points: { source: "constructor" } } }} sources={{}} />)
  expect(html).toContain("not connected")
  expect(html).toContain('role="alert"')
})

test("JSON frame enums must be literal strings and pointers well formed", () => {
  const frame = { title: "Card", tone: ["bold"], size: "wide" }
  expect(checkDefinition({ id: "test", kind: "stat", frame })).toBe(false)
  expect(checkDefinition({ id: "test", kind: "stat", frame: { title: "Card", size: ["full"] } })).toBe(false)
  expect(checkDefinition({ id: "test", kind: "stat", bindings: { value: { source: "live", path: "/bad~2path" } } })).toBe(false)
  expect(checkDefinition({ id: "test", kind: "stat", frame: { title: "Card", tone: "bold", size: "wide" } })).toBe(true)
})

test("an open JSON composer requires a closing action handler", () => {
  const html = renderToStaticMarkup(<LiveComponent definition={{ id: "composer", kind: "composer", props: { open: true, request: "" } }} />)
  expect(html).toContain("Connect the composer actions")
  expect(html).not.toContain("<dialog")
})

test("live values bind without becoming executable markup and failures hide stale figures", () => {
  const source = createDashboardSource({ value: "$12.00", caption: "<img src=x onerror=alert(1)>" })
  const definition = { id: "stat", kind: "stat" as const, bindings: { value: { source: "live", path: "/value" }, caption: { source: "live", path: "/caption" } } }
  const render = () => renderToStaticMarkup(<LiveComponent definition={definition} sources={{ live: source }} />)
  expect(render()).toContain("$12.00")
  expect(render()).toContain("&lt;img")
  expect(render()).not.toContain("<img")
  source.fail()
  expect(render()).toContain("Last updated")
  expect(render()).not.toContain("$12.00")
  source.publish({ value: "$18.00", caption: "Updated" })
  expect(render()).toContain("$18.00")
})
