import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { inspectParserCapabilities } from "../../../services/ingest/src/parser-capabilities"
import { ParserCapabilitiesView } from "../src/components/ParserCapabilities"

test("rendered inventory distinguishes host observations from parser readiness and accelerator capacity", async () => {
  const snapshot = await inspectParserCapabilities({ searchPath: "", diskDirectory: "/not-an-existing-inventory-filesystem",
    system: { platform: () => "darwin", arch: () => "arm64", totalmem: () => 8 * 1024 ** 3, freemem: () => 1024 ** 3 },
  })
  const html = renderToStaticMarkup(createElement(ParserCapabilitiesView, { state: { status: "observed", snapshot } }))
  expect(html).toContain("Observed on the ingestion host")
  expect(html).toContain("Availability does not establish parsing quality or readiness")
  expect(html).toContain("Process architecture</dt><dd>arm64")
  expect(html).toContain("Total system RAM</dt><dd>8 GiB")
  expect(html).toContain("OS-reported free RAM</dt><dd>1 GiB")
  expect(html).toContain("Disk available on ingest filesystem</dt><dd>Unknown")
  expect(html).toContain("Dedicated GPU VRAM</dt><dd>Unknown")
  expect(html).toContain("Unified memory availability</dt><dd>Unknown")
  expect(html).toContain("Not found on ingest PATH")
  expect(html).toContain("uvx is a launcher")
  expect(html).not.toContain("<button")
  expect(html).not.toContain("recommended parser")
  expect(html).not.toContain("Install")
})

test("loading and unavailable views do not invent absent components or zero capacity", () => {
  const loading = renderToStaticMarkup(createElement(ParserCapabilitiesView, { state: { status: "loading" } }))
  const unavailable = renderToStaticMarkup(createElement(ParserCapabilitiesView, { state: { status: "unavailable" } }))
  expect(loading).toContain("Reading ingestion host capabilities")
  expect(unavailable).toContain("Capabilities are unknown")
  for (const html of [loading, unavailable]) {
    expect(html).not.toContain("0 B")
    expect(html).not.toContain("Not found on ingest PATH")
    expect(html).not.toContain("<button")
  }
})

test("unknown memory stays visibly unknown while other observations remain available", async () => {
  const snapshot = await inspectParserCapabilities({ searchPath: "", diskDirectory: "/not-an-existing-inventory-filesystem",
    system: { platform: () => "linux", arch: () => "x64", totalmem: () => NaN, freemem: () => NaN },
  })
  const html = renderToStaticMarkup(createElement(ParserCapabilitiesView, { state: { status: "observed", snapshot } }))
  expect(html).toContain("Process OS</dt><dd>linux")
  expect(html).toContain("Total system RAM</dt><dd>Unknown")
  expect(html).toContain("OS-reported free RAM</dt><dd>Unknown")
  expect(html).not.toContain("0 B")
})
