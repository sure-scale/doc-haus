import { test, expect } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { pathToFileURL } from "node:url"
import { createPluginSources } from "../../packages/plugin/src/source"
import { getRate, computeCost, loadRateCatalogue, resolveRate } from "./pricing"

test("native costs retain explicit zero and keep an absent cache rate absent", () => {
  expect(resolveRate({ cost: { input: 2, output: 6 } })).toEqual({ input_per_mtok: 2, output_per_mtok: 6 })
  expect(resolveRate({ cost: { input: 0, output: 0, cache_read: 0 } })).toEqual({
    input_per_mtok: 0, output_per_mtok: 0, cache_read_per_mtok: 0,
  })
})

test("missing or incomplete native prices remain unknown", () => {
  for (const model of [undefined, {}, { cost: {} }, { cost: { input: 2 } }, { cost: { output: 6 } }]) {
    expect(resolveRate(model)).toBeUndefined()
  }
  expect(getRate("unknown-provider/mystery-model")).toBeUndefined()
  expect(getRate("unknown-provider/gpt-5.4-mini")).toBeUndefined()
})

test("complete overrides price custom models; partial custom overrides do not", () => {
  const custom = "synthetic-provider/custom"
  expect(getRate(custom, { [custom]: { input_per_mtok: 3, output_per_mtok: 9 } })).toEqual({ input_per_mtok: 3, output_per_mtok: 9 })
  expect(getRate(custom, { [custom]: { input_per_mtok: 3 } })).toBeUndefined()
  expect(getRate(custom, { [custom]: { output_per_mtok: 9 } })).toBeUndefined()
  expect(getRate(custom, { [custom]: { input_per_mtok: 0, output_per_mtok: 0 } })).toEqual({ input_per_mtok: 0, output_per_mtok: 0 })
})

test("overrides take precedence while preserving the other native fields", () => {
  const model = { cost: { input: 2, output: 6, cache_read: 0.25 } }
  expect(resolveRate(model, { input_per_mtok: 1 })).toEqual({ input_per_mtok: 1, output_per_mtok: 6, cache_read_per_mtok: 0.25 })
  for (const value of [-1, Number.NaN, Number.POSITIVE_INFINITY]) {
    expect(resolveRate(undefined, { input_per_mtok: value, output_per_mtok: 1 })).toBeUndefined()
    expect(resolveRate(undefined, { input_per_mtok: 1, output_per_mtok: value })).toBeUndefined()
    expect(resolveRate(model, { input_per_mtok: value })).toEqual({ input_per_mtok: 2, output_per_mtok: 6, cache_read_per_mtok: 0.25 })
  }
})

test("context tiers follow native strict threshold selection before overrides", () => {
  const model = { cost: { input: 2, output: 6, tiers: [
    { tier: { type: "context", size: 200_000 }, input: 4, output: 8 },
    { tier: { type: "context", size: 500_000 }, input: 6, output: 10 },
  ] } }
  expect(resolveRate(model, undefined, 200_000)?.input_per_mtok).toBe(2)
  expect(resolveRate(model, undefined, 200_001)?.input_per_mtok).toBe(4)
  expect(resolveRate(model, undefined, 500_001)?.input_per_mtok).toBe(6)
  expect(resolveRate(model, { input_per_mtok: 1 }, 500_001)).toEqual({ input_per_mtok: 1, output_per_mtok: 10 })
  expect(resolveRate({ cost: { input: 2, output: 6, context_over_200k: { input: 4, output: 8 } } }, undefined, 200_001)?.input_per_mtok).toBe(4)
})

test("missing or malformed catalogue files do not crash telemetry", () => {
  const dir = mkdtempSync(join(tmpdir(), "rate-catalogue-"))
  try {
    expect(loadRateCatalogue(join(dir, "missing.txt"))).toEqual({})
    for (const body of ["not-json", "null", "[]"]) {
      writeFileSync(join(dir, "catalogue.txt"), body)
      expect(loadRateCatalogue(join(dir, "catalogue.txt"))).toEqual({})
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
})

test("native plugin source loading resolves the bundled catalogue from an external entrypoint", async () => {
  const dir = mkdtempSync(join(tmpdir(), "native-pricing-plugin-"))
  const sources = createPluginSources(async () => {})
  try {
    const catalogue = JSON.parse(readFileSync(new URL("../../packages/core/src/models-dev/snapshot.txt", import.meta.url), "utf8"))
    const expected = catalogue.openai.models["gpt-5.4-mini"].cost
    writeFileSync(join(dir, "entry.ts"), `export { getRate, PRICING_SOURCE } from ${JSON.stringify(new URL("./pricing.ts", import.meta.url).href)}\n`)
    const loaded = await sources.read(pathToFileURL(join(dir, "entry.ts")).href)
    const module = loaded.module as { getRate: typeof getRate; PRICING_SOURCE: string }
    expect(module.PRICING_SOURCE).toBe("OpenCode bundled model catalogue")
    expect(module.getRate("openai/gpt-5.4-mini")).toEqual({
      input_per_mtok: expected.input, output_per_mtok: expected.output, cache_read_per_mtok: expected.cache_read,
    })
  } finally {
    sources.dispose()
    rmSync(dir, { recursive: true, force: true })
  }
})

test("computeCost scales tokens by the input rate", () => {
  const rate = { input_per_mtok: 2.0, output_per_mtok: 6.0 }
  expect(computeCost(500_000, rate)).toBeCloseTo(1.0, 6)
})

test("computeCost is zero for free-tier rates", () => {
  const free = { input_per_mtok: 0, output_per_mtok: 0 }
  expect(computeCost(1_000_000, free)).toBe(0)
})
