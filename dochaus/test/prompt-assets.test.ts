import { expect, test } from "bun:test"
import path from "node:path"
import fixtures from "./fixtures/prompt-assets.json"
import { findQuote } from "../lib/extract"

test("synthetic router and extraction assets resolve to product prompts and exact source text", async () => {
  for (const row of fixtures.router) {
    expect(row.candidates).toContain(row.expected)
    for (const candidate of row.candidates) expect(await Bun.file(path.join(import.meta.dir, "../agent", `${candidate}.md`)).exists()).toBe(true)
  }
  for (const row of fixtures.extract) expect(findQuote(fixtures.document, row.quote)).toBeDefined()
})

test("synthetic citation cases exercise the real matcher without accepting paraphrases", () => {
  for (const row of fixtures.citations) expect(findQuote(fixtures.document, row.quote) !== undefined).toBe(row.verified)
})

test("jurisdiction fixture identifies an explicit existing pack without inferring legal advice", async () => {
  expect(fixtures.jurisdiction.source).toBe("user_selected")
  expect(fixtures.jurisdiction.automaticInference).toBe(false)
  expect(await Bun.file(path.join(import.meta.dir, "../jurisdiction", fixtures.jurisdiction.code, "profile.json")).exists()).toBe(true)
})
