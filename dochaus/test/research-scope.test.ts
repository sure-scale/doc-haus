import { afterAll, beforeEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { normalizePreferences } from "../lib/drafting-preferences"
import { readFirmPreferences } from "../lib/preferences"
import { allowRedirect, researchScope } from "../lib/research"
import { LegalPlugin } from "../plugin/legal/policy"

const previous = process.env.WORKSPACE_ROOT
const workspace = mkdtempSync(path.join(tmpdir(), "research-scope-"))
const directory = path.join(workspace, "matter")
const preferences = path.join(workspace, ".preferences")
const file = path.join(preferences, "preferences.json")
mkdirSync(directory)
process.env.WORKSPACE_ROOT = workspace

beforeEach(() => rmSync(preferences, { recursive: true, force: true }))
afterAll(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  rmSync(workspace, { recursive: true, force: true })
})

function save(input: Record<string, unknown>) {
  mkdirSync(preferences, { recursive: true })
  writeFileSync(file, JSON.stringify(input))
}

async function prepareFetch() {
  const hooks = await LegalPlugin({ directory } as Parameters<typeof LegalPlugin>[0])
  await hooks["tool.execute.before"]!(
    { tool: "webfetch", sessionID: "research-scope", callID: "fetch" },
    { args: { url: "https://commentary.example/article", format: "markdown" } },
  )
}

test("fresh and legacy workspaces allow entire-web fetches and redirects", async () => {
  expect(researchScope()).toBe("open")
  await expect(prepareFetch()).resolves.toBeUndefined()
  expect(allowRedirect("https://records.gov/source", "https://commentary.example/article")).toBe(true)
  save({ firm: "Existing firm", courtlistenerMcpEnabled: false })
  expect(researchScope()).toBe("open")
  expect(normalizePreferences(readFirmPreferences()).webResearch).toBe("open")
  await expect(prepareFetch()).resolves.toBeUndefined()
})

test("explicit official scope remains restricted and saved changes apply on the next fetch", async () => {
  save({ webResearch: "official" })
  expect(researchScope()).toBe("official")
  await expect(prepareFetch()).rejects.toThrow("webfetch is limited")
  expect(allowRedirect("https://records.gov/source", "https://commentary.example/article")).toBe(false)
  save({ webResearch: "open" })
  expect(researchScope()).toBe("open")
  await expect(prepareFetch()).resolves.toBeUndefined()
  save({ webResearch: "official" })
  await expect(prepareFetch()).rejects.toThrow("webfetch is limited")
})

test("invalid saved scope cannot silently widen fetches or redirects", async () => {
  for (const webResearch of [null, false, 42, [], {}, "unsupported"]) {
    save({ webResearch })
    expect(researchScope()).toBe("official")
    expect(normalizePreferences(readFirmPreferences()).webResearch).toBe("official")
    await expect(prepareFetch()).rejects.toThrow("webfetch is limited")
    expect(allowRedirect("https://records.gov/source", "https://commentary.example/article")).toBe(false)
  }
})

test("malformed, non-object and unreadable saved preferences retain restricted scope and no MCP consent", async () => {
  mkdirSync(preferences, { recursive: true })
  for (const input of ["{malformed", "null", "[]", "false", '"open"']) {
    writeFileSync(file, input)
    expect(researchScope()).toBe("official")
    expect(normalizePreferences(readFirmPreferences()).courtlistenerMcpEnabled).toBe(false)
    await expect(prepareFetch()).rejects.toThrow("webfetch is limited")
  }
  rmSync(file)
  mkdirSync(file)
  expect(researchScope()).toBe("official")
  expect(normalizePreferences(readFirmPreferences()).courtlistenerMcpEnabled).toBe(false)
  await expect(prepareFetch()).rejects.toThrow("webfetch is limited")
})
