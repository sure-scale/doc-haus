import { afterEach, expect, test } from "bun:test"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Hono } from "hono"
import { onWordChatConsequencesDisabled, readWordChatPreferences, wordChatPreferencesFile, wordChatPreferencesRoutes, writeWordChatPreferences } from "./word-chat-preferences"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))
function fixture() {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-word-chat-preferences-"))
  directories.push(directory)
  const environment = { DOCHAUS_DATA_ROOT: directory }
  const app = new Hono().route("/word/chat-preferences", wordChatPreferencesRoutes(environment))
  return { environment, app, file: wordChatPreferencesFile(environment) }
}

test("installation defaults need no write, and a new reader sees validated incremental settings", async () => {
  const { environment, app, file } = fixture()
  expect(await (await app.request("/word/chat-preferences")).json()).toEqual({ defaultScope: "document", suggestions: true, consequences: true })
  expect(existsSync(file)).toBe(false)
  expect(writeWordChatPreferences({ defaultScope: "matter" }, environment)).toEqual({ defaultScope: "matter", suggestions: true, consequences: true })
  const sibling = path.join(path.dirname(file), "word-addin.json")
  writeFileSync(sibling, '{"origin":"https://localhost:3444"}\n')
  const saved = await app.request("/word/chat-preferences", { method: "PUT", headers: { "Content-Type": "application/json" }, body: '{"suggestions":false}' })
  expect(saved.status).toBe(200)
  expect(await saved.json()).toEqual({ defaultScope: "matter", suggestions: false, consequences: true })
  const reloaded = new Hono().route("/word/chat-preferences", wordChatPreferencesRoutes(environment))
  expect(await (await reloaded.request("/word/chat-preferences")).json()).toEqual(readWordChatPreferences(environment))
  expect(JSON.parse(readFileSync(file, "utf8"))).toEqual({ defaultScope: "matter", suggestions: false, consequences: true })
  expect(readFileSync(sibling, "utf8")).toBe('{"origin":"https://localhost:3444"}\n')
  expect(readdirSync(path.dirname(file)).sort()).toEqual(["word-addin.json", "word-chat-preferences.json"])
  if (process.platform !== "win32") expect(statSync(file).mode & 0o777).toBe(0o600)
})

test("unknown or invalid preferences cannot overwrite saved defaults, and corrupt storage is not silently reset", async () => {
  const { environment, app, file } = fixture()
  writeWordChatPreferences({ defaultScope: "matter", suggestions: false }, environment)
  const baseline = readFileSync(file, "utf8")
  for (const input of [null, [], { defaultScope: "selection" }, { suggestions: "false" }, { consequences: "false" }, { consequences: null }, { suggestionTarget: "document" }, { defaultScope: "document", suggestions: true, model: "unrelated" }]) {
    const response = await app.request("/word/chat-preferences", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input) })
    expect(response.status).toBe(400)
    expect(readFileSync(file, "utf8")).toBe(baseline)
  }
  writeFileSync(file, '{"private-content":broken')
  const read = await app.request("/word/chat-preferences")
  expect(read.status).toBe(503)
  expect(await read.text()).not.toContain("private-content")
  const write = await app.request("/word/chat-preferences", { method: "PUT", headers: { "Content-Type": "application/json" }, body: '{"suggestions":true}' })
  expect(write.status).toBe(503)
  expect(readFileSync(file, "utf8")).toBe('{"private-content":broken')
})

test("older saved Word defaults enable consequences without rewriting; toggling preserves the other defaults", async () => {
  const { environment, app, file } = fixture()
  mkdirSync(path.dirname(file), { recursive: true })
  const legacy = '{"defaultScope":"matter","suggestions":false}\n'
  writeFileSync(file, legacy)
  expect(readWordChatPreferences(environment)).toEqual({ defaultScope: "matter", suggestions: false, consequences: true })
  expect(readFileSync(file, "utf8")).toBe(legacy)
  const update = async (body: unknown) => app.request("/word/chat-preferences", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
  const off = await update({ consequences: false })
  expect(off.status).toBe(200)
  expect(await off.json()).toEqual({ defaultScope: "matter", suggestions: false, consequences: false })
  await update({ suggestions: true })
  expect(readWordChatPreferences(environment)).toEqual({ defaultScope: "matter", suggestions: true, consequences: false })
  await update({ consequences: true })
  expect(readWordChatPreferences(environment)).toEqual({ defaultScope: "matter", suggestions: true, consequences: true })
})

test("saving disable notifies active checks after persistence, while other updates do not", () => {
  const { environment } = fixture()
  let cancellations = 0
  let savedAtCancellation: boolean | undefined
  const stop = onWordChatConsequencesDisabled(() => {
    savedAtCancellation = readWordChatPreferences(environment).consequences
    cancellations++
  }, environment)
  try {
    writeWordChatPreferences({ suggestions: false }, environment)
    writeWordChatPreferences({ consequences: true }, environment)
    expect(cancellations).toBe(0)
    writeWordChatPreferences({ consequences: false }, fixture().environment)
    expect(cancellations).toBe(0)
    writeWordChatPreferences({ consequences: false }, environment)
    expect(cancellations).toBe(1)
    expect(savedAtCancellation).toBe(false)
  } finally { stop() }
  writeWordChatPreferences({ consequences: false }, environment)
  expect(cancellations).toBe(1)
})
