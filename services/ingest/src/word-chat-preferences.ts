import { Hono } from "hono"
import { closeSync, fsyncSync, mkdirSync, openSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs"
import path from "node:path"
import { randomUUID } from "node:crypto"
import { fileURLToPath } from "node:url"
import { dataPaths } from "../../engine/src/data-paths"
import type { WordChatPreferences } from "../../../dochaus/lib/word-chat-contract"

type Environment = Record<string, string | undefined>
const repository = fileURLToPath(new URL("../../../", import.meta.url))
const defaults: WordChatPreferences = { defaultScope: "document", suggestions: true }
class WordChatPreferencesError extends Error {
  constructor(message: string, readonly status: 400 | 503) { super(message) }
}

export function wordChatPreferencesFile(environment: Environment = process.env) {
  return path.join(dataPaths(repository, environment).runtime, "config", "word-chat-preferences.json")
}

function patch(input: unknown): Partial<WordChatPreferences> {
  if (!input || typeof input !== "object" || Array.isArray(input) ||
    Object.keys(input).some(key => key !== "defaultScope" && key !== "suggestions"))
    throw new WordChatPreferencesError("Provide only the Word chat default scope and suggestions setting.", 400)
  const value = input as Record<string, unknown>
  if (Object.hasOwn(value, "defaultScope") && value.defaultScope !== "document" && value.defaultScope !== "matter")
    throw new WordChatPreferencesError("Choose This document or This matter as the default scope.", 400)
  if (Object.hasOwn(value, "suggestions") && typeof value.suggestions !== "boolean")
    throw new WordChatPreferencesError("Choose whether to show suggestions.", 400)
  return value as Partial<WordChatPreferences>
}

export function readWordChatPreferences(environment: Environment = process.env): WordChatPreferences {
  try {
    const input: unknown = JSON.parse(readFileSync(wordChatPreferencesFile(environment), "utf8"))
    return { ...defaults, ...patch(input) }
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return { ...defaults }
    // Parser errors may repeat source bytes. Never expose saved configuration.
    throw new WordChatPreferencesError("Word chat defaults could not be read. Repair the saved settings before changing them.", 503)
  }
}

export function writeWordChatPreferences(input: unknown, environment: Environment = process.env): WordChatPreferences {
  const update = patch(input)
  const saved = { ...readWordChatPreferences(environment), ...update }
  const file = wordChatPreferencesFile(environment), temporary = `${file}.${randomUUID()}.tmp`
  // Synchronous read/merge/rename runs to completion on ingest's event loop.
  // Readers see a complete old or new file, and unrelated config is untouched.
  mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 })
  let descriptor: number | undefined
  try {
    descriptor = openSync(temporary, "wx", 0o600)
    writeFileSync(descriptor, JSON.stringify(saved, null, 2) + "\n")
    fsyncSync(descriptor)
    closeSync(descriptor)
    descriptor = undefined
    renameSync(temporary, file)
  } finally {
    if (descriptor !== undefined) closeSync(descriptor)
    try { unlinkSync(temporary) } catch { /* Already published or never created. */ }
  }
  return saved
}

/** Normal installation Settings API, deliberately outside the Word token bridge. */
export function wordChatPreferencesRoutes(environment: Environment = process.env) {
  const app = new Hono()
  app.onError((error, c) => c.json({ error: error instanceof WordChatPreferencesError ? error.message : "Word chat defaults could not be saved. Try again." }, error instanceof WordChatPreferencesError ? error.status : 503))
  app.get("/", c => c.json(readWordChatPreferences(environment)))
  app.put("/", async c => c.json(writeWordChatPreferences(await c.req.json().catch(() => undefined), environment)))
  return app
}
