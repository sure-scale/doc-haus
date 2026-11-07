import { Database } from "bun:sqlite"
import { existsSync, readFileSync, statSync } from "node:fs"
import path from "node:path"
import { effectiveDateFormat, isPosture, normalizePreferences, type DateFormat, type DraftingPreferences, type Posture } from "./drafting-preferences"
import { detectDateFormat } from "./markdown"
import { readFirmPreferences } from "./preferences"

// What the lawyer records on the matter about the client, read from its
// matter.json. The ingest service owns that file; the engine only reads it.
function readMatter(directory: string): { client?: unknown; posture?: unknown } {
  const file = path.join(directory, "matter.json")
  return existsSync(file) ? JSON.parse(readFileSync(file, "utf8")) : {}
}

// The lawyer's client as recorded on the matter: one line, control characters
// and white space collapsed as ingest writes it, or "" when the matter records none.
export function matterClient(directory: string) {
  const client = readMatter(directory).client
  return typeof client === "string" ? client.replace(/\p{Cc}/gu, " ").replace(/\s+/g, " ").trim() : ""
}

/** The matter's own Posture, or undefined when it records none (or a value
 * ingest would not write), so the firm's applies. Every chat turn reads it, so
 * an unreadable matter.json means no Posture rather than a failed turn. */
export function matterPosture(directory: string): Posture | undefined {
  let posture: unknown
  try {
    posture = readMatter(directory).posture
  } catch {
    return undefined
  }
  return isPosture(posture) ? posture : undefined
}

/** The Posture work on the matter takes: the matter's own, else the firm's. */
export function effectivePosture(directory: string): Posture {
  return matterPosture(directory) ?? normalizePreferences(readFirmPreferences()).posture
}

/** The date style the matter's sources use: the written form most of the
 * dates in their canonical text (the index ingest keeps in .dochaus/legal.db)
 * and in `request`, a drafting run's request, take (lib/markdown.ts
 * detectDateFormat); undefined when they write none or settle on none. Every
 * chat turn reads it, so the reading is kept until the index or the request
 * changes. */
export function matterDateFormat(directory: string, request = ""): DateFormat | undefined {
  const file = path.join(directory, ".dochaus", "legal.db")
  // The index is written ahead in its -wal file until a checkpoint, so both files date the reading.
  const stamp = [file, `${file}-wal`].map((name) => {
    try {
      const stat = statSync(name)
      return `${stat.mtimeMs}:${stat.size}`
    } catch {
      return ""
    }
  }).join("|")
  const kept = detected.get(directory)
  if (kept && kept.stamp === stamp && kept.request === request) return kept.value
  const value = detectDateFormat(`${sourcesText(file)}\n${request}`)
  detected.set(directory, { stamp, request, value })
  return value
}
const detected = new Map<string, { stamp: string; request: string; value: DateFormat | undefined }>()

/** The canonical text of every indexed source, or "" when there is no index
 * or it cannot be read: an unreadable index means no sources' style, not a
 * failed turn. */
function sourcesText(file: string) {
  if (!existsSync(file)) return ""
  let db: Database
  try {
    db = new Database(file, { readonly: true })
  } catch {
    return ""
  }
  try {
    return (db.query("SELECT text FROM chunks ORDER BY doc_name, rowid").all() as { text: string }[]).map((row) => row.text).join("\n")
  } catch {
    return ""
  } finally {
    db.close()
  }
}

/** The firm's preferences as work on the matter applies them: until the firm
 * sets its drafting preferences, with the date style the matter's sources use
 * (`sources`; a drafting run passes the style its plan froze, runDateFormat). */
export function matterDraftingPreferences(directory: string, sources: DateFormat | undefined = matterDateFormat(directory)): DraftingPreferences {
  const prefs = normalizePreferences(readFirmPreferences())
  return { ...prefs, dateFormat: effectiveDateFormat(prefs, sources) }
}
