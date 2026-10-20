import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { isPosture, normalizePreferences, type Posture } from "./drafting-preferences"
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
 * ingest would not write), so the firm's applies. */
export function matterPosture(directory: string): Posture | undefined {
  const posture = readMatter(directory).posture
  return isPosture(posture) ? posture : undefined
}

/** The Posture work on the matter takes: the matter's own, else the firm's. */
export function effectivePosture(directory: string): Posture {
  return matterPosture(directory) ?? normalizePreferences(readFirmPreferences()).posture
}
