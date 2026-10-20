import { existsSync, readFileSync } from "node:fs"
import path from "node:path"

// The lawyer's client as recorded on the matter, read from its matter.json: one
// line, control characters and white space collapsed as ingest writes it, or "" when the matter records none. The
// ingest service owns that file; the engine only reads it.
export function matterClient(directory: string) {
  const file = path.join(directory, "matter.json")
  if (!existsSync(file)) return ""
  const client = (JSON.parse(readFileSync(file, "utf8")) as { client?: unknown }).client
  return typeof client === "string" ? client.replace(/\p{Cc}/gu, " ").replace(/\s+/g, " ").trim() : ""
}
