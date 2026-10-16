import { readFileSync } from "node:fs"
import path from "node:path"

// The firm-wide preferences the ingest settings routes write
// (services/ingest/src/preferences.ts). Read from disk on every call so a
// settings save applies on the very next tool call, no engine restart.
export function readFirmPreferences(): Record<string, unknown> {
  const root = process.env.WORKSPACE_ROOT
  if (!root) return {}
  const file = path.join(root, ".preferences", "preferences.json")
  try {
    const input: unknown = JSON.parse(readFileSync(file, "utf8"))
    return input && typeof input === "object" && !Array.isArray(input)
      ? input as Record<string, unknown>
      : { webResearch: "official" }
  } catch (error) {
    // Fresh workspaces use defaults. A broken saved file retains the restricted
    // web scope; absent connector flags continue to revoke remote MCP consent.
    return (error as NodeJS.ErrnoException).code === "ENOENT" ? {} : { webResearch: "official" }
  }
}
