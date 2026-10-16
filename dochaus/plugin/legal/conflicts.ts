import { existsSync, readFileSync } from "node:fs"
import path from "node:path"

// The matter's conflicts intake record, read from its matter.json (the ingest
// service owns that file; POST /matters/:id/conflicts writes it). Absent on
// matters that predate the field.
export function readMatterConflicts(directory: string) {
  const file = path.join(directory, "matter.json")
  if (!existsSync(file)) return undefined
  return (
    JSON.parse(readFileSync(file, "utf8")) as {
      conflicts?: {
        status: "cleared" | "pending" | "bypassed"
        by?: string
        owner?: string
        due?: string
        reason?: string
      }
    }
  ).conflicts
}
