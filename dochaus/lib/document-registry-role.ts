import type { Database } from "bun:sqlite"
import path from "node:path"
import { canonicalArtifactIdentity } from "../capabilities/artifacts/mutation-lock"

export function documentRegistryIdentity(docPath: string) {
  return canonicalArtifactIdentity(path.dirname(docPath), path.basename(docPath))
}

/** Resolve the physical registry identity, rejecting contradictory roles. */
export function documentRegistryRole(db: Database, docPath: string): "source" | "deliverable" | null {
  const identity = documentRegistryIdentity(docPath)
  const rows = db.query("SELECT doc_path, role FROM documents").all() as { doc_path: string; role: string }[]
  const roles = [...new Set(rows.filter(row => documentRegistryIdentity(row.doc_path) === identity).map(row => row.role))]
  if (roles.length > 1) throw new Error("One physical matter document has conflicting source and deliverable registry roles.")
  if (roles.length && roles[0] !== "source" && roles[0] !== "deliverable") throw new Error("The document registry role is invalid.")
  return (roles[0] as "source" | "deliverable" | undefined) ?? null
}
