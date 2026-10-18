import { existsSync } from "node:fs"
import path from "node:path"
import { Hono, type Context } from "hono"
import { canonicalArtifactIdentity, withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { DirectoryLockContentionError } from "../../../dochaus/lib/directory-lock"
import { combinedParts, readLeanDocuments, type LeanDocument } from "../../../dochaus/workflows/lean/deliverables"
import { attachmentDocumentDisposition } from "./content-disposition"
import { getDocumentRole, openDb } from "./db"
import { deliverableDocumentName } from "./deliverable"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"

// The one file a combined document is delivered as with its requested memo
// (combinedDelivery): the file its delivery saved in the matter and registered
// in the run's summary (a "combined" row for that document), served as it is
// now. The run's registration writes it, ingest rebuilds it after a Word accept
// into its body or notes while the system's write is unchanged (server.ts), and a
// lawyer may accept changes into it directly; whichever wrote it last, the
// download is that matter file, byte for byte, and nothing is built here.
// Without a registered combined row nothing is offered: no file is guessed from
// the plan's name. The route reads the file under its mutation lock, so it never
// sees a write half done, and writes and registers nothing.

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

/** combinedParts when the run's summary registers the combined file itself (a
 * "combined" row for the same document under that name), every name plain;
 * undefined otherwise. The one rule for reporting, serving and rebuilding the
 * saved combined file. */
export function registeredCombinedParts(directory: string, rows: readonly LeanDocument[] = readLeanDocuments(directory)?.documents ?? []) {
  const parts = combinedParts(directory, [...rows])
  if (!parts || [parts.file, parts.body, parts.notes].some(name => deliverableDocumentName(name) !== name)) return
  return rows.some(row => row.kind === "combined" && row.document === parts.document && row.deliverable === parts.file) ? parts : undefined
}

export function combinedOutputRoutes(resolveMatter: (id: string) => string) {
  return new Hono().get("/matters/:id/documents/combined", async (c) => {
    const directory = resolveMatter(c.req.param("id"))
    const parts = registeredCombinedParts(directory)
    if (!parts) return c.json({ error: "no_combined_output" }, 404)
    let bytes: Uint8Array | undefined
    try {
      bytes = await withArtifactMutationLocksAsync(directory, [parts.file], () => registeredBytes(directory, parts.file))
    } catch (error) {
      if (error instanceof DirectoryLockContentionError) return busy(c)
      // A file that is a link out of the matter has no lock there and is never read.
      if (!insideMatter(directory, parts.file)) return missing(c)
      throw error
    }
    if (!bytes) return missing(c)
    return new Response(new Uint8Array(bytes), {
      headers: { "Content-Type": DOCX, "Content-Disposition": attachmentDocumentDisposition(parts.file) },
    })
  })
}

function missing(c: Context) {
  return c.json({ error: "combined_file_missing", message: "The combined file is no longer in the matter." }, 409)
}

/** Lock contention outlasted the lock's own wait: a long write holds the file. */
function busy(c: Context) {
  return c.json(
    { error: "combined_file_busy", message: "The combined file is being updated. Try again in a moment." },
    503,
    { "Retry-After": "5" },
  )
}

function insideMatter(directory: string, name: string) {
  return canonicalArtifactIdentity(directory, name).startsWith(`${canonicalArtifactIdentity(directory, ".")}${path.sep}`)
}

/** The current bytes of a registered deliverable; undefined when its file is
 * gone, is a link, or the name no longer holds a deliverable (a source that
 * took it). */
async function registeredBytes(directory: string, name: string) {
  let target: ReturnType<typeof resolveMatterDocumentTarget>
  try {
    target = resolveMatterDocumentTarget(directory, name)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) return
    throw error
  }
  if (!existsSync(target.docPath) || !existsSync(path.join(directory, ".dochaus", "legal.db"))) return
  const db = openDb(directory)
  try {
    if (getDocumentRole(db, target.docPath) !== "deliverable") return
  } finally {
    db.close()
  }
  return Bun.file(target.docPath).bytes()
}
