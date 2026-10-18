import { existsSync } from "node:fs"
import path from "node:path"
import { Hono, type Context } from "hono"
import { canonicalArtifactIdentity, withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { buildCombinedFile } from "../../../dochaus/lib/combined-docx"
import { DirectoryLockContentionError } from "../../../dochaus/lib/directory-lock"
import { combinedDelivery, combinedParts } from "../../../dochaus/workflows/lean/deliverables"
import { attachmentDocumentDisposition } from "./content-disposition"
import { getDocumentRole, openDb } from "./db"
import { deliverableDocumentName } from "./deliverable"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"

// The one file a combined document is delivered as with its requested memo
// (combinedDelivery): the file buildCombinedFile builds from the registered body
// ("<stem> - clean") and the registered notes, memo first, as the current plan
// names them (combinedParts); parts delivered under an earlier plan are never
// paired. Only those registered files carry Word-accepted changes and the notes
// addendum, so the run's own copies are never used. The delivery saves that file
// in the matter under its registered name (combinedParts' `file`), and a Word
// accept into either part rebuilds it (server.ts), so each download builds the
// same bytes as the saved file; a part that no longer builds is refused rather
// than the saved file served. The route itself writes and registers nothing.
// Both files are read together under their mutation locks and built after the
// locks are released. A Word accept on the body holds both locks across its
// body write, notes addendum and combined rebuild (server.ts), so a download
// never pairs an accepted body with the notes from before it. A lean
// re-delivery still registers the body and the notes one at a time, so a
// download between those two writes can pair the new body with the old notes.

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

/** buildCombinedFile's refusals of memo content the combined file cannot carry. */
const REFUSALS: Record<string, string> = {
  empty_memo: "The memo has no text to place before the draft.",
  unsupported_memo_numbering: "The memo uses list numbering the combined file cannot carry. Download the draft and the memo separately.",
  unsupported_memo_namespace: "The memo uses Word markup the combined file cannot carry. Download the draft and the memo separately.",
}
const REFUSED_CONTENT = "The memo holds content the combined file cannot carry. Download the draft and the memo separately."
const PART_LABEL = { body: "draft", memo: "memo" } as const

export function combinedOutputRoutes(resolveMatter: (id: string) => string) {
  return new Hono().get("/matters/:id/documents/combined", async (c) => {
    const directory = resolveMatter(c.req.param("id"))
    if (!combinedDelivery(directory)) return c.json({ error: "no_combined_output" }, 404)
    const parts = combinedParts(directory)
    // Plain file names only, as /documents/content takes them.
    if (!parts || ![parts.body, parts.notes].every(name => deliverableDocumentName(name) === name))
      return c.json({ error: "combined_parts_missing" }, 409)
    let read: { body: Uint8Array | undefined; memo: Uint8Array | undefined }
    try {
      read = await withArtifactMutationLocksAsync(directory, [parts.body, parts.notes], async () => ({
        body: await registeredBytes(directory, parts.body),
        memo: await registeredBytes(directory, parts.notes),
      }))
    } catch (error) {
      if (error instanceof DirectoryLockContentionError) return busy(c)
      // A part that is a link out of the matter has no lock there and is never read.
      if (![parts.body, parts.notes].every(name => insideMatter(directory, name))) return c.json({ error: "combined_parts_missing" }, 409)
      throw error
    }
    const { body, memo } = read
    if (!body || !memo) return c.json({ error: "combined_parts_missing" }, 409)
    const built = buildCombinedFile({ memo, body })
    if ("refused" in built) {
      if (built.part) return c.json({
        error: built.refused,
        part: built.part,
        message: `The ${PART_LABEL[built.part]} could not be read as a Word document. Open it in Word, save it again, and retry.`,
      }, 409)
      if (built.refused === "build_failed") throw new Error("The combined file could not be built.")
      return c.json({ error: built.refused, message: REFUSALS[built.refused] ?? REFUSED_CONTENT }, 409)
    }
    return new Response(new Uint8Array(built.bytes), {
      headers: { "Content-Type": DOCX, "Content-Disposition": attachmentDocumentDisposition(parts.file) },
    })
  })
}

/** Lock contention outlasted the lock's own wait: a long write holds a part. */
function busy(c: Context) {
  return c.json(
    { error: "combined_parts_busy", message: "The draft or the memo is being updated. Try again in a moment." },
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
