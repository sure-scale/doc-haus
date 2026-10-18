import type { Database } from "bun:sqlite"
import { createHash } from "node:crypto"
import path from "node:path"
import { existsSync, lstatSync } from "node:fs"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import {
  getDocumentRole,
  openDb,
  pendingRedlineCounts,
  refreshDeliverable,
  upsertDeliverable,
} from "./db"
import { commitDocumentWriteAsync } from "./atomic-document-write"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"
import { sha256File } from "../../../dochaus/lib/lineage"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { forgetSystemDocx, recordSystemDocx } from "../../../dochaus/lib/system-docx"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import {
  beginDocumentMutation,
  completeDocumentMutation,
  discardDocumentMutation,
  recordDocumentMutationCommit,
  recoverDocumentMutations,
} from "./document-mutation"

export class DeliverableConflictError extends Error {}

export function deliverableDocumentName(fileName: string) {
  const name = path.basename(fileName.trim())
  if (!name || name !== fileName.trim() || ![".docx", ".pdf"].includes(path.extname(name).toLowerCase())) return
  return name
}

// Generated outputs share the matter document registry with uploaded sources so
// the web can list and serve them, but they never enter extraction, chunking,
// embedding, structure analysis, or source-artifact generation.
// A write with `system` set is the system's own (a run's registration of its
// outputs, ingest's rebuild of a saved combined file): the registry transaction
// that writes it records the sha256 of the bytes stored (system-docx.ts). Every
// other write removes that record, so only an unchanged system write may be
// written over again by the system.
export async function writeDeliverable(
  matterDir: string,
  fileName: string,
  buffer: Buffer,
  options: {
    precondition?:
      | { kind: "create" }
      | { kind: "replace"; prior_sha256: string }
    system?: boolean
  } = {},
) {
  const name = deliverableDocumentName(fileName)
  if (!name) throw new DeliverableConflictError(`Deliverables must use a supported document filename and extension.`)
  return withArtifactMutationLocksAsync(matterDir, [name, ".dochaus/source-corpus"], async () => {
    await recoverDocumentMutations(matterDir)
    const target = deliverableTarget(matterDir, name)
    const db = openDb(matterDir)
    try {
      if (getDocumentRole(db, target.docPath) === "source")
        throw new DeliverableConflictError(`Cannot replace source document ${target.name} with a generated deliverable.`)
      const prior = sha256File(target.docPath)
      if (readWorkingDocx(matterDir, target.docPath))
        throw new DeliverableConflictError("This document has an accepted Word revision. Prepare changes or use the document revision API.")
      if (options.precondition?.kind === "create" && prior)
        throw new DeliverableConflictError(
          `Cannot create generated deliverable ${target.name} because the target already exists.`,
        )
      if (
        options.precondition?.kind === "replace" &&
        prior?.sha256 !== options.precondition.prior_sha256
      )
        throw new DeliverableConflictError(
          `Cannot replace generated deliverable ${target.name} because its prior identity changed.`,
        )
      const bytes = name.toLowerCase().endsWith(".docx") ? Buffer.from(withParagraphPointers(buffer)) : buffer
      return await persistDeliverable(db, matterDir, target, bytes, options.system)
    } finally {
      db.close()
    }
  })
}

// Copy canonical Word bytes exactly. App review proposals are separate state and
// deliberately stay on the original; no rendering, pointer injection or reindexing.
export function copyWordDocumentName(name: string) {
  return name === name.trim() && !/[\\/\0]/.test(name) && path.extname(name).toLowerCase() === ".docx"
    ? deliverableDocumentName(name)
    : undefined
}

export async function copyWordDocument(matterDir: string, sourceName: string, fileName: string) {
  const source = copyWordDocumentName(sourceName)
  const name = copyWordDocumentName(fileName)
  if (!source || !name) throw new DeliverableConflictError("Use Word document filenames within this matter, without paths.")
  deliverableTarget(matterDir, source)
  deliverableTarget(matterDir, name)
  return withArtifactMutationLocksAsync(matterDir, [source, name, ".dochaus/source-corpus"], async () => {
    await recoverDocumentMutations(matterDir)
    const original = deliverableTarget(matterDir, source)
    const target = deliverableTarget(matterDir, name)
    const db = openDb(matterDir)
    try {
      if (!getDocumentRole(db, original.docPath) || !existsSync(original.docPath) || !lstatSync(original.docPath).isFile())
        throw new DeliverableConflictError("The original must be an existing registered Word document in this matter.")
      if (existsSync(target.docPath) || getDocumentRole(db, target.docPath))
        throw new DeliverableConflictError(`Cannot create ${target.name}: that document already exists. Choose a different name.`)
      const bytes = Buffer.from(await Bun.file(original.docPath).bytes())
      const pendingProposalsExcluded = pendingRedlineCounts(db)[original.identity] ?? 0
      return { ...await persistDeliverable(db, matterDir, target, bytes), source: original.name, pendingProposalsExcluded }
    } finally {
      db.close()
    }
  })
}

async function persistDeliverable(db: Database, matterDir: string, target: { name: string; docPath: string }, bytes: Buffer, system = false) {
  const mutation = beginDocumentMutation(matterDir, target.name, "write", false)
  try {
    await commitDocumentWriteAsync(target.docPath, bytes, async () => {
      db.run("BEGIN IMMEDIATE")
      try {
        upsertDeliverable(db, target.docPath, target.name, Date.now())
        updateSystemDocx(db, matterDir, target.docPath, bytes, system)
        recordDocumentMutationCommit(db, mutation)
        db.run("COMMIT")
      } catch (error) {
        rollbackDocumentMutation(db)
        throw error
      }
    }, { backupPath: mutation.documentBackup })
    completeDocumentMutation(db, mutation)
    return { name: target.name, docPath: target.docPath, role: "deliverable" as const }
  } catch (error) {
    discardDocumentMutation(db, mutation)
    throw error
  }
}

// Persist an accepted edit to an existing deliverable without replacing its
// registry row or touching retrieval/source state. Pending redline rows remain
// available for the caller to mark accepted after the bytes are safely written.
export async function writeDeliverableRevision(
  matterDir: string,
  fileName: string,
  buffer: Buffer,
  options: { updateRegistry?: (db: Database) => void; priorSha256?: string; updateWorkingRevision?: (db: Database) => void; system?: boolean } = {},
) {
  const name = deliverableDocumentName(fileName)
  if (!name) throw new DeliverableConflictError(`Deliverables must use a supported document filename and extension.`)
  return withArtifactMutationLocksAsync(matterDir, [name, ".dochaus/source-corpus"], async () => {
    await recoverDocumentMutations(matterDir)
    const target = deliverableTarget(matterDir, name)
    const db = openDb(matterDir)
    try {
      if (getDocumentRole(db, target.docPath) !== "deliverable")
        throw new DeliverableConflictError(`${target.name} is not a generated deliverable.`)
      if (readWorkingDocx(matterDir,target.docPath) && (!options.priorSha256 || !options.updateWorkingRevision))
        throw new DeliverableConflictError("An accepted Word revision requires its exact prior hash and matching projection update.")
      if (options.priorSha256 && sha256File(target.docPath)?.sha256 !== options.priorSha256)
        throw new DeliverableConflictError("The accepted Word revision changed before publication.")
      const mutation = beginDocumentMutation(matterDir, target.name, "write", false)
      try {
        await commitDocumentWriteAsync(target.docPath, buffer, async () => {
          db.run("BEGIN IMMEDIATE")
          try {
            refreshDeliverable(db, target.docPath, target.name, Date.now())
            updateSystemDocx(db, matterDir, target.docPath, buffer, options.system)
            options.updateRegistry?.(db)
            options.updateWorkingRevision?.(db)
            recordDocumentMutationCommit(db, mutation)
            db.run("COMMIT")
          } catch (error) {
            rollbackDocumentMutation(db)
            throw error
          }
        }, { backupPath: mutation.documentBackup })
        completeDocumentMutation(db, mutation)
        return { name: target.name, docPath: target.docPath, role: "deliverable" as const }
      } catch (error) {
        discardDocumentMutation(db, mutation)
        throw error
      }
    } finally {
      db.close()
    }
  })
}

/** In the write's registry transaction: records the bytes of the system's own
 * write, or removes the record for any other write. */
function updateSystemDocx(db: Database, matterDir: string, docPath: string, bytes: Uint8Array, system = false) {
  if (system) recordSystemDocx(db, matterDir, docPath, createHash("sha256").update(bytes).digest("hex"))
  else forgetSystemDocx(db, matterDir, docPath)
}

function deliverableTarget(matterDir: string, fileName: string) {
  try {
    return resolveMatterDocumentTarget(matterDir, fileName)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) throw new DeliverableConflictError(error.message)
    throw error
  }
}

function rollbackDocumentMutation(db: Database) {
  try {
    db.run("ROLLBACK")
  } catch {
    // Preserve the mutation failure if SQLite already rolled back.
  }
}
