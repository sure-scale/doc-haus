import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { commitDocumentDeleteAsync } from "./atomic-document-write"
import { deleteDocument, getDocumentRole, listSourceDocuments, openDb } from "./db"
import { deliverableDocumentName } from "./deliverable"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"
import { matterDocumentName } from "./ingest"
import { stageSourceChange, writeSourceArtifacts } from "./source-artifacts"
import { workingDocxKey } from "../../../dochaus/lib/working-docx"
import { forgetSystemDocx } from "../../../dochaus/lib/system-docx"
import { detachVersionedDocument, purgeVersionBytes } from "./document-versions"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import {
  beginDocumentMutation,
  completeDocumentMutation,
  discardDocumentMutation,
  finalizeDocumentMutationLineage,
  recordDocumentMutationCommit,
  recordDocumentMutationLineageIntent,
  recoverDocumentMutations,
} from "./document-mutation"

export class RegisteredDocumentNotFoundError extends Error {}
export class DocumentMutationConflictError extends Error {}

export async function deleteMatterDocument(matterDir: string, requestedName: string) {
  const name = matterDocumentName(requestedName) ?? deliverableDocumentName(requestedName)
  if (!name) throw new DocumentMutationConflictError(`The document filename or extension is not supported.`)
  return withArtifactMutationLocksAsync(matterDir, [name, ".dochaus/source-corpus"], async () => {
    const recovery = await recoverDocumentMutations(matterDir)
    if (recovery.sourceRolledBack) await writeSourceArtifacts(matterDir)
    const target = documentTarget(matterDir, name)
    const db = openDb(matterDir, { commit: true })
    const role = getDocumentRole(db, target.docPath)
    if (!role) {
      db.close()
      throw new RegisteredDocumentNotFoundError(`The document is not registered in this matter.`)
    }
    const correlation = `delete-${Date.now().toString(36)}`
    const mutation = beginDocumentMutation(matterDir, target.name, "delete", role === "source")
    // The document's version history, detached in the delete's transaction; its
    // kept bytes are deleted once the delete commits, as the file is (#268).
    let detached: string[] = []
    try {
      if (role === "deliverable") {
        await commitDocumentDeleteAsync(target.docPath, async () => {
          db.run("BEGIN IMMEDIATE")
          try {
            deleteDocument(db, target.docPath)
            detached = detachVersionedDocument(db, matterDir, target.name)
            db.run("DELETE FROM meta WHERE key = ?",[workingDocxKey(matterDir,target.docPath)])
            forgetSystemDocx(db, matterDir, target.docPath)
            recordDocumentMutationCommit(db, mutation)
            db.run("COMMIT")
          } catch (error) {
            try {
              db.run("ROLLBACK")
            } catch {
              // Preserve the original mutation failure.
            }
            throw error
          }
        }, { backupPath: mutation.documentBackup })
        completeDocumentMutation(db, mutation)
        return { ok: true as const, name: target.name, role, ...purgeDetached(matterDir, target.name, detached) }
      }
      // The remaining sources' snapshot is staged before legal.db's write lock
      // is taken (see stageSourceChange), so the transaction holds the lock for
      // the registry delete and the snapshot's installation only.
      const staged = await stageSourceChange(matterDir, db, mutation.projectionDirectory, target.docPath)
      const artifacts = await commitDocumentDeleteAsync(target.docPath, async () => {
        db.run("BEGIN IMMEDIATE")
        let settled = false
        try {
          deleteDocument(db, target.docPath)
          detached = detachVersionedDocument(db, matterDir, target.name)
          recordDocumentMutationLineageIntent(
            mutation,
            (listSourceDocuments(db) as { name: string }[]).map((document) => document.name),
            correlation,
            "delete",
          )
          return await writeSourceArtifacts(matterDir, correlation, {
            db,
            stampLineage: false,
            staged,
            settle: () => {
              recordDocumentMutationCommit(db, mutation)
              db.run("COMMIT")
              settled = true
            },
          })
        } catch (error) {
          if (!settled) {
            try {
              db.run("ROLLBACK")
            } catch {
              // Preserve the mutation failure if SQLite already rolled back.
            }
          }
          throw error
        }
      }, { backupPath: mutation.documentBackup })
      finalizeDocumentMutationLineage(db, mutation, artifacts.files)
      completeDocumentMutation(db, mutation)
      return { ok: true as const, name: target.name, role, ...purgeDetached(matterDir, target.name, detached) }
    } catch (error) {
      discardDocumentMutation(db, mutation)
      throw error
    } finally {
      db.close()
    }
  })
}

// The delete stands; bytes left behind are logged and reported, no route
// serves a removed document's versions, and the startup sweep
// (sweepRemovedVersionBytes) deletes them.
function purgeDetached(matterDir: string, name: string, documentIds: readonly string[]): { purge_failed?: string } {
  let failed = false
  for (const id of documentIds) {
    try {
      purgeVersionBytes(matterDir, id)
    } catch (error) {
      failed = true
      console.error("[versions] removed document's versions not purged", diagnosticErrorMetadata(error))
    }
  }
  return failed ? { purge_failed: `The kept versions of ${name} could not be deleted from doc.haus; they are deleted when doc.haus next starts.` } : {}
}

function documentTarget(matterDir: string, fileName: string) {
  try {
    return resolveMatterDocumentTarget(matterDir, fileName)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) throw new DocumentMutationConflictError(error.message)
    throw error
  }
}
