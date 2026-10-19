import type { Database } from "bun:sqlite"
import { randomUUID } from "node:crypto"
import {
  closeSync,
  existsSync,
  fsyncSync,
  mkdirSync,
  openSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
  type Dirent,
} from "node:fs"
import path from "node:path"
import { listSourceDocuments, openDb } from "./db"
import { recoverSourceArtifactTransaction } from "./source-artifacts"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { stampIngestLineage, type ArtifactHash } from "./lineage"

type DocumentMutation = {
  version: 2
  mutation_id: string
  target_name: string
  operation: "write" | "delete"
  prior_exists: boolean
  affects_source_projection: boolean
}

type DocumentMutationEntry = { root: string; entry: Dirent }

export type ActiveDocumentMutation = DocumentMutation & {
  matterDir: string
  directory: string
  documentBackup: string
  projectionDirectory: string
  lineageIntent: string
  legacyLineageComplete: string
  databaseCommitted: string
  lineageSettled: string
}

export class DocumentMutationRecoveryError extends Error {}

export function beginDocumentMutation(
  matterDir: string,
  targetName: string,
  operation: DocumentMutation["operation"],
  affectsSourceProjection: boolean,
) {
  const mutation_id = randomUUID()
  const directory = path.join(matterDir, ".dochaus", "document-mutations", mutation_id)
  const mutation = {
    version: 2 as const,
    mutation_id,
    target_name: targetName,
    operation,
    prior_exists: existsSync(path.join(matterDir, targetName)),
    affects_source_projection: affectsSourceProjection,
  }
  mkdirSync(directory, { recursive: true })
  writeFileSync(path.join(directory, "transaction.json"), `${JSON.stringify(mutation, null, 2)}\n`, { flag: "wx" })
  return activeDocumentMutation(matterDir, directory, mutation)
}

export function recordDocumentMutationCommit(db: Database, mutation: ActiveDocumentMutation) {
  db.run(
    "INSERT INTO document_mutation_commits (mutation_id, target_name, committed_at) VALUES (?, ?, ?)",
    [mutation.mutation_id, mutation.target_name, Date.now()],
  )
}

export function recordDocumentMutationLineageIntent(
  mutation: ActiveDocumentMutation,
  documents: string[],
  correlation: string,
) {
  const content = `${JSON.stringify({ version: 1, documents, correlation }, null, 2)}\n`
  if (existsSync(mutation.lineageIntent)) {
    if (readFileSync(mutation.lineageIntent, "utf8") !== content)
      throw new Error(`Document mutation ${mutation.mutation_id} has conflicting lineage intent.`)
    return
  }
  writeFileSync(mutation.lineageIntent, content, { flag: "wx" })
}

// `files`: the hashes of the source snapshot this mutation installed, when the
// caller staged it (recovery has none and hashes the installed files).
export function finalizeDocumentMutationLineage(
  db: Database,
  mutation: ActiveDocumentMutation,
  files?: ReadonlyMap<string, ArtifactHash>,
) {
  if (!mutation.affects_source_projection || durableMutationPhaseExists(mutation, mutation.lineageSettled, "lineage-settled")) return
  settleDocumentMutationCommit(db, mutation)
  const intent = readDocumentMutationLineageIntent(db, mutation)
  db.run("PRAGMA wal_checkpoint(FULL)")
  stampIngestLineage(mutation.matterDir, intent.documents, intent.correlation, mutation.mutation_id, files)
  writeDurableMutationPhase(mutation, mutation.lineageSettled, "lineage-settled")
}

export function completeDocumentMutation(db: Database, mutation: ActiveDocumentMutation) {
  settleDocumentMutationCommit(db, mutation)
  if (
    mutation.affects_source_projection &&
    !durableMutationPhaseExists(mutation, mutation.lineageSettled, "lineage-settled")
  )
    throw new Error(`Document mutation ${mutation.mutation_id} has not finalized source lineage.`)
  rmSync(mutation.directory, { recursive: true, force: true })
}

export function discardDocumentMutation(db: Database, mutation: ActiveDocumentMutation) {
  const committed = documentMutationCommitted(db, mutation)
  if (committed) {
    if (mutation.affects_source_projection)
      recoverSourceArtifactTransaction(mutation.matterDir, mutation.projectionDirectory, true)
    return
  }
  restoreUncommittedDocument(mutation.matterDir, mutation)
  if (mutation.affects_source_projection)
    recoverSourceArtifactTransaction(mutation.matterDir, mutation.projectionDirectory, false)
  rmSync(mutation.directory, { recursive: true, force: true })
  deleteDocumentMutationCommit(db, mutation.mutation_id)
}

export async function recoverDocumentMutations(matterDir: string) {
  return withArtifactMutationLocksAsync(
    matterDir,
    [".dochaus/source-corpus"],
    async () => recoverDocumentMutationsLocked(matterDir),
  )
}

async function recoverDocumentMutationsLocked(matterDir: string) {
  const entries = ["document-mutations", "source-mutations"].flatMap((name) => {
    const root = path.join(matterDir, ".dochaus", name)
    if (!existsSync(root)) return []
    return readdirSync(root, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => ({ root, entry }))
  })
  if (!entries.length) return { recovered: 0, rolledBack: 0, sourceRolledBack: 0 }
  const targets = entries.flatMap(({ root, entry }) => {
    const mutation = readDocumentMutation(path.join(root, entry.name))
    return mutation && mutation.mutation_id === entry.name ? [mutation.target_name] : []
  })
  return withArtifactMutationLocksAsync(matterDir, targets, async () => recoverDocumentMutationsWithTargetLocks(matterDir, entries))
}

async function recoverDocumentMutationsWithTargetLocks(
  matterDir: string,
  entries: DocumentMutationEntry[],
) {
  const db = openDb(matterDir)
  try {
    return entries.reduce(
        (state, { root, entry }) => {
          const directory = path.join(root, entry.name)
          const mutation = readDocumentMutation(directory)
          if (!mutation || mutation.mutation_id !== entry.name) {
            const hasRecoverableState = existsSync(path.join(directory, "document.bak")) ||
              existsSync(path.join(directory, "projection")) ||
              existsSync(path.join(directory, "lineage-complete")) ||
              existsSync(path.join(directory, "database-committed.json")) ||
              existsSync(path.join(directory, "lineage-settled.json"))
            if (hasRecoverableState)
              throw new DocumentMutationRecoveryError(
                `Cannot safely recover damaged document mutation ${entry.name}.`,
              )
            rmSync(directory, { recursive: true, force: true })
            return { ...state, recovered: state.recovered + 1 }
          }
          const active = activeDocumentMutation(matterDir, directory, mutation)
          const committed = documentMutationCommitted(db, active)
          if (!committed) {
            restoreUncommittedDocument(
              matterDir,
              active,
            )
            if (active.affects_source_projection)
              recoverSourceArtifactTransaction(matterDir, active.projectionDirectory, false)
          } else if (active.affects_source_projection) {
            recoverSourceArtifactTransaction(matterDir, active.projectionDirectory, true)
          }
          if (committed && active.affects_source_projection)
            finalizeDocumentMutationLineage(db, active)
          if (committed && !active.affects_source_projection)
            settleDocumentMutationCommit(db, active)
          rmSync(directory, { recursive: true, force: true })
          return {
            recovered: state.recovered + 1,
            rolledBack: state.rolledBack + (committed ? 0 : 1),
            sourceRolledBack: state.sourceRolledBack +
              (!committed && mutation.affects_source_projection ? 1 : 0),
          }
        },
        { recovered: 0, rolledBack: 0, sourceRolledBack: 0 },
      )
  } finally {
    db.close()
  }
}

function readDocumentMutation(directory: string) {
  try {
    const value = JSON.parse(readFileSync(path.join(directory, "transaction.json"), "utf8")) as unknown
    if (!value || typeof value !== "object" || Array.isArray(value)) return
    const mutation = value as Record<string, unknown>
    if (
      (mutation.version !== 1 && mutation.version !== 2) ||
      typeof mutation.mutation_id !== "string" ||
      typeof mutation.target_name !== "string" ||
      path.basename(mutation.target_name) !== mutation.target_name ||
      (mutation.operation !== "write" && mutation.operation !== "delete") ||
      typeof mutation.prior_exists !== "boolean" ||
      (mutation.version === 2 && typeof mutation.affects_source_projection !== "boolean")
    ) return
    return {
      version: 2,
      mutation_id: mutation.mutation_id,
      target_name: mutation.target_name,
      operation: mutation.operation,
      prior_exists: mutation.prior_exists,
      affects_source_projection: mutation.version === 1 ? true : mutation.affects_source_projection,
    } as DocumentMutation
  } catch {
    return
  }
}

function activeDocumentMutation(
  matterDir: string,
  directory: string,
  mutation: DocumentMutation,
) {
  return {
    ...mutation,
    matterDir,
    directory,
    documentBackup: path.join(directory, "document.bak"),
    projectionDirectory: path.join(directory, "projection"),
    lineageIntent: path.join(directory, "lineage-intent.json"),
    legacyLineageComplete: path.join(directory, "lineage-complete"),
    databaseCommitted: path.join(directory, "database-committed.json"),
    lineageSettled: path.join(directory, "lineage-settled.json"),
  } satisfies ActiveDocumentMutation
}

function restoreUncommittedDocument(matterDir: string, mutation: ActiveDocumentMutation) {
  const target = path.join(matterDir, mutation.target_name)
  if (!mutation.prior_exists) {
    rmSync(target, { force: true })
    return
  }
  if (!existsSync(mutation.documentBackup)) return
  rmSync(target, { force: true })
  renameSync(mutation.documentBackup, target)
}

function documentMutationCommitted(db: Database, mutation: ActiveDocumentMutation) {
  return documentMutationCommitRecorded(db, mutation.mutation_id) ||
    durableMutationPhaseExists(mutation, mutation.databaseCommitted, "database-committed") ||
    legacyLineageCommitEvidence(mutation)
}

function deleteDocumentMutationCommit(db: Database, mutationId: string) {
  db.run("DELETE FROM document_mutation_commits WHERE mutation_id = ?", [mutationId])
  db.run("DELETE FROM source_mutation_commits WHERE mutation_id = ?", [mutationId])
}

function documentMutationCommitRecorded(db: Database, mutationId: string) {
  return Boolean(
    db.query("SELECT 1 AS present FROM document_mutation_commits WHERE mutation_id = ?")
      .get(mutationId) ??
    db.query("SELECT 1 AS present FROM source_mutation_commits WHERE mutation_id = ?")
      .get(mutationId),
  )
}

function settleDocumentMutationCommit(db: Database, mutation: ActiveDocumentMutation) {
  const recorded = documentMutationCommitRecorded(db, mutation.mutation_id)
  const durable = durableMutationPhaseExists(mutation, mutation.databaseCommitted, "database-committed")
  if (!recorded && !durable && !legacyLineageCommitEvidence(mutation))
    throw new Error(`Document mutation ${mutation.mutation_id} cannot settle before commit.`)
  if (!durable) writeDurableMutationPhase(mutation, mutation.databaseCommitted, "database-committed")
  if (recorded) deleteDocumentMutationCommit(db, mutation.mutation_id)
}

function legacyLineageCommitEvidence(mutation: ActiveDocumentMutation) {
  if (!mutation.affects_source_projection || !existsSync(mutation.legacyLineageComplete)) return false
  const value = readFileSync(mutation.legacyLineageComplete, "utf8")
  if (value.endsWith("\n") && !Number.isNaN(Date.parse(value.trim()))) return true
  throw new DocumentMutationRecoveryError(
    `Document mutation ${mutation.mutation_id} has invalid legacy lineage evidence.`,
  )
}

function writeDurableMutationPhase(
  mutation: ActiveDocumentMutation,
  file: string,
  phase: "database-committed" | "lineage-settled",
) {
  const content = mutationPhaseContent(mutation, phase)
  if (existsSync(file)) {
    if (readFileSync(file, "utf8") !== content)
      throw new DocumentMutationRecoveryError(
        `Document mutation ${mutation.mutation_id} has conflicting ${phase} evidence.`,
      )
    return
  }
  const descriptor = openSync(file, "wx")
  writeFileSync(descriptor, content)
  fsyncSync(descriptor)
  closeSync(descriptor)
  const directory = openSync(mutation.directory, "r")
  fsyncSync(directory)
  closeSync(directory)
}

function durableMutationPhaseExists(
  mutation: ActiveDocumentMutation,
  file: string,
  phase: "database-committed" | "lineage-settled",
) {
  if (!existsSync(file)) return false
  if (readFileSync(file, "utf8") === mutationPhaseContent(mutation, phase)) return true
  throw new DocumentMutationRecoveryError(
    `Document mutation ${mutation.mutation_id} has invalid ${phase} evidence.`,
  )
}

function mutationPhaseContent(
  mutation: ActiveDocumentMutation,
  phase: "database-committed" | "lineage-settled",
) {
  return `${JSON.stringify({
    version: 1,
    mutation_id: mutation.mutation_id,
    target_name: mutation.target_name,
    phase,
  }, null, 2)}\n`
}

function readDocumentMutationLineageIntent(db: Database, mutation: ActiveDocumentMutation) {
  if (!existsSync(mutation.lineageIntent))
    return {
      documents: (listSourceDocuments(db) as { name: string }[]).map((document) => document.name),
      correlation: `recovery-${mutation.mutation_id}`,
    }
  const parsed = JSON.parse(readFileSync(mutation.lineageIntent, "utf8")) as unknown
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed))
    throw new Error(`Document mutation ${mutation.mutation_id} has invalid lineage intent.`)
  const intent = parsed as Record<string, unknown>
  if (
    intent.version !== 1 ||
    !Array.isArray(intent.documents) ||
    intent.documents.some((document) => typeof document !== "string") ||
    typeof intent.correlation !== "string"
  ) throw new Error(`Document mutation ${mutation.mutation_id} has invalid lineage intent.`)
  return { documents: intent.documents as string[], correlation: intent.correlation }
}
