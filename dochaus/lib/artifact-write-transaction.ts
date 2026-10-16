import { createHash, randomUUID } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { readDurableJson, writeDurableJson } from "./durable-json"
import { sha256File } from "./lineage"
import { withDirectoryLock } from "./directory-lock"
import { withArtifactMutationLock } from "../capabilities/artifacts/mutation-lock"

export const ARTIFACT_WRITE_CHUNK_CHAR_LIMIT = 1_200
const ARTIFACT_WRITE_CHUNK_LIMIT = 1_000
const ARTIFACT_WRITE_BYTE_LIMIT = 2_000_000
const CHECKPOINT_TAIL_CHAR_LIMIT = 240

export type ArtifactWriteIdentity = {
  run_id: string
  stage: string
  contract_hash: string
}

type ArtifactWriteChunk = {
  sequence: number
  sha256: string
  previous_hash: string
  bytes: number
  label?: string
}

type ArtifactWriteTransaction = ArtifactWriteIdentity & {
  version: 1
  transaction_id: string
  target: string
  overwrite: boolean
  status: "open" | "committed"
  chunks: ArtifactWriteChunk[]
  content_sha256: string
  bytes: number
  target_initial_sha256: string | null
  target_initial_bytes: number | null
  committed_sha256?: string
  committed_at?: string
  created_at: string
  updated_at: string
}

export function beginArtifactWrite(
  directory: string,
  identity: ArtifactWriteIdentity,
  input: { target: string; overwrite: boolean; content?: string; label?: string },
) {
  assertChunk(input.content)
  mkdirSync(artifactWriteRunDirectory(directory, identity.run_id), { recursive: true })
  const locked_transaction_ids = new Set(
    listArtifactWriteStageTransactions(directory, identity, [input.target])
      .filter((transaction) => transaction.status === "open")
      .map((transaction) => transaction.transaction_id),
  )
  const transaction = withTransactionLocks(directory, identity.run_id, [...locked_transaction_ids].toSorted(), () =>
    withTargetLock(directory, identity.run_id, input.target, () =>
      beginArtifactWriteLocked(directory, identity, input, locked_transaction_ids),
    ),
  )
  if (input.content === undefined) return transaction
  return appendArtifactWrite(directory, identity, {
    transaction_id: transaction.transaction_id,
    sequence: 0,
    previous_hash: emptyContentHash(),
    content: input.content,
    label: input.label,
  })
}

export function appendArtifactWrite(
  directory: string,
  identity: ArtifactWriteIdentity,
  input: { transaction_id: string; sequence: number; previous_hash: string; content: string; label?: string },
) {
  assertChunk(input.content)
  return withTransactionLock(directory, identity.run_id, input.transaction_id, () =>
    appendArtifactWriteLocked(directory, requireArtifactWriteTransaction(directory, identity, input.transaction_id), input),
  )
}

export function artifactWriteStatus(directory: string, identity: ArtifactWriteIdentity, transaction_id: string) {
  const transaction = requireArtifactWriteTransaction(directory, identity, transaction_id)
  return {
    transaction_id: transaction.transaction_id,
    target: transaction.target,
    status: transaction.status,
    next_sequence: transaction.chunks.length,
    checkpoint_hash: transaction.content_sha256,
    bytes: transaction.bytes,
    completed_labels: [
      ...new Set(transaction.chunks.map((chunk) => chunk.label).filter((label): label is string => Boolean(label))),
    ],
    tail: readArtifactWriteContent(directory, transaction).slice(-CHECKPOINT_TAIL_CHAR_LIMIT),
  }
}

export function commitArtifactWrite(
  directory: string,
  identity: ArtifactWriteIdentity,
  transaction_id: string,
  previous_hash: string,
) {
  return withTransactionLock(directory, identity.run_id, transaction_id, () => {
    const target = requireArtifactWriteTransaction(directory, identity, transaction_id).target
    return withTargetLock(directory, identity.run_id, target, () =>
      withArtifactMutationLock(directory, target, () => {
        const transaction = requireArtifactWriteTransaction(directory, identity, transaction_id)
        if (transaction.content_sha256 !== previous_hash)
          throw new Error(
            `Artifact write ${transaction_id} changed before commit. Its current checkpoint is ` +
              `${transaction.content_sha256}; read status and retry with that checkpoint.`,
          )
        const content = readArtifactWriteContent(directory, transaction)
        const candidate = contentHash(content)
        if (candidate !== transaction.content_sha256)
          throw new Error(`Artifact write ${transaction_id} failed its assembled-content hash check.`)
        const file = path.join(directory, transaction.target)
        const current = sha256File(file)
        if (transaction.status === "committed") {
          if (current?.sha256 !== transaction.committed_sha256)
            throw new Error(
              `${transaction.target} changed after transaction ${transaction_id} committed. ` +
                `Do not replay the old transaction over a newer artifact.`,
            )
          return { transaction, mutated: false }
        }
        if (current?.sha256 === candidate)
          return { transaction: markArtifactWriteCommitted(directory, transaction, candidate), mutated: true }
        if (artifactWriteTargetChanged(directory, transaction))
          throw new Error(
            `${transaction.target} changed after transaction ${transaction_id} began. ` +
              `Abort this transaction rather than replacing a newer artifact.`,
          )
        const temporary = `${file}.${process.pid}.${randomUUID()}.tmp`
        try {
          writeFileSync(temporary, content, { flag: "wx" })
          renameSync(temporary, file)
        } finally {
          if (existsSync(temporary)) rmSync(temporary, { force: true })
        }
        return { transaction: markArtifactWriteCommitted(directory, transaction, candidate), mutated: true }
      }),
    )
  })
}

export function abortArtifactWrite(directory: string, identity: ArtifactWriteIdentity, transaction_id: string) {
  return withTransactionLock(directory, identity.run_id, transaction_id, () => {
    const transaction = requireArtifactWriteTransaction(directory, identity, transaction_id)
    if (transaction.status === "committed")
      throw new Error(`Artifact write ${transaction_id} is already committed and cannot be aborted.`)
    rmSync(artifactWriteTransactionDirectory(directory, transaction), { recursive: true, force: true })
    return transaction
  })
}

function appendArtifactWriteLocked(
  directory: string,
  transaction: ArtifactWriteTransaction,
  input: { sequence: number; previous_hash: string; content: string; label?: string },
) {
  if (transaction.status !== "open")
    throw new Error(`Artifact write ${transaction.transaction_id} is already committed.`)
  if (transaction.bytes === 0 && transaction.chunks.length === 0 && artifactWriteTargetChanged(directory, transaction))
    throw new Error(
      `${transaction.target} changed before the first chunk was staged. ` +
        `Call begin again so the empty transaction can restart from the current target snapshot.`,
    )
  if (!Number.isSafeInteger(input.sequence) || input.sequence < 0)
    throw new Error("Artifact write sequence must be a non-negative integer.")
  const chunkHash = contentHash(input.content)
  const existing = transaction.chunks[input.sequence]
  if (existing) {
    if (
      existing.sha256 === chunkHash &&
      existing.previous_hash === input.previous_hash &&
      existing.label === normalizedLabel(input.label)
    )
      return transaction
    throw new Error(
      `Artifact write sequence ${input.sequence} was already stored with different content or checkpoint metadata.`,
    )
  }
  if (input.sequence !== transaction.chunks.length)
    throw new Error(
      `Artifact write ${transaction.transaction_id} expects sequence ${transaction.chunks.length}, not ${input.sequence}.`,
    )
  if (input.previous_hash !== transaction.content_sha256)
    throw new Error(
      `Artifact write ${transaction.transaction_id} is at checkpoint ${transaction.content_sha256}, not ` +
        `${input.previous_hash}. Read status before appending.`,
    )
  if (transaction.chunks.length >= ARTIFACT_WRITE_CHUNK_LIMIT)
    throw new Error(
      `Artifact write ${transaction.transaction_id} reached its ${ARTIFACT_WRITE_CHUNK_LIMIT}-chunk limit.`,
    )
  const bytes = Buffer.byteLength(input.content, "utf8")
  if (transaction.bytes + bytes > ARTIFACT_WRITE_BYTE_LIMIT)
    throw new Error(`Artifact write ${transaction.transaction_id} exceeds its ${ARTIFACT_WRITE_BYTE_LIMIT}-byte limit.`)
  writeArtifactWriteChunk(directory, transaction, input.sequence, input.content)
  const next: ArtifactWriteTransaction = {
    ...transaction,
    chunks: [
      ...transaction.chunks,
      {
        sequence: input.sequence,
        sha256: chunkHash,
        previous_hash: input.previous_hash,
        bytes,
        ...(normalizedLabel(input.label) ? { label: normalizedLabel(input.label) } : {}),
      },
    ],
    content_sha256: contentHash(`${readArtifactWriteContent(directory, transaction)}${input.content}`),
    bytes: transaction.bytes + bytes,
    updated_at: new Date().toISOString(),
  }
  writeArtifactWriteTransaction(directory, next)
  return next
}

function beginArtifactWriteLocked(
  directory: string,
  identity: ArtifactWriteIdentity,
  input: { target: string; overwrite: boolean },
  locked_transaction_ids: ReadonlySet<string>,
) {
  const open = listArtifactWriteStageTransactions(directory, identity, [input.target]).filter(
    (transaction) => transaction.status === "open",
  )
  const discarded = open.filter(
    (transaction) =>
      transaction.bytes === 0 &&
      transaction.chunks.length === 0 &&
      (transaction.contract_hash !== identity.contract_hash || artifactWriteTargetChanged(directory, transaction)),
  )
  const unlocked = discarded.find((transaction) => !locked_transaction_ids.has(transaction.transaction_id))
  if (unlocked)
    throw new Error(
      `Artifact write state for ${input.target} changed while a new transaction was beginning. ` +
        `Retry begin so transaction ${unlocked.transaction_id} can be reconciled safely.`,
    )
  const conflicting = open.find(
    (transaction) => transaction.contract_hash !== identity.contract_hash && !discarded.includes(transaction),
  )
  if (conflicting)
    throw new Error(
      `Open artifact write ${conflicting.transaction_id} for ${input.target} belongs to a different input contract ` +
        `and contains durable progress. Resume or explicitly abort that transaction before starting the new contract.`,
    )
  discarded.forEach((transaction) => discardArtifactWriteTransaction(directory, transaction))
  const existing = open.find(
    (transaction) => transaction.contract_hash === identity.contract_hash && !discarded.includes(transaction),
  )
  if (existing) {
    if (existing.overwrite !== input.overwrite)
      throw new Error(
        `An open artifact write already exists for ${input.target} with a different overwrite policy. ` +
          `Resume transaction ${existing.transaction_id}, or abort it before starting another one.`,
      )
    return existing
  }

  const target = path.join(directory, input.target)
  if (existsSync(target) && !input.overwrite)
    throw new Error(
      `${input.target} already exists in this matter. Set overwrite=true only when this stage is explicitly allowed to replace it.`,
    )
  const initial = sha256File(target)
  const transaction: ArtifactWriteTransaction = {
    version: 1,
    ...identity,
    transaction_id: `aw_${randomUUID()}`,
    target: input.target,
    overwrite: input.overwrite,
    status: "open",
    chunks: [],
    content_sha256: emptyContentHash(),
    bytes: 0,
    target_initial_sha256: initial?.sha256 ?? null,
    target_initial_bytes: initial?.bytes ?? null,
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  }
  mkdirSync(artifactWriteTransactionDirectory(directory, transaction), { recursive: true })
  writeArtifactWriteTransaction(directory, transaction)
  return transaction
}

function listArtifactWriteStageTransactions(
  directory: string,
  identity: Pick<ArtifactWriteIdentity, "run_id" | "stage">,
  targets: readonly string[],
) {
  const root = artifactWriteRunDirectory(directory, identity.run_id)
  if (!existsSync(root)) return []
  const allowed = new Set(targets)
  return readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && /^aw_[a-f0-9-]+$/i.test(entry.name))
    .map((entry) => readArtifactWriteTransactionFile(path.join(root, entry.name, "transaction.json")))
    .filter((transaction): transaction is ArtifactWriteTransaction => Boolean(transaction))
    .filter((transaction) => transaction.stage === identity.stage && allowed.has(transaction.target))
}

function requireArtifactWriteTransaction(directory: string, identity: ArtifactWriteIdentity, transaction_id: string) {
  assertTransactionID(transaction_id)
  const transaction = readArtifactWriteTransactionFile(
    path.join(artifactWriteRunDirectory(directory, identity.run_id), transaction_id, "transaction.json"),
  )
  if (!transaction) throw new Error(`Artifact write transaction ${transaction_id} does not exist in the current run.`)
  if (
    transaction.run_id !== identity.run_id ||
    transaction.stage !== identity.stage ||
    transaction.contract_hash !== identity.contract_hash
  )
    throw new Error(
      `Artifact write transaction ${transaction_id} belongs to a different run, stage, or input contract. ` +
        `Start a new transaction for the current output contract.`,
    )
  return transaction
}

function readArtifactWriteTransactionFile(file: string): ArtifactWriteTransaction | undefined {
  const parsed = readDurableJson(file)
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return
  const transaction = parsed as Partial<ArtifactWriteTransaction>
  if (
    transaction.version !== 1 ||
    typeof transaction.transaction_id !== "string" ||
    typeof transaction.run_id !== "string" ||
    typeof transaction.stage !== "string" ||
    typeof transaction.contract_hash !== "string" ||
    typeof transaction.target !== "string" ||
    typeof transaction.overwrite !== "boolean" ||
    !["open", "committed"].includes(transaction.status ?? "") ||
    !Array.isArray(transaction.chunks) ||
    typeof transaction.content_sha256 !== "string" ||
    typeof transaction.bytes !== "number" ||
    typeof transaction.created_at !== "string" ||
    !Number.isFinite(Date.parse(transaction.created_at)) ||
    typeof transaction.updated_at !== "string" ||
    !Number.isFinite(Date.parse(transaction.updated_at))
  )
    return
  return transaction as ArtifactWriteTransaction
}

function readArtifactWriteContent(directory: string, transaction: ArtifactWriteTransaction) {
  return transaction.chunks
    .map((chunk) => {
      const content = readFileSync(artifactWriteChunkFile(directory, transaction, chunk.sequence), "utf8")
      if (contentHash(content) !== chunk.sha256)
        throw new Error(
          `Artifact write ${transaction.transaction_id} chunk ${chunk.sequence} failed its content hash check.`,
        )
      return content
    })
    .join("")
}

function writeArtifactWriteChunk(
  directory: string,
  transaction: ArtifactWriteTransaction,
  sequence: number,
  content: string,
) {
  const file = artifactWriteChunkFile(directory, transaction, sequence)
  mkdirSync(path.dirname(file), { recursive: true })
  const temporary = `${file}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(temporary, content, { flag: "wx" })
    renameSync(temporary, file)
  } finally {
    if (existsSync(temporary)) rmSync(temporary, { force: true })
  }
}

function writeArtifactWriteTransaction(directory: string, transaction: ArtifactWriteTransaction) {
  writeDurableJson(
    path.join(artifactWriteTransactionDirectory(directory, transaction), "transaction.json"),
    transaction,
  )
}

function markArtifactWriteCommitted(
  directory: string,
  transaction: ArtifactWriteTransaction,
  committed_sha256: string,
) {
  const now = new Date().toISOString()
  const next: ArtifactWriteTransaction = {
    ...transaction,
    status: "committed",
    committed_sha256,
    committed_at: now,
    updated_at: now,
  }
  writeArtifactWriteTransaction(directory, next)
  return next
}

function artifactWriteRunDirectory(directory: string, run_id: string) {
  return path.join(directory, ".dochaus", "artifact-writes", safeSegment(run_id))
}

function artifactWriteTransactionDirectory(directory: string, transaction: ArtifactWriteTransaction) {
  return path.join(artifactWriteRunDirectory(directory, transaction.run_id), transaction.transaction_id)
}

function discardArtifactWriteTransaction(directory: string, transaction: ArtifactWriteTransaction) {
  const current = artifactWriteTransactionDirectory(directory, transaction)
  const discarded = `${current}.discarded.${process.pid}.${randomUUID()}`
  renameSync(current, discarded)
  rmSync(discarded, { recursive: true, force: true })
}

function artifactWriteChunkFile(directory: string, transaction: ArtifactWriteTransaction, sequence: number) {
  return path.join(artifactWriteTransactionDirectory(directory, transaction), "chunks", `${sequence}.txt`)
}

function artifactWriteTargetChanged(directory: string, transaction: ArtifactWriteTransaction) {
  const current = sha256File(path.join(directory, transaction.target))
  if (!transaction.target_initial_sha256) return Boolean(current)
  return (
    !current ||
    current.sha256 !== transaction.target_initial_sha256 ||
    current.bytes !== transaction.target_initial_bytes
  )
}

function withTransactionLock<T>(directory: string, run_id: string, transaction_id: string, fn: () => T) {
  assertTransactionID(transaction_id)
  return withLock(path.join(artifactWriteRunDirectory(directory, run_id), transaction_id, ".lock"), fn)
}

function withTransactionLocks<T>(
  directory: string,
  run_id: string,
  transaction_ids: readonly string[],
  fn: () => T,
): T {
  const transaction_id = transaction_ids[0]
  if (!transaction_id) return fn()
  return withTransactionLock(directory, run_id, transaction_id, () =>
    withTransactionLocks(directory, run_id, transaction_ids.slice(1), fn),
  )
}

function withTargetLock<T>(directory: string, run_id: string, target: string, fn: () => T) {
  const digest = createHash("sha256").update(target).digest("hex")
  return withLock(path.join(artifactWriteRunDirectory(directory, run_id), "locks", digest), fn)
}

function withLock<T>(lock: string, fn: () => T) {
  return withDirectoryLock(
    lock,
    "This artifact write is already being updated. Read its status and retry once the active call finishes.",
    fn,
  )
}

function assertChunk(content: string | undefined) {
  if (content === undefined) return
  if (content.length > ARTIFACT_WRITE_CHUNK_CHAR_LIMIT)
    throw new Error(
      `Artifact write chunks must be at most ${ARTIFACT_WRITE_CHUNK_CHAR_LIMIT} characters; this chunk has ${content.length}.`,
    )
}

function assertTransactionID(transaction_id: string) {
  if (!/^aw_[a-f0-9-]+$/i.test(transaction_id)) throw new Error("Artifact write transaction ID is invalid.")
}

function normalizedLabel(label: string | undefined) {
  const value = label?.trim()
  if (!value) return undefined
  if (value.length > 120) throw new Error("Artifact write labels must be at most 120 characters.")
  return value
}

function safeSegment(value: string) {
  return value.replace(/[^a-zA-Z0-9._-]+/g, "_")
}

function contentHash(content: string) {
  return createHash("sha256").update(content).digest("hex")
}

function emptyContentHash() {
  return contentHash("")
}
