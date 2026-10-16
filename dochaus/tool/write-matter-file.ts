import { tool } from "./legacy-definition"
import { existsSync, writeFileSync } from "node:fs"
import path from "node:path"
import { withArtifactMutationLock } from "../capabilities/artifacts/mutation-lock"
import {
  abortArtifactWrite,
  appendArtifactWrite,
  ARTIFACT_WRITE_CHUNK_CHAR_LIMIT,
  artifactWriteStatus,
  beginArtifactWrite,
  commitArtifactWrite,
  type ArtifactWriteIdentity,
} from "../lib/artifact-write-transaction"
import { matterMarkdownName } from "./matter-markdown-name"

// Matter-scoped markdown writer for drafting intermediates. Agents use this for
// generated .md artifacts only; canonical .docx output still goes through render.

export default tool({
  description:
    `Write one matter artifact. Inside a typed drafting stage, a complete write must stay at or below ${ARTIFACT_WRITE_CHUNK_CHAR_LIMIT} content characters. ` +
    `For a larger typed-stage output, use the transactional actions begin, append, status, and commit: begin with the target and an optional first chunk; ` +
    `append exact consecutive chunks using the returned transaction ID, next sequence, and checkpoint hash; status resumes an interrupted write; ` +
    `commit atomically publishes the exact staged bytes after full validation. Exact repeated chunks are idempotent. Abort discards an uncommitted transaction. ` +
    `Complete writes outside a typed drafting output contract retain their normal behavior.`,
  args: {
    action: tool.schema
      .enum(["write", "begin", "append", "status", "commit", "abort"])
      .optional()
      .describe('"write" (default) for a complete small artifact, or a transactional action for a larger artifact'),
    name: tool.schema.string().min(1).optional().describe("Matter artifact basename; required for write and begin"),
    content: tool.schema.string().optional().describe("Complete small artifact content, or one transactional chunk"),
    overwrite: tool.schema.boolean().optional().describe("Allow replacing an existing file (default false)"),
    transaction_id: tool.schema.string().optional().describe("Transaction ID returned by begin or status"),
    sequence: tool.schema
      .number()
      .int()
      .min(0)
      .optional()
      .describe("Exact next chunk sequence returned by begin or status"),
    previous_hash: tool.schema
      .string()
      .optional()
      .describe("Exact checkpoint hash returned by begin, append, or status"),
    label: tool.schema
      .string()
      .max(120)
      .optional()
      .describe("Optional stable label for the complete logical unit in this chunk"),
  },
  async execute(args, ctx) {
    const action = args.action ?? "write"
    if (action !== "write") {
      const identity = artifactWriteIdentity(args)
      if (!identity)
        throw new Error(
          "Transactional artifact writes require a drafting stage with a declared output contract. " +
            "Use this tool only for an output assigned to the current stage.",
        )
      if (action === "begin") {
        const name = requireMatterName(args.name)
        const transaction = beginArtifactWrite(ctx.directory, identity, {
          target: name,
          overwrite: args.overwrite === true,
          content: args.content,
          label: args.label,
        })
        return artifactWriteResult(ctx.directory, identity, transaction.transaction_id, "Staged artifact write")
      }
      if (action === "append") {
        if (args.content === undefined) throw new Error("Artifact write append requires content.")
        if (args.sequence === undefined) throw new Error("Artifact write append requires sequence.")
        const transaction_id = requireTransactionID(args.transaction_id)
        const transaction = appendArtifactWrite(ctx.directory, identity, {
          transaction_id,
          sequence: args.sequence,
          previous_hash: requireCheckpointHash(args.previous_hash),
          content: args.content,
          label: args.label,
        })
        return artifactWriteResult(ctx.directory, identity, transaction.transaction_id, "Staged artifact chunk")
      }
      if (action === "status")
        return artifactWriteResult(
          ctx.directory,
          identity,
          requireTransactionID(args.transaction_id),
          "Artifact write status",
        )
      if (action === "abort") {
        const transaction = abortArtifactWrite(ctx.directory, identity, requireTransactionID(args.transaction_id))
        return {
          title: `Aborted ${transaction.target}`,
          output: `Discarded uncommitted artifact write ${transaction.transaction_id} for ${transaction.target}.`,
          metadata: { artifactWriteAction: "abort", artifactWriteTransaction: transaction.transaction_id },
        }
      }
      const result = commitArtifactWrite(
        ctx.directory,
        identity,
        requireTransactionID(args.transaction_id),
        requireCheckpointHash(args.previous_hash),
      )
      return {
        title: `Wrote ${result.transaction.target}`,
        output: result.mutated
          ? `Atomically committed ${result.transaction.target} to this matter directory.`
          : `${result.transaction.target} was already committed by this transaction.`,
        metadata: {
          document: result.transaction.target,
          bytes: result.transaction.bytes,
          artifactWriteAction: "commit",
          artifactWriteTransaction: result.transaction.transaction_id,
          artifactWriteCommitted: result.mutated,
        },
      }
    }

    if (args.content === undefined) throw new Error("A complete artifact write requires content.")
    if (artifactWriteDirectLimit(args) && args.content.length > ARTIFACT_WRITE_CHUNK_CHAR_LIMIT)
      throw new Error(
        `Complete artifact writes are limited to ${ARTIFACT_WRITE_CHUNK_CHAR_LIMIT} content characters. ` +
          `Use begin, append, and commit for a larger artifact.`,
      )
    if (!args.name) throw new Error("A complete artifact write requires a target name.")
    const parsed = matterMarkdownName(args.name)
    if (!parsed.ok) return parsed.message
    const name = parsed.name

    const targetPath = path.join(ctx.directory, name)
    const written = withArtifactMutationLock(ctx.directory, name, () => {
      if (existsSync(targetPath) && !args.overwrite) return false
      writeFileSync(targetPath, args.content!.endsWith("\n") ? args.content! : `${args.content}\n`)
      return true
    })
    if (!written)
      return `${name} already exists in this matter. Pass overwrite=true only if this stage is explicitly allowed to update it.`

    return {
      title: `Wrote ${name}`,
      output: `Wrote ${name} to this matter directory.`,
      metadata: { document: name, bytes: Buffer.byteLength(args.content, "utf8") },
    }
  },
})

function artifactWriteIdentity(args: object) {
  const identity = (args as { __artifact_write_identity?: ArtifactWriteIdentity }).__artifact_write_identity
  if (!identity?.run_id || !identity.stage || !identity.contract_hash) return
  return identity
}

function artifactWriteDirectLimit(args: object) {
  return (args as { __artifact_write_direct_limit?: boolean }).__artifact_write_direct_limit === true
}

function artifactWriteResult(
  directory: string,
  identity: ArtifactWriteIdentity,
  transaction_id: string,
  title: string,
) {
  const checkpoint = artifactWriteStatus(directory, identity, transaction_id)
  return {
    title,
    output:
      `Artifact write checkpoint:\n` +
      `transaction_id: ${checkpoint.transaction_id}\n` +
      `target: ${checkpoint.target}\n` +
      `next_sequence: ${checkpoint.next_sequence}\n` +
      `checkpoint_hash: ${checkpoint.checkpoint_hash}\n` +
      `bytes: ${checkpoint.bytes}\n` +
      `completed_labels: ${checkpoint.completed_labels.join(", ") || "none"}\n` +
      `tail: ${JSON.stringify(checkpoint.tail)}`,
    metadata: {
      artifactWriteAction: "checkpoint",
      artifactWriteTransaction: checkpoint.transaction_id,
      artifactWriteTarget: checkpoint.target,
      artifactWriteNextSequence: checkpoint.next_sequence,
      artifactWriteCheckpointHash: checkpoint.checkpoint_hash,
    },
  }
}

function requireMatterName(value: string | undefined) {
  if (!value) throw new Error("Artifact write requires a target name.")
  const parsed = matterMarkdownName(value)
  if (!parsed.ok) throw new Error(parsed.message)
  return parsed.name
}

function requireTransactionID(value: string | undefined) {
  if (!value) throw new Error("Artifact write requires transaction_id.")
  return value
}

function requireCheckpointHash(value: string | undefined) {
  if (!value) throw new Error("Artifact write requires previous_hash from the latest checkpoint.")
  return value
}
