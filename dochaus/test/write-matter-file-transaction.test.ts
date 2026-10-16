import { describe, expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  abortArtifactWrite,
  appendArtifactWrite,
  ARTIFACT_WRITE_CHUNK_CHAR_LIMIT,
  artifactWriteStatus,
  beginArtifactWrite,
  commitArtifactWrite,
  type ArtifactWriteIdentity,
} from "../lib/artifact-write-transaction"
import writeMatterFile from "../tool/write-matter-file"

const identity: ArtifactWriteIdentity = {
  run_id: "run_transaction_test",
  stage: "test-stage",
  contract_hash: "contract-v1",
}

function matter() {
  return mkdtempSync(path.join(tmpdir(), "artifact-write-transaction-"))
}

describe("transactional matter artifact writes", () => {
  test("does not register the obsolete content-bearing drafting tool", () => {
    expect(existsSync(path.join(import.meta.dir, "../tool/write-drafting-artifact.ts"))).toBe(false)
  })

  test("assembles a large artifact as exact bytes and publishes only at commit", () => {
    const dir = matter()
    const chunks = Array.from(
      { length: 20 },
      (_, index) => `${index}:${"x".repeat(ARTIFACT_WRITE_CHUNK_CHAR_LIMIT - 100)}\n`,
    )
    let transaction = beginArtifactWrite(dir, identity, {
      target: "large-output.md",
      overwrite: false,
      content: chunks[0],
      label: "part-0",
    })

    expect(existsSync(path.join(dir, "large-output.md"))).toBe(false)
    for (const [sequence, content] of chunks.slice(1).entries())
      transaction = appendArtifactWrite(dir, identity, {
        transaction_id: transaction.transaction_id,
        sequence: sequence + 1,
        previous_hash: transaction.content_sha256,
        content,
        label: `part-${sequence + 1}`,
      })

    expect(artifactWriteStatus(dir, identity, transaction.transaction_id)).toEqual(
      expect.objectContaining({ next_sequence: chunks.length, bytes: Buffer.byteLength(chunks.join("")) }),
    )
    expect(existsSync(path.join(dir, "large-output.md"))).toBe(false)

    const committed = commitArtifactWrite(dir, identity, transaction.transaction_id, transaction.content_sha256)
    expect(committed.mutated).toBe(true)
    expect(readFileSync(path.join(dir, "large-output.md"), "utf8")).toBe(chunks.join(""))
  })

  test("keeps interrupted chunks durable without satisfying the target artifact", () => {
    const dir = matter()
    const transaction = beginArtifactWrite(dir, identity, {
      target: "interrupted.md",
      overwrite: false,
      content: "# Partial\n",
      label: "heading",
    })
    const checkpoint = artifactWriteStatus(dir, identity, transaction.transaction_id)

    expect(existsSync(path.join(dir, "interrupted.md"))).toBe(false)
    expect(checkpoint.next_sequence).toBe(1)
    expect(checkpoint.completed_labels).toEqual(["heading"])
    expect(checkpoint.tail).toBe("# Partial\n")
    expect(checkpoint.status).toBe("open")
  })

  test("makes exact chunk retries idempotent and rejects conflicting or out-of-order retries", () => {
    const dir = matter()
    const transaction = beginArtifactWrite(dir, identity, {
      target: "ordered.md",
      overwrite: false,
      content: "first",
      label: "first",
    })
    const duplicate = appendArtifactWrite(dir, identity, {
      transaction_id: transaction.transaction_id,
      sequence: 0,
      previous_hash: transaction.chunks[0]!.previous_hash,
      content: "first",
      label: "first",
    })
    expect(duplicate.chunks).toHaveLength(1)
    expect(() =>
      appendArtifactWrite(dir, identity, {
        transaction_id: transaction.transaction_id,
        sequence: 0,
        previous_hash: transaction.chunks[0]!.previous_hash,
        content: "different",
        label: "first",
      }),
    ).toThrow("different content or checkpoint metadata")
    expect(() =>
      appendArtifactWrite(dir, identity, {
        transaction_id: transaction.transaction_id,
        sequence: 2,
        previous_hash: transaction.content_sha256,
        content: "third",
      }),
    ).toThrow("expects sequence 1")
  })

  test("rejects oversized chunks and a stale input contract", () => {
    const dir = matter()
    expect(() =>
      beginArtifactWrite(dir, identity, {
        target: "oversized.md",
        overwrite: false,
        content: "x".repeat(ARTIFACT_WRITE_CHUNK_CHAR_LIMIT + 1),
      }),
    ).toThrow(`must be at most ${ARTIFACT_WRITE_CHUNK_CHAR_LIMIT} characters`)

    const transaction = beginArtifactWrite(dir, identity, {
      target: "scoped.md",
      overwrite: false,
      content: "safe",
    })
    expect(() =>
      artifactWriteStatus(
        dir,
        { ...identity, contract_hash: "changed-input-contract" },
        transaction.transaction_id,
      ),
    ).toThrow("different run, stage, or input contract")
  })

  test("compare-and-swap protects both existing and newly created targets", () => {
    const existingDir = matter()
    writeFileSync(path.join(existingDir, "replace.md"), "original")
    const replacement = beginArtifactWrite(existingDir, identity, {
      target: "replace.md",
      overwrite: true,
      content: "replacement",
    })
    writeFileSync(path.join(existingDir, "replace.md"), "newer")
    expect(() =>
      commitArtifactWrite(existingDir, identity, replacement.transaction_id, replacement.content_sha256),
    ).toThrow("changed after transaction")
    expect(readFileSync(path.join(existingDir, "replace.md"), "utf8")).toBe("newer")

    const absentDir = matter()
    const fresh = beginArtifactWrite(absentDir, identity, {
      target: "fresh.md",
      overwrite: false,
      content: "candidate",
    })
    writeFileSync(path.join(absentDir, "fresh.md"), "external")
    expect(() => commitArtifactWrite(absentDir, identity, fresh.transaction_id, fresh.content_sha256)).toThrow(
      "changed after transaction",
    )
    expect(readFileSync(path.join(absentDir, "fresh.md"), "utf8")).toBe("external")
  })

  test("restarts only an empty open transaction when its target changed", () => {
    const dir = matter()
    writeFileSync(path.join(dir, "restart.md"), "original")
    const empty = beginArtifactWrite(dir, identity, {
      target: "restart.md",
      overwrite: true,
    })
    writeFileSync(path.join(dir, "restart.md"), "newer")
    expect(() =>
      appendArtifactWrite(dir, identity, {
        transaction_id: empty.transaction_id,
        sequence: 0,
        previous_hash: empty.content_sha256,
        content: "stale candidate",
      }),
    ).toThrow("Call begin again")

    const restarted = beginArtifactWrite(dir, identity, {
      target: "restart.md",
      overwrite: true,
    })
    expect(restarted.transaction_id).not.toBe(empty.transaction_id)
    expect(restarted.target_initial_sha256).not.toBe(empty.target_initial_sha256)
    expect(restarted.target_initial_bytes).toBe(Buffer.byteLength("newer"))
    expect(() => artifactWriteStatus(dir, identity, empty.transaction_id)).toThrow(
      "does not exist in the current run",
    )

    const progressed = appendArtifactWrite(dir, identity, {
      transaction_id: restarted.transaction_id,
      sequence: 0,
      previous_hash: restarted.content_sha256,
      content: "candidate",
    })
    writeFileSync(path.join(dir, "restart.md"), "newest")
    const resumed = beginArtifactWrite(dir, identity, {
      target: "restart.md",
      overwrite: true,
    })
    expect(resumed.transaction_id).toBe(progressed.transaction_id)
    expect(resumed.bytes).toBe(Buffer.byteLength("candidate"))
    expect(() => commitArtifactWrite(dir, identity, resumed.transaction_id, resumed.content_sha256)).toThrow(
      "changed after transaction",
    )
    expect(readFileSync(path.join(dir, "restart.md"), "utf8")).toBe("newest")
  })

  test("reconciles empty prior-contract writes without discarding durable progress", () => {
    const dir = matter()
    const oldIdentity = { ...identity, contract_hash: "contract-old" }
    const newIdentity = { ...identity, contract_hash: "contract-new" }
    writeFileSync(path.join(dir, "epoch.md"), "unchanged")
    const empty = beginArtifactWrite(dir, oldIdentity, {
      target: "epoch.md",
      overwrite: true,
    })

    const current = beginArtifactWrite(dir, newIdentity, {
      target: "epoch.md",
      overwrite: true,
    })
    expect(current.transaction_id).not.toBe(empty.transaction_id)
    expect(current.contract_hash).toBe(newIdentity.contract_hash)
    expect(() => artifactWriteStatus(dir, oldIdentity, empty.transaction_id)).toThrow(
      "does not exist in the current run",
    )

    const progressed = beginArtifactWrite(dir, oldIdentity, {
      target: "preserved.md",
      overwrite: false,
      content: "durable partial output",
    })
    expect(() =>
      beginArtifactWrite(dir, newIdentity, {
        target: "preserved.md",
        overwrite: false,
      }),
    ).toThrow("belongs to a different input contract and contains durable progress")
    expect(artifactWriteStatus(dir, oldIdentity, progressed.transaction_id)).toEqual(
      expect.objectContaining({
        transaction_id: progressed.transaction_id,
        next_sequence: 1,
        bytes: Buffer.byteLength("durable partial output"),
      }),
    )
  })

  test("commits once, reports an exact replay, and supports abort before publication", () => {
    const dir = matter()
    const transaction = beginArtifactWrite(dir, identity, {
      target: "once.md",
      overwrite: false,
      content: "body",
    })
    expect(commitArtifactWrite(dir, identity, transaction.transaction_id, transaction.content_sha256).mutated).toBe(
      true,
    )
    expect(artifactWriteStatus(dir, identity, transaction.transaction_id)).toEqual(
      expect.objectContaining({ transaction_id: transaction.transaction_id, target: "once.md", status: "committed" }),
    )
    expect(commitArtifactWrite(dir, identity, transaction.transaction_id, transaction.content_sha256).mutated).toBe(
      false,
    )
    expect(() => abortArtifactWrite(dir, identity, transaction.transaction_id)).toThrow("already committed")
    writeFileSync(path.join(dir, "once.md"), "newer")
    expect(() =>
      commitArtifactWrite(dir, identity, transaction.transaction_id, transaction.content_sha256),
    ).toThrow("changed after transaction")

    const abandoned = beginArtifactWrite(dir, identity, {
      target: "abandoned.md",
      overwrite: false,
      content: "partial",
    })
    abortArtifactWrite(dir, identity, abandoned.transaction_id)
    expect(existsSync(path.join(dir, "abandoned.md"))).toBe(false)
    expect(() => artifactWriteStatus(dir, identity, abandoned.transaction_id)).toThrow(
      "does not exist in the current run",
    )
  })

  test("ignores a transaction with a corrupt timestamp", () => {
    const dir = matter()
    const transaction = beginArtifactWrite(dir, identity, {
      target: "corrupt-time.md",
      overwrite: false,
      content: "partial",
    })
    const file = path.join(
      dir,
      ".dochaus",
      "artifact-writes",
      identity.run_id,
      transaction.transaction_id,
      "transaction.json",
    )
    const parsed = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>
    writeFileSync(file, JSON.stringify({ ...parsed, created_at: "not-a-timestamp" }))
    expect(() => artifactWriteStatus(dir, identity, transaction.transaction_id)).toThrow(
      "does not exist in the current run",
    )
  })

  test("keeps direct small writes compatible and refuses oversized direct writes", async () => {
    const dir = matter()
    await writeMatterFile.execute({ action: "write", name: "small.md", content: "small", overwrite: false }, {
      directory: dir,
    } as never)
    expect(readFileSync(path.join(dir, "small.md"), "utf8")).toBe("small\n")
    await expect(
      writeMatterFile.execute(
        {
          action: "write",
          name: "too-large.md",
          content: "x".repeat(ARTIFACT_WRITE_CHUNK_CHAR_LIMIT + 1),
          overwrite: false,
          __artifact_write_direct_limit: true,
        } as never,
        { directory: dir } as never,
      ),
    ).rejects.toThrow("Use begin, append, and commit")

    await writeMatterFile.execute(
      {
        action: "write",
        name: "non-pipeline-large.md",
        content: "x".repeat(ARTIFACT_WRITE_CHUNK_CHAR_LIMIT + 1),
        overwrite: false,
      } as never,
      { directory: dir } as never,
    )
    expect(readFileSync(path.join(dir, "non-pipeline-large.md"), "utf8")).toHaveLength(
      ARTIFACT_WRITE_CHUNK_CHAR_LIMIT + 2,
    )
  })
})
