// Ingest-side lineage appender (#96). Self-contained — the ingest service cannot
// import dochaus/lib — producing the SAME row shape as dochaus/lib/lineage.ts.
// An ingest stamps one PERMANENT row per present source artifact; a rebuild of
// the source projection stamps only the artifacts it changed. `supersedes` links
// each row to the prior row for that artifact. Canonical schema documented in
// dochaus/lib/lineage.ts.
//
// A row's hash and size are those of the artifact's bytes. The caller may pass
// them for the files it just installed (`files`, see stageSourceArtifacts):
// reading thousands of chunk files back only to hash bytes it wrote a moment
// ago held the event loop for seconds on a large matter. Any present artifact
// not in `files` (legal.db, or a file the caller did not write) is read and
// hashed here.
import { createHash } from "node:crypto"
import { appendFileSync, existsSync, mkdirSync, readFileSync } from "node:fs"
import path from "node:path"
import { withArtifactMutationLock } from "../../../dochaus/capabilities/artifacts/mutation-lock"

const CROCKFORD = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"

function ulid(): string {
  const ts = Date.now()
  let time = ""
  let t = ts
  for (let i = 0; i < 10; i++) {
    time = CROCKFORD[t % 32] + time
    t = Math.floor(t / 32)
  }
  const randomBytes = new Uint8Array(16)
  globalThis.crypto.getRandomValues(randomBytes)
  let random = ""
  for (let i = 0; i < 16; i++) random += CROCKFORD[randomBytes[i]! % 32]!
  return time + random
}

function lineageFile(matterDir: string): string {
  return path.join(matterDir, ".dochaus", "lineage.jsonl")
}

// The latest lineage row per artifact: its entry id and recorded content hash.
function priorEntryByArtifact(matterDir: string): Map<string, { entry_id: string; content_sha256: unknown }> {
  const map = new Map<string, { entry_id: string; content_sha256: unknown }>()
  const file = lineageFile(matterDir)
  if (!existsSync(file)) return map
  for (const line of readFileSync(file, "utf8").split(/\r?\n/)) {
    if (!line.trim()) continue
    try {
      const row = JSON.parse(line) as { artifact?: string; entry_id?: string; content_sha256?: unknown }
      if (typeof row.artifact === "string" && typeof row.entry_id === "string")
        map.set(row.artifact, { entry_id: row.entry_id, content_sha256: row.content_sha256 })
    } catch {
      /* skip corrupt line */
    }
  }
  return map
}

export type ArtifactHash = { sha256: string; bytes: number }

function hashFile(absPath: string): ArtifactHash | null {
  if (!existsSync(absPath)) return null
  const buf = readFileSync(absPath)
  return { sha256: createHash("sha256").update(buf).digest("hex"), bytes: buf.length }
}

function artifactsOnDisk(matterDir: string): string[] {
  const top = [".dochaus/legal.db", "source-index.json", "source-map.json", "source-map.md", "source-ledger.json", "matter-identifiers.json", "matter-identifiers.md"]
    .filter((artifact) => existsSync(path.join(matterDir, artifact)))
  const glob = (pat: string) => [...new Bun.Glob(pat).scanSync({ cwd: matterDir, onlyFiles: true })].map((p) => p.replace(/\\/g, "/"))
  return [...top, ...glob("source-chunks/*.txt"), ...glob("sources/*.txt")]
}

// Returns the number of artifacts stamped.
export function stampIngestLineage(
  matterDir: string,
  documents: string[],
  correlation: string,
  mutationID?: string,
  files?: ReadonlyMap<string, ArtifactHash>,
): number {
  return withArtifactMutationLock(matterDir, ".dochaus/lineage.jsonl", () => {
    const completed = mutationID ? completedMutationArtifacts(matterDir, mutationID) : new Set<string>()
    return appendArtifactRows(matterDir, {
      artifacts: artifactsOnDisk(matterDir),
      op: "ingest",
      run: `ingest:${correlation}`,
      provenance: {
        origin: "upload",
        documents,
        ...(mutationID ? { document_mutation_id: mutationID } : {}),
      },
      stamp: (artifact, sha256) => !completed.has(`${artifact}\0${sha256}`),
      files,
    })
  })
}

// Stamps a rebuild of the source projection from the registry (the startup
// self-heal). The rebuild re-derives files from sources already ingested and
// uploads nothing, so each row is a whole-file "write" with origin
// "source-rebuild". Only artifacts whose bytes differ from their latest lineage
// row (or that have none) are stamped, so every file the rebuild changed gets a
// current row and unchanged ones get none. legal.db is never written by a
// rebuild and is not stamped. Returns the number of artifacts stamped.
export function stampSourceRebuildLineage(
  matterDir: string,
  documents: string[],
  correlation: string,
  files?: ReadonlyMap<string, ArtifactHash>,
): number {
  return withArtifactMutationLock(matterDir, ".dochaus/lineage.jsonl", () =>
    appendArtifactRows(matterDir, {
      artifacts: artifactsOnDisk(matterDir).filter((artifact) => artifact !== ".dochaus/legal.db"),
      op: "write",
      run: `ingest:${correlation}`,
      provenance: { origin: "source-rebuild", documents },
      // A file with no lineage row yet is stamped too, so it becomes ledgered.
      stamp: (_artifact, sha256, prior) => prior?.content_sha256 !== sha256,
      files,
    }))
}

// Appends one PERMANENT row per present artifact the caller chooses to stamp.
// Runs under the caller's lineage lock.
function appendArtifactRows(
  matterDir: string,
  input: {
    artifacts: string[]
    op: "ingest" | "write"
    run: string
    provenance: Record<string, unknown>
    stamp: (artifact: string, sha256: string, prior: { content_sha256: unknown } | undefined) => boolean
    files?: ReadonlyMap<string, ArtifactHash>
  },
) {
  mkdirSync(path.join(matterDir, ".dochaus"), { recursive: true })
  const prior = priorEntryByArtifact(matterDir)
  const ts = new Date().toISOString()
  const lines = input.artifacts.flatMap((artifact) => {
    const hash = input.files?.get(artifact) ?? hashFile(path.join(matterDir, artifact))
    if (!hash || !input.stamp(artifact, hash.sha256, prior.get(artifact))) return []
    return [
      JSON.stringify({
        entry_id: ulid(),
        ts,
        tier: "PERMANENT",
        artifact,
        op: input.op,
        run: input.run,
        stage: null,
        subagent: "ingest",
        content_sha256: hash.sha256,
        bytes: hash.bytes,
        supersedes: prior.get(artifact)?.entry_id ?? null,
        provenance: input.provenance,
      }),
    ]
  })
  if (lines.length) {
    const file = lineageFile(matterDir)
    const existing = existsSync(file) ? readFileSync(file, "utf8") : ""
    const prefix = existing.length && !existing.endsWith("\n") ? "\n" : ""
    appendFileSync(file, prefix + lines.join("\n") + "\n")
  }
  return lines.length
}

function completedMutationArtifacts(matterDir: string, mutationID: string) {
  const file = lineageFile(matterDir)
  if (!existsSync(file)) return new Set<string>()
  return new Set(
    readFileSync(file, "utf8")
      .split(/\r?\n/)
      .flatMap((line) => {
        try {
          const row = JSON.parse(line) as {
            artifact?: string
            content_sha256?: string
            provenance?: { document_mutation_id?: string }
          }
          if (
            row.provenance?.document_mutation_id !== mutationID ||
            typeof row.artifact !== "string" ||
            typeof row.content_sha256 !== "string"
          ) return []
          return [`${row.artifact}\0${row.content_sha256}`]
        } catch {
          return []
        }
      }),
  )
}
