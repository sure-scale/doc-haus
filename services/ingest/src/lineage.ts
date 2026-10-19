// Ingest-side lineage appender (#96). Self-contained — the ingest service cannot
// import dochaus/lib — producing the SAME row shape as dochaus/lib/lineage.ts.
// An ingest (upload or delete) and a rebuild of the source projection each
// stamp one PERMANENT row per present source artifact whose bytes differ from
// its latest row; an unchanged artifact gets no row. `supersedes` links each
// row to the prior row for that artifact. Only the first row a stamp appends
// lists the source documents. Canonical schema documented in
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

type Ledger = ReturnType<typeof readLedger>

// The ledger as a stamp needs it, from one read under the lineage lock: the
// latest row per artifact (its entry id and recorded content hash) and whether
// the ledger lacks a final newline. The ledger only grows; reading and parsing
// it once per stamp instead of up to three times keeps that stretch of the
// event loop short.
function readLedger(matterDir: string) {
  const file = lineageFile(matterDir)
  const text = existsSync(file) ? readFileSync(file, "utf8") : ""
  const prior = new Map<string, { entry_id: string; content_sha256: unknown }>()
  for (const line of text.split(/\r?\n/)) {
    if (!line.trim()) continue
    try {
      const row = JSON.parse(line) as { artifact?: string; entry_id?: string; content_sha256?: unknown }
      if (typeof row.artifact === "string" && typeof row.entry_id === "string")
        prior.set(row.artifact, { entry_id: row.entry_id, content_sha256: row.content_sha256 })
    } catch {
      /* skip corrupt line */
    }
  }
  return { prior, unterminated: text.length > 0 && !text.endsWith("\n") }
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

// Stamps a source upload or delete (or a recovery that restaged it) as an
// ingest, legal.db included. A finalize retried after a stop, even one that
// stopped mid-append, finds the rows it already wrote as the latest ones and
// stamps only what is still unrecorded. Returns the number of artifacts stamped.
export function stampIngestLineage(
  matterDir: string,
  documents: string[],
  correlation: string,
  mutationID?: string,
  files?: ReadonlyMap<string, ArtifactHash>,
): number {
  return withArtifactMutationLock(matterDir, ".dochaus/lineage.jsonl", () =>
    appendArtifactRows(matterDir, readLedger(matterDir), {
      artifacts: artifactsOnDisk(matterDir),
      op: "ingest",
      run: `ingest:${correlation}`,
      provenance: {
        origin: "upload",
        documents,
        ...(mutationID ? { document_mutation_id: mutationID } : {}),
      },
      files,
    }))
}

// Stamps a rebuild of the source projection from the registry (the startup
// self-heal). The rebuild re-derives files from sources already ingested and
// uploads nothing, so each row is a whole-file "write" with origin
// "source-rebuild". legal.db is never written by a rebuild and is not stamped.
// Returns the number of artifacts stamped.
export function stampSourceRebuildLineage(
  matterDir: string,
  documents: string[],
  correlation: string,
  files?: ReadonlyMap<string, ArtifactHash>,
): number {
  return withArtifactMutationLock(matterDir, ".dochaus/lineage.jsonl", () =>
    appendArtifactRows(matterDir, readLedger(matterDir), {
      artifacts: artifactsOnDisk(matterDir).filter((artifact) => artifact !== ".dochaus/legal.db"),
      op: "write",
      run: `ingest:${correlation}`,
      provenance: { origin: "source-rebuild", documents },
      files,
    }))
}

// Appends one PERMANENT row per present artifact whose bytes differ from its
// latest row; a file with no lineage row yet is stamped too, so it becomes
// ledgered. The first row lists `documents`; later rows keep the rest of the
// provenance, so a stamp records the document list once rather than once per
// artifact. Runs under the caller's lineage lock, which `ledger` was read under.
function appendArtifactRows(
  matterDir: string,
  ledger: Ledger,
  input: {
    artifacts: string[]
    op: "ingest" | "write"
    run: string
    provenance: { origin: string; documents: string[]; document_mutation_id?: string }
    files?: ReadonlyMap<string, ArtifactHash>
  },
) {
  mkdirSync(path.join(matterDir, ".dochaus"), { recursive: true })
  const { documents: _documents, ...later } = input.provenance
  const ts = new Date().toISOString()
  const lines: string[] = []
  for (const artifact of input.artifacts) {
    const hash = input.files?.get(artifact) ?? hashFile(path.join(matterDir, artifact))
    const prior = ledger.prior.get(artifact)
    if (!hash || prior?.content_sha256 === hash.sha256) continue
    lines.push(JSON.stringify({
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
      supersedes: prior?.entry_id ?? null,
      provenance: lines.length ? later : input.provenance,
    }))
  }
  if (lines.length) {
    const prefix = ledger.unterminated ? "\n" : ""
    appendFileSync(lineageFile(matterDir), prefix + lines.join("\n") + "\n")
  }
  return lines.length
}
