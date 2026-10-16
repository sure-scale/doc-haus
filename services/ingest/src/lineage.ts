// Ingest-side lineage appender (#96). Self-contained — the ingest service cannot
// import dochaus/lib — producing the SAME row shape as dochaus/lib/lineage.ts.
// Stamps one PERMANENT row per present source artifact, with `supersedes` linking
// to the prior row for that artifact. Canonical schema documented in
// dochaus/lib/lineage.ts.
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

function priorEntryByArtifact(matterDir: string): Map<string, string> {
  const map = new Map<string, string>()
  const file = lineageFile(matterDir)
  if (!existsSync(file)) return map
  for (const line of readFileSync(file, "utf8").split(/\r?\n/)) {
    if (!line.trim()) continue
    try {
      const row = JSON.parse(line) as { artifact?: string; entry_id?: string }
      if (typeof row.artifact === "string" && typeof row.entry_id === "string") map.set(row.artifact, row.entry_id)
    } catch {
      /* skip corrupt line */
    }
  }
  return map
}

function hashFile(absPath: string): { sha256: string; bytes: number } | null {
  if (!existsSync(absPath)) return null
  const buf = readFileSync(absPath)
  return { sha256: createHash("sha256").update(buf).digest("hex"), bytes: buf.length }
}

function artifactsOnDisk(matterDir: string): string[] {
  const top = [".dochaus/legal.db", "source-index.json", "source-map.json", "source-map.md", "source-ledger.json", "matter-identifiers.json", "matter-identifiers.md"]
  const glob = (pat: string) => [...new Bun.Glob(pat).scanSync({ cwd: matterDir, onlyFiles: true })].map((p) => p.replace(/\\/g, "/"))
  return [...top, ...glob("source-chunks/*.txt"), ...glob("sources/*.txt")]
}

// Returns the number of artifacts stamped.
export function stampIngestLineage(
  matterDir: string,
  documents: string[],
  correlation: string,
  mutationID?: string,
): number {
  return withArtifactMutationLock(matterDir, ".dochaus/lineage.jsonl", () => {
    mkdirSync(path.join(matterDir, ".dochaus"), { recursive: true })
    const prior = priorEntryByArtifact(matterDir)
    const run = `ingest:${correlation}`
    const completed = mutationID ? completedMutationArtifacts(matterDir, mutationID) : new Set<string>()
    const provenance = {
      origin: "upload" as const,
      documents,
      ...(mutationID ? { document_mutation_id: mutationID } : {}),
    }
    const ts = new Date().toISOString()
    const lines = artifactsOnDisk(matterDir).flatMap((artifact) => {
      const hash = hashFile(path.join(matterDir, artifact))
      if (!hash || completed.has(`${artifact}\0${hash.sha256}`)) return []
      return [
        JSON.stringify({
          entry_id: ulid(),
          ts,
          tier: "PERMANENT",
          artifact,
          op: "ingest",
          run,
          stage: null,
          subagent: "ingest",
          content_sha256: hash.sha256,
          bytes: hash.bytes,
          supersedes: prior.get(artifact) ?? null,
          provenance,
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
  })
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
