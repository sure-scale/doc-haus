// Per-stage drafting context manifest (#77). Append-only, matter-local, and
// content-light: records filenames, hashes, ranges, queries, and result ids,
// never source bodies or retrieved excerpts.
import { appendFileSync, existsSync, mkdirSync, readFileSync } from "node:fs"
import path from "node:path"
import { sha256File, ulid } from "./lineage"
import { withArtifactMutationLock } from "../capabilities/artifacts/mutation-lock"
import type { ResearchSourceSnapshotRef } from "./research-source-snapshot"

export type ContextManifestEvent =
  | "stage_input"
  | "artifact_read"
  | "search"
  | "source_fetch"
  | "retrieval_attempt"
  | "stage_output"
  | "context_prepared"

export type ContextArtifactHash = {
  artifact: string
  sha256: string | null
  bytes: number | null
}

export type ContextSearchResult = {
  document: string
  section: string
  structural_path?: string
  char_start?: number
  char_end?: number
  score?: number
}

export type ContextManifestRow = {
  entry_id: string
  ts: string
  run: string
  stage: string | null
  subagent: string
  session_id: string
  event: ContextManifestEvent
  upstream_artifacts?: string[]
  output_artifacts?: string[]
  input_hashes?: ContextArtifactHash[]
  artifact?: string
  artifact_hash?: ContextArtifactHash
  range?: { start: number; end: number | null }
  research_rows?: number
  query?: string
  k?: number
  document?: string
  results?: ContextSearchResult[]
  url?: string
  retrieved?: boolean
  source_sha256?: string
  source_snapshot?: ResearchSourceSnapshotRef
  authority_aliases?: string[]
  missing_terms?: string[]
  validation_version?: number
  assigned_scope?: string
  stage_epoch?: number
  drafting_action_id?: string
  drafting_claim_id?: string
  drafting_action_generation?: number
  research_tier?: string
  submitted_research_tier?: string
  source_evidence_entries?: string[]
  unsupported_authority_claims?: string[]
  quotation_evidence?: ResearchQuotationEvidence[]
  request_message_id?: string
  request_parent_message_id?: string
  owner_session_id?: string
  context_sha256?: string
  context_bytes?: number
  strategy_readiness_projection?:
    | { kind: "absent" }
    | { kind: "applied"; assessment_id: string; result_hash: string; context_hash: string }
    | { kind: "conserved"; receipt_id: string }
  context_delivery?: {
    kind: "complete_inline_packet"
    packet_sha256: string
    packet_bytes: number
    artifacts: ContextArtifactHash[]
  }
  contract_hash?: string
  deferred_context?: { kind: string; owner_stage: string | null }[]
}

// Literal span evidence only. Semantic/legal support is assessed separately by
// the existing research and review workflow, never inferred from text equality.
export type ResearchQuotationEvidence = {
  quote_sha256: string
  source_snapshot: string
  locator: string
  textual_status: "verified" | "unverified"
  reason?: "missing_source" | "source_unavailable" | "source_incomplete" | "locator_unavailable" | "quote_mismatch" | "ambiguous_quote"
  evidence_entry_id?: string
  span?: { start: number; end: number; excerpt_sha256: string }
}

export function contextManifestFile(directory: string): string {
  return path.join(directory, ".dochaus", "context-manifest.jsonl")
}

export function readContextManifest(directory: string): ContextManifestRow[] {
  const file = contextManifestFile(directory)
  if (!existsSync(file)) return []
  try {
    return readFileSync(file, "utf8")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => parseContextManifestRow(line))
      .filter((row): row is ContextManifestRow => row !== null)
  } catch {
    return []
  }
}

export function appendContextManifestRow(
  directory: string,
  partial: Omit<ContextManifestRow, "entry_id" | "ts">,
): ContextManifestRow {
  const row: ContextManifestRow = { ...partial, entry_id: ulid(), ts: new Date().toISOString() }
  withArtifactMutationLock(directory, ".dochaus/context-manifest.jsonl", () => {
    mkdirSync(path.join(directory, ".dochaus"), { recursive: true })
    appendFileSync(contextManifestFile(directory), `${JSON.stringify(row)}\n`)
  })
  return row
}

export function artifactHashes(directory: string, artifacts: readonly string[]): ContextArtifactHash[] {
  return artifacts.map((artifact) => {
    const hash = sha256File(path.join(directory, artifact))
    return {
      artifact,
      sha256: hash?.sha256 ?? null,
      bytes: hash?.bytes ?? null,
    }
  })
}

function parseContextManifestRow(line: string): ContextManifestRow | null {
  try {
    const parsed = JSON.parse(line) as unknown
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return null
    const row = parsed as Record<string, unknown>
    if (typeof row.entry_id !== "string") return null
    return row as unknown as ContextManifestRow
  } catch {
    return null
  }
}
