import { mkdirSync, readdirSync, statSync, writeFileSync } from "node:fs"
import path from "node:path"
import { citationRowsForArtifact } from "./citation-extraction"
import { buildSupportIndex } from "./citation-support"

export type CitationCheckStatus = "verified" | "unresolved" | "unsupported"

export type CitationCheckArtifact = {
  file: string
  line: number
  column: number
  character_start: number
  character_end: number
}

export type CitationSupport = {
  kind: "authority-ledger" | "official-fetch" | "research-row" | "source-artifact" | "source-index" | "source-map"
  id: string
  status: Exclude<CitationCheckStatus, "unsupported">
  artifact?: string
  url?: string
  research_row_id?: string
}

export type CitationPinpoint = {
  base_key: string
  suffix: string
  reason: string
}

export type CitationCheckRow = {
  citation_text: string
  normalized_key: string
  kind: "case" | "reporter" | "rule-or-statute" | "url"
  artifact: CitationCheckArtifact
  status: CitationCheckStatus
  linked_source_ids: string[]
  linked_research_ids: string[]
  support: CitationSupport[]
  pinpoint?: CitationPinpoint
}

export type CitationCheckGroup = {
  normalized_key: string
  citation_texts: string[]
  count: number
  status: CitationCheckStatus
  artifact_locations: CitationCheckArtifact[]
  linked_source_ids: string[]
  linked_research_ids: string[]
}

export type CitationCheckReport = {
  schema_version: 1
  report_file: string
  artifacts: string[]
  summary: {
    total: number
    verified: number
    unresolved: number
    unsupported: number
    groups: number
  }
  groups: CitationCheckGroup[]
  rows: CitationCheckRow[]
}

type CitationCheckOptions = {
  directory: string
  artifacts?: readonly string[]
  output?: string
}

const DEFAULT_REPORT_FILE = "citation-check-report.json"

export function buildCitationCheckReport(options: CitationCheckOptions): CitationCheckReport {
  const artifacts = citationArtifacts(options.directory, options.artifacts)
  const supportIndex = buildSupportIndex(options.directory)
  const rows = artifacts
    .flatMap((artifact) => citationRowsForArtifact(options.directory, artifact, supportIndex))
    .sort(compareRows)
  const groups = citationGroups(rows)
  return {
    schema_version: 1,
    report_file: options.output ?? DEFAULT_REPORT_FILE,
    artifacts,
    summary: {
      total: rows.length,
      verified: rows.filter((row) => row.status === "verified").length,
      unresolved: rows.filter((row) => row.status === "unresolved").length,
      unsupported: rows.filter((row) => row.status === "unsupported").length,
      groups: groups.length,
    },
    groups,
    rows,
  }
}

export function writeCitationCheckReport(options: CitationCheckOptions): CitationCheckReport {
  const report = buildCitationCheckReport(options)
  mkdirSync(options.directory, { recursive: true })
  writeFileSync(path.join(options.directory, report.report_file), `${JSON.stringify(report, null, 2)}\n`)
  return report
}

function citationArtifacts(directory: string, requested?: readonly string[]): string[] {
  const files = requested
    ? requested.map((file) => path.basename(file))
    : readdirSync(directory)
        .filter((file) => isCitationArtifact(directory, file))
        .sort((a, b) => a.localeCompare(b))
  return [...new Set(files)].filter((file) => safeMatterFile(directory, file))
}

function isCitationArtifact(directory: string, file: string) {
  if (!safeMatterFile(directory, file)) return false
  if (!/\.(?:md|txt)$/i.test(file)) return false
  if (/^(?:source-|authority-ledger|citation-check-report|verification-report|matter-identifiers)/i.test(file)) return false
  if (/^(?:drafting-workflow-state|research-source-audit|render-status-report)/i.test(file)) return false
  if (!/(?:draft|agreement|memo|research|review|issue-ledger|ledger|advice|analysis|packet|register|spine|work-product)/i.test(file))
    return false
  return true
}

function safeMatterFile(directory: string, file: string) {
  if (file.includes("/") || file.includes("\\")) return false
  const resolved = path.resolve(directory, file)
  if (resolved === directory || !resolved.startsWith(path.resolve(directory) + path.sep)) return false
  try {
    return statSync(resolved).isFile()
  } catch {
    return false
  }
}

export function formatCitationCheckSummary(report: CitationCheckReport): string {
  const s = report.summary
  const lines: string[] = [
    `**${s.total}** citation reference${s.total === 1 ? "" : "s"} across **${report.artifacts.length}** artifact${report.artifacts.length === 1 ? "" : "s"}: ` +
      `**${s.verified}** verified, **${s.unresolved}** unresolved, **${s.unsupported}** unsupported.`,
  ]
  const unsupported = report.rows.filter((row) => row.status === "unsupported")
  const unresolved = report.rows.filter((row) => row.status === "unresolved")
  if (unsupported.length) {
    lines.push("", `### Unsupported (${unsupported.length})`, "")
    for (const row of unsupported) {
      lines.push(`- \`${row.artifact.file}:${row.artifact.line}\` "${row.citation_text}"`)
    }
  }
  if (unresolved.length) {
    lines.push("", `### Unresolved (${unresolved.length})`, "")
    for (const row of unresolved) {
      const linked = [...row.linked_source_ids, ...row.linked_research_ids].join(", ") || "(none)"
      const pinpoint = row.pinpoint ? ` — ${row.pinpoint.reason}` : ""
      lines.push(`- \`${row.artifact.file}:${row.artifact.line}\` "${row.citation_text}" → ${linked}${pinpoint}`)
    }
  }
  if (!unsupported.length && !unresolved.length && s.total > 0) {
    lines.push("", "All citations have durable source support.")
  }
  if (s.total === 0) {
    lines.push("", "No citation-like references found in scanned artifacts.")
  }
  lines.push("", `Full report: \`${report.report_file}\``)
  return lines.join("\n")
}

function citationGroups(rows: CitationCheckRow[]): CitationCheckGroup[] {
  const groups = new Map<string, CitationCheckRow[]>()
  for (const row of rows) groups.set(row.normalized_key, [...(groups.get(row.normalized_key) ?? []), row])
  return [...groups.entries()]
    .map(([normalized_key, grouped]) => ({
      normalized_key,
      citation_texts: [...new Set(grouped.map((row) => row.citation_text))].sort((a, b) => a.localeCompare(b)),
      count: grouped.length,
      status: groupStatus(grouped),
      artifact_locations: grouped.map((row) => row.artifact),
      linked_source_ids: [...new Set(grouped.flatMap((row) => row.linked_source_ids))].sort((a, b) =>
        a.localeCompare(b),
      ),
      linked_research_ids: [...new Set(grouped.flatMap((row) => row.linked_research_ids))].sort((a, b) =>
        a.localeCompare(b),
      ),
    }))
    .sort((a, b) => a.normalized_key.localeCompare(b.normalized_key))
}

function groupStatus(rows: CitationCheckRow[]): CitationCheckStatus {
  if (rows.some((row) => row.status === "unsupported")) return "unsupported"
  if (rows.some((row) => row.status === "unresolved")) return "unresolved"
  return "verified"
}

function compareRows(a: CitationCheckRow, b: CitationCheckRow) {
  return (
    a.artifact.file.localeCompare(b.artifact.file) ||
    a.artifact.line - b.artifact.line ||
    a.artifact.column - b.artifact.column ||
    a.citation_text.localeCompare(b.citation_text)
  )
}
