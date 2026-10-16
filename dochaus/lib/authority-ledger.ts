// Durable matter authority ledger (#84). Extracts authority references as
// provenance rows only; it never infers legal conclusions from a reference.
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { isOfficialSource } from "./research"

export type AuthorityStatus = "verified" | "tentative" | "unresolved"
export type AuthorityKind = "url" | "citation" | "research-row"

export type AuthorityLedgerSource = {
  artifact?: string
  tool?: string
  session_id?: string
  line?: number
}

export type AuthorityLedgerEntry = {
  id: string
  first_seen: string
  last_seen: string
  raw: string
  kind: AuthorityKind
  status: AuthorityStatus
  source: AuthorityLedgerSource
  jurisdiction_hint: string | null
  unresolved_reason?: string
}

type AuthoritySeed = Omit<AuthorityLedgerEntry, "id" | "first_seen" | "last_seen">

const STATUS_RANK: Record<AuthorityStatus, number> = { unresolved: 0, tentative: 1, verified: 2 }

export function authorityLedgerFile(directory: string): string {
  return path.join(directory, "authority-ledger.json")
}

export function readAuthorityLedger(directory: string): AuthorityLedgerEntry[] {
  const file = authorityLedgerFile(directory)
  if (!existsSync(file)) return []
  try {
    const parsed = JSON.parse(readFileSync(file, "utf8")) as unknown
    if (Array.isArray(parsed)) return parsed.filter(isAuthorityLedgerEntry)
    if (isRecord(parsed) && Array.isArray(parsed.entries)) return parsed.entries.filter(isAuthorityLedgerEntry)
    return []
  } catch {
    return []
  }
}

export function mergeAuthorityLedgerEntries(directory: string, seeds: readonly AuthoritySeed[]): AuthorityLedgerEntry[] {
  if (!seeds.length) return readAuthorityLedger(directory)
  const now = new Date().toISOString()
  const entries = new Map(readAuthorityLedger(directory).map((entry) => [authorityKey(entry.raw), entry]))
  for (const seed of seeds.filter((entry) => entry.raw.trim())) {
    const raw = trimAuthorityRaw(seed.raw)
    const key = authorityKey(raw)
    const existing = entries.get(key)
    const nextStatus =
      existing && STATUS_RANK[existing.status] > STATUS_RANK[seed.status] ? existing.status : seed.status
    entries.set(key, {
      ...seed,
      raw,
      id: existing?.id ?? authorityID(raw),
      first_seen: existing?.first_seen ?? now,
      last_seen: now,
      status: nextStatus,
    })
  }
  const rows = [...entries.values()].sort((a, b) => a.raw.localeCompare(b.raw) || a.id.localeCompare(b.id))
  writeAuthorityLedger(directory, rows)
  return rows
}

export function authorityEntriesFromSourceFetch(input: {
  url: string
  tool: string
  sessionID: string
}): AuthoritySeed[] {
  const url = normalizeURL(input.url)
  if (!url) return []
  return [
    {
      raw: url,
      kind: "url",
      status: input.tool === "official-source" || isOfficialSource(url) ? "verified" : "tentative",
      source: { tool: input.tool, session_id: input.sessionID },
      jurisdiction_hint: null,
    },
  ]
}

export function authorityEntriesFromResearch(
  content: string,
  source: AuthorityLedgerSource,
): AuthoritySeed[] {
  return parseResearchRows(content).flatMap((row): AuthoritySeed[] => {
    const tier = stringField(row, "tier").toLowerCase()
    const status: AuthorityStatus =
      tier === "verified" || tier === "corroborated" ? "verified" : tier === "open" ? "unresolved" : "tentative"
    const sourceRows = researchRowSources(row)
      .map((researchSource) => ({
        raw: normalizeURL(stringField(researchSource, "url")) || stringField(researchSource, "title"),
        kind: normalizeURL(stringField(researchSource, "url")) ? "url" as const : "citation" as const,
        status: status === "unresolved" ? "tentative" as const : status,
        source,
        jurisdiction_hint: null,
      }))
      .filter((entry) => entry.raw)
    if (sourceRows.length && status !== "unresolved") return sourceRows
    return [
      {
        raw: stringField(row, "question") || stringField(row, "requirement") || stringField(row, "note"),
        kind: "research-row" as const,
        status: "unresolved" as const,
        source,
        jurisdiction_hint: null,
        unresolved_reason: "open research row lacks a verified source",
      },
      ...sourceRows,
    ].filter((entry) => entry.raw)
  })
}

export function authorityEntriesFromArtifact(
  content: string,
  source: AuthorityLedgerSource,
): AuthoritySeed[] {
  return [
    ...matchEntries(content, /https?:\/\/[^\s<>"')\]]+/g, "url", source),
    ...matchEntries(
      content,
      /\b(?:[A-Z][A-Za-z&.'-]+(?:\s+[A-Z][A-Za-z&.'-]+){0,5}\s+)?(?:Act|Code|Rules?|Regulations?|Statutes?)\s*§+\s*[A-Za-z0-9][A-Za-z0-9().-]*/g,
      "citation",
      source,
    ),
    ...matchEntries(
      content,
      /\b[A-Z][A-Za-z0-9&.'-]+(?:\s+[A-Z][A-Za-z0-9&.'-]+){0,8}\s+v\.?\s+[A-Z][A-Za-z0-9&.'-]+(?:\s+[A-Z][A-Za-z0-9&.'-]+){0,8}\b/g,
      "citation",
      source,
    ),
  ]
}

function writeAuthorityLedger(directory: string, entries: AuthorityLedgerEntry[]): void {
  mkdirSync(directory, { recursive: true })
  writeFileSync(authorityLedgerFile(directory), `${JSON.stringify({ schema_version: 1, entries }, null, 2)}\n`)
}

function matchEntries(
  content: string,
  pattern: RegExp,
  kind: AuthorityKind,
  source: AuthorityLedgerSource,
): AuthoritySeed[] {
  return [...content.matchAll(pattern)]
    .map((match) => match[0].replace(/[),.;:]+$/, "").trim())
    .filter(Boolean)
    .map((raw) => ({
      raw: kind === "url" ? normalizeURL(raw) || raw : raw,
      kind,
      status: "tentative" as const,
      source: { ...source, line: lineFor(content, raw) },
      jurisdiction_hint: null,
    }))
}

function parseResearchRows(content: string): Record<string, unknown>[] {
  const match = /```research\s*\n?([\s\S]*?)```/i.exec(content)
  if (!match?.[1]) return []
  try {
    const parsed = JSON.parse(match[1]) as unknown
    return Array.isArray(parsed) ? parsed.filter(isRecord) : []
  } catch {
    return []
  }
}

function researchRowSources(row: Record<string, unknown>) {
  return Array.isArray(row.sources) ? row.sources.filter(isRecord) : []
}

function lineFor(content: string, raw: string): number | undefined {
  const index = content.indexOf(raw)
  if (index < 0) return undefined
  return content.slice(0, index).split(/\r?\n/).length
}

function authorityID(raw: string): string {
  return `auth_${createHash("sha256").update(authorityKey(raw)).digest("hex").slice(0, 20)}`
}

function authorityKey(raw: string): string {
  return raw.toLowerCase().replace(/\s+/g, " ").trim()
}

function trimAuthorityRaw(raw: string): string {
  const normalized = raw.replace(/\s+/g, " ").trim()
  return normalized.length <= 500 ? normalized : `${normalized.slice(0, 500).trimEnd()}...`
}

function normalizeURL(url: string): string {
  if (!url.trim()) return ""
  try {
    const parsed = new URL(url)
    parsed.hash = ""
    parsed.searchParams.sort()
    return parsed.toString()
  } catch {
    return url.trim()
  }
}

function stringField(record: Record<string, unknown>, key: string): string {
  const value = record[key]
  return typeof value === "string" ? value.trim() : ""
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

function isAuthorityLedgerEntry(value: unknown): value is AuthorityLedgerEntry {
  if (!isRecord(value)) return false
  return typeof value.id === "string" && typeof value.raw === "string" && typeof value.status === "string"
}
