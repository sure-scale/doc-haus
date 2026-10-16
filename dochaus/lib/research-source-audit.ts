import path from "node:path"
import { fetchSourceText, readableText } from "./research-source-fetch"

const GENERIC_TITLE_TOKENS = new Set([
  "about",
  "agency",
  "agreement",
  "annual",
  "appendix",
  "article",
  "bulletin",
  "case",
  "chapter",
  "code",
  "commission",
  "department",
  "division",
  "document",
  "enforcement",
  "final",
  "financial",
  "form",
  "guidance",
  "guide",
  "information",
  "interpretation",
  "letter",
  "matter",
  "notice",
  "official",
  "order",
  "policy",
  "press",
  "public",
  "question",
  "questions",
  "regulation",
  "regulatory",
  "release",
  "report",
  "rule",
  "rules",
  "section",
  "service",
  "services",
  "staff",
  "state",
  "statement",
  "statute",
  "updated",
])

const ENTITY_SUFFIX_TOKENS = new Set([
  "co",
  "corp",
  "corporation",
  "inc",
  "incorporated",
  "llc",
  "llp",
  "lp",
  "ltd",
  "limited",
  "plc",
  "the",
])

type ResearchRow = Record<string, unknown>

export type ResearchSource = {
  row: number
  question: string
  tier: string
  title: string
  url: string
  kind: string
}

export type SourceCheck = {
  source: ResearchSource
  ok: boolean
  reason: string
  evidence: string
}

export type ResearchSourceAudit = {
  status: "PASS" | "PARTIAL" | "BLOCKED"
  checked: number
  failures: SourceCheck[]
  warnings: SourceCheck[]
}

export function parseJsonFence(text: string, tag: string) {
  const match = new RegExp(`\`\`\`${tag}\\s*\\n?([\\s\\S]*?)\`\`\``, "i").exec(text)
  if (!match?.[1]) return []
  const parsed = JSON.parse(match[1]) as unknown
  return Array.isArray(parsed) ? parsed : []
}

export function collectResearchSources(rows: unknown[]) {
  const sources = rows.flatMap((row, index) => {
    if (!isRecord(row)) return []
    if (!Array.isArray(row.sources)) return []
    return row.sources.flatMap((source) => {
      if (!isRecord(source)) return []
      const url = stringValue(source.url)
      if (!url) return []
      return [
        {
          row: index + 1,
          question: stringValue(row.question),
          tier: stringValue(row.tier),
          title: stringValue(source.title),
          url,
          kind: stringValue(source.kind),
        },
      ]
    })
  })
  const seen = new Set<string>()
  return sources.filter((source) => {
    const key = `${source.row}\n${source.url}\n${source.title}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

export function evaluateSourceIdentity(source: ResearchSource, fetchedText: string) {
  const haystack = tokenSet(readableText(fetchedText))
  const compactHaystack = `${compact(readableText(fetchedText))} ${compact(source.url)}`
  const entityTokens = entityTitleTokens(source.title)
  const identifiers = legalIdentifiers(source.title)
  const titleTokens = distinctiveTokens(source.title)
  const releaseNumber = releaseNumberFrom(source.title)
  const failed = (reason: string, evidence: string) => ({ source, ok: false, reason, evidence })
  const passed = (evidence: string) => ({ source, ok: true, reason: "matched", evidence })

  if (releaseNumber && !haystack.has(releaseNumber))
    return failed("release or identifier in source title was not found in fetched text", `missing ${releaseNumber}`)

  if (entityTokens.length) {
    const required = entityTokens[0]!
    if (!haystack.has(required))
      return failed(
        "named source identity in title was not found in fetched text",
        `missing ${required}; entity tokens: ${entityTokens.join(", ")}`,
      )
    const hits = entityTokens.filter((token) => haystack.has(token))
    if (entityTokens.length > 1 && hits.length < Math.min(2, entityTokens.length))
      return failed(
        "too few named source identity tokens matched fetched text",
        `matched ${hits.join(", ") || "none"} of ${entityTokens.join(", ")}`,
      )
    return passed(`matched identity tokens: ${hits.join(", ")}`)
  }

  if (identifiers.length) {
    const hits = identifiers.filter((identifier) => compactHaystack.includes(identifier))
    if (hits.length) return passed(`matched legal identifier: ${hits.join(", ")}`)
    return failed(
      "legal identifier in source title was not found in fetched text",
      `missing ${identifiers.join(", ")}`,
    )
  }

  if (titleTokens.length >= 2) {
    const hits = titleTokens.filter((token) => haystack.has(token))
    if (hits.length < Math.min(2, titleTokens.length))
      return failed(
        "source title tokens did not match fetched text",
        `matched ${hits.join(", ") || "none"} of ${titleTokens.slice(0, 8).join(", ")}`,
      )
    return passed(`matched title tokens: ${hits.join(", ")}`)
  }

  return {
    source,
    ok: true,
    reason: "weak-title",
    evidence: "source title had too few distinctive tokens for deterministic identity matching",
  }
}

export async function auditResearchFiles(args: {
  directory: string
  researchFile?: string
  maxFetches?: number
}) {
  const researchPath = path.join(args.directory, basename(args.researchFile ?? "research.md"))
  const researchRows = parseJsonFence(await Bun.file(researchPath).text(), "research")
  const sources = collectResearchSources(researchRows).slice(0, Math.max(1, Math.min(args.maxFetches ?? 30, 50)))
  const checks = await Promise.all(sources.map((source) => checkSource(source)))
  return summarizeSourceChecks(checks)
}

export function summarizeSourceChecks(checks: SourceCheck[]) {
  const rowsWithPass = new Set(checks.filter((check) => check.ok).map((check) => check.source.row))
  const failures = checks.filter(
    (check) =>
      !check.ok &&
      verifiedTier(check.source.tier) &&
      (identityFailure(check.reason) || !rowsWithPass.has(check.source.row)),
  )
  const warnings = checks.filter((check) => !check.ok && !failures.includes(check))
  return {
    status: failures.length ? "BLOCKED" : warnings.length ? "PARTIAL" : "PASS",
    checked: checks.length,
    failures,
    warnings,
  } satisfies ResearchSourceAudit
}

export function renderResearchSourceAudit(audit: ResearchSourceAudit) {
  return `# Research Source Audit

\`\`\`research_source_audit
${JSON.stringify(audit, null, 2)}
\`\`\`
`
}

async function checkSource(source: ResearchSource): Promise<SourceCheck> {
  const result = await fetchSourceText(source.url)
  if (result.status === "blocked")
    return {
      source,
      ok: false,
      reason: `could not fetch source: ${result.reason}`,
      evidence: `${result.descriptor.finalUrl} (${result.descriptor.host})`,
    }
  const identity = evaluateSourceIdentity(source, result.text)
  if (result.status === "partial")
    return { ...identity, evidence: `${identity.evidence} [partial fetch: ${result.reason}]` }
  return identity
}

function entityTitleTokens(title: string) {
  const normalized = title.replace(/[–—]/g, " - ")
  const inMatter = /in the matter of\s+([^,;(]+)/i.exec(normalized)
  const orderTitle = /(?:order|release|decision|settlement)\s*[-:]\s*([^,;(]+)/i.exec(normalized)
  const segment = inMatter?.[1] ?? orderTitle?.[1] ?? ""
  return distinctiveTokens(segment).filter((token) => !ENTITY_SUFFIX_TOKENS.has(token))
}

function distinctiveTokens(text: string) {
  return [...new Set(tokenize(text).filter((token) => token.length > 3 && !GENERIC_TITLE_TOKENS.has(token)))]
}

function releaseNumberFrom(title: string) {
  return /rel(?:ease)?\.?\s*no\.?\s*[a-z]{0,6}[-\s]?(\d{3,})/i.exec(title)?.[1]
}

function legalIdentifiers(title: string) {
  const raw = [
    ...title.matchAll(/\b\d+[a-z]?(?:[-.]\d+[a-z]?)+(?:\([a-z0-9]+\))?/gi),
    ...title.matchAll(/\b\d+[a-z]?\([a-z0-9]+\)(?:\([a-z0-9]+\))*[-.]?\d*[a-z]?\b/gi),
    ...title.matchAll(/\b(?:\d+\s+[A-Z]{2,6}\s+)?part\s+\d+[a-z]?\b/gi),
  ]
    .map((match) => match[0])
    .flatMap((identifier) => [
      identifier,
      identifier.replace(/\([a-z0-9]+\).*$/i, ""),
      identifier.replace(/\([a-z0-9]+\)/gi, ""),
      identifier.replace(/^.*\bpart\s+(\d+[a-z]?)$/i, "part$1"),
    ])
    .map(compact)
    .filter((identifier) => identifier.length >= 3)
  return [...new Set(raw)]
}

function tokenSet(text: string) {
  return new Set(tokenize(text))
}

function tokenize(text: string) {
  return text
    .toLowerCase()
    .replace(/&amp;/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
}

function compact(text: string) {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, "")
}

function verifiedTier(tier: string) {
  return tier === "verified" || tier === "corroborated"
}

function identityFailure(reason: string) {
  return reason.includes("identity") || reason.includes("title tokens") || reason.includes("legal identifier")
}

function basename(name: string) {
  return path.basename(name.endsWith(".md") ? name : `${name}.md`)
}

function stringValue(value: unknown) {
  return typeof value === "string" ? value.trim() : ""
}

function isRecord(value: unknown): value is ResearchRow {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
