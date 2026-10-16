// Verification-loop orchestrator. Six value-agnostic gates run over a single
// deliverable in a matter directory. Each gate is a function (ctx) → result;
// gates may read additional files from ctx.directory and may perform HTTP
// fetches (research-sources calls auditResearchFiles which fetches URLs).
// The orchestrator writes verification-report.md and returns a structured
// report. Quote matching reuses lib/quote-verify.verifyQuote (single source of
// truth shared with tool/verify-quote); value extraction reuses
// lib/coverage-primitives.extractValues.

import { existsSync, readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { verifyQuote } from "./quote-verify"
import { extractValues } from "./coverage-primitives"
import { auditResearchFiles, type SourceCheck } from "./research-source-audit"
import { writeCitationCheckReport, type CitationCheckReport, type CitationCheckRow } from "./citation-check"
import { readDurableJson } from "./durable-json"
import { readLineage, readRunId } from "./lineage"

export type GateId =
  | "citation-anchors"
  | "value-grounding"
  | "research-sources"
  | "render-readiness"
  | "redline-anchors"
  | "privilege-export"

export type GateStatus = "pass" | "fail" | "not-applicable"

export type GateFinding = {
  severity: "fail" | "warn"
  description: string
  detail?: string
}

export type GateResult = {
  id: GateId
  status: GateStatus
  findings: GateFinding[]
  summary: string
}

export type VerificationReport = {
  status: "PASS" | "PARTIAL" | "BLOCKED"
  gates: Record<GateId, GateResult>
  deliverable: string | null
  reportFile: string
}

type RunOptions = {
  directory: string
  deliverable?: string
  gates?: GateId[]
  output?: string
}

export const ALL_GATES: GateId[] = [
  "citation-anchors",
  "value-grounding",
  "research-sources",
  "render-readiness",
  "redline-anchors",
  "privilege-export",
]

export async function runVerificationGates(options: RunOptions): Promise<VerificationReport> {
  const gates = options.gates ?? ALL_GATES
  const deliverable = options.deliverable ?? discoverDeliverable(options.directory)
  const body = deliverable ? readFileSync(path.join(options.directory, deliverable), "utf8") : ""

  const ctx: GateContext = { directory: options.directory, deliverable, body }
  const results = {} as Record<GateId, GateResult>
  let threw = false

  for (const id of gates) {
    try {
      results[id] = await runGate(id, ctx)
    } catch (err) {
      // A gate that throws signals a runtime/transport failure, not a drafting
      // finding. Record a fail-result so the report still shows the cause, and
      // escalate the overall status to BLOCKED so the caller knows the loop
      // could not complete (distinct from PARTIAL, which means a gate ran and
      // produced findings).
      threw = true
      const message = err instanceof Error ? err.message : String(err)
      results[id] = {
        id,
        status: "fail",
        summary: `gate threw: ${message}`,
        findings: [{ severity: "fail", description: "gate threw", detail: message }],
      }
    }
  }

  const status = computeOverall(results, threw)
  const reportFile = options.output ?? "verification-report.md"
  const reportText = formatReport(results, deliverable, status)
  await Bun.write(path.join(options.directory, reportFile), reportText)

  return { status, gates: results, deliverable, reportFile }
}

export function computeOverall(
  results: Record<GateId, GateResult>,
  threw: boolean,
): "PASS" | "PARTIAL" | "BLOCKED" {
  if (threw) return "BLOCKED"
  const found = Object.values(results)
  if (found.some((r) => r.status === "fail")) return "PARTIAL"
  return "PASS"
}

type GateContext = { directory: string; deliverable: string | null; body: string }

async function runGate(id: GateId, ctx: GateContext): Promise<GateResult> {
  switch (id) {
    case "citation-anchors":
      return citationAnchorsGate(ctx)
    case "value-grounding":
      return valueGroundingGate(ctx)
    case "research-sources":
      return await researchSourcesGate(ctx)
    case "render-readiness":
      return renderReadinessGate(ctx)
    case "redline-anchors":
      return redlineAnchorsGate(ctx)
    case "privilege-export":
      return privilegeExportGate(ctx)
  }
}

// ─── citation-anchors ─────────────────────────────────────────
// Extract every bracketed span the author intends as a verbatim quotation and
// confirm each appears in some indexed matter document. Matching reuses
// lib/quote-verify.verifyQuote — the same path tool/verify-quote takes — so the
// gate inherits single-chunk and cross-boundary matching without reimplementing
// the DB query.
function citationAnchorsGate(ctx: GateContext): GateResult {
  const report = writeCitationCheckReport({ directory: ctx.directory })
  const supportFindings = citationSupportFindings(report)
  const supportFailures = supportFindings.filter((finding) => finding.severity === "fail")
  const supportSummary = citationSupportSummary(report)
  if (!ctx.deliverable && report.rows.length === 0) return notApplicable("citation-anchors", "no deliverable discovered")

  const anchors = ctx.deliverable ? extractQuoteAnchors(ctx.body) : []
  const missing: string[] = []
  for (const anchor of anchors) {
    const result = verifyQuote({ directory: ctx.directory, quote: anchor })
    if (result.reason === "no-db") {
      if (supportFailures.length) return fail("citation-anchors", supportSummary, supportFindings)
      return {
        ...notApplicable("citation-anchors", `no indexed documents in this matter (no .dochaus/legal.db); ${supportSummary}`),
        findings: supportFindings,
      }
    }
    // DB exists but holds zero indexed documents (transient ingest state, parse
    // failures). Treating this as not-applicable avoids a misleading drafting
    // FAIL that is really an indexing problem.
    if (result.reason === "no-docs") {
      if (supportFailures.length) return fail("citation-anchors", supportSummary, supportFindings)
      return {
        ...notApplicable("citation-anchors", `no documents indexed in this matter yet; ${supportSummary}`),
        findings: supportFindings,
      }
    }
    // A span that normalizes to empty (e.g. bracketed whitespace) is not a
    // quote the author intends to verify; skip rather than flag.
    if (result.reason === "empty-quote") continue
    if (!result.found) missing.push(anchor)
  }
  if (missing.length || supportFailures.length) {
    const missingFindings = missing.map((anchor) => ({
      severity: "fail" as const,
      description: "quoted anchor not found verbatim",
      detail: anchor,
    }))
    return fail(
      "citation-anchors",
      [
        missing.length ? `${missing.length} anchor(s) not found verbatim` : "",
        supportSummary,
      ].filter(Boolean).join("; "),
      [
        ...missingFindings,
        ...supportFindings,
      ],
    )
  }
  return {
    ...pass(
      "citation-anchors",
      anchors.length ? `${anchors.length} anchor(s) verified verbatim; ${supportSummary}` : supportSummary,
    ),
    findings: supportFindings,
  }
}

function citationSupportSummary(report: CitationCheckReport): string {
  if (report.rows.length === 0) return `no citation-like references found; wrote ${report.report_file}`
  return (
    `${report.summary.total} citation-like reference(s): ` +
    `${report.summary.verified} verified, ${report.summary.unresolved} unresolved, ` +
    `${report.summary.unsupported} unsupported; wrote ${report.report_file}`
  )
}

function citationSupportFindings(report: CitationCheckReport): GateFinding[] {
  return report.rows
    .filter((row) => row.status !== "verified")
    .map((row) => ({
      severity: row.status === "unsupported" ? "fail" as const : "warn" as const,
      description: row.status === "unsupported" ? "unsupported citation" : "unresolved citation",
      detail: citationSupportFindingDetail(row, report.report_file),
    }))
}

function citationSupportFindingDetail(row: CitationCheckRow, reportFile: string): string {
  const location = `${row.artifact.file}:${row.artifact.line}:${row.artifact.column}`
  if (row.status === "unsupported")
    return `${location} "${row.citation_text}" has no durable source support (${reportFile})`
  return `${location} "${row.citation_text}" remains unresolved; linked support: ${
    [...row.linked_source_ids, ...row.linked_research_ids].join(", ") || "(none)"
  } (${reportFile})`
}

// Bracketed spans of 20+ chars are candidate quote/citation anchors. Intentional
// `[insert ...]` placeholders are excluded so the gate only verifies spans the
// author intends as verbatim quotations.
function extractQuoteAnchors(body: string): string[] {
  const out: string[] = []
  for (const m of body.matchAll(/\[([^\]]{20,})\]/g)) {
    const anchor = m[1]
    if (!anchor || /^\s*insert\b/i.test(anchor)) continue
    out.push(anchor)
  }
  return out
}

// ─── value-grounding ──────────────────────────────────────────
// Extract every dollar amount, percentage, ISO date, and year-shaped value
// from the deliverable body. Each must either (a) appear in a source-term
// artifact (binding-decisions.md, key-dates.md) or (b) sit inside a
// bracketed [insert ...] placeholder. Fail otherwise.
function valueGroundingGate(ctx: GateContext): GateResult {
  if (!ctx.deliverable) return notApplicable("value-grounding", "no deliverable discovered")

  const sourceTerms = collectSourceTerms(ctx.directory)
  const placeholderRanges = findPlaceholderRanges(ctx.body)
  const values = extractAllValues(ctx.body)

  const missing: string[] = []
  for (const { value, index } of values) {
    if (sourceTerms.includes(value)) continue
    if (placeholderRanges.some(([start, end]) => index >= start && index < end)) continue
    missing.push(value)
  }

  if (missing.length) {
    return fail(
      "value-grounding",
      `${missing.length} value(s) trace to no source term or placeholder`,
      missing,
    )
  }
  return pass("value-grounding", `${values.length} value(s) verified against source terms or placeholders`)
}

function collectSourceTerms(directory: string): string[] {
  const out: string[] = []
  for (const file of ["binding-decisions.md", "key-dates.md"]) {
    const p = path.join(directory, file)
    if (!existsSync(p)) continue
    const text = readFileSync(p, "utf8")
    out.push(...extractValues(text))
    out.push(...extractIsoDates(text))
    out.push(...extractYears(text))
  }
  return out
}

function extractAllValues(body: string): { value: string; index: number }[] {
  const out: { value: string; index: number }[] = []
  const pushAll = (re: RegExp, src: string) => {
    for (const m of src.matchAll(re)) out.push({ value: m[0], index: m.index ?? 0 })
  }
  pushAll(/\$[\d,]+(?:\.\d+)?/g, body)
  pushAll(/\d+(?:\.\d+)?%/g, body)
  pushAll(/\d{4}-\d{2}-\d{2}/g, body)
  pushAll(/\b(?:19|20)\d{2}\b/g, body)
  return out
}

function extractIsoDates(text: string): string[] {
  return [...text.matchAll(/\d{4}-\d{2}-\d{2}/g)].map((m) => m[0])
}

function extractYears(text: string): string[] {
  return [...text.matchAll(/\b(?:19|20)\d{2}\b/g)].map((m) => m[0])
}

function findPlaceholderRanges(body: string): [number, number][] {
  const ranges: [number, number][] = []
  for (const m of body.matchAll(/\[insert[^\]]*\]/g)) {
    if (m.index !== undefined) ranges.push([m.index, m.index + m[0].length])
  }
  return ranges
}

// ─── research-sources ─────────────────────────────────────────
async function researchSourcesGate(ctx: GateContext): Promise<GateResult> {
  const researchPath = path.join(ctx.directory, "research.md")
  if (!existsSync(researchPath)) {
    return notApplicable("research-sources", "no research.md in this matter")
  }
  const audit = await auditResearchFiles({
    directory: ctx.directory,
    researchFile: "research.md",
    maxFetches: 10,
  })
  if (audit.status === "BLOCKED") {
    return fail(
      "research-sources",
      `${audit.failures.length} blocking failure(s); ${audit.warnings.length} warning(s)`,
      audit.failures.map(sourceCheckFinding),
    )
  }
  return pass(
    "research-sources",
    `audited ${audit.checked} source(s); ${audit.warnings.length} warning(s)`,
  )
}

function sourceCheckFinding(check: SourceCheck): GateFinding {
  return {
    severity: "fail",
    description: check.reason,
    detail: `${check.source.title} (${check.source.url}): ${check.evidence}`,
  }
}

// ─── render-readiness ─────────────────────────────────────────
// Flag bare [___] or [] blanks. Intentional [insert ...] placeholders PASS.
function renderReadinessGate(ctx: GateContext): GateResult {
  if (!ctx.deliverable) return notApplicable("render-readiness", "no deliverable discovered")
  const blanks: string[] = []
  for (const m of ctx.body.matchAll(/\[\s*(?:_{2,}|\s*)\]/g)) {
    blanks.push(m[0])
  }
  if (blanks.length) {
    return fail("render-readiness", `${blanks.length} bare bracket blank(s) remain`, blanks)
  }
  return pass("render-readiness", "no ungrounded bare blanks")
}

// ─── redline-anchors ──────────────────────────────────────────
// Pending redlines ledger is matter-state; if absent, not-applicable.
function redlineAnchorsGate(ctx: GateContext): GateResult {
  const ledgerPath = path.join(ctx.directory, "pending-redlines.json")
  if (!existsSync(ledgerPath)) {
    return notApplicable("redline-anchors", "no pending-redlines ledger")
  }
  let ledger: { anchor?: string; document?: string }[] = []
  try {
    ledger = JSON.parse(readFileSync(ledgerPath, "utf8"))
  } catch {
    return notApplicable("redline-anchors", "pending-redlines ledger unreadable")
  }
  if (!ledger.length) return notApplicable("redline-anchors", "no pending redlines")

  const orphans: string[] = []
  for (const entry of ledger) {
    if (!entry.anchor || !entry.document) continue
    const docPath = path.join(ctx.directory, entry.document)
    if (!existsSync(docPath)) {
      orphans.push(`anchor "${entry.anchor.slice(0, 40)}..." → document missing`)
      continue
    }
    const docText = readFileSync(docPath, "utf8")
    if (!docText.includes(entry.anchor)) {
      orphans.push(`anchor "${entry.anchor.slice(0, 40)}..." not found in ${entry.document}`)
    }
  }
  if (orphans.length) {
    return fail("redline-anchors", `${orphans.length} orphaned/ambiguous anchor(s)`, orphans)
  }
  return pass("redline-anchors", `${ledger.length} pending redline(s) still anchored`)
}

// ─── privilege-export ─────────────────────────────────────────
// Three states: no export-requested entries at all → not-applicable (nothing is
// queued for export, so privilege cannot leak through a production); exports
// queued but none privileged → pass; ≥1 entry both privileged AND
// export-requested → fail.
function privilegeExportGate(ctx: GateContext): GateResult {
  const logPath = path.join(ctx.directory, "redaction-log.md")
  if (!existsSync(logPath)) {
    return notApplicable("privilege-export", "no redaction-log.md in this matter")
  }
  const text = readFileSync(logPath, "utf8")
  const exportRequested = extractListEntries(text).filter((e) => /export-requested/i.test(e))
  if (exportRequested.length === 0) {
    return notApplicable("privilege-export", "no entries flagged export-requested in this matter")
  }
  const blocked = exportRequested.filter((e) => /privileged/i.test(e))
  if (blocked.length) {
    return fail(
      "privilege-export",
      `${blocked.length} privileged entry/entries marked for export`,
      blocked,
    )
  }
  return pass("privilege-export", `${exportRequested.length} export-requested entry/entries, none privileged`)
}

function extractListEntries(text: string): string[] {
  return [...text.matchAll(/^-\s+(.+)/gim)].flatMap((match) =>
    match[1] ? [match[1].trim()] : []
  )
}

// ─── shared helpers ───────────────────────────────────────────
export function discoverDeliverable(directory: string): string | null {
  const recorded = recordedDeliverable(directory)
  if (recorded) return recorded

  const sourceDocuments = sourceDocumentNames(directory)
  const internalArtifacts = new Set([
    "citation-check-report.md",
    "coverage-report.md",
    "drafting-workflow-state.md",
    "review-precheck.md",
    "verification-report.md",
  ])
  const candidates = readdirSync(directory)
    .filter((file) => /\.md$/i.test(file))
    .filter((file) => !internalArtifacts.has(file))
    .filter((file) => !sourceDocuments.has(file))
    .filter((file) =>
      !/^(?:section-\d+|research-q[^/]*|source-authority-research-q[^/]*|review-packet-[^/]*)\.md$/i.test(file),
    )
    .map((file) => ({ file, bytes: readFileSync(path.join(directory, file)).byteLength }))
    .filter((candidate) => candidate.bytes > 0)
    .sort((a, b) => b.bytes - a.bytes || a.file.localeCompare(b.file))
  return candidates[0]?.file ?? null
}

function recordedDeliverable(directory: string) {
  const currentRun = readRunId(directory)
  const ledger = readDurableJson(path.join(directory, ".dochaus-deliverable.json"))
  if (isRecord(ledger)) {
    const ledgerRun = typeof ledger.run_id === "string" ? ledger.run_id : undefined
    if (!currentRun || !ledgerRun || currentRun === ledgerRun) {
      const assembledBody = existingMarkdown(directory, ledger.assembledBody)
      if (assembledBody) return assembledBody
    }
  }
  return readLineage(directory)
    .filter((row) => row.stage === "assemble" && row.op !== "delete")
    .filter((row) => !currentRun || row.run === currentRun)
    .map((row) => existingMarkdown(directory, row.artifact))
    .filter((file): file is string => Boolean(file))
    .at(-1) ?? null
}

function sourceDocumentNames(directory: string) {
  const index = readDurableJson(path.join(directory, "source-index.json"))
  if (!isRecord(index) || !Array.isArray(index.sources)) return new Set<string>()
  return new Set(
    index.sources
      .filter(isRecord)
      .map((source) => (typeof source.document === "string" ? path.basename(source.document) : ""))
      .filter(Boolean),
  )
}

function existingMarkdown(directory: string, value: unknown) {
  if (typeof value !== "string" || value !== path.basename(value) || !/\.md$/i.test(value)) return null
  return existsSync(path.join(directory, value)) ? value : null
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value)
}

function pass(id: GateId, summary: string): GateResult {
  return { id, status: "pass", findings: [], summary }
}

function fail(id: GateId, summary: string, detail: string[] | GateFinding[]): GateResult {
  const findings: GateFinding[] = detail.map((d) =>
    typeof d === "string"
      ? { severity: "fail" as const, description: d, detail: d }
      : d,
  )
  return { id, status: "fail", findings, summary }
}

function notApplicable(id: GateId, summary: string): GateResult {
  return { id, status: "not-applicable", findings: [], summary }
}

function formatReport(
  results: Record<GateId, GateResult>,
  deliverable: string | null,
  status: "PASS" | "PARTIAL" | "BLOCKED",
): string {
  const values = Object.values(results)
  const passCount = values.filter((r) => r.status === "pass").length
  const failCount = values.filter((r) => r.status === "fail").length
  const naCount = values.filter((r) => r.status === "not-applicable").length
  const lines: string[] = [
    "# Verification report",
    "",
    `Deliverable: ${deliverable ?? "(none discovered)"}`,
    `Overall: ${status} — ${passCount} pass, ${failCount} fail, ${naCount} not-applicable`,
    "",
  ]
  for (const id of ALL_GATES) {
    const r = results[id]
    if (!r) continue
    const heading = `## ${id}: ${r.status.toUpperCase()}`
    const body = r.summary ? `${heading} (${r.summary})` : heading
    lines.push(body)
    for (const f of r.findings) {
      lines.push(`- ${f.detail ?? f.description}`)
    }
  }
  return lines.join("\n") + "\n"
}
