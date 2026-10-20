import { parseJsonObject } from "./change-checks"
import { PID_REFERENCE, TOOL_SYNTAX_ISSUE, printedToolCall, record, type ParagraphRow } from "./checks"
import { requestedMemo } from "./deliverables"
import { readLeanPlan, reviewStages, type DocID } from "./plan"
import { openItem, quoteFound, quoteMissItems, type SourceQuote } from "./quotes"
import { readLeanJson, shortHash, type OpenItem, type ReviewStageID } from "./state"

// Typed review findings. Each review stage records its findings once, in
// <document>/<stage>.json, and code gives each one an id. A finding names the
// document it concerns: the one under review, another planned document or,
// when the request asks for a memo, the memo. The change plan of that document
// adjudicates it when it is the reviewed document or one still to be drafted;
// a finding for an earlier document, which is final by then, or for the memo
// is disclosed in the notes. The register (register.ts) is a view over these
// records and the change plans; nothing here is recorded twice.

const FINDING_ORIGINS = ["request", "source_conflict", "missing_fact", "missing_protection", "law", "drafting"] as const
export type FindingOrigin = (typeof FINDING_ORIGINS)[number]
/** The target of a finding that belongs in the accompanying memo. */
export const MEMO_TARGET = "memo"
export type FindingTarget = DocID | typeof MEMO_TARGET
type ReviewFinding = { id: string; document: FindingTarget; pids?: string[]; origin: FindingOrigin; finding: string; instead?: string; sources: SourceQuote[] }
/** A playbook review also records its coverage of the guidance's positions. */
type ReviewRecord = { findings: ReviewFinding[]; coverage?: string }
/** A finding with the document and stage whose review raised it. */
export type RaisedFinding = ReviewFinding & { raised_in: DocID; stage: ReviewStageID }

export const REVIEW_FINDINGS_CONTRACT = 'Reply with exactly one JSON object: {"findings":[{"document":"id of the planned document the finding concerns; leave out for this document","pids":["pid of each paragraph of this draft the finding affects, or where missing wording belongs; leave out for another document"],"origin":"request|source_conflict|missing_fact|missing_protection|law|drafting","finding":"what is wrong and why","instead":"what those paragraphs should say instead, in proposed wording where you can give it; leave out for a decision only the lawyer can make","sources":[{"source_id":"exact matter document name, request or approved authority URL","quote":"exact supporting words"}]}]}. ' +
  "origin says where the finding comes from: request (the draft departs from or misses what the request or the lawyer's answers ask), source_conflict (it misstates or conflicts with a matter source, or the sources disagree), missing_fact (it needs a fact or decision the matter does not give), missing_protection (it lacks a protection a careful drafter would include), law (its wording does not work under the governing law) or drafting (ambiguous, inconsistent or wrongly cross-referenced wording). " +
  "A decision only the lawyer can make is a finding with origin missing_fact saying what must be decided; there is no other list. " +
  "sources may be empty when the finding rests on your view of the law or general practice. Return an empty findings array only when there is nothing to correct or decide."

/** Checks one review reply. Code gives each finding its id from the reviewed
 * document, the stage and the finding's target, paragraphs and text; a repeated
 * finding is kept once. Quotations the sources do not contain are open items,
 * as before. */
export function acceptReviewReply(directory: string, document: DocID, stage: ReviewStageID, reply: string, paragraphs: ParagraphRow[]):
  { issues: string[]; record?: ReviewRecord; openItems: OpenItem[] } {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE], openItems: [] }
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`${REVIEW_FINDINGS_CONTRACT} The reply could not be read: ${parsed.error}.`], openItems: [] }
  const value = parsed.value
  const playbook = stage === "review-playbook"
  if (!Array.isArray(value.findings) || (playbook && !nonempty(value.coverage)))
    return { issues: [playbook ? 'Return "coverage" (your coverage of the guidance positions, in Markdown) and a "findings" array.' : 'Return a "findings" array; it may be empty.'], openItems: [] }
  const targets = findingTargets(directory, document)
  const known = new Set(paragraphs.map(row => row.pid))
  const unknownPids = (at: string, text: string) => {
    const unknown = [...new Set([...text.matchAll(PID_REFERENCE)].map(match => match[1]!).filter(ref => !known.has(ref)))]
    return unknown.length ? [`${at}: these paragraph references do not exist in the draft: ${unknown.map(ref => `[${ref}]`).join(", ")}.`] : []
  }
  const issues: string[] = playbook ? unknownPids("coverage", String(value.coverage)) : []
  const findings = new Map<string, ReviewFinding>()
  value.findings.forEach((row, index) => {
    const at = `findings[${index}]`
    if (!record(row) || !nonempty(row.finding) || !FINDING_ORIGINS.includes(row.origin as FindingOrigin)) {
      issues.push(`${at} needs a finding and an origin (${FINDING_ORIGINS.join(", ")}).`)
      return
    }
    const target = nonempty(row.document) ? row.document.trim() : document
    if (!targets.includes(target)) issues.push(`${at}.document must be one of ${targets.map(id => `"${id}"`).join(", ")}, or left out for this document.`)
    // A pid points into this draft only; none is kept for another document. A
    // single "pid" is read as one of the pids.
    const pids = target === document ? [...new Set([...(Array.isArray(row.pids) ? row.pids : [row.pids]), row.pid].filter(nonempty)
      .map(pid => pid.trim().replace(/^\[(.*)\]$/, "$1")))] : []
    pids.filter(pid => !known.has(pid)).forEach(pid => issues.push(`${at}: [${pid}] is not a paragraph of the draft. Use only the [pid] values listed with the draft.`))
    const instead = nonempty(row.instead) ? row.instead.trim() : undefined
    issues.push(...unknownPids(at, row.finding), ...(instead ? unknownPids(`${at}.instead`, instead) : []))
    const sources = row.sources ?? []
    if (!Array.isArray(sources) || !sources.every(source => record(source) && nonempty(source.source_id) && nonempty(source.quote))) {
      issues.push(`${at}.sources must list {"source_id","quote"} pairs with exact nonempty quotes, or be empty.`)
      return
    }
    const finding: ReviewFinding = {
      id: `f-${shortHash(`${document}/${stage}\0${target}\0${pids.toSorted().join(",")}\0${row.finding.trim().toLowerCase().replace(/\s+/g, " ")}`)}`,
      document: target as FindingTarget,
      ...(pids.length ? { pids } : {}),
      origin: row.origin as FindingOrigin,
      finding: row.finding.trim(),
      ...(instead ? { instead } : {}),
      sources: sources.map(source => ({ source_id: String(source.source_id).trim(), quote: String(source.quote).trim() })),
    }
    if (!findings.has(finding.id)) findings.set(finding.id, finding)
  })
  if (issues.length) return { issues, openItems: [] }
  const rows = [...findings.values()]
  const openItems = [
    ...rows.flatMap(row => [
      ...row.sources.filter(source => !quoteFound(directory, source)).map(source =>
        openItem(stage, "quote_miss", `Quoted text was not found verbatim in source "${source.source_id}".`, { source_id: source.source_id, quote: source.quote })),
      ...quoteMissItems(directory, stage, row.finding),
      ...(row.instead ? quoteMissItems(directory, stage, row.instead) : []),
    ]),
    ...(playbook ? quoteMissItems(directory, stage, String(value.coverage)) : []),
  ]
  return { issues: [], record: { findings: rows, ...(playbook ? { coverage: String(value.coverage).trim() } : {}) }, openItems }
}

/** The saved review of one document and stage; undefined when missing or not
 * of the recorded shape. */
export function readReviewRecord(directory: string, document: DocID, stage: ReviewStageID): ReviewRecord | undefined {
  const value = readLeanJson(directory, `${document}/${stage}.json`)
  if (!record(value) || !Array.isArray(value.findings)) return
  const findings = value.findings.filter((row): row is ReviewFinding => record(row) && nonempty(row.id) && nonempty(row.document) &&
    nonempty(row.finding) && FINDING_ORIGINS.includes(row.origin as FindingOrigin) && Array.isArray(row.sources) && (row.pids === undefined || Array.isArray(row.pids)))
  if (findings.length !== value.findings.length) return
  return { findings, ...(typeof value.coverage === "string" ? { coverage: value.coverage } : {}) }
}

/** Every saved review finding of the run, by planned document and review stage. */
export function raisedFindings(directory: string): RaisedFinding[] {
  return (readLeanPlan(directory)?.documents ?? []).flatMap(({ id }) => reviewStages(directory, id).flatMap(stage =>
    (readReviewRecord(directory, id, stage)?.findings ?? []).map(row => ({ ...row, raised_in: id, stage }))))
}

/** Where a finding goes: to its document's change plan when that document is
 * the reviewed one or a later one, to the notes under its document when that
 * document is earlier and so final, or to the notes for the memo. A later
 * document's finding never reaches an earlier plan, so keys run one way. */
export function findingRoute(order: Map<string, number>, row: RaisedFinding): "plan" | "earlier" | "memo" {
  if (row.document === MEMO_TARGET) return "memo"
  return (order.get(row.document) ?? 0) < (order.get(row.raised_in) ?? 0) ? "earlier" : "plan"
}

/** Each planned document's place in plan order. */
export function documentOrder(directory: string) {
  return new Map((readLeanPlan(directory)?.documents ?? []).map((plan, index) => [plan.id as string, index]))
}

/** The findings `document`'s change plan adjudicates: its own reviews'
 * findings about it, then earlier documents' findings about it. */
export function findingsFor(directory: string, document: DocID): RaisedFinding[] {
  const order = documentOrder(directory)
  const rows = raisedFindings(directory).filter(row => row.document === document && findingRoute(order, row) === "plan")
  return [...rows.filter(row => row.raised_in === document), ...rows.filter(row => row.raised_in !== document)]
}

/** Earlier documents' findings about `document`, which its draft is shown. */
export function forwardedFindings(directory: string, document: DocID) {
  return findingsFor(directory, document).filter(row => row.raised_in !== document)
}

/** Findings no change plan adjudicates: one about an earlier document, filed
 * under that document, and one for the memo. Each is an item the notes disclose. */
export function routedFindingItems(directory: string): OpenItem[] {
  const order = documentOrder(directory)
  return raisedFindings(directory).flatMap(row => {
    const route = findingRoute(order, row)
    return route === "plan" ? [] : [findingItem(row, route === "earlier" ? row.document as DocID : undefined)]
  })
}

/** The omission item that discloses a finding, filed under `document` when given. */
export function findingItem(row: RaisedFinding, document?: DocID) {
  return openItem("change-plan", "omission", findingDetail(row).replace(PID_REFERENCE, "(a paragraph)"), { ...(document ? { document } : {}), finding: row.id })
}

/** A finding's text with what it says the document should say instead. */
function findingDetail(row: ReviewFinding) {
  return row.instead ? `${row.finding} Instead: ${row.instead}` : row.finding
}

/** Findings as a prompt lists them: each with its id, the review that raised
 * it and, for another document's review, that document. Paragraph pointers
 * stay only on the reviewed document's own findings; a draft is shown no ids. */
export function findingsText(findings: RaisedFinding[], document: DocID, options: { ids?: false } = {}) {
  return JSON.stringify(findings.map(row => {
    const own = row.raised_in === document
    const text = (value: string) => own ? value : value.replace(PID_REFERENCE, "(earlier paragraph)")
    return {
      ...(options.ids === false ? {} : { id: row.id }), review: row.stage, ...(own ? {} : { raised_by: row.raised_in }),
      ...(row.pids ? { pids: row.pids } : {}), origin: row.origin, finding: text(row.finding), ...(row.instead ? { instead: text(row.instead) } : {}), sources: row.sources,
    }
  }), null, 2)
}

/** One saved review as later prompts show it, without finding ids, which only
 * a change plan uses, and without paragraph pointers when shown to another
 * document. */
export function reviewText(directory: string, document: DocID, stage: ReviewStageID, options: { strip?: boolean } = {}) {
  const review = readReviewRecord(directory, document, stage)
  if (!review) return undefined
  const text = JSON.stringify({ ...(review.coverage !== undefined ? { coverage: review.coverage } : {}),
    findings: review.findings.map(({ id: _id, pids, ...row }) => ({ ...(pids && !options.strip ? { pids } : {}), ...row })) }, null, 2)
  return options.strip ? text.replace(PID_REFERENCE, "(earlier paragraph)") : text
}

/** Where a finding raised in `document` may be directed: any planned document
 * and, when the request asks for a memo, the memo. */
function findingTargets(directory: string, document: DocID): string[] {
  const planned = readLeanPlan(directory)?.documents.map(row => row.id) ?? [document]
  return [...planned, ...(requestedMemo(directory) ? [MEMO_TARGET] : [])]
}

function nonempty(value: unknown): value is string {
  return typeof value === "string" && Boolean(value.trim())
}
