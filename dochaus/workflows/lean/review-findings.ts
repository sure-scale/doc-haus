import type { CounterpartyItem } from "../../lib/docx-markup"
import { parseJsonObject } from "./change-checks"
import { PID_REFERENCE, TOOL_SYNTAX_ISSUE, printedToolCall, record, type ParagraphRow } from "./checks"
import { counterpartyFindings } from "./counterparty"
import { requestedMemo } from "./deliverables"
import { planActingFor, readLeanPlan, reviewStages, runPosture, type DocID } from "./plan"
import { openItem, quoteFound, quoteMissItems, type SourceQuote } from "./quotes"
import { readLeanJson, shortHash, type OpenItem, type ReviewStageID } from "./state"

// Typed review findings. Each review stage records its findings once, in
// <document>/<stage>.json, and code gives each one an id. A finding names the
// document it concerns: the one under review, another planned document or,
// when the request asks for a memo, the memo. The change plan of that document
// adjudicates it when it is the reviewed document or one still to be drafted;
// a finding for an earlier document, which is final by then, or for the memo
// is disclosed in the notes. The other side's changes and comments in the
// original a document revises join its findings (counterparty.ts). The register
// (register.ts) is a view over these records and the change plans; nothing here
// is recorded twice.
//
// Each finding also says how far its change moves rights or risk toward one
// party, measured against what a neutral, mid-market document of the kind,
// prepared for neither party, would contain, whichever side the run takes. The run's
// Posture decides in code what that permits: unless it acts for a side under
// Client-favorable, a beyond-market finding is held back. It reaches no change
// plan and the notes disclose it as a recommendation naming the side it favours.

const FINDING_ORIGINS = ["request", "source_conflict", "missing_fact", "missing_protection", "law", "drafting"] as const
export type FindingOrigin = (typeof FINDING_ORIGINS)[number]
const FINDING_BALANCES = ["neutral", "market", "beyond_market"] as const
export type FindingBalance = (typeof FINDING_BALANCES)[number]
/** The target of a finding that belongs in the accompanying memo. */
export const MEMO_TARGET = "memo"
export type FindingTarget = DocID | typeof MEMO_TARGET
type ReviewFinding = { id: string; document: FindingTarget; pids?: string[]; origin: FindingOrigin; balance?: FindingBalance; favours?: string; finding: string; instead?: string; sources: SourceQuote[] }
/** A playbook review also records its coverage of the guidance's positions. */
type ReviewRecord = { findings: ReviewFinding[]; coverage?: string }
/** A finding with the document and stage whose review raised it. One of
 * stage and origin "counterparty" is the other side's change or comment in the
 * original the document revises (counterparty.ts); no review reply can give it. */
export type RaisedFinding = Omit<ReviewFinding, "origin"> & { origin: FindingOrigin | "counterparty"; raised_in: DocID; stage: ReviewStageID | "counterparty"; counterparty?: CounterpartyItem }

export const REVIEW_FINDINGS_CONTRACT = 'Reply with exactly one JSON object: {"findings":[{"document":"id of the planned document the finding concerns; leave out for this document","pids":["pid of each paragraph of this draft the finding affects, or where missing wording belongs; leave out for another document"],"origin":"request|source_conflict|missing_fact|missing_protection|law|drafting","balance":"neutral|market|beyond_market","favours":"the party the change favours; leave out when neutral","finding":"what is wrong and why","instead":"what those paragraphs should say instead, in proposed wording where you can give it; leave out for a decision only the lawyer can make","sources":[{"source_id":"exact matter document name, request or approved authority URL","quote":"exact supporting words"}]}]}. ' +
  "origin says where the finding comes from: request (the draft departs from or misses what the request or the lawyer's answers ask), source_conflict (it misstates or conflicts with a matter source, or the sources disagree), missing_fact (it needs a fact or decision the matter does not give), missing_protection (it lacks a protection a careful drafter would include), law (its wording does not work under the governing law) or drafting (ambiguous, inconsistent or wrongly cross-referenced wording). " +
  "balance says how far the change moves rights, obligations, remedies or risk toward one party, measured against what a neutral, mid-market document of this kind, prepared for neither party, would contain, whichever party the lawyer acts for: " +
  "neutral (toward neither, or it carries out the request, the lawyer's answers or a position the sources record as instructed or agreed), " +
  "market (toward one party, as far as such a document commonly goes) or beyond_market (further than that, including terms a party may seek in negotiation that such a document does not commonly contain). " +
  "favours names that party. " +
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
    if (!record(row) || !nonempty(row.finding) || !FINDING_ORIGINS.includes(row.origin as FindingOrigin) || !FINDING_BALANCES.includes(row.balance as FindingBalance)) {
      issues.push(`${at} needs a finding, an origin (${FINDING_ORIGINS.join(", ")}) and a balance (${FINDING_BALANCES.join(", ")}).`)
      return
    }
    const balance = row.balance as FindingBalance
    const favours = balance !== "neutral" && nonempty(row.favours) ? row.favours.trim().replace(/\s+/g, " ") : undefined
    if (balance !== "neutral" && !favours) issues.push(`${at}.favours must name the party the change favours when its balance is market or beyond_market.`)
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
      balance,
      ...(favours ? { favours } : {}),
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
    nonempty(row.finding) && FINDING_ORIGINS.includes(row.origin as FindingOrigin) && Array.isArray(row.sources) && (row.pids === undefined || Array.isArray(row.pids)) &&
    (row.balance === undefined || FINDING_BALANCES.includes(row.balance as FindingBalance)))
  if (findings.length !== value.findings.length) return
  return { findings, ...(typeof value.coverage === "string" ? { coverage: value.coverage } : {}) }
}

/** Every saved review finding of the run, by planned document and review
 * stage, each document's followed by its counterparty items. */
export function raisedFindings(directory: string): RaisedFinding[] {
  return (readLeanPlan(directory)?.documents ?? []).flatMap(({ id }) => [...reviewStages(directory, id).flatMap(stage =>
    (readReviewRecord(directory, id, stage)?.findings ?? []).map(row => ({ ...row, raised_in: id, stage }))), ...counterpartyFindings(directory, id)])
}

/** Whether the run holds back beyond-market findings: it does unless it acts
 * for a side under the Client-favorable Posture frozen into its plan. */
export function holdsBeyondMarket(directory: string) {
  return !planActingFor(directory).stated || runPosture(directory) !== "client-favorable"
}

/** Origins the Posture never holds back: an instruction, a source the draft
 * misstates and a decision only the lawyer can make. */
const KEPT_ORIGINS: RaisedFinding["origin"][] = ["request", "source_conflict", "missing_fact"]

/** Where a finding goes: held back when the run holds back what goes beyond
 * market (`holds`) and the finding does, unless the firm's chosen guidance or
 * a kept origin raised it; otherwise to its document's change plan when that
 * document is the reviewed one or a later one, to the notes under its document
 * when that document is earlier and so final, or to the notes for the memo. A
 * later document's finding never reaches an earlier plan, so keys run one way. */
export function findingRoute(order: Map<string, number>, row: RaisedFinding, holds: boolean): "held" | "plan" | "earlier" | "memo" {
  if (holds && row.balance === "beyond_market" && row.stage !== "review-playbook" && !KEPT_ORIGINS.includes(row.origin)) return "held"
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
  const holds = holdsBeyondMarket(directory)
  const rows = raisedFindings(directory).filter(row => row.document === document && findingRoute(order, row, holds) === "plan")
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
  const holds = holdsBeyondMarket(directory)
  return raisedFindings(directory).flatMap(row => {
    const route = findingRoute(order, row, holds)
    return route === "plan" || route === "held" ? [] : [findingItem(row, route === "earlier" ? row.document as DocID : undefined)]
  })
}

/** Findings raised in `document` that the run held back, as the recommendation
 * dispositions the notes and package review read beside its change plan's. */
export function heldBackDispositions(directory: string, document: DocID) {
  return heldBackFindings(directory).filter(row => row.raised_in === document).map(row => ({
    disposition: "recommendation" as const, kind: "suggestion" as const, detail: findingDetail(row).replace(PID_REFERENCE, "(a paragraph)"), reason: heldBackReason(directory, row),
    documents: [row.document === MEMO_TARGET ? row.raised_in : row.document as DocID], sources: row.sources, limitations: [] as string[],
  }))
}

/** Why a held-back finding was not drafted: the side it favours and the run's Posture. */
export function heldBackReason(directory: string, row: RaisedFinding) {
  return `Not drafted: it favours ${row.favours ?? "one party"} beyond market practice, and ${planActingFor(directory).stated
    ? `the run's Posture is ${runPosture(directory) === "conservative" ? "Conservative" : "Balanced"}` : "the run acts for neither party"}.`
}

/** How many review findings the run held back, for the run header. */
export function heldBackCount(directory: string) {
  return heldBackFindings(directory).length
}

function heldBackFindings(directory: string) {
  const order = documentOrder(directory)
  return holdsBeyondMarket(directory) ? raisedFindings(directory).filter(row => findingRoute(order, row, true) === "held") : []
}

/** The omission item that discloses a finding, filed under `document` when given. */
export function findingItem(row: RaisedFinding, document?: DocID) {
  return openItem("change-plan", "omission", findingDetail(row).replace(PID_REFERENCE, "(a paragraph)"), { ...(document ? { document } : {}), finding: row.id })
}

/** A finding's text with what it says the document should say instead. */
function findingDetail(row: RaisedFinding) {
  return row.instead ? `${row.finding} Instead: ${row.instead}` : row.finding
}

/** Findings as a prompt lists them: each with its id, the review that raised
 * it and, for another document's review, that document. Paragraph pointers
 * stay only on the reviewed document's own findings; a draft is shown no ids.
 * Every "<" is written as its JSON escape, so no finding's text (the other
 * side's included) can close the tag the prompt lists them in. */
export function findingsText(findings: RaisedFinding[], document: DocID, options: { ids?: false } = {}) {
  return JSON.stringify(findings.map(row => {
    const own = row.raised_in === document
    const text = (value: string) => own ? value : value.replace(PID_REFERENCE, "(earlier paragraph)")
    return {
      ...(options.ids === false ? {} : { id: row.id }), review: row.stage, ...(own ? {} : { raised_by: row.raised_in }),
      ...(row.pids ? { pids: row.pids } : {}), origin: row.origin, finding: text(row.finding), ...(row.instead ? { instead: text(row.instead) } : {}), sources: row.sources,
    }
  }), null, 2).replaceAll("<", "\\u003c")
}

/** One saved review as later prompts show it, without finding ids, which only
 * a change plan uses, and without paragraph pointers when shown to another
 * document. */
export function reviewText(directory: string, document: DocID, stage: ReviewStageID, options: { strip?: boolean } = {}) {
  const review = readReviewRecord(directory, document, stage)
  if (!review) return undefined
  // A held-back finding reaches no writer; balance only routes.
  const held = new Set(heldBackFindings(directory).map(row => row.id))
  const text = JSON.stringify({ ...(review.coverage !== undefined ? { coverage: review.coverage } : {}),
    findings: review.findings.filter(row => !held.has(row.id)).map(({ id: _id, pids, balance: _balance, favours: _favours, ...row }) => ({ ...(pids && !options.strip ? { pids } : {}), ...row })) }, null, 2)
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
