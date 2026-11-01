import { normalize } from "../../lib/quote-verify"
import { nameWithProvenance, sourceProvenance } from "../../lib/source-provenance"
import { TOOL_SYNTAX_ISSUE, printedToolCall, record } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { instructionSections, listedActIssues, mergeCoverageActs, namingTexts } from "./package-coverage"
import { nameKey, namedIn } from "./plan"
import { retryPrompt, tagged } from "./prompts"
import { REQUEST_SOURCE, adoptedSources, excerpt, matterSourceIDs, proseQuoteIssues, quoteFound, quoteMissIssue } from "./quotes"
import { approvedResearch, researchHandoff } from "./research"
import type { StageRun } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, type LeanDeps } from "./stages"
import { readLeanJson, writeLeanText, type LeanStageState, type LeanState } from "./state"

// The act record: what the lawyer's instructions direct, read in focused jobs
// that each answer one question. WHAT lists the legal acts the instructions
// direct; it runs as ACT_RECORD_SAMPLES samples in sessions of their own, and
// their verified lists are joined (mergeCoverageActs), so an act one sample
// misses another can still list. VERIFY then reads that union once, every
// candidate with its words and source: each is a legal act, the same legal step
// as another candidate it keeps, or not an act, of one kind with a reason.
// Applied in code, a candidate that is not an act and a repeat stay in the
// record (verify) and only the kept acts go on; a VERIFY that never passes
// keeps the whole union, so a failed check never drops an act. DELIVERABLES
// lists what the instructions ask to be prepared or delivered, alongside VERIFY
// and blind to it. WHO names, for each listed act, the parties whose
// approval authorises it and those who carry it out; HOW, shown each act's WHO,
// names the instrument that effects it, its signatories and what must exist or
// be taken first, and records any conflict between the instructions and the
// research or sources over who authorises or takes the act, which document
// carries it or who signs it: both sides' words, each verified, never resolved.
// Each job is one checked exchange: a deterministic check
// verifies every quote in its source and every party in the record, and its
// issues go back to the same session for a bounded retry. An empty list is a
// valid answer, so it is never retried. A job that never passes leaves its field
// null for every act and the record keeps that job's issues; the acts stand.
// Nothing here judges the legal substance or supplies a party. Whether an act's
// instrument is executed by the parties whose authority it needs is read from
// the record when it is used (signerFlags), never stored. A recorded conflict
// reaches the lawyer as an open item (package-coverage.ts, actRecordItems).
//
// Every job's prompt starts with the same pack (actRecordPack): the task line,
// the instructions, the approved research, the source names and the rules. Only
// the tail after it differs, and every job runs as the same agent under the run's
// owner session, so a provider can reuse the shared prefix. Matter documents are
// named, never inlined: the jobs retrieve what they need.

export const LEAN_ACTS_SCHEMA = "dochaus.lean.acts/3"
/** The one agent every act-record job runs as. */
export const ACT_RECORD_AGENT = "lean-act-record"
/** WHAT samples, run in parallel and joined. */
export const ACT_RECORD_SAMPLES = 4

/** One entry of evidence: an exact source name and the words that support a value. */
export type ActEvidence = { source: string; quote: string }
/** One act the instructions direct: what is to be done, the words that direct it and their source. */
export type ListedAct = { act: string; source: string; quote: string }
/** Who must authorise an act and who carries it out; [] where the record does
 * not settle a list, with `reason` saying what it leaves unsettled. */
export type ActWho = { authorised_by: string[]; performed_by: string[]; basis: ActEvidence[]; reason: string | null }
/** Something that must exist or be taken before an act: a listed act by its index, or null for anything else. */
export type ActRequirement = { act: number | null; need: string; basis: ActEvidence[] }
/** Where the instructions direct who authorises or takes an act, which document
 * carries it or who signs it, and the approved research or a matter document
 * requires otherwise: the instructions' words, the words that require
 * otherwise and what differs. Stated for the lawyer, never resolved. */
export type ActConflict = { instruction: ActEvidence; authority: ActEvidence; difference: string }
/** The instrument that effects an act (null when the record does not settle
 * it), its signatories, what the act requires first and any conflict between
 * the instructions and the authority (null when there is none). */
export type ActHow = { instrument: string | null; executed_by: string[]; requires: ActRequirement[]; basis: ActEvidence[]; reason: string | null; conflict: ActConflict | null }
/** A listed act with its WHO and HOW, each null when that job never passed. */
export type RecordedAct = ListedAct & { who: ActWho | null; how: ActHow | null }
/** One deliverable the instructions ask for, with the words that ask for it. */
export type Deliverable = { deliverable: string; source: string; quote: string }
export type ActRecordJob = "what" | "verify" | "deliverables" | "who" | "how"
/** Why a listed candidate is not a legal act: the preparation of a document,
 * the advisers' own work on the matter, or what a document is to provide. */
export const NON_ACT_KINDS = ["document", "firm_work", "term"] as const
export type NonActKind = typeof NON_ACT_KINDS[number]
/** VERIFY's reading of one candidate, by its index in the union. */
export type ActDecision = { decision: "act" } | { decision: "same"; same_as: number } | { decision: "not_act"; kind: NonActKind; reason: string }
/** The candidates VERIFY did not keep as acts: each one not an act, with its
 * kind and reason, and each repeat of a kept act, with that act's index in
 * `acts` (`into`), its words kept as further support for that act. */
export type ActVerify = { excluded: (ListedAct & { kind: NonActKind; reason: string })[]; merged: (ListedAct & { into: number })[] }
/** "listed": at least one act; "none": a WHAT sample passed and no candidate
 * is kept as an act; "unchecked": no WHAT sample passed its check, so no act is
 * recorded and `issues` holds the first sample's. `verify` is null when VERIFY
 * was not asked (no candidate) or never passed, and then `acts` is the whole
 * union. `deliverables` is null when that job never passed. `issues` holds the
 * last issues of each job that never passed. */
export type ActRecord = {
  schema: typeof LEAN_ACTS_SCHEMA
  status: "listed" | "none" | "unchecked"
  acts: RecordedAct[]
  verify: ActVerify | null
  deliverables: Deliverable[] | null
  issues: { job: ActRecordJob; issues: string[] }[]
}
/** One act's signer check: "ok" when its instrument's signatories include every
 * party whose authority it needs, "flag" when they leave one out (`missing`),
 * and "unchecked" when the record names no authorising party or no HOW. */
export type SignerFlag = { act: number; status: "ok" | "flag" | "unchecked"; missing: string[] }

type Checked<T> = { issues: string[]; value?: T }

const BASIS_SHAPE = '[{"source":"...","quote":"..."}]'
const WHAT_SHAPE = 'Reply with the JSON object only: {"acts":[{"act":"...","source":"...","quote":"..."}]}.'
const VERIFY_SHAPE = 'Reply with the JSON object only: {"acts":[{"index":0,"verify":{"decision":"act","same_as":null,"kind":null,"reason":null}}]}.'
const DELIVERABLES_SHAPE = 'Reply with the JSON object only: {"deliverables":[{"deliverable":"...","source":"...","quote":"..."}]}.'
const WHO_SHAPE = `Reply with the JSON object only: {"acts":[{"index":0,"who":{"authorised_by":["..."],"performed_by":["..."],"basis":${BASIS_SHAPE},"reason":null}}]}.`
const CONFLICT_SHAPE = '{"instruction":{"source":"...","quote":"..."},"authority":{"source":"...","quote":"..."},"difference":"..."}'
const HOW_SHAPE = `Reply with the JSON object only: {"acts":[{"index":0,"how":{"instrument":"...","executed_by":["..."],"requires":[{"act":null,"need":"...","basis":${BASIS_SHAPE}}],"basis":${BASIS_SHAPE},"reason":null,"conflict":null}}]}.`
const EVERY_ACT = "Give every listed act exactly once, by its index."

const WHAT_JOB = [
  "List every legal act that the request, the lawyer's updates or answers direct to be taken or documented: in the requested document or in any other document, now or later.",
  "A matter document directs an act only where the request adopts it as instructions; an obligation a matter document merely describes is not an instructed act. Do not list an act the instructions say is not to be taken at all.",
  'For each act give "act": what is to be done, briefly; "source": the exact source name of the words that direct it; "quote": those words, copied verbatim from that source.',
].join("\n")

const VERIFY_JOB = [
  "Each candidate below was listed by a reading of the instructions as a legal act, with the words that direct it and their source. Decide whether each is one.",
  "A legal act is a step with legal effect, taken by a party, body, person, court or authority, that someone must authorise and some document must carry. " +
  "Test each candidate by asking who authorises it and which document carries it: a legal act has a sensible answer to both. " +
  "It is something done, not what a document is to say or the work of preparing one. Whose legal step it is decides, not who does the physical work: a party's step remains an act when its advisers carry it out.",
  'For each candidate give "verify": {"decision": ..., "same_as": ..., "kind": ..., "reason": ...}.',
  '- "decision": "act" for a legal act; "same" for the same legal step as another candidate given as "act", only worded or sourced differently; "not_act" for anything else.',
  '- "same_as": for "same", the index of the candidate it repeats; otherwise null.',
  '- "kind": for "not_act", "document" for the preparation of a document rather than a step the document takes or records, "firm_work" for the advisers\' own work on the matter, ' +
  'or "term" for what a document is to provide; otherwise null.',
  '- "reason": for "not_act", why, briefly; otherwise null. Write any words you quote as "exact words" [source: exact source name].',
].join("\n")

const DELIVERABLES_JOB = [
  "List every document or other deliverable that the request, the lawyer's updates or answers ask to be prepared or delivered. A matter document asks for one only where the request adopts it as instructions.",
  'For each give "deliverable": what is to be prepared, briefly; "source": the exact source name of the words that ask for it; "quote": those words, copied verbatim from that source.',
].join("\n")

const EVIDENCE_RULE = 'Each "basis" entry is {"source": an exact source name, "quote": the exact words that support the value, copied verbatim from that source}. ' +
  "Give every non-null value at least one basis entry. A value the evidence does not support is null (for parties, []) with its reason; never supply a party, instrument or requirement by assumption."

const WHO_JOB = [
  'For each listed act below, give "who": {"authorised_by": [...], "performed_by": [...], "basis": [evidence], "reason": ...}.',
  '- "authorised_by": every party or body whose approval or agreement legally authorises the act, without which the act is not effective.',
  '- "performed_by": every party who carries the act out or signs for it.',
  "- Name each party as the sources name it; give [] where the record does not settle it. Who must authorise or take an act is settled by the law and the record, not by which document the instructions put it in: " +
  "where the legal research or a source says a party of some kind must authorise or take an act of that kind, name the party of that kind the sources name.",
  '- "basis": the evidence for both lists.',
  '- "reason": null when both are given; otherwise what the record leaves unsettled.',
  `- ${EVIDENCE_RULE}`,
].join("\n")

const HOW_JOB = [
  'For each listed act below, give "how": {"instrument": ..., "executed_by": [...], "requires": [...], "basis": [evidence], "reason": ..., "conflict": ...}. Each act\'s WHO is given with it below.',
  '- "instrument": the instrument or document that effects the act, and the part of it that takes the act. ' +
  "Where WHO names no authorising party, still give the instrument that effects the act. null when the record does not settle it.",
  '- "executed_by": every signatory of that instrument, using the WHO name for a party already named there, or [] when the record does not settle it.',
  '- "requires": the other listed acts, or other approvals or things, that must exist or be taken before this act can be effectively taken, ' +
  'each as {"act": its index, or null when it is not a listed act, "need": what is required, briefly, "basis": [evidence]}; [] when the record sets none.',
  '- "basis": the evidence for each value given.',
  '- "reason": null when instrument and executed_by are both given; otherwise what the record leaves unsettled.',
  '- "conflict": null, unless the instructions direct who authorises or takes the act, which document carries it or who signs it, and the approved legal research or a matter document requires otherwise. ' +
  `Then give ${CONFLICT_SHAPE}: "instruction", one evidence entry with the instructions' words that direct it; "authority", one evidence entry with the research or matter document words that require otherwise; ` +
  '"difference", what differs, briefly. Give "instrument" and "executed_by" as the instructions direct them. The record states the difference for the lawyer; do not resolve it.',
  `- ${EVIDENCE_RULE}`,
].join("\n")

/** The material every act-record prompt starts with, byte for byte: built only
 * from the run's saved instructions, approved research and source names, in a
 * fixed order, with no per-job value. */
export function actRecordPack(directory: string) {
  const provenance = sourceProvenance(directory)
  return [
    "This is a lean drafting run. The material below serves several focused questions about the legal acts the lawyer's instructions direct; the question this reply answers follows it.",
    ...instructionSections(directory),
    tagged("legal_research", "Approved legal research: each finding's proposition, its source URL and its exact supporting words. Missing research is a limitation, never an invitation to rely on model memory.\n" +
      researchHandoff(directory)),
    tagged("sources", [
      "Matter documents (use these exact names with search-document and read-document, and as source names):",
      ...matterSourceIDs(directory).map(name => `- ${nameWithProvenance(name, provenance)}`),
      `To quote the request, the lawyer's updates or the lawyer's own answers, use the source name "${REQUEST_SOURCE}". To quote approved legal research, use its finding's "url" as the source name.`,
      ...adoptedSources(directory).map(name => `To quote the recommendation adopted for question ${name.slice("adopted:".length)}, which the lawyer delegated, use the source name "${name}"; its words are the drafter's judgment, not the lawyer's.`),
    ].join("\n")),
    tagged("rules", [
      "- Read the matter documents you rely on with search-document or read-document.",
      "- Copy each quote verbatim from the source it names, as one continuous passage, and check it with verify-quote.",
      "- Never supply an act, party, instrument or requirement by assumption.",
      "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks the reply.",
    ].join("\n")),
  ].filter(Boolean).join("\n\n")
}

/** A reply's JSON value, bare or fenced; undefined when it is not JSON. */
function replyJson(reply: string): unknown {
  try { return JSON.parse(reply.trim().replace(/^```(?:json)?\s*/, "").replace(/\s*```$/, "")) } catch {}
}

const text = (row: Record<string, unknown>, key: string) => typeof row[key] === "string" ? (row[key] as string).trim() : ""

/** A listing reply's rows: each a nonempty `field`, "source" and "quote", the
 * source a matter document, "request" or an adopted recommendation and the quote found in it
 * (listedActIssues). Every defect is retry feedback; the rows are returned only
 * when there is none, each once (by its text, source and quote, as the quote
 * checks normalise them). An empty list is valid. */
function parseListing(reply: string, directory: string, list: "acts" | "deliverables", field: "act" | "deliverable", shape: string): Checked<ListedAct[]> {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const value = replyJson(reply)
  if (!record(value) || !Array.isArray(value[list])) return { issues: [shape] }
  const issues: string[] = []
  const rows: ListedAct[] = []
  ;(value[list] as unknown[]).forEach((row, index) => {
    if (!record(row) || !text(row, field) || !text(row, "source") || !text(row, "quote"))
      return void issues.push(`${list}[${index}] needs nonempty "${field}", "source" and "quote" strings. ${shape}`)
    const listed = { act: text(row, field), source: text(row, "source"), quote: text(row, "quote") }
    issues.push(...listedActIssues(directory, `${list}[${index}] ("${excerpt(listed.act, 80)}")`, { ...listed, parties: [] }, () => []))
    rows.push(listed)
  })
  if (issues.length) return { issues }
  const seen = new Set<string>()
  return { issues, value: rows.filter(row => {
    const key = JSON.stringify([normalize(row.act).toLowerCase(), row.source, normalize(row.quote)])
    if (seen.has(key)) return false
    seen.add(key)
    return true
  }) }
}

/** Reads one WHAT sample's reply: the legal acts the instructions direct. */
export function parseActList(reply: string, directory: string): Checked<ListedAct[]> {
  return parseListing(reply, directory, "acts", "act", WHAT_SHAPE)
}

/** Reads the DELIVERABLES reply: what the instructions ask to be prepared or delivered. */
export function parseDeliverables(reply: string, directory: string): Checked<Deliverable[]> {
  const checked = parseListing(reply, directory, "deliverables", "deliverable", DELIVERABLES_SHAPE)
  return checked.value ? { issues: [], value: checked.value.map(row => ({ deliverable: row.act, source: row.source, quote: row.quote })) } : { issues: checked.issues }
}

/** Sources a basis entry may name: the instructions, an adopted recommendation, a matter document or an approved research finding's URL. */
function evidenceSources(directory: string) {
  return [...new Set([REQUEST_SOURCE, ...adoptedSources(directory), ...matterSourceIDs(directory), ...approvedResearch(directory).map(row => row.url)])]
}

/** A WHO or HOW reply's entries by index: each listed act exactly once. The
 * index defects are retry feedback alongside each entry's own. */
function indexedEntries(reply: string, count: number, field: "verify" | "who" | "how", shape: string): { issues: string[]; rows?: Map<number, Record<string, unknown>> } {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const value = replyJson(reply)
  if (!record(value) || !Array.isArray(value.acts)) return { issues: [`${shape} ${EVERY_ACT}`] }
  const issues: string[] = []
  const rows = new Map<number, Record<string, unknown>>()
  const again = new Set<number>()
  value.acts.forEach((row: unknown, at) => {
    if (!record(row) || !record(row[field]))
      return void issues.push(`acts[${at}] needs "index" as the index of a listed act and "${field}" as an object. ${shape}`)
    const index = row.index
    if (typeof index !== "number" || !Number.isInteger(index) || index < 0 || index >= count)
      return void issues.push(`acts[${at}] gives "index": ${JSON.stringify(index) ?? "nothing"}, which is not the index of a listed act. ` +
        (count > 1 ? `Use the indexes 0 to ${count - 1}.` : "Use the index 0."))
    if (rows.has(index)) return void again.add(index)
    rows.set(index, row[field] as Record<string, unknown>)
  })
  const listing = (indexes: number[]) => `act${indexes.length > 1 ? "s" : ""} ${indexes.join(", ")}`
  if (again.size) issues.push(`More than one entry for ${listing([...again].toSorted((a, b) => a - b))}. ${EVERY_ACT}`)
  const missing = Array.from({ length: count }, (_, index) => index).filter(index => !rows.has(index))
  if (missing.length) issues.push(`No entry for ${listing(missing)}. ${EVERY_ACT}`)
  return { issues, rows }
}

/** Reads the VERIFY reply over the union's `candidates`: every candidate once,
 * by its index, with a decision; a repeat ("same") names another candidate
 * given as "act"; a candidate that is not an act gives one of NON_ACT_KINDS and
 * a reason, each quotation in which is found in its source (proseQuoteIssues).
 * Every defect is retry feedback; the decisions are returned, in index order,
 * only when there is none. */
export function parseActVerify(reply: string, directory: string, candidates: ListedAct[]): Checked<ActDecision[]> {
  const read = indexedEntries(reply, candidates.length, "verify", VERIFY_SHAPE)
  if (!read.rows) return { issues: read.issues }
  const issues = [...read.issues]
  const values = new Map<number, ActDecision>()
  for (const [index, verify] of read.rows) {
    const where = `candidate ${index} ("${excerpt(candidates[index].act, 80)}")`
    if (verify.decision === "act") values.set(index, { decision: "act" })
    else if (verify.decision === "same") {
      const same = verify.same_as
      if (typeof same !== "number" || !Number.isInteger(same) || same < 0 || same >= candidates.length || same === index)
        issues.push(`${where}: "verify.same_as" is ${JSON.stringify(same) ?? "missing"}, which is not the index of another candidate. For "same", give the index of the candidate given as "act" that it repeats.`)
      else values.set(index, { decision: "same", same_as: same })
    } else if (verify.decision === "not_act") {
      const reason = text(verify, "reason")
      if (!NON_ACT_KINDS.includes(verify.kind as NonActKind))
        issues.push(`${where}: "verify.kind" is ${JSON.stringify(verify.kind) ?? "missing"}. For "not_act", give one of ${NON_ACT_KINDS.map(kind => `"${kind}"`).join(", ")}.`)
      if (!reason) issues.push(`${where}: "verify.reason" must say briefly why the candidate is not a legal act.`)
      const quoted = reason ? proseQuoteIssues(directory, `${where}: "verify.reason"`, reason, { format: "plain" }) : []
      issues.push(...quoted)
      if (reason && !quoted.length && NON_ACT_KINDS.includes(verify.kind as NonActKind)) values.set(index, { decision: "not_act", kind: verify.kind as NonActKind, reason })
    } else issues.push(`${where}: "verify.decision" is ${JSON.stringify(verify.decision) ?? "missing"}. Give "act", "same" or "not_act".`)
  }
  for (const [index, decision] of values) if (decision.decision === "same" && values.get(decision.same_as)?.decision !== "act" && read.rows.has(decision.same_as))
    issues.push(`candidate ${index} ("${excerpt(candidates[index].act, 80)}"): "verify.same_as" is ${decision.same_as}, which is not given as "act". ` +
      'A repeat names the candidate kept as the act: give that candidate "decision": "act", or name the one that is.')
  return issues.length ? { issues } : { issues, value: candidates.map((_, index) => values.get(index)!) }
}

/** VERIFY's decisions applied to the union, in its order: the kept acts, then
 * each candidate not kept, never dropped: a repeat by the index of the act it
 * repeats among the kept acts, a candidate that is not an act with its kind
 * and reason. */
export function applyActVerify(candidates: ListedAct[], decisions: ActDecision[]): { acts: ListedAct[]; verify: ActVerify } {
  const kept = candidates.flatMap((_, index) => decisions[index].decision === "act" ? [index] : [])
  return {
    acts: kept.map(index => candidates[index]),
    verify: {
      excluded: decisions.flatMap((decision, index) => decision.decision === "not_act" ? [{ ...candidates[index], kind: decision.kind, reason: decision.reason }] : []),
      merged: decisions.flatMap((decision, index) => decision.decision === "same" ? [{ ...candidates[index], into: kept.indexOf(decision.same_as) }] : []),
    },
  }
}

/** A list of parties, each a nonempty name the instructions or a matter
 * document names (namedIn), each once (by nameKey). */
function partyList(where: string, field: string, value: unknown, texts: () => string[], reasonField: string): Checked<string[]> {
  if (!Array.isArray(value) || !value.every(party => typeof party === "string" && party.trim()))
    return { issues: [`${where}: "${field}" must be a list of party names, each a nonempty string ([] when the record does not settle it).`] }
  const named = (value as string[]).map(party => party.trim())
  const parties = named.filter((party, at) => named.findIndex(other => nameKey(other) === nameKey(party)) === at)
  const issues = parties.filter(party => !texts().some(source => namedIn(source, party))).map(party =>
    `${where}: "${field}" names the party "${party}", but neither the request, the lawyer's updates and answers nor any matter document names that party. ` +
    `Name each party as the sources name it, or give [] and say in "${reasonField}" what the record leaves unsettled.`)
  return issues.length ? { issues } : { issues, value: parties }
}

/** A "basis" list: each entry an allowed source and words found in it
 * (quoteFound, which also reads approved research by URL). */
function basisList(directory: string, where: string, field: string, value: unknown, sources: () => string[]): Checked<ActEvidence[]> {
  if (!Array.isArray(value) || !value.every(entry => record(entry) && text(entry, "source") && text(entry, "quote")))
    return { issues: [`${where}: "${field}" must be a list of {"source": an exact source name, "quote": the exact supporting words}, each nonempty.`] }
  const basis = (value as Record<string, unknown>[]).map(entry => ({ source: text(entry, "source"), quote: text(entry, "quote") }))
  const issues = basis.flatMap((entry, at) => {
    if (!sources().includes(entry.source))
      return [`${where}: "${field}[${at}].source" is "${entry.source}", which is not a source of this matter. Use one of ${sources().map(name => `"${name}"`).join(", ")}.`]
    if (quoteFound(directory, { quote: entry.quote, source_id: entry.source })) return []
    const place = `${where}: "${field}[${at}].quote"`
    return [/^https?:\/\//.test(entry.source)
      ? `${place} quotes "${excerpt(entry.quote)}" as from "${entry.source}", but those words are not in that approved research source. Copy the exact words of its finding.`
      : quoteMissIssue(directory, place, { quote: entry.quote, source_id: entry.source }, "json", { adopted: true })]
  })
  return issues.length ? { issues } : { issues, value: basis }
}

/** A "reason": null or a string; required, as what the record leaves
 * unsettled, when `unsettled` names a value left empty. */
function reasonValue(where: string, field: string, value: unknown, unsettled: string[]): Checked<string | null> {
  if (value !== undefined && value !== null && typeof value !== "string") return { issues: [`${where}: "${field}" must be null or a string.`] }
  const reason = typeof value === "string" ? value.trim() : ""
  if (!reason && unsettled.length) return { issues: [`${where}: ${unsettled.join(" and ")}, so "${field}" must say what the record leaves unsettled.`] }
  return { issues: [], value: reason || null }
}

/** Reads the WHO reply over the listed `acts`: every act once, by its index,
 * each party named in the record, each basis quote found in its source, a basis
 * for any party given and a reason for any list left empty. Every defect is
 * retry feedback; the values are returned, in index order, only when there is
 * none. */
export function parseActWho(reply: string, directory: string, acts: ListedAct[]): Checked<ActWho[]> {
  const read = indexedEntries(reply, acts.length, "who", WHO_SHAPE)
  if (!read.rows) return { issues: read.issues }
  let texts: string[] | undefined
  let sources: string[] | undefined
  const naming = () => texts ??= namingTexts(directory)
  const allowed = () => sources ??= evidenceSources(directory)
  const issues = [...read.issues]
  const values = new Map<number, ActWho>()
  for (const [index, who] of read.rows) {
    const where = `act ${index} ("${excerpt(acts[index].act, 80)}")`
    const authorised = partyList(where, "who.authorised_by", who.authorised_by, naming, "who.reason")
    const performed = partyList(where, "who.performed_by", who.performed_by, naming, "who.reason")
    const basis = basisList(directory, where, "who.basis", who.basis, allowed)
    const reason = reasonValue(where, "who.reason", who.reason, [
      ...(authorised.value?.length === 0 ? ['"who.authorised_by" is []'] : []), ...(performed.value?.length === 0 ? ['"who.performed_by" is []'] : []),
    ])
    issues.push(...authorised.issues, ...performed.issues, ...basis.issues, ...reason.issues)
    if ((authorised.value?.length || performed.value?.length) && basis.value?.length === 0)
      issues.push(`${where}: "who" names parties, but "who.basis" is empty. Give at least one basis entry with the exact words that support them, or give [] with a reason where no evidence does.`)
    if (authorised.value && performed.value && basis.value && !reason.issues.length)
      values.set(index, { authorised_by: authorised.value, performed_by: performed.value, basis: basis.value, reason: reason.value ?? null })
  }
  return issues.length ? { issues } : { issues, value: acts.map((_, index) => values.get(index)!) }
}

/** The "requires" list of act `self`: each a nonempty need, a listed act other
 * than `self` or null, and at least one basis entry found in its source. */
function requirementList(directory: string, where: string, value: unknown, self: number, count: number, sources: () => string[]): Checked<ActRequirement[]> {
  if (!Array.isArray(value))
    return { issues: [`${where}: "how.requires" must be a list of {"act": the index of a listed act or null, "need": "...", "basis": ${BASIS_SHAPE}} ([] when the record sets none).`] }
  const issues: string[] = []
  const requires: ActRequirement[] = []
  value.forEach((entry: unknown, at) => {
    const field = `how.requires[${at}]`
    if (!record(entry) || !text(entry, "need"))
      return void issues.push(`${where}: "${field}" needs "need" as a nonempty string saying briefly what is required, "act" as the index of a listed act or null, and "basis".`)
    const act = entry.act ?? null
    const listed = count > 1 ? `the listed acts are 0 to ${count - 1}` : "the only listed act is 0"
    if (act !== null && (typeof act !== "number" || !Number.isInteger(act) || act < 0 || act >= count))
      issues.push(`${where}: "${field}.act" is ${JSON.stringify(act)}, which is not the index of a listed act (${listed}). Give the index of another listed act, or null when the requirement is not a listed act.`)
    else if (act === self)
      issues.push(`${where}: "${field}.act" is ${self}, the act itself. An act cannot require itself: give the index of another listed act it needs, or null with "need" saying what else is required.`)
    const basis = basisList(directory, where, `${field}.basis`, entry.basis, sources)
    issues.push(...basis.issues)
    if (basis.value?.length === 0)
      issues.push(`${where}: "${field}.basis" is empty. Give at least one basis entry with the exact words that set the requirement, or leave the requirement out when no evidence supports it.`)
    requires.push({ act: act as number | null, need: text(entry, "need"), basis: basis.value ?? [] })
  })
  return issues.length ? { issues } : { issues, value: requires }
}

/** A "conflict": null (or left out) when there is none; otherwise the
 * instructions' words, from the request, an adopted recommendation or a matter
 * document, and the words that require otherwise, from approved research or another matter document,
 * each one entry found in its source (basisList), and a nonempty difference. */
function conflictValue(directory: string, where: string, value: unknown, sources: () => string[]): Checked<ActConflict | null> {
  if (value === undefined || value === null) return { issues: [], value: null }
  const shape = `${where}: "how.conflict" must be null when the instructions and the authority agree, or ${CONFLICT_SHAPE}.`
  if (!record(value) || !record(value.instruction) || !record(value.authority) || !text(value, "difference")) return { issues: [shape] }
  const documents = matterSourceIDs(directory)
  const adopted = adoptedSources(directory)
  const instruction = basisList(directory, where, "how.conflict.instruction", [value.instruction], () => [REQUEST_SOURCE, ...adopted, ...documents])
  const authority = basisList(directory, where, "how.conflict.authority", [value.authority], () => sources().filter(source => source !== REQUEST_SOURCE && !adopted.includes(source)))
  const issues = [...instruction.issues, ...authority.issues]
  if (!issues.length && instruction.value![0]!.source === authority.value![0]!.source)
    issues.push(`${where}: "how.conflict" gives "${authority.value![0]!.source}" as the source of both sides. "instruction" quotes the instructions; "authority" quotes the approved research or another matter document that requires otherwise.`)
  return issues.length ? { issues } : { issues, value: { instruction: instruction.value![0]!, authority: authority.value![0]!, difference: text(value, "difference") } }
}

/** The first circle of listed acts each requiring the next, searched in index
 * order, as [a, b, ..., a]; undefined when there is none. */
function requiresCycle(edges: number[][]): number[] | undefined {
  const seen = edges.map(() => 0)
  const path: number[] = []
  const visit = (node: number): number[] | undefined => {
    seen[node] = 1
    path.push(node)
    for (const next of edges[node]) {
      if (seen[next] === 1) return [...path.slice(path.indexOf(next)), next]
      const found = seen[next] === 0 ? visit(next) : undefined
      if (found) return found
    }
    seen[node] = 2
    path.pop()
    return undefined
  }
  for (const node of edges.keys()) {
    const found = seen[node] === 0 ? visit(node) : undefined
    if (found) return found
  }
  return undefined
}

/** Reads the HOW reply over the listed `acts`: every act once, by its index,
 * the instrument a nonempty string or null, each signatory named in the record,
 * each requirement a need with its basis and a listed act other than itself or
 * null, any conflict with both sides' words found in their sources
 * (conflictValue), each basis quote found in its source, a basis for any instrument or
 * signatory given and a reason for any left unsettled. Requirements between
 * listed acts must not form a circle. Every defect is retry feedback; the values
 * are returned, in index order, only when there is none. */
export function parseActHow(reply: string, directory: string, acts: ListedAct[]): Checked<ActHow[]> {
  const read = indexedEntries(reply, acts.length, "how", HOW_SHAPE)
  if (!read.rows) return { issues: read.issues }
  let texts: string[] | undefined
  let sources: string[] | undefined
  const naming = () => texts ??= namingTexts(directory)
  const allowed = () => sources ??= evidenceSources(directory)
  const issues = [...read.issues]
  const values = new Map<number, ActHow>()
  const edges: number[][] = acts.map(() => [])
  for (const [index, how] of read.rows) {
    const where = `act ${index} ("${excerpt(acts[index].act, 80)}")`
    const instrument: Checked<string | null> = how.instrument === undefined || how.instrument === null ? { issues: [], value: null }
      : typeof how.instrument === "string" && how.instrument.trim() ? { issues: [], value: how.instrument.trim() }
      : { issues: [`${where}: "how.instrument" must name the instrument that effects the act and the part of it that takes the act, or be null when the record does not settle it.`] }
    const executed = partyList(where, "how.executed_by", how.executed_by, naming, "how.reason")
    const requires = requirementList(directory, where, how.requires, index, acts.length, allowed)
    const basis = basisList(directory, where, "how.basis", how.basis, allowed)
    const conflict = conflictValue(directory, where, how.conflict, allowed)
    const reason = reasonValue(where, "how.reason", how.reason, [
      ...(instrument.value === null ? ['"how.instrument" is null'] : []), ...(executed.value?.length === 0 ? ['"how.executed_by" is []'] : []),
    ])
    issues.push(...instrument.issues, ...executed.issues, ...requires.issues, ...basis.issues, ...conflict.issues, ...reason.issues)
    if ((instrument.value || executed.value?.length) && basis.value?.length === 0)
      issues.push(`${where}: "how" names ${instrument.value ? "an instrument" : "signatories"}, but "how.basis" is empty. Give at least one basis entry with the exact words that support each value given, or give null or [] with a reason where no evidence does.`)
    for (const need of requires.value ?? []) if (need.act !== null && need.act !== index) edges[index].push(need.act)
    if (instrument.value !== undefined && executed.value && requires.value && basis.value && conflict.value !== undefined && !reason.issues.length)
      values.set(index, { instrument: instrument.value, executed_by: executed.value, requires: requires.value, basis: basis.value, reason: reason.value ?? null, conflict: conflict.value })
  }
  const cycle = requiresCycle(edges)
  if (cycle) issues.push(`The "how.requires" entries form a circle: ${cycle.slice(0, -1).map((act, at) => `act ${act} requires act ${cycle[at + 1]}`).join(", ")}. ` +
    "Acts cannot each have to be taken before the other. Keep only the requirements the record supports, so that no act requires itself through other acts.")
  return issues.length ? { issues } : { issues, value: acts.map((_, index) => values.get(index)!) }
}

/** Each act's signer check, read from the record and never stored: the
 * parties whose authority the act needs (WHO "authorised_by") against the
 * signatories of its instrument (HOW "executed_by"), compared by nameKey so a
 * legal form's spelling does not count. */
export function signerFlags(record: ActRecord): SignerFlag[] {
  return record.acts.map((act, index) => {
    const authorisers = act.who?.authorised_by ?? []
    if (!authorisers.length || !act.how) return { act: index, status: "unchecked", missing: [] }
    const signers = new Set(act.how.executed_by.map(nameKey))
    const missing = authorisers.filter(party => !signers.has(nameKey(party)))
    return { act: index, status: missing.length ? "flag" : "ok", missing }
  })
}

const strings = (value: unknown) => Array.isArray(value) && value.every(row => typeof row === "string")
const evidence = (value: unknown) => Array.isArray(value) && value.every(row => record(row) && typeof row.source === "string" && typeof row.quote === "string")
const nullableText = (value: unknown) => value === null || typeof value === "string"
const sourced = (row: unknown, field: string): row is Record<string, unknown> => record(row) && typeof row[field] === "string" && typeof row.source === "string" && typeof row.quote === "string"
const isWho = (value: unknown) => value === null || record(value) && strings(value.authorised_by) && strings(value.performed_by) && evidence(value.basis) && nullableText(value.reason)
const isConflict = (value: unknown) => value === null || record(value) && evidence([value.instruction]) && evidence([value.authority]) && typeof value.difference === "string"
const isVerify = (value: unknown) => value === null || record(value) &&
  Array.isArray(value.excluded) && value.excluded.every(row => sourced(row, "act") && NON_ACT_KINDS.includes(row.kind as NonActKind) && typeof row.reason === "string") &&
  Array.isArray(value.merged) && value.merged.every(row => sourced(row, "act") && Number.isInteger(row.into))
const isHow = (value: unknown) => value === null || record(value) && nullableText(value.instrument) && strings(value.executed_by) && evidence(value.basis) && nullableText(value.reason) &&
  Array.isArray(value.requires) && value.requires.every(need => record(need) && (need.act === null || Number.isInteger(need.act)) && typeof need.need === "string" && evidence(need.basis)) &&
  isConflict(value.conflict)

/** The saved act record; undefined when there is none or it is not in this schema's shape. */
export function readActRecord(directory: string): ActRecord | undefined {
  const value = readLeanJson(directory, "acts.json")
  if (!record(value) || value.schema !== LEAN_ACTS_SCHEMA || !["listed", "none", "unchecked"].includes(String(value.status))) return
  if (!Array.isArray(value.acts) || !value.acts.every(act => sourced(act, "act") && isWho(act.who) && isHow(act.how)) || !isVerify(value.verify)) return
  if (!(value.deliverables === null || Array.isArray(value.deliverables) && value.deliverables.every(row => sourced(row, "deliverable")))) return
  if (!Array.isArray(value.issues) || !value.issues.every(row => record(row) && typeof row.job === "string" && strings(row.issues))) return
  return value as ActRecord
}

const ACT_RECORD_USE = "The act record is a reading of the instructions to weigh, not an instruction or evidence: use it when deciding which documents carry which acts and who executes them; " +
  "where the record and the sources differ, the sources govern and the difference becomes an open item."

/** The act record as the spine is shown it, built only from acts.json in a
 * fixed order: each listed act with its WHO, HOW, any conflict and signer flag, then the
 * deliverables. Evidence for WHO and HOW is left to retrieval. Empty when the
 * record is missing or unchecked, so the spine plans from the sources alone. */
export function actRecordBlock(directory: string) {
  const saved = readActRecord(directory)
  if (!saved || saved.status === "unchecked") return ""
  const names = (parties: string[]) => parties.length ? parties.map(party => JSON.stringify(party)).join(", ") : "not settled"
  const unsettled = (reason: string | null) => reason ? ` Unsettled: ${reason}` : ""
  const flags = signerFlags(saved)
  const acts = saved.status === "none" ? ["The act record found no act that the instructions direct."] : saved.acts.flatMap((act, index) => {
    const { who, how } = act
    const requires = how?.requires.map(need => need.act === null ? JSON.stringify(need.need) : `act ${need.act + 1} (${JSON.stringify(need.need)})`) ?? []
    const flag = flags[index]
    return [
      `${index + 1}. ${JSON.stringify(act.act)}`,
      `   Directed by ${act.source}: ${JSON.stringify(act.quote)}`,
      who ? `   WHO: authorised by ${names(who.authorised_by)}; performed by ${names(who.performed_by)}.${unsettled(who.reason)}` : "   WHO: not recorded.",
      how ? `   HOW: instrument ${how.instrument === null ? "not settled" : JSON.stringify(how.instrument)}; executed by ${names(how.executed_by)}${requires.length ? `; requires ${requires.join(", ")}` : ""}.${unsettled(how.reason)}` : "   HOW: not recorded.",
      ...(how?.conflict ? [`   Conflict, an open item for the lawyer and not resolved here: the instructions say ${JSON.stringify(how.conflict.instruction.quote)} (${how.conflict.instruction.source}); ` +
        `${how.conflict.authority.source} says ${JSON.stringify(how.conflict.authority.quote)}. ${how.conflict.difference}`] : []),
      ...(flag.status === "flag" ? [`   Signer flag: ${names(flag.missing)} must authorise the act but ${flag.missing.length > 1 ? "are" : "is"} not among its instrument's recorded signatories.`] : []),
    ]
  })
  const deliverables = saved.deliverables === null ? ["Deliverables: not recorded."]
    : saved.deliverables.length ? ["Deliverables:", ...saved.deliverables.map(row => `- ${JSON.stringify(row.deliverable)}, asked for by ${row.source}: ${JSON.stringify(row.quote)}`)]
    : ["Deliverables: none listed."]
  return ["Act record:", ACT_RECORD_USE, ...acts, ...deliverables].join("\n")
}

/** Runs the act-record jobs and writes acts.json. The WHAT samples run in
 * parallel; DELIVERABLES then runs alongside VERIFY, WHO and HOW in turn: VERIFY
 * over the union, WHO over the acts it keeps (the whole union when it never
 * passes), and HOW, shown each act's WHO (null when WHO never passed). With no
 * candidate, VERIFY is not asked; with no kept act, WHO and HOW are not asked.
 * Each job is one checked exchange in a session of its own: up to
 * LEAN_MAX_RETRIES retries hand its check's issues back to the same session.
 * The step always settles as done: a job that never passes is recorded, not
 * failed. One act_record event records the counts, for diagnosis only. Errors
 * from the session calls, including a pause, propagate and leave the record
 * unwritten. */
export async function runActRecordStage(directory: string, state: LeanState, deps: LeanDeps, progress: (patch: Partial<LeanStageState>) => void): Promise<StageRun> {
  const pack = actRecordPack(directory)
  const sessionIDs: string[] = []
  const calls = { count: 0, attempts: 0 }
  const job = async <T>(label: string, tail: string[], check: (reply: string) => Checked<T>): Promise<Checked<T>> => {
    const id = await deps.sessions.createSession({
      directory, parentID: state.owner_session_id, agent: ACT_RECORD_AGENT, title: `Lean drafting: act record: ${label}`, model: state.model, permissions: LEAN_CHILD_PERMISSIONS,
    })
    sessionIDs.push(id)
    progress({ ...(sessionIDs.length === 1 ? { session_id: id } : {}), session_ids: [...sessionIDs] })
    const step = `act record ${label}`
    const ask = async (text: string, attempt: number): Promise<Checked<T>> => {
      appendLeanEvent(directory, { stage: "acts", event: attempt ? "retry" : "model_call", step, attempt, session_id: id })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID: id, agent: ACT_RECORD_AGENT, text })
      const answered = performance.now()
      const checked = check(reply)
      const timing = { wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) }
      calls.count += 1
      calls.attempts = Math.max(calls.attempts, attempt + 1)
      progress({ attempts: calls.attempts, calls: calls.count, feedback: checked.issues })
      appendLeanEvent(directory, { stage: "acts", event: "checker", step, attempt, issues: checked.issues, ...timing })
      if (!checked.issues.length || attempt >= LEAN_MAX_RETRIES) return checked
      return ask(retryPrompt("acts", checked.issues), attempt + 1)
    }
    return ask([pack, ...tail].join("\n\n"), 0)
  }
  const contract = (...lines: string[]) => tagged("output_contract", lines.map(line => `- ${line}`).join("\n"))
  const shown = (acts: object[]) => tagged("acts", JSON.stringify(acts, null, 2))

  const samples = await Promise.all(Array.from({ length: ACT_RECORD_SAMPLES }, (_, sample) =>
    job(`what ${sample + 1}`, [tagged("job", WHAT_JOB), contract(`${WHAT_SHAPE} Reply {"acts":[]} when the instructions direct no legal act.`)], reply => parseActList(reply, directory))))
  const passed = samples.flatMap(sample => sample.value ? [sample.value] : [])
  const union = mergeCoverageActs(...passed)
  const view = (acts: ListedAct[]) => acts.map((act, index) => ({ index, ...act }))

  const [deliverables, described] = await Promise.all([
    job("deliverables", [tagged("job", DELIVERABLES_JOB), contract(`${DELIVERABLES_SHAPE} Reply {"deliverables":[]} when the instructions ask for none.`)], reply => parseDeliverables(reply, directory)),
    (async () => {
      if (!union.length) return { listed: union }
      const verify = await job("verify", [tagged("job", VERIFY_JOB), shown(view(union)), contract(`${VERIFY_SHAPE} Give every candidate exactly once, by its index.`)],
        reply => parseActVerify(reply, directory, union))
      const verified = verify.value ? applyActVerify(union, verify.value) : undefined
      const listed = verified?.acts ?? union
      if (!listed.length) return { verify, verified, listed }
      const who = await job("who", [tagged("job", WHO_JOB), shown(view(listed)), contract(`${WHO_SHAPE} ${EVERY_ACT}`)], reply => parseActWho(reply, directory, listed))
      const how = await job("how", [tagged("job", HOW_JOB), shown(view(listed).map(act => ({ ...act, who: who.value?.[act.index] ?? null }))), contract(`${HOW_SHAPE} ${EVERY_ACT}`)],
        reply => parseActHow(reply, directory, listed))
      return { verify, verified, listed, who, how }
    })(),
  ])
  const { verify, verified, listed, who, how } = described

  const issues: ActRecord["issues"] = []
  if (!passed.length) issues.push({ job: "what", issues: samples[0].issues })
  if (verify && !verify.value) issues.push({ job: "verify", issues: verify.issues })
  if (!deliverables.value) issues.push({ job: "deliverables", issues: deliverables.issues })
  if (who && !who.value) issues.push({ job: "who", issues: who.issues })
  if (how && !how.value) issues.push({ job: "how", issues: how.issues })
  const result: ActRecord = {
    schema: LEAN_ACTS_SCHEMA,
    status: !passed.length ? "unchecked" : listed.length ? "listed" : "none",
    acts: listed.map((act, index) => ({ ...act, who: who?.value?.[index] ?? null, how: how?.value?.[index] ?? null })),
    verify: verified?.verify ?? null,
    deliverables: deliverables.value ?? null,
    issues,
  }
  writeLeanText(directory, "acts.json", `${JSON.stringify(result, null, 2)}\n`)
  const checkedJob = (row: Checked<unknown> | undefined) => !row ? "not_asked" : row.value ? "passed" : "failed"
  appendLeanEvent(directory, {
    stage: "acts", event: "act_record", outcome: result.status,
    counts: {
      acts: result.acts.length, deliverables: result.deliverables?.length ?? 0, what_samples: samples.length, what_passed: passed.length, candidates: union.length,
      excluded: result.verify?.excluded.length ?? 0, merged: result.verify?.merged.length ?? 0, conflicts: result.acts.filter(act => act.how?.conflict).length,
    },
    jobs: { what: passed.length ? "passed" : "failed", verify: checkedJob(verify), deliverables: checkedJob(deliverables), who: checkedJob(who), how: checkedJob(how) },
  })
  return { status: "done", outputs: ["acts.json"], openItems: [] }
}
