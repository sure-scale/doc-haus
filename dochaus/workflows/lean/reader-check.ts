import { normalizeLegend, normalizePreferences } from "../../lib/drafting-preferences"
import { furnitureLegend } from "../../lib/furniture-legend"
import { INSTRUCTION_COMMENT_PATTERN, OPEN_ITEM_PATTERN, PLACEHOLDER_PATTERN } from "../../lib/marker-patterns"
import type { DocumentForm } from "../../lib/page-furniture"
import { readFirmPreferences } from "../../lib/preferences"
import { normalize } from "../../lib/quote-verify"
import { TOOL_SYNTAX_ISSUE, finalBodyParagraphs, parseJsonObject, printedToolCall, readParagraphs, record, type ParagraphRow } from "./checks"
import { parallel } from "./checklist-flow"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { FIRST_DOCUMENT, planActingFor, plannedDocument, type DocID, type PlannedDocument } from "./plan"
import { retryPrompt, tagged } from "./prompts"
import { excerpt, matterSourceIDs, openItem } from "./quotes"
import { LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, stepSessionTitle, type LeanDeps, type PermissionRule } from "./stages"
import { readLeanJson, readOpenItems, shortHash, writeLeanText, type LeanState, type OpenItem } from "./state"
import { rowsBody } from "./template-coverage"
import { templateSource } from "./template-contract"
import { sourceText } from "./template-structure"

// Reader check. A document's text speaks to its reader: the parties who sign
// an agreement, the person a letter is addressed to. A sentence the drafter
// writes for the lawyer instead (what to verify before sending, how to serve
// it, how to handle the evidence, that this is a draft, an assumption the
// drafting rests on) would go out in the document the reader gets. Once a
// document's paragraphs exist, one focused job per drafted part, each in a
// fresh session of its own and side by side, reads that part alone and names
// each such sentence (readerCheckPrompt, parseReaderHits). Code verifies each
// hit against the part's own paragraphs: its paragraph must be in the part and
// its sentence in that paragraph word for word, and a hit that sits inside an
// open-item marker, a placeholder or quoted words, or that the firm's legend,
// a legend the request asks for or the lawyer's own original supplies, is not
// a hit. A settled part is saved under the content key of its prompt
// (reader-check.json), so a resume does not ask it again. The verified hits
// are open items of kind "reader": in the draft, required repairs for the
// change cycle (repairs.ts), which rewrites the sentence for the reader and
// moves the point for the lawyer to a visible marker or the drafting notes; at
// apply, every current part uses the current prompt's content key, so unchanged
// parts reuse their result and an old audience contract cannot carry stale
// hits forward. A hit left in the final body is disclosed by the
// drafting notes or holds delivery (gate.ts). The check never fails a draft: a
// reply that never passes keeps its verified hits on the last attempt.

export const LEAN_READER_CHECK_SCHEMA = "dochaus.lean.reader-check/1"
/** Recheck current body parts when the reader's audience contract changes. */
export const READER_CHECK_CONTRACT = "dochaus.lean.reader-check/audiences-1"

/** The job reads only the part it is shown, so neither it nor its session may call a tool. */
export const READER_CHECK_PERMISSIONS: PermissionRule[] = [{ action: "*", resource: "*", effect: "deny" }]

export type ReaderHit = { pid: string; sentence: string; addressed_to: string; keep: string }
export type ReaderPart = { part: number; rows: ParagraphRow[] }
type Checked<T> = { issues: string[]; value?: T }
/** One checked exchange with a reader job: prompt, check, and the step's own retries. */
export type ReaderAsk = (label: string, text: string, check: (reply: string, final: boolean) => Checked<ReaderHit[]>) => Promise<Checked<ReaderHit[]>>

/** Whom a document of each form speaks to. */
const READER: Record<DocumentForm, string> = {
  agreement: "the parties who sign it",
  letter: "the person it is addressed to",
  "advice letter": "the client it is addressed to",
  advice: "the person it advises, as its purpose says",
  statement: "those it is given to",
  other: "the person it is addressed to, or the parties who sign it",
}

const SHAPE = 'Reply with the JSON object only: {"hits":[{"pid":"<paragraph id>","sentence":"<the operative sentence or contiguous passage, word for word>","addressed_to":"<whom it speaks to>","keep":"<what it tells the reader, to keep>"}]}; {"hits":[]} when there are no wrong-audience operative passages.'

const QUESTION = [
  "Read the two audiences separately: operative prose speaks to the document's reader; visible markers carry instructions or missing inputs for the lawyer.",
  "Is each operative sentence addressed to the document's reader, and still true and sensible on the day the document is sent or signed?",
  "A sentence written for the lawyer instead (what to verify, how to serve or send the document, how to handle evidence, that this is a draft, a working assumption) is a hit.",
  "Read each lawyer instruction in its internal role. An instruction to confirm a missing input belongs to that audience and is not itself a reader hit; recognising its audience does not resolve its factual gap or approve release.",
  "A marker does not shelter adjacent prose: check every operative passage, including prose sharing a sentence or paragraph with a marker. Do not flag recipient-facing prose merely because a neighbouring marker asks the lawyer to act before sending.",
  "The part is the canonical text. Where a paragraph has markers, operative_prose and lawyer_instructions separate its exact contiguous passages by audience. Do not join passages across a marker or include a marker in a hit.",
  "Quote only the actual wrong-audience operative sentence or contiguous passage, word for word from its paragraph; say whom it speaks to and what it tells the reader that a rewrite keeps.",
].join("\n")

/** One part's reader job: the document, the lines code prints on its pages, the part's paragraphs, the question and the output contract. */
export function readerCheckPrompt(plan: PlannedDocument, actingFor: string, rows: ParagraphRow[], furniture: string[]) {
  const form = plan.form ?? "other"
  const split = rows.map(row => ({ pid: row.pid, ...readerAudiences(row.text) })).filter(row => row.instructions.length)
  return [
    tagged("document", JSON.stringify({ form, title: plan.title, purpose: plan.purpose, acting_for: actingFor || "not stated", reader: READER[form] }, null, 2)),
    ...(furniture.length ? [tagged("furniture", `Lines code prints on every page from the firm's settings:\n${furniture.join("\n")}`)] : []),
    tagged("part", rows.map(row => `[${row.pid}] ${row.text}`).join("\n\n")),
    ...(split.length ? [
      tagged("operative_prose", JSON.stringify(split.map(row => ({ pid: row.pid, passages: row.operative })), null, 2)),
      tagged("lawyer_instructions", JSON.stringify(split.map(row => ({ pid: row.pid, passages: row.instructions })), null, 2)),
    ] : []),
    tagged("question", QUESTION),
    tagged("output_contract", SHAPE),
  ].join("\n\n")
}

/** Reads one part's reply against the part's own paragraphs. Each hit's
 * paragraph must be in the part and its sentence in that paragraph word for
 * word; a hit inside a marker, placeholder or quoted words, or taken from
 * `kept` (the furniture, the lawyer's original), is left out without a word.
 * On the last attempt the verified hits stand whatever else the reply got
 * wrong. */
export function parseReaderHits(reply: string, rows: ParagraphRow[], kept: string[] = [], { final = false } = {}): Checked<ReaderHit[]> {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`${SHAPE} ${parsed.error}`] }
  const hits = parsed.value.hits
  if (!Array.isArray(hits)) return { issues: [SHAPE] }
  const texts = new Map(rows.map(row => [row.pid, row.text]))
  const issues: string[] = []
  const verified: ReaderHit[] = []
  const seen = new Set<string>()
  hits.forEach((hit, index) => {
    const at = `hits[${index}]`
    if (!record(hit) || typeof hit.pid !== "string" || typeof hit.sentence !== "string" || typeof hit.addressed_to !== "string" || typeof hit.keep !== "string") {
      issues.push(`${at}: ${SHAPE}`)
      return
    }
    const text = texts.get(hit.pid)
    if (text === undefined) {
      issues.push(`${at}: pid "${hit.pid}" is not a paragraph of this part.`)
      return
    }
    const sentence = normalize(hit.sentence)
    if (!sentence || !normalize(text).includes(sentence)) {
      issues.push(`${at} ([${hit.pid}]): the sentence is not in that paragraph word for word. Quote it exactly as the paragraph reads.`)
      return
    }
    const key = `${hit.pid}\0${sentence.toLowerCase()}`
    if (seen.has(key) || silent(text, sentence) || kept.some(block => flat(block).includes(flat(sentence)))) return
    const normalized = normalize(text)
    const markers = markerSpans(normalized)
    if (!occurrences(normalized, sentence).some(([from, to]) => !markers.some(([start, end]) => from < end && start < to))) {
      issues.push(`${at} ([${hit.pid}]): the sentence crosses a lawyer-instruction marker. Quote only the wrong-audience operative words, as one contiguous passage of the original paragraph; do not join words across a marker.`)
      return
    }
    seen.add(key)
    verified.push({ pid: hit.pid, sentence: hit.sentence.trim(), addressed_to: hit.addressed_to.trim(), keep: hit.keep.trim() })
  })
  return issues.length && !final ? { issues } : { issues, value: verified }
}

type Span = readonly [from: number, to: number]
const MARKERS = [OPEN_ITEM_PATTERN, PLACEHOLDER_PATTERN, INSTRUCTION_COMMENT_PATTERN]

/** Merge overlapping marker matches before separating the two audiences. */
function markerSpans(text: string): Span[] {
  const spans = MARKERS.flatMap(pattern => [...text.matchAll(new RegExp(pattern.source, pattern.flags))]
    .map(match => [match.index, match.index + match[0].length] as Span)).toSorted(([a], [b]) => a - b)
  return spans.reduce<Span[]>((merged, span) => {
    const previous = merged.at(-1)
    if (previous && span[0] <= previous[1]) merged[merged.length - 1] = [previous[0], Math.max(previous[1], span[1])]
    else merged.push(span)
    return merged
  }, [])
}

/** Exact contiguous passages keep a direct mapping to the original paragraph. */
function readerAudiences(text: string) {
  const operative: string[] = [], instructions: string[] = []
  let at = 0
  for (const [from, to] of markerSpans(text)) {
    const passage = text.slice(at, from).trim()
    if (passage) operative.push(passage)
    instructions.push(text.slice(from, to))
    at = to
  }
  const tail = text.slice(at).trim()
  if (tail) operative.push(tail)
  return { operative, instructions }
}

function occurrences(text: string, sentence: string): Span[] {
  return [...text.matchAll(new RegExp(sentence.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"))]
    .map(match => [match.index, match.index + match[0].length])
}

/** Whether every occurrence of `sentence` in the paragraph sits inside a silent span. */
function silent(text: string, sentence: string) {
  const normalized = normalize(text)
  const quotes = [...normalized.matchAll(/"[^"\n]*"/g)].map(match => [match.index, match.index + match[0].length] as Span)
  const spans = [...markerSpans(normalized), ...quotes]
  return occurrences(normalized, sentence).every(([from, to]) => spans.some(([start, end]) => start <= from && to <= end))
}

/** Lowercased words between single spaces, so a sentence matches whole words. */
function flat(text: string) {
  return ` ${text.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim()} `
}

/** Open items for the verified hits: in the draft, required repairs for the change cycle; in the final body, items the drafting notes disclose. */
export function readerItems(stage: "draft" | "apply", document: DocID, hits: ReaderHit[]): OpenItem[] {
  return hits.map(hit => {
    const to = hit.addressed_to ? ` It speaks to ${hit.addressed_to}.` : ""
    const keep = hit.keep ? ` keeping what it tells the reader (${hit.keep})` : ""
    return openItem(stage, "reader",
      `A sentence written for the lawyer, not the document's reader: "${excerpt(hit.sentence)}".${to} Rewrite it for the reader${keep}; put the point for the lawyer in a visible [open item: ...] marker or leave it to the drafting notes.`,
      { document, pid: hit.pid, statement: hit.sentence })
  })
}

/** A part job's content key: the prompt it reads, so a changed part or question asks again. */
export const readerCheckKey = (prompt: string) => shortHash(prompt, 40)

/** The parts settled so far, by content key. */
export function readReaderChecks(directory: string, document: DocID): Record<string, ReaderHit[]> {
  const value = readLeanJson(directory, `${document}/reader-check.json`)
  if (!record(value) || value.schema !== LEAN_READER_CHECK_SCHEMA || !record(value.parts)) return {}
  return Object.fromEntries(Object.entries(value.parts).filter(([, hits]) => Array.isArray(hits) && hits.every(isHit))) as Record<string, ReaderHit[]>
}

function isHit(value: unknown): value is ReaderHit {
  return record(value) && typeof value.pid === "string" && typeof value.sentence === "string" && typeof value.addressed_to === "string" && typeof value.keep === "string"
}

/** Saves one settled part as soon as its job passes, so a resume skips it. */
export function saveReaderCheck(directory: string, document: DocID, key: string, hits: ReaderHit[]) {
  writeLeanText(directory, `${document}/reader-check.json`, `${JSON.stringify({ schema: LEAN_READER_CHECK_SCHEMA, parts: { ...readReaderChecks(directory, document), [key]: hits } }, null, 2)}\n`)
}

/** A document's body by planned part, as rowsBody attributes its rows; a row of no part (the title line) is in none. */
export function readerParts(plan: PlannedDocument, draft: ParagraphRow[], final?: ParagraphRow[]): ReaderPart[] {
  const rows = final ?? draft
  const { partOf } = rowsBody(plan, draft, final)
  const parts = new Map<number, ParagraphRow[]>()
  rows.forEach((row, at) => {
    const part = partOf[at]
    if (part !== undefined) parts.set(part, [...(parts.get(part) ?? []), row])
  })
  return [...parts].toSorted(([a], [b]) => a - b).map(([part, rows]) => ({ part, rows }))
}

/** Each part's job, side by side; a part settled before is not asked again.
 * A part whose reply never passed its check adds no hit and is not saved, so
 * a resume asks it again. */
export async function checkReaderParts(directory: string, document: DocID, parts: ReaderPart[], ask: ReaderAsk): Promise<ReaderHit[]> {
  const plan = plannedDocument(directory, document)
  if (!plan || !parts.length) return []
  const acting = planActingFor(directory).party
  const furniture = furnitureLines(plan)
  const kept = [...furniture, revisedOriginal(directory, plan)].filter(Boolean)
  const hits: ReaderHit[][] = parts.map(() => [])
  await parallel(parts, async ({ part, rows }, at) => {
    const prompt = readerCheckPrompt(plan, acting, rows, furniture)
    const key = readerCheckKey(prompt)
    const saved = readReaderChecks(directory, document)[key]
    if (saved) {
      hits[at] = saved
      return
    }
    const checked = await ask(`reader: part ${part + 1}/${plan.parts.length}`, prompt, (reply, final) => parseReaderHits(reply, rows, kept, { final }))
    if (!checked.value) return
    hits[at] = checked.value
    saveReaderCheck(directory, document, key, checked.value)
  })
  return hits.flat()
}

/** A checked retirement must also remove that exact hit from the ordinary
 * current-part cache, or a later apply can restore it. Preserve other hits
 * and historical keys. A missing cache is no whole-part audience verdict: its
 * uninspected paragraphs still need the ordinary reader job. */
export function retireCurrentReaderHits(directory: string, document: DocID, retired: { pid: string; statement: string }[]) {
  const plan = plannedDocument(directory, document)
  if (!plan || !retired.length) return
  const parts = readerParts(plan, readParagraphs(directory, document), finalBodyParagraphs(directory, document))
  const acting = planActingFor(directory).party
  const furniture = furnitureLines(plan)
  const saved = readReaderChecks(directory, document)
  const falseHits = new Set(retired.map(row => `${row.pid}\0${normalize(row.statement)}`))
  for (const { rows } of parts) {
    const key = readerCheckKey(readerCheckPrompt(plan, acting, rows, furniture))
    const hits = saved[key]
    if (!hits) continue
    const kept = hits.filter(hit => !falseHits.has(`${hit.pid}\0${normalize(hit.sentence)}`))
    if (kept.length !== hits.length) saveReaderCheck(directory, document, key, kept)
  }
}

/** The lines code prints on a document's pages: the firm's legend for its form and a legend the request asks for. */
function furnitureLines(plan: PlannedDocument) {
  return [furnitureLegend(plan.form ?? "other", normalizePreferences(readFirmPreferences())), normalizeLegend(plan.furniture?.legend)].filter(Boolean)
}

function revisedOriginal(directory: string, plan: PlannedDocument) {
  const source = plan.mode === "revise" ? templateSource(plan) : undefined
  return source && matterSourceIDs(directory).includes(source) ? sourceText(directory, source) : ""
}

/** The apply step checks every current part under the current prompt contract;
 * unchanged parts reuse its content-keyed result. Without a session port, only
 * standing draft hits in parts the changes left alone can be carried. */
export async function applyReaderItems(directory: string, state: LeanState, document: DocID, final: ParagraphRow[], amended: Set<string>, deps?: Pick<LeanDeps, "sessions">): Promise<OpenItem[]> {
  const plan = plannedDocument(directory, document)
  if (!plan) return []
  const parts = readerParts(plan, readParagraphs(directory, document), final)
  if (deps) return readerItems("apply", document, await checkReaderParts(directory, document, parts, readerAsk(directory, state, deps, document)))
  const changed = parts.filter(({ rows }) => rows.some(row => amended.has(row.pid)))
  const rechecked = new Set(changed.flatMap(({ rows }) => rows.map(row => row.pid)))
  const texts = new Map(final.map(row => [row.pid, normalize(row.text)]))
  const carried = readOpenItems(directory)
    .filter(item => item.stage === "draft" && item.kind === "reader" && (item.ref.document ?? FIRST_DOCUMENT) === document
      && item.ref.pid && !rechecked.has(item.ref.pid) && item.ref.statement && texts.get(item.ref.pid)?.includes(normalize(item.ref.statement)))
    .map(item => openItem("apply", "reader", item.detail, item.ref))
  return carried
}

/** The apply step's exchange with one reader job: a fresh session of its own
 * and up to LEAN_MAX_RETRIES retries that hand the check's issues back to it. */
function readerAsk(directory: string, state: LeanState, deps: Pick<LeanDeps, "sessions">, document: DocID): ReaderAsk {
  const agent = LEAN_STAGE_AGENTS.reader
  return async (label, text, check) => {
    const id = await deps.sessions.createSession({
      directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, `${document}.apply`, `apply: ${label}`), model: state.model, permissions: READER_CHECK_PERMISSIONS,
    })
    let prompt = text
    for (let attempt = 0; ; attempt += 1) {
      appendLeanEvent(directory, { stage: "apply", document, event: attempt ? "retry" : "model_call", step: label, attempt, session_id: id })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID: id, agent, text: prompt })
      const answered = performance.now()
      const checked = check(reply, attempt >= LEAN_MAX_RETRIES)
      appendLeanEvent(directory, { stage: "apply", document, event: "checker", step: label, attempt, issues: checked.issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (!checked.issues.length || attempt >= LEAN_MAX_RETRIES) return checked
      prompt = retryPrompt("apply", checked.issues)
    }
  }
}
