import { normalize } from "../../lib/quote-verify"
import { parseJsonObject } from "./checks"
import { parallel } from "./checklist-flow"
import { record } from "./checks"
import { figures, numbers, SPINE_KINDS, unquotedFigures } from "./evidence"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { OutputLimitError } from "./native-stage-result"
import { retryPrompt, tagged } from "./prompts"
import type { SourceQuote } from "./quotes"
import { SPINE_ATOMS } from "./spine-placement"
import { LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, SPINE_CONTRACT, stepSessionTitle, type LeanDeps, type PermissionRule } from "./stages"
import { readLeanJson, shortHash, writeLeanText, type LeanStageState, type LeanState } from "./state"

// Once the spine reply passes its checks, and before the step's files are
// written, one job per entry asks whether the entry states one point or several
// and, if several, splits it into atoms that each state one. Code checks every
// atom: a kind from the spine's set, evidence and limitations taken by index
// from the entry's own (so every quotation stays one the spine check verified),
// a cited quotation stating each figure the entry's quotations state, and no
// figure, quoted passage or capitalised term the entry does not hold. An entry
// that is already atomic, or whose job never passes, stays as it is.
// spine.json stays the accepted reply, its entries as the spine wrote them, for
// the jobs that read the whole spine (outline, parts, reviews, notes, package
// review); spine-atoms.json holds the atoms in the entries' order, for the jobs
// that read one entry at a time (level, placement). Each finished split is
// saved under its entry's content, so a re-run of the step splits only entries
// not split before.

const FILE = "spine-split.json"

export type SpineEntry = { kind: string; detail: string; evidence: SourceQuote[]; limitations: string[] } & Record<string, unknown>
type Saved = { splits: Record<string, SpineEntry[]> }

/** The job needs no tool: it reads only the entry it is shown. */
export const SPINE_SPLIT_PERMISSIONS: PermissionRule[] = [{ action: "*", resource: "*", effect: "deny" }]

/** One spine entry, its evidence and its limitations, each listed by index. */
export function spineSplitPrompt(entry: SpineEntry) {
  const listed = (rows: string[]) => rows.length ? rows.map((row, index) => `[${index}] ${row}`).join("\n") : "None."
  return [
    "This is a lean drafting run. Decide one thing: whether the drafting spine entry below states one point or several and, if several, split it into atoms that each state one point.",
    tagged("entry", `kind: ${entry.kind}\ndetail: ${entry.detail}`),
    tagged("evidence", listed(entry.evidence.map(row => `${row.source_id}: "${row.quote}"`))),
    tagged("limitations", listed(entry.limitations)),
    tagged("output_contract", [
      "- A point is one proposition, with its drafting consequence, that a later job can draft or check on its own. Parts that only make sense together are one point.",
      '- When the entry states one point, reply {"atomic":true}.',
      `- Otherwise reply {"atoms":[{"kind":"${SPINE_KINDS.join("|")}","detail":"...","evidence":[0],"limitations":[]}]} with one atom per point, in the order the entry states them.`,
      "- Each atom's detail uses the entry's own words and adds nothing to them; together the atoms keep everything the entry says.",
      "- evidence and limitations list the indexes above. Each atom cites the evidence that supports it, at least one when the entry has any, and every evidence and limitation goes to at least one atom. When the entry has no evidence, each atom carries at least one limitation.",
      "- Answer from the entry above; no lookup is needed. Reply with the JSON object only.",
    ].join("\n")),
  ].join("\n\n")
}

const SHAPE = 'Reply with the JSON object only: {"atomic":true} or {"atoms":[{"kind":"...","detail":"...","evidence":[0],"limitations":[]}]}.'

/** Reads the reply to spineSplitPrompt: the entry itself when it is atomic or
 * the reply gives one atom, otherwise its atoms, each built from the entry's
 * own evidence and limitations. Every other reply is retried. */
export function parseSpineSplit(reply: string, parent: SpineEntry): { issues: string[]; atoms?: SpineEntry[] } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`${SHAPE} ${parsed.error}`] }
  if (parsed.value.atomic === true && parsed.value.atoms === undefined) return { issues: [], atoms: [parent] }
  const rows = parsed.value.atoms
  if (!Array.isArray(rows) || !rows.length) return { issues: [SHAPE] }
  // One atom says the entry states one point: the entry stands as written, not the atom's rewording.
  if (rows.length === 1) return { issues: [], atoms: [parent] }
  const known = entryText(parent)
  const used = { evidence: new Set<number>(), limitations: new Set<number>() }
  const issues: string[] = []
  const atoms = rows.flatMap((row: unknown, index: number): SpineEntry[] => {
    const at = `atoms[${index}]`
    if (!record(row) || typeof row.kind !== "string" || typeof row.detail !== "string" || !row.detail.trim() || !Array.isArray(row.evidence) || !Array.isArray(row.limitations)) {
      issues.push(`${at} needs kind, detail, evidence [indexes] and limitations [indexes]. ${SHAPE}`)
      return []
    }
    const own: string[] = []
    if (!(SPINE_KINDS as readonly string[]).includes(row.kind)) own.push(`${at} kind must be ${SPINE_KINDS.join(", ")}.`)
    const evidence = indexes(row.evidence, parent.evidence.length, `${at}.evidence`, "evidence", own)
    const limitations = indexes(row.limitations, parent.limitations.length, `${at}.limitations`, "limitation", own)
    if (parent.evidence.length && !evidence.length) own.push(`${at} cites no evidence; a limitation does not replace evidence. Cite the index of the evidence that supports it, 0 to ${parent.evidence.length - 1}.`)
    if (!parent.evidence.length && !limitations.length) own.push(`${at} has no evidence, so it must carry at least one limitation index.`)
    own.push(...added(row.detail, known).map(term => `${at} says ${JSON.stringify(term)}, which the entry does not; use only the entry's own words.`))
    // A figure the entry's quotes state is stated by a quote the atom cites, as the spine check asks of the entry.
    const unstated = unquotedFigures(row.detail, parent.evidence.map(ref => ref.quote))
    own.push(...unquotedFigures(row.detail, evidence.map(item => parent.evidence[item]!.quote)).filter(figure => !unstated.includes(figure)).map(figure => `${at} states ${JSON.stringify(figure)}, which none of its evidence states; cite the index of the evidence that states it.`))
    evidence.forEach(item => used.evidence.add(item))
    limitations.forEach(item => used.limitations.add(item))
    issues.push(...own)
    return own.length ? [] : [{ ...parent, kind: row.kind, detail: row.detail.trim(), evidence: evidence.map(item => parent.evidence[item]!), limitations: limitations.map(item => parent.limitations[item]!) }]
  })
  parent.evidence.forEach((_, item) => { if (!used.evidence.has(item)) issues.push(`Evidence [${item}] goes to no atom; cite it in the atom it supports.`) })
  parent.limitations.forEach((_, item) => { if (!used.limitations.has(item)) issues.push(`Limitation [${item}] goes to no atom; give it to the atom it qualifies.`) })
  if (atoms.length === rows.length) {
    // Kept in whatever form an atom writes it.
    const kept = new Set(atoms.flatMap(atom => numbers(atom.detail)))
    for (const figure of new Set(figures(parent.detail))) if (!kept.has(figure)) issues.push(`The entry's figure ${JSON.stringify(figure)} is in no atom; keep it in the atom that states it.`)
    const details = atoms.map(atom => normalize(atom.detail).toLowerCase())
    details.forEach((detail, index) => { if (details.indexOf(detail) !== index) issues.push(`atoms[${index}] repeats atoms[${details.indexOf(detail)}]; each atom states a different point.`) })
  }
  return issues.length ? { issues } : { issues: [], atoms }
}

// Distinct in-range integer indexes in ascending order; an issue for any other value.
function indexes(values: unknown[], length: number, at: string, name: string, issues: string[]) {
  const bad = values.filter(value => !Number.isInteger(value) || (value as number) < 0 || (value as number) >= length)
  if (bad.length) issues.push(length ? `${at} lists ${JSON.stringify(bad)}: use ${name} indexes 0 to ${length - 1}.` : `${at} lists ${JSON.stringify(bad)}, but the entry has no ${name}; leave it empty.`)
  return [...new Set(values.filter(value => !bad.includes(value)) as number[])].toSorted((a, b) => a - b)
}

// The entry's words: its detail, its quotations and its limitations.
function entryText(entry: SpineEntry) {
  return normalize([entry.detail, ...entry.evidence.map(row => row.quote), ...entry.limitations].join("\n"))
}

/** What an atom's detail says that the entry's words do not, where code can
 * tell: a figure the entry states in no form, a quoted passage, or a
 * capitalised word not opening a sentence (a name or defined term) that the
 * entry never uses. */
function added(detail: string, known: string) {
  const text = normalize(detail)
  const lower = known.toLowerCase()
  const knownNumbers = new Set(numbers(known))
  const opensSentence = (index: number) => /(^|[.!?:;("'[]\s*)$/.test(text.slice(0, index))
  const terms = [
    ...figures(text).filter(figure => !knownNumbers.has(figure)),
    ...[...text.matchAll(/"([^"]+)"/g)].map(match => match[1]!.trim()).filter(passage => passage && !lower.includes(passage.toLowerCase())),
    ...[...text.matchAll(/(?<![\p{L}\p{N}])\p{Lu}[\p{L}\p{N}-]*/gu)].filter(match => !opensSentence(match.index) && !lower.includes(match[0].toLowerCase())).map(match => match[0]),
  ]
  return [...new Set(terms)]
}

/** Splits each entry of the checked spine reply into atoms, in parallel within
 * the concurrency limit, and writes the reply as spine.json and the atoms, in
 * the entries' order, as spine-atoms.json. A saved split of the same entry is
 * reused. A job that never passes its check, or whose reply stops at the output
 * limit, keeps its entry. Any other error, including a pause, propagates once
 * the jobs in flight settle; the splits finished by then stay saved and neither
 * file is written. */
export async function splitSpineEntries(directory: string, state: LeanState, deps: LeanDeps, spine: Record<string, unknown>, progress: (patch: Partial<LeanStageState>) => void): Promise<string[]> {
  const entries = spine.entries as SpineEntry[]
  const keys = entries.map(entry => shortHash(JSON.stringify([SPINE_CONTRACT, entry]), 40))
  const splits = savedSplits(directory)
  const save = () => writeLeanText(directory, FILE, `${JSON.stringify({ splits } satisfies Saved, null, 2)}\n`)
  const sessionIDs: string[] = []
  const counts = { entries: entries.length, reused: 0, atomic: 0, split: 0, kept: 0 }
  let failure: { error: unknown } | undefined
  await parallel(entries.map((entry, index) => ({ entry, index })), async ({ entry, index }) => {
    if (failure) return
    if (splits[keys[index]!]) {
      counts.reused += 1
      return
    }
    try {
      const atoms = await splitEntry(directory, state, deps, entry, index + 1, id => {
        sessionIDs.push(id)
        progress({ session_ids: [...sessionIDs] })
      })
      counts[!atoms ? "kept" : atoms.length > 1 ? "split" : "atomic"] += 1
      splits[keys[index]!] = atoms ?? [entry]
      save()
    } catch (error) {
      failure ??= { error }
    }
  })
  if (failure) throw failure.error
  for (const key of Object.keys(splits)) if (!keys.includes(key)) delete splits[key]
  save()
  const atoms = keys.flatMap(key => splits[key]!)
  appendLeanEvent(directory, { stage: "spine", event: "spine_split", step: "split entries", counts: { ...counts, atoms: atoms.length } })
  writeLeanText(directory, SPINE_ATOMS, `${JSON.stringify({ entries: atoms }, null, 2)}\n`)
  writeLeanText(directory, "spine.json", `${JSON.stringify(spine, null, 2)}\n`)
  return ["spine.json", SPINE_ATOMS]
}

function savedSplits(directory: string): Record<string, SpineEntry[]> {
  const saved = readLeanJson(directory, FILE)
  if (!record(saved) || !record(saved.splits)) return {}
  return Object.fromEntries(Object.entries(saved.splits).filter(([, atoms]) => Array.isArray(atoms) && atoms.length)) as Record<string, SpineEntry[]>
}

/** One entry's job in a session of its own: its atoms (the entry itself when
 * atomic), or undefined when its reply never passes or stops at the output limit. */
async function splitEntry(directory: string, state: LeanState, deps: LeanDeps, entry: SpineEntry, number: number, started: (sessionID: string) => void): Promise<SpineEntry[] | undefined> {
  const agent = LEAN_STAGE_AGENTS.split
  const at = { stage: "spine" as const, step: `split entry ${number}` }
  const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, "spine", `spine split: entry ${number}`), model: state.model, permissions: SPINE_SPLIT_PERMISSIONS })
  started(sessionID)
  let issues: string[] = []
  for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt += 1) {
    appendLeanEvent(directory, { ...at, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
    const asked = performance.now()
    let reply: string
    try {
      reply = await deps.sessions.prompt({ directory, sessionID, agent, text: attempt ? retryPrompt("spine", issues) : spineSplitPrompt(entry) })
    } catch (error) {
      // A reply cut off before any answer keeps the entry, as one that never passes does.
      if (error instanceof OutputLimitError) return undefined
      throw error
    }
    const answered = performance.now()
    const checked = parseSpineSplit(reply, entry)
    issues = checked.issues
    appendLeanEvent(directory, { ...at, event: "checker", attempt, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
    if (checked.atoms) return checked.atoms
  }
  return undefined
}
