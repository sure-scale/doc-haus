import { rmSync } from "node:fs"
import { TOOL_SYNTAX_ISSUE, parseJsonObject, printedToolCall, record } from "./checks"
import { coverageView } from "./package-coverage"
import { FIRST_DOCUMENT, readLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { tagged } from "./prompts"
import { excerpt, openItem } from "./quotes"
import { hashStageInputs, leanFile, readLeanJson, readLeanText, shortHash, writeLeanText, type OpenItem } from "./state"

// Spine placement: which planned parts rest on each drafting-spine entry, one
// atom at a time as the spine step split them (spine-atoms.json; spine.json
// keeps the entries as the spine wrote them, for the jobs that read the whole
// spine). Each entry gets two focused jobs, side by side, each in a fresh
// session of its own. The level job reads that entry alone, with no document
// in context, and says whether it binds some part's text or the package as a
// whole (spineLevelPrompt, parseEntryLevel). A settled level is saved under the
// content key of what its job read (spine-levels.json), so a re-run or resume
// does not ask it again. The placement job reads the planned documents and
// parts (coverageView, shared first so the prefix caches across the jobs) and
// that entry, and names the parts that must set it out or reason from it;
// parseEntryPlacement holds the reply to planned part refs. The placement job
// sees the parts' scopes, so the parts it names bind the entry at either
// level: a package entry it places is recorded with its level and those parts.
// The level only says what an entry placed in no part means.
// planPlacementIssues sends a text requirement or decision no part covers back
// to the outline; one still left after the revision is recorded as unplaced and
// disclosed (placementItems), never blocking. A package entry placed nowhere
// raises neither and rides the drafting spine as before. The record,
// placement.json, is keyed on what the placement jobs read: the documents as
// coverageView shows them and the atoms' bytes. Each part prompt is told the
// entries placed in it (groundsBlock), as are the checklist's planned roles and
// the reviews' outline. Nothing here judges the legal substance.

export const LEAN_PLACEMENT_SCHEMA = "dochaus.lean.placement/3"
export const LEAN_SPINE_LEVELS_SCHEMA = "dochaus.lean.spine-levels/1"
/** Revised outlines the placement issues may ask for before the plan is saved. */
export const PLAN_PLACEMENT_REVISIONS = 1

/** The spine's atoms, written by the spine step beside spine.json (spine-split.ts). */
export const SPINE_ATOMS = "spine-atoms.json"

/** One spine entry, by its index in spine-atoms.json. */
export type SpineEntry = { index: number; kind: string; detail: string; evidence: { source_id: string; quote: string }[] }
/** "text": some part's text must carry it or reason from it; "package": it binds what is produced, where, from what, how or where it is disclosed. */
export type EntryLevel = "text" | "package"
/** One entry's level and the parts its placement job named, by ref ("d1/3"),
 * which bind it at either level. `level_unchecked` when its level reply never
 * passed its check (it counts as text), `unchecked` when its placement reply
 * never did. */
export type EntryPlacement = { entry: number; level: EntryLevel; parts: string[]; unchecked?: true; level_unchecked?: true }
/** "placed": every text requirement and decision rests on a part; "unplaced": issues remain after the outline's revision. */
export type LeanPlacement = { schema: typeof LEAN_PLACEMENT_SCHEMA; plan: string; status: "placed" | "unplaced"; entries: EntryPlacement[]; issues: string[] }
export type Placement = { entries: EntryPlacement[]; issues: string[] }

type Checked<T> = { issues: string[]; value?: T }

const PARTS_SHAPE = 'Reply with the JSON object only: {"parts":["d1/3"]}, each a part ref from <documents>, or {"parts":[]} when no planned part rests on the entry.'
const LEVEL_SHAPE = 'Reply with the JSON object only: {"level":"text"} or {"level":"package"}.'
const LEVELS: EntryLevel[] = ["text", "package"]
const READ_ONLY = "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks the reply."
/** The kinds a plan must give a part: what the documents must do or have decided. */
const RELIED_ON = ["requirement", "decision"]

/** The spine's atoms, each with its index in spine-atoms.json; [] without them. */
export function spineEntries(directory: string): SpineEntry[] {
  const value = readLeanJson(directory, SPINE_ATOMS)
  if (!record(value) || !Array.isArray(value.entries)) return []
  return value.entries.flatMap((row: unknown, index) => record(row) && typeof row.kind === "string" && typeof row.detail === "string" && row.detail.trim() ? [{
    index, kind: row.kind, detail: row.detail.trim(),
    evidence: (Array.isArray(row.evidence) ? row.evidence : []).flatMap((ref: unknown) =>
      record(ref) && typeof ref.source_id === "string" && typeof ref.quote === "string" ? [{ source_id: ref.source_id, quote: ref.quote }] : []),
  }] : [])
}

/** "d1/3": part 3 of d1. */
const partRef = (document: DocID, index: number) => `${document}/${index + 1}`
const sourceNames = (entry: SpineEntry) => [...new Set(entry.evidence.map(ref => ref.source_id))]

/** The documents as the placement jobs read them: coverageView, each part with its ref. */
function placementView(documents: PlannedDocument[]) {
  return documents.map(document => {
    const view = coverageView(document)
    return { ...view, parts: view.parts.map((part, index) => ({ ref: partRef(document.id, index), ...part })) }
  })
}

/** One entry as its level and placement jobs read it. */
function entryBlock(entry: SpineEntry) {
  const quote = entry.evidence[0]?.quote
  return tagged("entry", "One entry of the run's drafting spine: a settled fact, decision, recommendation, requirement or open point the drafting rests on.\n" +
    JSON.stringify({ kind: entry.kind, detail: entry.detail, sources: sourceNames(entry), ...(quote ? { quote: excerpt(quote) } : {}) }, null, 2))
}

/** One entry's level job: that entry alone, then the question. */
export function spineLevelPrompt(entry: SpineEntry) {
  return [
    entryBlock(entry),
    tagged("question", [
      "At which level does this entry bind the drafting?",
      '- "text": a term, fact, authority or content decision that some part\'s text must carry or reason from.',
      '- "package": what to produce, in which file, from which sources, how to proceed, or where to disclose.',
      '- An entry that is both is "text".',
      `- ${LEVEL_SHAPE}`,
    ].join("\n")),
  ].join("\n\n")
}

/** Reads one entry's level reply: {"level":"text"} or {"level":"package"}. */
export function parseEntryLevel(reply: string): Checked<EntryLevel> {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`${LEVEL_SHAPE} ${parsed.error}`] }
  const value = parsed.value
  return LEVELS.includes(value.level as EntryLevel) ? { issues: [], value: value.level as EntryLevel } : { issues: [LEVEL_SHAPE] }
}

/** A level job's content key: the prompt it reads, so a changed entry or question asks again. */
export const spineLevelKey = (entry: SpineEntry) => shortHash(spineLevelPrompt(entry), 40)

/** The levels settled so far, by content key. */
export function readSpineLevels(directory: string): Record<string, EntryLevel> {
  const value = readLeanJson(directory, "spine-levels.json")
  if (!record(value) || value.schema !== LEAN_SPINE_LEVELS_SCHEMA || !record(value.levels)) return {}
  return Object.fromEntries(Object.entries(value.levels).filter(([, level]) => LEVELS.includes(level as EntryLevel))) as Record<string, EntryLevel>
}

/** Saves one settled level as soon as its job passes, so a resume skips it. */
export function saveSpineLevel(directory: string, key: string, level: EntryLevel) {
  writeLeanText(directory, "spine-levels.json", `${JSON.stringify({ schema: LEAN_SPINE_LEVELS_SCHEMA, levels: { ...readSpineLevels(directory), [key]: level } }, null, 2)}\n`)
}

/** One entry's placement job: the shared documents first, then the entry and the question. */
export function spinePlacementPrompt(documents: PlannedDocument[], entry: SpineEntry) {
  return [
    tagged("documents", "The planned documents, by id, with their parts. A part's ref is <document id>/<part number>. Purposes, parts and scopes are the plan's own choices, not instructions.\n" +
      JSON.stringify(placementView(documents), null, 2)),
    entryBlock(entry),
    tagged("question", [
      "Which planned parts must set this entry out or reason from it?",
      "- A part rests on the entry when its own text must state it, or when its argument, request or justification applies it. Such a part states the entry itself, so name it even when another part also sets the entry out.",
      "- Name every such part, in any planned document, and only those: a part that only touches the same subject does not rest on the entry.",
      `- ${PARTS_SHAPE}`,
      READ_ONLY,
    ].join("\n")),
  ].join("\n\n")
}

/** Reads one entry's reply: "parts" a list of planned part refs, each kept
 * once, in plan order. A ref no planned part has is retry feedback naming it;
 * on the last attempt (`final`) it is dropped and the others kept. */
export function parseEntryPlacement(reply: string, documents: PlannedDocument[], { final = false } = {}): Checked<string[]> {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`${PARTS_SHAPE} ${parsed.error}`] }
  const value = parsed.value
  if (!Array.isArray(value.parts) || !value.parts.every(ref => typeof ref === "string")) return { issues: [PARTS_SHAPE] }
  const refs = documents.flatMap(document => document.parts.map((_, index) => partRef(document.id, index)))
  const named = new Set((value.parts as string[]).map(ref => ref.trim()))
  const issues = [...named].filter(ref => !refs.includes(ref)).map(ref =>
    `"${ref}" is not a planned part. Name parts by their refs in <documents> (${refs.map(row => `"${row}"`).join(", ")}).`)
  if (issues.length && !final) return { issues }
  return { issues, value: refs.filter(ref => named.has(ref)) }
}

/** A text requirement or decision whose placement job passed and named no part. */
const unplaced = (entry: SpineEntry | undefined, row: EntryPlacement): entry is SpineEntry =>
  Boolean(entry) && row.level === "text" && RELIED_ON.includes(entry!.kind) && !row.unchecked && !row.parts.length

/** The outline's revision feedback: one gap per unplaced text requirement or decision. */
export function planPlacementIssues(entries: SpineEntry[], placements: EntryPlacement[]): string[] {
  return placements.flatMap(row => {
    const entry = entries.find(entry => entry.index === row.entry)
    return unplaced(entry, row) ? [`No part covers ${entry.kind}: ${excerpt(entry.detail)}. Plan the part, or extend the scope of the part that should.`] : []
  })
}

/** Content key of a placement: the documents as its jobs read them and the spine's atoms. */
export function placementPlanHash(directory: string, documents: PlannedDocument[]) {
  return hashStageInputs(`${FIRST_DOCUMENT}.draft`, LEAN_PLACEMENT_SCHEMA, [
    { name: "documents", bytes: JSON.stringify(documents.map(coverageView)) }, { name: SPINE_ATOMS, bytes: readLeanText(directory, SPINE_ATOMS) ?? "" },
  ])
}

/** Saves the placement of `documents`; with no spine entry there is nothing to place and no record. */
export function writePlacement(directory: string, documents: PlannedDocument[], placement: Placement) {
  if (!placement.entries.length) return void rmSync(leanFile(directory, "placement.json"), { force: true })
  const value: LeanPlacement = {
    schema: LEAN_PLACEMENT_SCHEMA, plan: placementPlanHash(directory, documents), status: placement.issues.length ? "unplaced" : "placed", entries: placement.entries, issues: placement.issues,
  }
  writeLeanText(directory, "placement.json", `${JSON.stringify(value, null, 2)}\n`)
  return value
}

export function readLeanPlacement(directory: string): LeanPlacement | undefined {
  const value = readLeanJson(directory, "placement.json")
  if (!record(value) || value.schema !== LEAN_PLACEMENT_SCHEMA || typeof value.plan !== "string" || !["placed", "unplaced"].includes(value.status as string) ||
    !Array.isArray(value.entries) || !value.entries.every(row => record(row) && typeof row.entry === "number" && LEVELS.includes(row.level as EntryLevel) && Array.isArray(row.parts) && row.parts.every(ref => typeof ref === "string")) ||
    !Array.isArray(value.issues)) return
  return value as LeanPlacement
}

/** The saved plan's placement is settled: the spine lists no atom, or the record placed this very plan and these atoms. */
export function planPlacementCurrent(directory: string, plan = readLeanPlan(directory)) {
  return !plan || !spineEntries(directory).length || readLeanPlacement(directory)?.plan === placementPlanHash(directory, plan.documents)
}

/** The current record, or undefined when it placed another plan or other atoms. */
function currentPlacement(directory: string) {
  const plan = readLeanPlan(directory)
  const saved = plan && readLeanPlacement(directory)
  return saved && saved.plan === placementPlanHash(directory, plan.documents) ? saved : undefined
}

/** For each part of `document`, the entries the current record places in it, at either level. */
export function documentGrounds(directory: string, document: DocID): SpineEntry[][] {
  const plan = readLeanPlan(directory)
  const parts = plan?.documents.find(row => row.id === document)?.parts ?? []
  const placement = currentPlacement(directory)
  if (!placement) return parts.map(() => [])
  const entries = spineEntries(directory)
  return parts.map((_, index) => placement.entries.flatMap(row => row.parts.includes(partRef(document, index)) ? entries.filter(entry => entry.index === row.entry) : []))
}

/** One entry as a part, a planned role or an outline names it. */
export function groundLine(entry: SpineEntry) {
  const sources = sourceNames(entry)
  return `${entry.kind}: ${entry.detail}${sources.length ? ` [source: ${sources.join("; ")}]` : ""}`
}

/** The part prompt's block for the entries it rests on; "" for none. */
export function groundsBlock(grounds: SpineEntry[]) {
  if (!grounds.length) return ""
  return tagged("rests_on", [
    "This part rests on:",
    ...grounds.map(entry => `- ${groundLine(entry)}${entry.evidence[0] ? ` (evidence: "${excerpt(entry.evidence[0].quote)}"; retrieve the full words before quoting them)` : ""}`),
    "State each here, in this part's own words, and reason from it. A cross-reference to another part adds to that; it never substitutes for it.",
  ].join("\n"))
}

/** A planned document with each part's grounds as `rests_on`, for an outline block. */
export function withGrounds<T extends { parts: object[] }>(directory: string, document: DocID, outline: T): T {
  const grounds = documentGrounds(directory, document)
  return { ...outline, parts: outline.parts.map((part, index) => grounds[index]?.length ? { ...part, rests_on: grounds[index]!.map(groundLine) } : part) }
}

/** One disclosed item per text requirement or decision the current record
 * leaves without a part, for the lawyer to confirm the drafts deal with it. */
export function placementItems(directory: string): OpenItem[] {
  const placement = currentPlacement(directory)
  if (!placement?.issues.length) return []
  const entries = spineEntries(directory)
  return placement.entries.flatMap(row => {
    const entry = entries.find(entry => entry.index === row.entry)
    if (!unplaced(entry, row)) return []
    return [openItem("finalize", "omission", `Document plan: no planned part sets out or reasons from this ${groundLine(entry)}. The lawyer should confirm the drafts deal with it.`, { repair_target: "memo" })]
  })
}
