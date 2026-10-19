import { Database } from "bun:sqlite"
import { createHash } from "node:crypto"
import { existsSync, lstatSync, readFileSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../../capabilities/artifacts/mutation-lock"
import { buildCombinedFile } from "../../lib/combined-docx"
import { readDocxCommentParagraphs, readDocxRevisionProperties } from "../../lib/docx-comments"
import { withParagraphPointers } from "../../lib/drafting-amendment"
import { sha256File } from "../../lib/lineage"
import { ownedDeliverable, readSystemDocx, replaceableDeliverable } from "../../lib/system-docx"
import { readWorkingDocx } from "../../lib/working-docx"
import { appendLeanEvent } from "./events"
import { readLeanPlan, type DocID, type LeanPlan } from "./plan"
import { openItem } from "./quotes"
import type { LeanDeps } from "./stages"
import { leanFile, readLeanJson, writeLeanBytes, writeLeanText, type OpenItem } from "./state"

// Documents a lean run produces: for each planned document its final body and
// redline, and once for the run the drafting notes. When the run stops at its final gate, the
// finished files are listed in .dochaus/lean/documents.json, the summary
// drafting-status reports. When the gate passes, the Word files are also
// registered as matter deliverables, so they appear in the matter's document
// list next to the sources. Registration replaces a file only when it still holds
// exactly the bytes this run registered before (or the same bytes again), so it
// never overwrites a matter document the run did not write or a lawyer changed.
// Once the lawyer accepts a Word revision of a file this run registered, the run
// keeps that file and name: a newer output is neither written over the revision
// nor forked to "<name> (2)"; its row says so and the output stays in the run.
// A body's redline, which each accept into that body rebuilds from it, stays
// with the body it shows the same way, while the file under its name is still a
// registered deliverable redline of that body.
// When the gate passes, the notes delivered are the gate's delivered notes,
// which list any unresolved checklist findings and then, in a section of their
// own, any requirements left open (in memo mode the memo first, then both
// sections; otherwise both sections first, then the drafting notes), under the
// drafting notes' own kind and name.
// With a combined document (combinedDelivery), the file built from its two
// registered parts (buildCombinedFile) is then registered under the requested
// name. The system replaces a file there only while the summary lists it as
// the combined file, it still holds exactly the bytes the system last wrote
// there, by the registry's record of its own writes (system-docx.ts), and no
// proposal on it is pending review (replaceableDeliverable, checked again under
// the write's lock); a run's save and ingest's rebuild after an accept into a
// part both record theirs. Any other file under the name is kept, and listed
// as an earlier combined file when a delivery saved it there, and the new file
// takes the next free name, as a body does. A save refused under the write's
// lock takes the next free name too; when the lawyer's request for the file then
// gives an item the gate's notes did not disclose (requestedFilesItems), the run
// fails visibly once the summary lists it. When the build is refused, the file under the name
// holds the lawyer's accepted Word revision or no name is free, it is not
// saved: a finalize item the notes disclose says so (combinedFileItems).

export type LeanDocumentKind = "body" | "redline" | "drafting_notes" | "drafting_notes_markdown" | "combined"

export type LeanDocument = {
  kind: LeanDocumentKind
  /** The planned document of a body or redline; null for the run's notes. */
  document: DocID | null
  /** The planned title of a body or redline; null for the run's notes. */
  title: string | null
  /** Matter-relative path of the run's own copy; for the combined file, of the
   * matter file, as the run keeps no copy of it (while the matter has not taken
   * it, the name registration gave it). */
  path: string
  sha256: string
  /** Matter-relative path of the registered deliverable, once registered. */
  deliverable: string | null
  /** Set when `deliverable` holds the lawyer's accepted Word revision and these
   * bytes differ: this output was not registered over it, or under another name. */
  conflict?: "accepted_revision"
}

export type LeanDocuments = { schema: "dochaus.lean.documents/2"; status: "blocked" | "completed"; documents: LeanDocument[] }

/** The gate's notes delivered in place of the drafting notes; see deliveredNotes. */
export const DELIVERED_NOTES = { docx: "delivered-notes.docx", markdown: "delivered-notes.md" }

type Produced = { kind: Exclude<LeanDocumentKind, "combined">; document: DocID | null; title: string | null; file: string; name: string | null; requested?: boolean }

/** The Word content type of every delivered file, the combined download's included. */
export const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

// The fixed wording of every delivery row: what code sets, and what a row made
// before delivery cannot show.
const PLAIN_CODE_SETS = "Code sets this file's name (registration's naming logic on the current plan and matter files; registration may add \" (n)\" to keep an existing matter document) and its Word format."
const COMBINED_CODE_SETS = "Code sets this file's name (registration's naming logic on the current plan; registration adds \" (n)\" to keep another file that has the name), its Word format and its part order: the memo, a section break, then the body from a new page. Once the final gate passes, code builds it from the two registered parts and saves it in the matter; the download serves the saved file. The system replaces a file under the name only while it still holds, unchanged, the combined file the system last saved there, with no change pending review. When the build is refused, the file of that name holds the lawyer's accepted Word changes, or no numbered name is free, it is not saved, and a finalize item says so."
const NOT_COVERED = "Not covered: Word changes accepted after delivery, changes the final gate makes to the notes, and rendered layout."
const NOT_WRITTEN = "not yet written"

/** A current view of the same files and names registration consumes. These
 * names come from the model's plan, not independently verified user intent;
 * registration may suffix a name to preserve an existing matter document. With
 * a combined document (combinedDelivery), a last row states the one file built
 * from two of these rows once they are registered, under the name registration
 * gave or will give it (combinedFileName) and in the order registration builds
 * it. With `build`, that row also
 * gives each part's sha256 and paragraph count and the properties of the file
 * buildCombinedFile builds from the run's current copies of both parts (or its
 * refusal code, or which part is not yet written). The build is held in memory
 * only; nothing is written. */
export function deliveryArtifactEvidence(directory: string, options: { build?: boolean } = {}) {
  const rows = producedFiles(directory).flatMap(row => row.name === null ? [] : [{
    kind: row.kind,
    document: row.document,
    path: path.relative(directory, leanFile(directory, row.file)),
    planned_filename: `${row.name}.docx`,
    file_status: existsSync(leanFile(directory, row.file)) ? "present" : NOT_WRITTEN,
    delivery_registration: "pending; code registers this file once the final gate passes",
    code_sets: PLAIN_CODE_SETS,
    not_covered: NOT_COVERED,
  }])
  const combined = titledCombinedDelivery(directory)
  const notes = rows.find(row => row.kind === "drafting_notes")
  const body = combined && rows.find(row => row.kind === "body" && row.document === combined.document)
  if (!combined || !notes || !body) return rows
  const parts = [
    { order: 1, kind: notes.kind, role: "the requested memo", path: notes.path, planned_filename: notes.planned_filename, file_status: notes.file_status },
    { order: 2, kind: body.kind, document: body.document, role: "the clean body", path: body.path, planned_filename: body.planned_filename, file_status: body.file_status },
  ]
  const built = options.build ? builtCombined(directory, parts) : undefined
  return [...rows, {
    kind: "combined" as const,
    document: combined.document,
    download_filename: combinedFileName(directory, combined),
    content_type: DOCX_CONTENT_TYPE,
    files_delivered: 1,
    contains: built?.contains ?? parts,
    ...(built ? { built_at_check: built.built_at_check } : {}),
    code_sets: COMBINED_CODE_SETS,
    not_covered: NOT_COVERED,
  }]
}

/** The combined file as buildCombinedFile builds it now from the parts' run
 * copies, in memory: its sha256, its non-empty accepted paragraphs and whether
 * they are the memo's followed by the body's; else the build's refusal code,
 * or which part is not yet written. Each written part gains its sha256 and
 * paragraph count (left out when the part cannot be read). */
function builtCombined<T extends { kind: string; path: string; file_status: string }>(directory: string, parts: T[]) {
  const bytes = parts.map(part => part.file_status === NOT_WRITTEN ? undefined : new Uint8Array(readFileSync(path.join(directory, part.path))))
  const texts = bytes.map(value => value && paragraphTexts(value))
  const contains = parts.map((part, index) => {
    const value = bytes[index], text = texts[index]
    return value ? { ...part, sha256: sha256(value), ...(text ? { nonempty_paragraphs: text.length } : {}) } : part
  })
  const [memo, body] = bytes
  if (!memo || !body) {
    const missing = parts.filter((_, index) => !bytes[index]).map(part => part.kind === "body" ? "body" : "memo")
    return { contains, built_at_check: { not_built: `The ${missing.join(" and ")} part${missing.length > 1 ? "s are" : " is"} not yet written at check time.` } }
  }
  const built = buildCombinedFile({ memo, body })
  if ("refused" in built) return { contains, built_at_check: { refused: built.refused } }
  const file = built.bytes
  const text = paragraphTexts(file)
  const [memoText, bodyText] = texts
  return { contains, built_at_check: {
    sha256: sha256(file),
    ...(text ? { nonempty_paragraphs: text.length } : {}),
    ...(text && memoText && bodyText ? { paragraphs_equal_memo_then_body: JSON.stringify(text) === JSON.stringify([...memoText, ...bodyText]) } : {}),
  } }
}

function sha256(bytes: Uint8Array) {
  return createHash("sha256").update(bytes).digest("hex")
}

/** Non-empty accepted paragraph texts, in order; undefined when unreadable. */
function paragraphTexts(bytes: Uint8Array) {
  try {
    return readDocxCommentParagraphs(bytes).map(row => row.paragraph_text)
  } catch {
    return undefined
  }
}

const MAX_NAME_SUFFIX = 9

export function readLeanDocuments(directory: string): LeanDocuments | undefined {
  const value = readLeanJson(directory, "documents.json")
  if (typeof value !== "object" || value === null || !("documents" in value) || !Array.isArray(value.documents)) return
  return value as LeanDocuments
}

/** Writes the documents summary. With `register` set (the gate passed), the Word
 * files also become matter deliverables through `deps.registerDeliverable`. A
 * registration failure throws: the files exist but the lawyer cannot see them. */
export async function recordLeanDocuments(directory: string, deps: LeanDeps, register: boolean): Promise<LeanDocuments> {
  const earlier = readLeanDocuments(directory)?.documents ?? []
  const previous = earlierRows(earlier)
  const documents: LeanDocument[] = []
  // A stop at the gate after a delivery of the gate's notes lists those notes,
  // which stay delivered while unchanged, as a body does.
  const notes = path.relative(directory, leanFile(directory, DELIVERED_NOTES.docx))
  const produced = producedFiles(directory, register || earlier.some(row => row.kind === "drafting_notes" && row.path === notes && row.deliverable !== null))
  const closed = closedNames(directory, produced)
  const items = produced.filter(row => existsSync(leanFile(directory, row.file)))
  const priorRow = (item: Produced) => previous.get(`${item.kind}:${item.title ?? item.name}`)
  // The rows whose files a part may replace: this run's earlier parts of the
  // same title. Never a combined file: only its own save, which refuses a file
  // with pending proposals, rebuilds it; a part planned under its name takes a free one.
  const owns = (title: string | null) => [...previous.values()].filter(row => row.title === title && row.kind !== "combined")
  // Each document's body kept at the lawyer's accepted revision, found before any
  // row is recorded, so its redline's keep does not depend on the rows' order.
  const keptBodies = new Map(items.flatMap(item => {
    const kept = register && item.kind === "body" && item.document && item.name !== null ? acceptedRevisionName(directory, item.name, priorRow(item)) : undefined
    return kept ? [[item.document!, kept] as const] : []
  }))
  for (const item of items) {
    const source = readFileSync(leanFile(directory, item.file))
    const clean = item.document ? leanFile(directory, `${item.document}/body-applied.docx`) : undefined
    // Normalize legacy generated copies before their delivery identity is read.
    // Publication must not silently give the registry different bytes/hash.
    const bytes = register && item.name !== null ? withParagraphPointers(source,
      item.kind === "redline" && clean && existsSync(clean) ? readFileSync(clean) : undefined) : source
    if (bytes !== source) writeLeanBytes(directory, item.file, bytes)
    const sha256 = createHash("sha256").update(bytes).digest("hex")
    const prior = priorRow(item)
    const body = item.kind === "redline" && item.document ? keptBodies.get(item.document) : undefined
    const kept = register && item.name !== null
      ? body ? keptRedlineName(directory, item.name, prior, body) : acceptedRevisionName(directory, item.name, prior) : undefined
    const deliverable = kept ?? (register && item.name !== null && deps.registerDeliverable
      ? await registerOnce(directory, deps.registerDeliverable, item.name, bytes, sha256, prior, owns(item.title), closed(item, item.name))
      : prior?.sha256 === sha256 ? prior.deliverable : null)
    documents.push({
      kind: item.kind, document: item.document, title: item.title, path: path.relative(directory, leanFile(directory, item.file)), sha256, deliverable,
      ...(kept && sha256File(path.join(directory, kept))?.sha256 !== sha256 ? { conflict: "accepted_revision" as const } : {}),
    })
  }
  const delivery = titledCombinedDelivery(directory)
  const combined = delivery && await combinedRow(directory, deps, register, delivery, documents, earlier)
  const outputs = documents.flatMap(row => row.deliverable ?? [])
  const own = combined ? [...combined.kept, ...(combined.row ? [combined.row] : [])] : []
  // The combined files a delivery under another plan saved (another title or
  // name, or no combined delivery now) stay listed while they are in the
  // matter, before this run's (`.at(-1)`), so no file a delivery saved as a
  // combined file is ever listed as a draft.
  documents.push(...keptCombined(directory, earlier.filter(row => row.kind === "combined" && !(delivery && combinedOf(delivery, row))), [...outputs, ...own.map(row => row.path)]), ...own)
  const summary: LeanDocuments = { schema: "dochaus.lean.documents/2", status: register ? "completed" : "blocked", documents }
  writeLeanText(directory, "documents.json", `${JSON.stringify(summary, null, 2)}\n`)
  if (register) appendLeanEvent(directory, { stage: "run", event: "delivered", outputs: [...outputs, ...(combined?.row?.deliverable ? [combined.row.deliverable] : [])] })
  // The parts are delivered and listed; the run then fails visibly.
  if (combined?.failure) throw new Error(combined.failure)
  if (combined?.moved) {
    // A save refused under the write's lock moved the file after the gate's
    // notes were checked: a request whose file the lawyer's words or the notes
    // name is checked against the name it was saved under, read from this
    // summary. One file asked for without a name, which the notes did not
    // place, is still the one file delivered; the next finalize discloses where
    // it went. Imported here, as that check reads this module.
    const { requestedFilesItems } = await import("./requested-files")
    const told = gateDisclosures(directory)
    const item = requestedFilesItems(directory, told).find(item => !told.some(row => row.id === item.id))
    if (item) throw new Error(item.detail)
  }
  return summary
}

/** The items the gate's record lists as disclosed by its notes. */
function gateDisclosures(directory: string) {
  const gate = readLeanJson(directory, "gate.json") as { items?: OpenItem[] } | undefined
  return gate?.items?.filter(row => row.status === "disclosed") ?? []
}

/** Whether the gate's record lists `item` as disclosed by its notes. */
function gateDisclosed(directory: string, item: OpenItem) {
  return gateDisclosures(directory).some(row => row.id === item.id)
}

/** The summary's earlier rows by kind and title. Document ids are plan
 * positions, so a re-plan can give an id to another document; an earlier row
 * is a file's only when its title matches (for the untitled notes, its
 * deliverable name, which follows a document's name). */
function earlierRows(rows: LeanDocument[]) {
  return new Map(rows.map(row => [`${row.kind}:${row.title ?? row.deliverable?.replace(/(?: \(\d+\))?\.docx$/, "") ?? null}`, row]))
}

/** The summary's combined rows of this combined document (its title, under a
 * name registration gives it: "<name>.docx" or a " (n)" suffix), in order: any
 * earlier combined files a delivery kept, then the run's own (`.at(-1)`), the
 * file it saved or the name a failed save gave it. */
function earlierCombined(delivery: CombinedDelivery, rows: LeanDocument[]) {
  return rows.filter(row => combinedOf(delivery, row)).map(row => ({ ...row, document: delivery.document }))
}

/** Whether `row` is a combined row of this combined document: its title, under
 * a name registration gives it. */
function combinedOf(delivery: CombinedDelivery, row: LeanDocument) {
  return row.kind === "combined" && row.title === delivery.title &&
    deliverableNames(delivery.name).some(name => name.toLocaleLowerCase() === row.path.toLocaleLowerCase())
}

/** The combined file's summary rows: `kept`, the earlier combined files still in
 * the matter under names the run's file does not take (keptCombined), and
 * `row`, the run's own. Once the gate passed, the file built from the two parts
 * as registered (registeredParts on this run's `documents`) is registered under
 * the first name it may take (combinedOutcome): replacing the file the system
 * saved there, which must still be the system's under the write's lock, or
 * under the next free name, keeping the file there. A file not
 * saved keeps the run's earlier row (`prior`) without a deliverable, or, when
 * the matter failed to take it, a row without a deliverable under the name
 * registration gave it, so a retry takes that name again, as a body's row keeps
 * its name. At a stop before delivery, the earlier row is kept, delivered only
 * while both parts are and its file still holds the system's last write of it
 * (ownedDeliverable). `failure` says why the file
 * is not saved when the gate's notes did not disclose it (combinedFileItems);
 * `moved`, that a save refused under the write's lock saved it under a later name. */
async function combinedRow(directory: string, deps: LeanDeps, register: boolean, delivery: CombinedDelivery, documents: LeanDocument[], rows: LeanDocument[]): Promise<{ kept: LeanDocument[]; row?: LeanDocument; failure?: string; moved?: true }> {
  const earlier = earlierCombined(delivery, rows)
  const prior = earlier.at(-1)
  const others = [...rows, ...documents].filter(row => row.kind !== "combined")
  const kept = (name: string | undefined) => keptCombined(directory, earlier, [...(name ? [name] : []), ...others.flatMap(row => row.deliverable ?? [])])
  const parts = registeredParts(directory, documents)
  if (!register || !deps.registerDeliverable)
    return { kept: kept(prior?.path), row: prior && { ...prior, deliverable: parts && present(directory, prior.path) && ownedDeliverable(directory, prior.path) ? prior.deliverable : null } }
  const unsaved = (failure?: string) => ({ kept: kept(prior?.path), row: prior && { ...prior, deliverable: null }, ...(failure ? { failure } : {}) })
  const read = (file: string) => new Uint8Array(readFileSync(path.join(directory, file)))
  const outcome = combinedOutcome(directory, delivery, parts && { memo: read(parts.notes), body: read(parts.body) }, earlier, others)
  if (!outcome) return unsaved(`The combined file "${combinedName(delivery, prior)}" is not saved in the matter: the memo and "${delivery.title}" are not both registered.`)
  if ("item" in outcome) return unsaved(gateDisclosed(directory, outcome.item) ? undefined : outcome.item.detail)
  const row = (name: string, deliverable: string | null): LeanDocument => ({ kind: "combined", document: delivery.document, title: delivery.title, path: name, sha256: outcome.sha256, deliverable })
  const save = async (name: string, current: string | null) => {
    try {
      // A file the system replaces must still be its own, checked again under the write's lock.
      return await deps.registerDeliverable!(directory, name, outcome.bytes, current, current ? { owned: true } : undefined)
    } catch (error) {
      return error instanceof Error ? error : new Error(String(error))
    }
  }
  for (let next: CombinedName | void = outcome.first; next; next = outcome.rest.next().value) {
    const registered = await save(next.name, next.current)
    // The parts are registered; a file the matter cannot take is the run's visible failure, not a lost summary.
    if (registered instanceof Error) return { kept: kept(next.name), row: row(next.name, null), failure: `The combined file "${next.name}" is not saved in the matter: ${registered.message}` }
    if (registered) return { kept: kept(registered), row: row(registered, registered), ...(next !== outcome.first ? { moved: true } : {}) }
  }
  return { kept: kept(outcome.first.name), row: row(outcome.first.name, null), failure: `The combined file "${outcome.first.name}" is not saved in the matter: the matter refused it.` }
}

/** The earlier combined files (`earlier`) still in the matter: each row's
 * delivered file, unless a source or one of the names `taken` (the run's own
 * file, or another output's), listed under its name with its deliverable while
 * the file is still the delivery's (earlierCombinedFile). A row without a
 * deliverable lists no file: whatever holds that name now, a delivery did not
 * save it there. The run keeps these files; it never removes a matter document. */
function keptCombined(directory: string, earlier: LeanDocument[], taken: string[]) {
  const closed = new Set(taken.map(name => name.toLocaleLowerCase()))
  const kept = new Map<string, LeanDocument>()
  for (const row of earlier) {
    // A row without a deliverable names no file a delivery saved: the file now under that name is not the run's.
    const name = row.deliverable
    if (!name || closed.has(name.toLocaleLowerCase()) || !present(directory, name) || sourceName(directory, name) || !earlierCombinedFile(directory, name)) continue
    kept.delete(name.toLocaleLowerCase())
    kept.set(name.toLocaleLowerCase(), { ...row, path: name, deliverable: name })
  }
  return [...kept.values()]
}

/** Whether the matter file `name` a delivery saved as a combined file is still
 * that file: the system's last write of it (ownedDeliverable), or a revision of
 * it written through doc.haus (an accept, bound to an accepted Word revision or
 * not, or an upload), which removes the record of the system's write
 * (readSystemDocx). A file changed outside doc.haus, whose bytes no longer
 * match the record of the system's write or of the accepted Word revision
 * (readWorkingDocx), is not, nor is one whose record cannot be read. Reads only.
 * Known gap, a label only: after an accept not bound to a Word revision, an
 * upload, or a delete through doc.haus, no record is left to compare, so a
 * later save outside doc.haus under that name is still listed as the earlier
 * combined file. No write follows from that label: no part replaces a
 * combined file, and a rebuild needs the current plan's registered parts and
 * the system's own last write (replaceableDeliverable). Closing it would need
 * a new record of the file's bytes. */
function earlierCombinedFile(directory: string, name: string) {
  try {
    if (ownedDeliverable(directory, name)) return true
    if (readSystemDocx(directory, name) !== undefined) return false
    const file = path.join(directory, name)
    const accepted = readWorkingDocx(directory, file)
    return !accepted || accepted.source_sha256 === sha256File(file)?.sha256
  } catch {
    return false
  }
}

/** Whether a file, link or anything else has the matter name `name`. */
export function present(directory: string, name: string) {
  try {
    lstatSync(path.join(directory, name))
    return true
  } catch {
    return false
  }
}

/** The finalize item saying the combined file will not be saved, as
 * registration will find it (combinedOutcome), from the run's current copies of
 * its parts, or of a part kept at the lawyer's accepted Word revision, the file
 * registration keeps for it. None without a combined delivery, nor, unless the
 * lawyer's accepted Word revision holds the name, while a part is not yet
 * written. Builds in memory only; writes nothing. */
export function combinedFileItems(directory: string): OpenItem[] {
  const delivery = titledCombinedDelivery(directory)
  const outcome = delivery && predictedCombined(directory, delivery).outcome
  return outcome && "item" in outcome ? [outcome.item] : []
}

/** What registration will do with the combined file (combinedOutcome), as
 * combinedFileItems predicts it from the summary's earlier rows. Each part's
 * bytes are read only so the build's refusal is found as registration finds it:
 * the run's current copy, or the matter file a part is kept at. */
function predictedCombined(directory: string, delivery: CombinedDelivery) {
  const rows = readLeanDocuments(directory)?.documents ?? []
  const previous = earlierRows(rows)
  const copy = (item: Produced | undefined) => {
    if (!item?.name) return
    const kept = acceptedRevisionName(directory, item.name, previous.get(`${item.kind}:${item.title ?? item.name}`))
    const file = kept ? path.join(directory, kept) : leanFile(directory, item.file)
    return existsSync(file) ? new Uint8Array(readFileSync(file)) : undefined
  }
  const produced = producedFiles(directory)
  const memo = copy(produced.find(row => row.kind === "drafting_notes"))
  const body = copy(produced.find(row => row.kind === "body" && row.document === delivery.document))
  const earlier = earlierCombined(delivery, rows)
  return { prior: earlier.at(-1), outcome: combinedOutcome(directory, delivery, memo && body ? { memo, body } : undefined, earlier, rows.filter(row => row.kind !== "combined")) }
}

/** The combined file's registered name, in the matter's spelling: the name its
 * delivery saved it under or the next delivery saves it under (registration's
 * collision decision, predicted as combinedFileItems predicts it), else, while
 * that cannot be told or the file is not saved, combinedName. */
function combinedFileName(directory: string, delivery: CombinedDelivery) {
  const { prior, outcome } = predictedCombined(directory, delivery)
  return outcome && "first" in outcome ? outcome.first.name : combinedName(delivery, prior)
}

/** The combined file's matter name: the run's earlier one (`prior`) while that
 * is a name registration gives it ("<name>.docx" or a " (n)" suffix), else
 * "<name>.docx". */
function combinedName(delivery: CombinedDelivery, prior: LeanDocument | undefined) {
  const names = new Set(deliverableNames(delivery.name).map(name => name.toLocaleLowerCase()))
  return prior && names.has(prior.path.toLocaleLowerCase()) ? prior.path : `${delivery.name}.docx`
}

/** The names, lowercased, the combined file's registration may not take: the
 * names closedNames closes to its document's body, the names other planned
 * outputs are requested under, whether or not their files are written yet; but
 * not the combined file's own requested name. */
function combinedClosed(directory: string, delivery: CombinedDelivery) {
  const own = `${delivery.name}.docx`.toLocaleLowerCase()
  const produced = producedFiles(directory)
  const body = produced.find(row => row.kind === "body" && row.document === delivery.document)
  return new Set(body?.name ? [...closedNames(directory, produced)(body, body.name)].filter(name => name !== own) : [])
}

type CombinedName = { name: string; current: string | null }

/** The names the combined file may be saved under, in the order registration
 * tries them (the run's earlier name, then "<name>.docx", "<name> (2).docx" and
 * on, without combinedClosed), each with the sha256 of the file it replaces
 * (null when no file has the name). A name with a file is usable only while the
 * system may replace that file: the summary lists it as this combined file
 * (`earlier`) and as no other output (`others`), it is not a source, it still
 * holds exactly the bytes the system last wrote there (the registry's record)
 * and no proposal on it is pending review (replaceableDeliverable). Any other
 * file there is kept. Read lazily. */
function* combinedNames(directory: string, delivery: CombinedDelivery, earlier: LeanDocument[], others: LeanDocument[]): Generator<CombinedName, void> {
  const names = openNames(delivery.name, combinedClosed(directory, delivery))
  const prior = earlier.at(-1)?.path
  const tried = new Set<string>()
  for (const name of [...(prior && names.some(name => name.toLocaleLowerCase() === prior.toLocaleLowerCase()) ? [prior] : []), ...names]) {
    const key = name.toLocaleLowerCase()
    if (tried.has(key)) continue
    tried.add(key)
    if (!present(directory, name)) {
      yield { name, current: null }
      continue
    }
    const combined = earlier.some(row => row.path.toLocaleLowerCase() === key) && !others.some(row => row.deliverable?.toLocaleLowerCase() === key)
    const current = combined && !sourceName(directory, name) && replaceableDeliverable(directory, name) ? sha256File(path.join(directory, name))?.sha256 : undefined
    if (current) yield { name, current }
  }
}

/** What registration saves as the combined file: the bytes buildCombinedFile
 * gives from `parts`, with paragraph pointers as registration gives them, and
 * the names it may take (combinedNames), the first as `first`, the rest as
 * `rest`; or the finalize item saying why nothing is saved: the file under the
 * combined name (combinedName) holds the lawyer's accepted Word revision, the
 * build is refused, or no name is free. Undefined when `parts` is and no
 * accepted revision holds the name. Reads the matter; writes nothing. */
function combinedOutcome(directory: string, delivery: CombinedDelivery, parts: { memo: Uint8Array; body: Uint8Array } | undefined, earlier: LeanDocument[], others: LeanDocument[]):
  { bytes: Uint8Array; sha256: string; first: CombinedName; rest: Generator<CombinedName, void> } | { item: OpenItem } | undefined {
  const name = combinedName(delivery, earlier.at(-1))
  const file = path.join(directory, name)
  const unsaved = (reason: string) => ({ item: openItem("finalize", "omission", `The combined file "${name}" (the memo, then "${delivery.title}") is not saved in the matter: ${reason}. ` +
    `The memo and "${delivery.title}" are delivered as separate files.`, { repair_target: "memo" }) })
  if (present(directory, name) && readWorkingDocx(directory, file)) return unsaved("the file of that name holds the lawyer's accepted Word changes, which the run does not write over")
  if (!parts) return
  const built = buildCombinedFile(parts)
  if ("refused" in built) return unsaved(built.part ? `the ${built.part === "memo" ? "memo" : "draft"} could not be read as a Word document` : COMBINED_REFUSALS[built.refused] ?? "the file could not be built")
  const rest = combinedNames(directory, delivery, earlier, others)
  const first = rest.next()
  if (first.done) return unsaved("another file already has that name, which the run keeps (it replaces only a combined file it saved there that still holds what it saved, with no change pending review), " +
    `and other files already have each numbered name the run can give the new file, up to "${delivery.name} (${MAX_NAME_SUFFIX}).docx"`)
  // Paragraph pointers as registration gives them, the replaced file's where the text is unchanged.
  const bytes = withParagraphPointers(built.bytes, first.value.current ? readFileSync(path.join(directory, first.value.name)) : undefined)
  return { bytes, sha256: sha256(bytes), first: first.value, rest }
}

/** A refused build's reason, in plain words, by buildCombinedFile's code. */
const COMBINED_REFUSALS: Record<string, string> = {
  empty_memo: "the memo has no text to place before the draft",
  unsupported_memo_numbering: "the memo uses list numbering the combined file cannot carry",
  unsupported_memo_namespace: "the memo uses Word markup the combined file cannot carry",
  unsupported_memo_dependency: "the memo holds footnotes, endnotes, pictures, embedded objects or links to other parts of its file, which the combined file cannot carry",
}

/** Whether the matter registry lists the matter file `name` as a source. */
function sourceName(directory: string, name: string) {
  try {
    return registeredRole(directory, name) === "source"
  } catch {
    return false
  }
}

/** Drops the summary rows of documents no longer planned. Rows are matched by
 * title, since a plan edit can give a kept document another id. A file already
 * registered for a dropped document stays in the matter: the run never removes
 * a matter document. A combined file a delivery saved keeps its row, which
 * lists it as an earlier combined file while it is in the matter
 * (recordLeanDocuments). */
export function forgetUnplannedDocuments(directory: string) {
  const summary = readLeanDocuments(directory)
  const titles = new Set(readLeanPlan(directory)?.documents.map(row => row.title))
  if (!summary || !titles.size) return
  const documents = summary.documents.filter(row => row.title === null || titles.has(row.title) || (row.kind === "combined" && row.deliverable !== null))
  if (documents.length < summary.documents.length) writeLeanText(directory, "documents.json", `${JSON.stringify({ ...summary, documents }, null, 2)}\n`)
}

/** Every file the run may deliver, in plan order, with its deliverable name
 * stem (null for the Markdown notes, which are never registered). A document
 * the request gave a file name is delivered under that name; any other under
 * its title. A combined document's body is "<stem> - clean", as a redline
 * deliverable's is, leaving its name to the combined file (combinedDelivery).
 * Notes use their selected name or the first document's name. Once
 * the gate passed (`delivered`), its delivered notes, when it wrote them,
 * replace the drafting notes. `deliveryPlan` is the saved plan unless given. */
function producedFiles(directory: string, delivered = false, deliveryPlan = readLeanPlan(directory)): Produced[] {
  const plan = (deliveryPlan?.documents ?? []).map(row => ({
    ...row,
    stem: deliveryStem(row),
  }))
  const notes = delivered && existsSync(leanFile(directory, DELIVERED_NOTES.docx)) ? DELIVERED_NOTES : { docx: "drafting-notes.docx", markdown: "drafting-notes.md" }
  return [
    ...plan.flatMap((row): Produced[] => [
      { kind: "body", document: row.id, title: row.title, file: `${row.id}/body-applied.docx`, name: (row.deliverable ?? "body") === "body" ? row.stem : `${row.stem} - clean`, requested: Boolean(row.file) && (row.deliverable ?? "body") === "body" },
      { kind: "redline", document: row.id, title: row.title, file: `${row.id}/redline.docx`, name: row.deliverable === "redline" ? row.stem : `${row.stem} - redline`, requested: Boolean(row.file) && row.deliverable === "redline" },
    ]),
    // A reviewed note can receive its requested name; otherwise use the usual default.
    { kind: "drafting_notes", document: null, title: null, file: notes.docx, name: deliveryPlan?.notes_file ? fileNameStem(deliveryPlan.notes_file.replace(/\.docx$/i, "")) : `${plan[0]?.stem ?? "Draft"} - drafting notes`, requested: Boolean(deliveryPlan?.notes_file) },
    { kind: "drafting_notes_markdown", document: null, title: null, file: notes.markdown, name: null },
  ]
}

/** The names, lowercased, registration gives a plan's files in a matter that
 * holds none of them: each named file of producedFiles, in order, takes the
 * first name registration tries (openNames) that no earlier file took, as the
 * earlier file's registered copy would refuse it; null when none is left. The
 * combined file (combinedDelivery) keeps its name, which no other file takes.
 * `wanted` is the name before any suffix; `requested`, whether the plan
 * requests it.
 * Nothing is read from or written to the matter. */
export function plannedRegistration(directory: string, plan: LeanPlan) {
  const key = (name: string) => `${name}.docx`.toLocaleLowerCase()
  const produced = producedFiles(directory, false, plan)
  const closed = closedNames(directory, produced, plan)
  const taken = new Set<string>()
  const files = produced.flatMap((item): PlannedName[] => {
    if (item.name === null) return []
    const name = openNames(item.name, closed(item, item.name)).map(name => name.toLocaleLowerCase()).find(name => !taken.has(name)) ?? null
    if (name) taken.add(name)
    return [{ kind: item.kind, document: item.document, wanted: key(item.name), name, requested: Boolean(item.requested) }]
  })
  const combined = combinedDelivery(directory, plan)
  return combined ? [...files, { kind: "combined", document: combined.document, wanted: key(combined.name), name: key(combined.name), requested: true }] : files
}

type PlannedName = { kind: LeanDocumentKind; document: DocID | null; wanted: string; name: string | null; requested: boolean }

/** The names, lowercased, a file's registration may not take: a name another
 * file was requested under; a combined document's "<stem>.docx", the combined
 * file's (combinedDelivery); and, for a requested memo, every planned
 * document's name. The memo is named after a planned document (memoNotesFile),
 * so it can repeat a body, clean or redline name; it then yields and takes the
 * next suffix. `plan` is the saved plan unless given. */
function closedNames(directory: string, produced: Produced[], plan = readLeanPlan(directory)) {
  const key = (name: string) => `${name}.docx`.toLocaleLowerCase()
  const memo = requestedMemo(directory, plan) !== undefined
  const yields = (row: Produced) => memo && row.document === null
  const requested = produced.flatMap(row => row.requested && row.name && !yields(row) ? [key(row.name)] : [])
  const documents = produced.flatMap(row => row.document !== null && row.name ? [key(row.name)] : [])
  const combined = combinedDelivery(directory, plan)
  return (item: Produced, name: string) => new Set([
    ...requested.filter(other => !item.requested || other !== key(name)),
    ...(combined ? [key(combined.name)] : []),
    ...(yields(item) ? documents : []),
  ])
}

/** This output's earlier deliverable name when the lawyer has since accepted a
 * Word revision there. The registry refuses to write over such a revision; trying
 * the next free name would silently move the run's delivery away from it. */
function acceptedRevisionName(directory: string, stem: string, prior: LeanDocument | undefined) {
  const name = prior?.deliverable
  if (!name || !deliverableNames(stem).includes(name)) return
  return readWorkingDocx(directory, path.join(directory, name)) ? name : undefined
}

/** The earlier deliverable name of a redline whose body is kept so (`body`, the
 * body's name), while the file there is still a redline of that body
 * (redlineOfBody). Any other file under that name, such as a source or one the
 * lawyer put there, is not the run's redline; the redline then registers like
 * any other output (registerOnce), under a free name. */
function keptRedlineName(directory: string, stem: string, prior: LeanDocument | undefined, body: string) {
  const name = prior?.deliverable
  return name && deliverableNames(stem).includes(name) && redlineOfBody(directory, name, body) ? name : undefined
}

/** Whether the matter file `name` is still the run's redline of the body file
 * `body`, as each accept into the body rebuilds it: it is registered as a
 * deliverable, holds tracked changes of any kind (text, paragraph-mark or
 * formatting) and reads as the body once they are accepted. */
export function redlineOfBody(directory: string, name: string, body: string) {
  try {
    if (registeredRole(directory, name) !== "deliverable") return false
    const redline = readFileSync(path.join(directory, name))
    const tracked = Object.values(readDocxRevisionProperties(redline).revision_elements).some(count => count > 0)
    const text = (bytes: Uint8Array) => JSON.stringify(readDocxCommentParagraphs(bytes).map(row => row.paragraph_text))
    return tracked && text(redline) === text(readFileSync(path.join(directory, body)))
  } catch {
    return false
  }
}

/** The matter registry's role for the matter file `name` ("source" or
 * "deliverable"), read-only; undefined when it is unregistered or its rows disagree. */
function registeredRole(directory: string, name: string) {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return
  const db = new Database(file, { readonly: true })
  try {
    const identity = canonicalArtifactIdentity(directory, name)
    const roles = new Set((db.query("SELECT doc_path, role FROM documents").all() as { doc_path: string; role: string }[])
      .filter(row => canonicalArtifactIdentity(path.dirname(row.doc_path), path.basename(row.doc_path)) === identity)
      .map(row => row.role))
    return roles.size === 1 ? [...roles][0] : undefined
  } finally {
    db.close()
  }
}

function deliverableNames(stem: string) {
  return Array.from({ length: MAX_NAME_SUFFIX }, (_, index) => `${stem}${index ? ` (${index + 1})` : ""}.docx`)
}

/** The names registration tries for a file, in order: "<stem>.docx",
 * "<stem> (2).docx" and on, without its closed names (closedNames). */
function openNames(stem: string, closed: Set<string>) {
  return deliverableNames(stem).filter(name => !closed.has(name.toLocaleLowerCase()))
}

/** Tries this run's earlier deliverable name first, then "<title>.docx",
 * "<title> (2).docx" and on, skipping closed names and names not usable
 * (usableNames). */
async function registerOnce(
  directory: string,
  registerDeliverable: NonNullable<LeanDeps["registerDeliverable"]>,
  stem: string,
  bytes: Uint8Array,
  sha256: string,
  prior: LeanDocument | undefined,
  owned: LeanDocument[],
  closed: Set<string>,
): Promise<string> {
  for (const { name, current } of usableNames(directory, stem, sha256, prior, owned, closed)) {
    const registered = await registerDeliverable(directory, name, bytes, current)
    if (registered) return registered
  }
  throw new Error(`No free deliverable name for "${stem}.docx" after ${MAX_NAME_SUFFIX} tries; rename or remove an existing matter document.`)
}

/** The names registerOnce may register bytes of `sha256` under, in the order
 * it tries them (this run's earlier deliverable name first, then openNames),
 * each with the sha256 of the file it replaces (null when absent). A name is
 * usable when its file is absent (create), holds these exact bytes, or still
 * holds the bytes this run registered there (`owned`). Read lazily, name by name. */
function* usableNames(directory: string, stem: string, sha256: string, prior: LeanDocument | undefined, owned: LeanDocument[], closed: Set<string>) {
  const names = openNames(stem, closed)
  for (const name of new Set([...(prior?.deliverable && names.includes(prior.deliverable) ? [prior.deliverable] : []), ...names])) {
    const target = path.join(directory, name)
    const current = existsSync(target) ? createHash("sha256").update(readFileSync(target)).digest("hex") : null
    if (current === null || current === sha256 || owned.some(row => row.deliverable === name && current === row.sha256)) yield { name, current }
  }
}

/** A planned document's delivery name stem: the file name the request gives
 * it, else its title. */
export function deliveryStem(row: { title: string; file?: string }) {
  return row.file ? fileNameStem(row.file.replace(/\.(?:docx?|pdf|rtf|odt|txt|md)$/i, "")) : deliverableTitle(`# ${row.title}`)
}

/** The notes file of a memo the request asks for by name: "<document's delivery
 * stem> - <memo name>.docx", as delivery will name it. The document is the one
 * the memo goes with: the combined document, when the plan has exactly one (as
 * combinedDelivery); else the first document the request gave a file name; else
 * the first document. The document's stem is shortened, not the memo name,
 * when the whole is too long. */
export function memoNotesFile(documents: { title: string; file?: string; deliverable?: string }[], memo: string) {
  const name = fileNameStem(memo)
  const document = memoAnchor(documents) ?? { title: "" }
  return `${fileNameStem(`${characters(deliveryStem(document), Math.max(1, NAME_STEM_MAX - [...name].length - 3)).trim()} - ${name}`)}.docx`
}

/** The document a requested memo is named after (see memoNotesFile). */
export function memoAnchor<T extends { file?: string; deliverable?: string }>(documents: T[]): T | undefined {
  const combined = documents.filter(row => row.deliverable === "combined")
  return (combined.length === 1 ? combined[0] : documents.find(row => row.file)) ?? documents[0]
}

/** The memo the notes are written as, when the request asks for one: the notes
 * file it is delivered as and the outline's name for it, both from the plan.
 * A notes file without a memo name (a delivery selection's rename) is no memo.
 * `plan` is the saved plan unless given. */
export function requestedMemo(directory: string, plan = readLeanPlan(directory)): { file: string; name: string } | undefined {
  if (!plan?.notes_file || !plan.memo) return
  return { file: plan.notes_file, name: plan.memo }
}

/** The one file a combined document is delivered as with the requested memo:
 * the document's id and its requested delivery name stem, so "<name>.docx" is
 * the requested file. The file is built from the registered body and notes
 * (buildCombinedFile) and registered under that name once they are, or under a
 * " (n)" suffix while another file the system may not replace holds it
 * (recordLeanDocuments). Undefined unless
 * the plan has exactly one combined document and a requested memo. `plan` is
 * the saved plan unless given. */
export function combinedDelivery(directory: string, plan = readLeanPlan(directory)): { document: DocID; name: string } | undefined {
  const combined = plan?.documents.filter(row => row.deliverable === "combined") ?? []
  if (combined.length !== 1 || !requestedMemo(directory, plan)) return
  return { document: combined[0]!.id, name: deliveryStem(combined[0]!) }
}

type CombinedDelivery = { document: DocID; name: string; title: string }

/** combinedDelivery with the combined document's planned title. */
function titledCombinedDelivery(directory: string): CombinedDelivery | undefined {
  const plan = readLeanPlan(directory)
  const delivery = combinedDelivery(directory, plan)
  const document = delivery && plan?.documents.find(row => row.id === delivery.document)
  return delivery && document ? { ...delivery, title: document.title } : undefined
}

/** The combined file's parts as the current plan delivers them: the registered
 * names of its body ("<stem> - clean") and of the requested memo's notes, and
 * `file`, the combined file's name: the summary's last combined row's (the
 * run's own; earlier rows are files a delivery kept), its deliverable, or,
 * while the matter has not taken the file, the name registration gave it.
 * Each summary row must be that file's (kind, document and title) under a name
 * registration gives its current stem: the stem or a " (n)" suffix, in any case
 * (the registry keeps an existing file's spelling). Rows delivered under
 * another plan (another deliverable, file or memo name) never match, and a run
 * stopped at its gate lists a part it changed without a name. Plain file names
 * only; reads the plan and summary (or the given `rows`), not the registry or
 * the files. Undefined unless combinedDelivery holds, both parts are listed and
 * the summary has a combined row. */
export function combinedParts(directory: string, rows = readLeanDocuments(directory)?.documents ?? []): { document: DocID; name: string; file: string; body: string; notes: string } | undefined {
  const parts = registeredParts(directory, rows)
  if (!parts) return
  const names = new Set(deliverableNames(parts.name).map(name => name.toLocaleLowerCase()))
  const row = rows.findLast(row => row.kind === "combined" && row.document === parts.document && row.title === parts.title && names.has(row.path.toLocaleLowerCase()))
  return row ? { document: parts.document, name: parts.name, file: row.deliverable ?? row.path, body: parts.body, notes: parts.notes } : undefined
}

/** combinedParts without the combined file: the registered names of the
 * combined document's two parts, and its planned title. */
function registeredParts(directory: string, rows: LeanDocument[]) {
  const delivery = combinedDelivery(directory)
  if (!delivery) return
  const registered = (part: Produced | undefined) => {
    if (!part?.name) return
    const names = new Set(deliverableNames(part.name).map(name => name.toLocaleLowerCase()))
    return rows.find(row => row.kind === part.kind && row.document === part.document && row.title === part.title &&
      names.has(row.deliverable?.toLocaleLowerCase() ?? ""))?.deliverable ?? undefined
  }
  const produced = producedFiles(directory)
  const part = produced.find(row => row.kind === "body" && row.document === delivery.document)
  const body = registered(part)
  const notes = registered(produced.find(row => row.kind === "drafting_notes"))
  return part && body && notes ? { ...delivery, title: part.title, body, notes } : undefined
}

/** File-name stem from the draft's "# " title. */
export function deliverableTitle(markdown: string) {
  return fileNameStem(/^#\s+(.+)$/m.exec(markdown)?.[1] ?? "")
}

const NAME_STEM_MAX = 120

/** Characters a file name cannot carry, and Markdown code marks, are dropped. */
function fileNameStem(text: string) {
  const stem = characters(text.replace(/[\\/:*?"<>|`\u0000-\u001f]|\p{Cs}/gu, " ").replace(/\s+/g, " ").replace(/^[.\s]+|[.\s]+$/g, ""), NAME_STEM_MAX).trim()
  return stem || "Draft"
}

/** The first `count` characters, counted by code point so a cut never splits a
 * surrogate pair (an ill-formed name no header or file system can carry). */
function characters(text: string, count: number) {
  return Array.from(text).slice(0, count).join("")
}
