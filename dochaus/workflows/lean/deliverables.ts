import { createHash } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { combineMemoAndBody } from "../../lib/combined-docx"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { withParagraphPointers } from "../../lib/drafting-amendment"
import { sha256File } from "../../lib/lineage"
import { readWorkingDocx } from "../../lib/working-docx"
import { appendLeanEvent } from "./events"
import { readLeanPlan, type DocID } from "./plan"
import type { LeanDeps } from "./stages"
import { leanFile, readLeanJson, writeLeanBytes, writeLeanText } from "./state"

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
// When the gate passes with unresolved checklist findings, the notes delivered
// are the gate's delivered notes (in memo mode the memo first, then those
// findings; otherwise those findings first, then the drafting notes), under the drafting notes' own kind and name.

export type LeanDocumentKind = "body" | "redline" | "drafting_notes" | "drafting_notes_markdown"

export type LeanDocument = {
  kind: LeanDocumentKind
  /** The planned document of a body or redline; null for the run's notes. */
  document: DocID | null
  /** The planned title of a body or redline; null for the run's notes. */
  title: string | null
  /** Matter-relative path of the run's own copy. */
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

type Produced = { kind: LeanDocumentKind; document: DocID | null; title: string | null; file: string; name: string | null; requested?: boolean }

/** The Word content type of every delivered file, the combined download's included. */
export const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

// The fixed wording of every delivery row: what code sets, and what a row made
// before delivery cannot show.
const PLAIN_CODE_SETS = "Code sets this file's name (registration's naming logic on the current plan and matter files; registration may add \" (n)\" to keep an existing matter document) and its Word format."
const COMBINED_CODE_SETS = "Code sets this file's name (registration's naming logic on the current plan; each download uses exactly this name), its Word format and its part order: the memo, a section break, then the body from a new page. Each download builds it from the two registered parts; it is never registered itself."
const NOT_COVERED = "Not covered: Word changes accepted after delivery, changes the final gate makes to the notes, and rendered layout."
const NOT_WRITTEN = "not yet written"

/** A current view of the same files and names registration consumes. These
 * names come from the model's plan, not independently verified user intent;
 * registration may suffix a name to preserve an existing matter document. With
 * a combined document (combinedDelivery), a last row states the one file built
 * from two of these rows at download, under the name and in the order the
 * download builds it. With `build`, that row also gives each part's sha256 and
 * paragraph count and the properties of the file combineMemoAndBody builds from
 * the run's current copies of both parts (or its refusal code, or which part is
 * not yet written). The build is held in memory only; nothing is written. */
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
  const combined = combinedDelivery(directory)
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
    download_filename: `${combined.name}.docx`,
    content_type: DOCX_CONTENT_TYPE,
    files_delivered: 1,
    contains: built?.contains ?? parts,
    ...(built ? { built_at_check: built.built_at_check } : {}),
    code_sets: COMBINED_CODE_SETS,
    not_covered: NOT_COVERED,
  }]
}

/** The combined file as combineMemoAndBody builds it now from the parts' run
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
  let file: Uint8Array
  try {
    file = combineMemoAndBody({ memo, body })
  } catch (error) {
    return { contains, built_at_check: { refused: combineRefusalCode(error) } }
  }
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

/** combineMemoAndBody's code for a build it refuses, as the combined download
 * reports it: its own codes, invalid_docx_xml for any XML error, and
 * invalid_docx_package for a zip error (fflate's numeric code). */
function combineRefusalCode(error: unknown) {
  const message = error instanceof Error ? error.message : ""
  if (message.startsWith("invalid_docx_xml:")) return "invalid_docx_xml"
  if (/^(?:unsupported_memo_\w+|empty_memo|invalid_docx_package)$/.test(message)) return message
  if (typeof (error as { code?: unknown } | null)?.code === "number") return "invalid_docx_package"
  return "build_failed"
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
  // Document ids are plan positions, so a re-plan can give an id to another
  // document; an earlier row is this file's only when its title matches (for
  // the untitled notes, its deliverable name, which follows a document's name).
  const previous = new Map((readLeanDocuments(directory)?.documents ?? [])
    .map(row => [`${row.kind}:${row.title ?? row.deliverable?.replace(/(?: \(\d+\))?\.docx$/, "") ?? null}`, row]))
  const documents: LeanDocument[] = []
  const produced = producedFiles(directory, register)
  const closed = closedNames(directory, produced)
  for (const item of produced.filter(row => existsSync(leanFile(directory, row.file)))) {
    const source = readFileSync(leanFile(directory, item.file))
    const clean = item.document ? leanFile(directory, `${item.document}/body-applied.docx`) : undefined
    // Normalize legacy generated copies before their delivery identity is read.
    // Publication must not silently give the registry different bytes/hash.
    const bytes = register && item.name !== null ? withParagraphPointers(source,
      item.kind === "redline" && clean && existsSync(clean) ? readFileSync(clean) : undefined) : source
    if (bytes !== source) writeLeanBytes(directory, item.file, bytes)
    const sha256 = createHash("sha256").update(bytes).digest("hex")
    const prior = previous.get(`${item.kind}:${item.title ?? item.name}`)
    const kept = register && item.name !== null ? acceptedRevisionName(directory, item.name, prior) : undefined
    const deliverable = kept ?? (register && item.name !== null && deps.registerDeliverable
      ? await registerOnce(directory, deps.registerDeliverable, item.name, bytes, sha256, prior, [...previous.values()].filter(row => row.title === item.title), closed(item, item.name))
      : prior?.sha256 === sha256 ? prior.deliverable : null)
    documents.push({
      kind: item.kind, document: item.document, title: item.title, path: path.relative(directory, leanFile(directory, item.file)), sha256, deliverable,
      ...(kept && sha256File(path.join(directory, kept))?.sha256 !== sha256 ? { conflict: "accepted_revision" as const } : {}),
    })
  }
  const summary: LeanDocuments = { schema: "dochaus.lean.documents/2", status: register ? "completed" : "blocked", documents }
  writeLeanText(directory, "documents.json", `${JSON.stringify(summary, null, 2)}\n`)
  if (register) appendLeanEvent(directory, { stage: "run", event: "delivered", outputs: documents.flatMap(row => row.deliverable ?? []) })
  return summary
}

/** Drops the summary rows of documents no longer planned. Rows are matched by
 * title, since a plan edit can give a kept document another id. A file already
 * registered for a dropped document stays in the matter: the run never removes
 * a matter document. */
export function forgetUnplannedDocuments(directory: string) {
  const summary = readLeanDocuments(directory)
  const titles = new Set(readLeanPlan(directory)?.documents.map(row => row.title))
  if (!summary || !titles.size) return
  const documents = summary.documents.filter(row => row.title === null || titles.has(row.title))
  if (documents.length < summary.documents.length) writeLeanText(directory, "documents.json", `${JSON.stringify({ ...summary, documents }, null, 2)}\n`)
}

/** Every file the run may deliver, in plan order, with its deliverable name
 * stem (null for the Markdown notes, which are never registered). A document
 * the request gave a file name is delivered under that name; any other under
 * its title. A combined document's body is "<stem> - clean", as a redline
 * deliverable's is, leaving its name to the combined file (combinedDelivery).
 * Notes use their selected name or the first document's name. Once
 * the gate passed (`delivered`), its delivered notes, when it wrote them,
 * replace the drafting notes. */
function producedFiles(directory: string, delivered = false): Produced[] {
  const deliveryPlan = readLeanPlan(directory)
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

/** The names, lowercased, a file's registration may not take: a name another
 * file was requested under; a combined document's "<stem>.docx", the file built
 * at download (combinedDelivery); and, for a requested memo, every planned
 * document's name. The memo is named after a planned document (memoNotesFile),
 * so it can repeat a body, clean or redline name; it then yields and takes the
 * next suffix. */
function closedNames(directory: string, produced: Produced[]) {
  const key = (name: string) => `${name}.docx`.toLocaleLowerCase()
  const memo = requestedMemo(directory) !== undefined
  const yields = (row: Produced) => memo && row.document === null
  const requested = produced.flatMap(row => row.requested && row.name && !yields(row) ? [key(row.name)] : [])
  const documents = produced.flatMap(row => row.document !== null && row.name ? [key(row.name)] : [])
  const combined = combinedDelivery(directory)
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

function deliverableNames(stem: string) {
  return Array.from({ length: MAX_NAME_SUFFIX }, (_, index) => `${stem}${index ? ` (${index + 1})` : ""}.docx`)
}

/** Tries this run's earlier deliverable name first, then "<title>.docx",
 * "<title> (2).docx" and on, skipping closed names (closedNames). A name is
 * usable when its file is absent (create), holds these exact bytes, or still
 * holds the bytes this run registered there. */
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
  const names = deliverableNames(stem)
  for (const name of [...new Set([...(prior?.deliverable && names.includes(prior.deliverable) ? [prior.deliverable] : []), ...names])]) {
    if (closed.has(name.toLocaleLowerCase())) continue
    const target = path.join(directory, name)
    const current = existsSync(target) ? createHash("sha256").update(readFileSync(target)).digest("hex") : null
    const usable = current === null || current === sha256 || owned.some(row => row.deliverable === name && current === row.sha256)
    if (!usable) continue
    const registered = await registerDeliverable(directory, name, bytes, current)
    if (registered) return registered
  }
  throw new Error(`No free deliverable name for "${stem}.docx" after ${MAX_NAME_SUFFIX} tries; rename or remove an existing matter document.`)
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
 * A notes file without a memo name (a delivery selection's rename) is no memo. */
export function requestedMemo(directory: string): { file: string; name: string } | undefined {
  const plan = readLeanPlan(directory)
  if (!plan?.notes_file || !plan.memo) return
  return { file: plan.notes_file, name: plan.memo }
}

/** The one file a combined document is delivered as with the requested memo:
 * the document's id and its requested delivery name stem, so "<name>.docx" is
 * the requested file. The file is built at download time from the registered
 * body and notes and is never written or registered. Undefined unless the plan
 * has exactly one combined document and a requested memo. */
export function combinedDelivery(directory: string): { document: DocID; name: string } | undefined {
  const combined = readLeanPlan(directory)?.documents.filter(row => row.deliverable === "combined") ?? []
  if (combined.length !== 1 || !requestedMemo(directory)) return
  return { document: combined[0]!.id, name: deliveryStem(combined[0]!) }
}

/** The combined file's parts as the current plan delivers them: the registered
 * names of its body ("<stem> - clean") and of the requested memo's notes. Each
 * summary row must be that part's (kind, document and title) under a name
 * registration gives its current stem: the stem or a " (n)" suffix, in any case
 * (the registry keeps an existing file's spelling). Rows delivered under another
 * plan (another deliverable, file or memo name) never match, and a run stopped
 * at its gate lists a part it changed without a name. Plain file names only;
 * reads the plan and summary, not the registry or the files. Undefined unless
 * combinedDelivery holds and both parts are listed. */
export function combinedParts(directory: string): { document: DocID; name: string; body: string; notes: string } | undefined {
  const delivery = combinedDelivery(directory)
  if (!delivery) return
  const rows = readLeanDocuments(directory)?.documents ?? []
  const registered = (part: Produced | undefined) => {
    if (!part?.name) return
    const names = new Set(deliverableNames(part.name).map(name => name.toLocaleLowerCase()))
    return rows.find(row => row.kind === part.kind && row.document === part.document && row.title === part.title &&
      names.has(row.deliverable?.toLocaleLowerCase() ?? ""))?.deliverable ?? undefined
  }
  const produced = producedFiles(directory)
  const body = registered(produced.find(row => row.kind === "body" && row.document === delivery.document))
  const notes = registered(produced.find(row => row.kind === "drafting_notes"))
  return body && notes ? { ...delivery, body, notes } : undefined
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
