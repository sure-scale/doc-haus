import { createHash } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
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

type Produced = { kind: LeanDocumentKind; document: DocID | null; title: string | null; file: string; name: string | null; requested?: boolean }

/** A current view of the same files and names registration consumes. These
 * names come from the model's plan, not independently verified user intent;
 * registration may suffix a name to preserve an existing matter document. */
export function deliveryArtifactEvidence(directory: string) {
  return producedFiles(directory).flatMap(row => row.name === null ? [] : [{
    kind: row.kind,
    document: row.document,
    path: path.relative(directory, leanFile(directory, row.file)),
    planned_filename: `${row.name}.docx`,
    file_status: existsSync(leanFile(directory, row.file)) ? "present" : "pending",
    delivery_registration: "pending",
  }])
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
  // the untitled notes, its deliverable name, which follows the first title).
  const previous = new Map((readLeanDocuments(directory)?.documents ?? [])
    .map(row => [`${row.kind}:${row.title ?? row.deliverable?.replace(/(?: \(\d+\))?\.docx$/, "") ?? null}`, row]))
  const documents: LeanDocument[] = []
  const produced = producedFiles(directory)
  const reserved = new Set(produced.filter(row => row.requested && row.name).map(row => `${row.name}.docx`.toLocaleLowerCase()))
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
      ? await registerOnce(directory, deps.registerDeliverable, item.name, bytes, sha256, prior, [...previous.values()].filter(row => row.title === item.title), reserved, item.requested === true)
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
 * its title. Notes use their selected name or the first document's name. */
function producedFiles(directory: string): Produced[] {
  const deliveryPlan = readLeanPlan(directory)
  const plan = (deliveryPlan?.documents ?? []).map(row => ({
    ...row,
    stem: row.file ? fileNameStem(row.file.replace(/\.(?:docx?|pdf|rtf|odt|txt|md)$/i, "")) : deliverableTitle(`# ${row.title}`),
  }))
  return [
    ...plan.flatMap((row): Produced[] => [
      { kind: "body", document: row.id, title: row.title, file: `${row.id}/body-applied.docx`, name: row.deliverable === "redline" ? `${row.stem} - clean` : row.stem, requested: Boolean(row.file) && row.deliverable !== "redline" },
      { kind: "redline", document: row.id, title: row.title, file: `${row.id}/redline.docx`, name: row.deliverable === "redline" ? row.stem : `${row.stem} - redline`, requested: Boolean(row.file) && row.deliverable === "redline" },
    ]),
    // A reviewed note can receive its requested name; otherwise use the usual default.
    { kind: "drafting_notes", document: null, title: null, file: "drafting-notes.docx", name: deliveryPlan?.notes_file ? fileNameStem(deliveryPlan.notes_file.replace(/\.docx$/i, "")) : `${plan[0]?.stem ?? "Draft"} - drafting notes`, requested: Boolean(deliveryPlan?.notes_file) },
    { kind: "drafting_notes_markdown", document: null, title: null, file: "drafting-notes.md", name: null },
  ]
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
 * "<title> (2).docx" and on. A name is usable when its file is absent (create),
 * holds these exact bytes, or still holds the bytes this run registered there. */
async function registerOnce(
  directory: string,
  registerDeliverable: NonNullable<LeanDeps["registerDeliverable"]>,
  stem: string,
  bytes: Uint8Array,
  sha256: string,
  prior: LeanDocument | undefined,
  owned: LeanDocument[],
  reserved: Set<string>,
  requested: boolean,
): Promise<string> {
  const names = deliverableNames(stem)
  for (const name of [...new Set([...(prior?.deliverable && names.includes(prior.deliverable) ? [prior.deliverable] : []), ...names])]) {
    if (reserved.has(name.toLocaleLowerCase()) && (!requested || name.toLocaleLowerCase() !== `${stem}.docx`.toLocaleLowerCase())) continue
    const target = path.join(directory, name)
    const current = existsSync(target) ? createHash("sha256").update(readFileSync(target)).digest("hex") : null
    const usable = current === null || current === sha256 || owned.some(row => row.deliverable === name && current === row.sha256)
    if (!usable) continue
    const registered = await registerDeliverable(directory, name, bytes, current)
    if (registered) return registered
  }
  throw new Error(`No free deliverable name for "${stem}.docx" after ${MAX_NAME_SUFFIX} tries; rename or remove an existing matter document.`)
}

/** File-name stem from the draft's "# " title. */
export function deliverableTitle(markdown: string) {
  return fileNameStem(/^#\s+(.+)$/m.exec(markdown)?.[1] ?? "")
}

/** Characters a file name cannot carry, and Markdown code marks, are dropped. */
function fileNameStem(text: string) {
  const stem = text.replace(/[\\/:*?"<>|`\u0000-\u001f]/g, " ").replace(/\s+/g, " ").replace(/^[.\s]+|[.\s]+$/g, "").slice(0, 120).trim()
  return stem || "Draft"
}
