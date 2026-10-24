import { currentDocumentArtifacts } from "./current-document-artifacts"
import path from "node:path"
import { realpathSync } from "node:fs"
import { createHash } from "node:crypto"
import { readRegularFile } from "./regular-file"
import { readDocxCommentParagraphs, readDocxComments, readDocxRevisionText } from "./docx-comments"
import { projectedChangeParagraphs } from "../workflows/lean/change-pipeline"
import { readChangeList, readParagraphs } from "../workflows/lean/checks"
import { readLeanPlan } from "../workflows/lean/plan"
import { UNTRUSTED_DOCUMENT_NOTICE } from "./untrusted"

export type CurrentDocumentView = "current_body" | "comments" | "deleted_text" | "proposed_body"
/** `label` is the list number Word prints before a body paragraph ("2.1"); it is not part of `text`. */
export type CurrentDocumentEntry = { index: number; artifact: string; sha256: string; label?: string; text: string; paragraph_index?: number; pids?: string[]; comment_id?: string; markers?: unknown; representation?: "generated_proposal"; initial_pid?: string; proposed_part_index?: number; basis?: string[] }
type CurrentView = { status: "available"; document: string; view: CurrentDocumentView; entries: CurrentDocumentEntry[] } | { status: "unavailable"; document: string; view: CurrentDocumentView; reason: string }

/** Actual Word and the explicitly requested pending projection remain distinct evidence views. */
export function currentDocumentView(directory: string, document: string, view: CurrentDocumentView): CurrentView {
  const plan = readLeanPlan(directory)
  const declared = [...(plan?.documents ?? []).map(row => ({ id: row.id, file: row.file ?? row.id })), ...(plan ? [{ id: "notes", file: plan.notes_file ?? "drafting-notes.docx" }] : [])]
  const matches = declared.filter(row => row.id === document || row.file === document)
  if (matches.length !== 1) return { status: "unavailable", document, view, reason: "Current document must have one exact plan ID or filename match. Use the declared current-document inventory; uploaded sources use view=source." }
  const selected = matches[0]!
  const read = (artifact: string) => {
    const file = path.resolve(directory, artifact)
    const parent = path.relative(realpathSync(directory), realpathSync(path.dirname(file)))
    if (parent === ".." || parent.startsWith(`..${path.sep}`) || path.isAbsolute(parent)) throw Error("Artifact is outside the matter.")
    return readRegularFile(file)
  }
  try {
    if (view === "proposed_body") {
      if (selected.id === "notes") return { status: "unavailable", document, view, reason: "Drafting notes have no pending paragraph-amendment projection." }
      const doc = selected.id as `d${number}`
      const changes = readChangeList(directory, doc), initial = readParagraphs(directory, doc)
      if (!changes || !initial.length || changes.amendments.some(row => initial.find(paragraph => paragraph.pid === row.pid)?.sha256 !== row.existing_sha256)) return { status: "unavailable", document, view, reason: "Pending changes are missing or do not match the current initial paragraph identities/content. No proposed-body search was performed." }
      const artifact = `.dochaus/lean/${doc}/changes.json`, basis = [`.dochaus/lean/${doc}/paragraphs.json`, artifact]
      const inputs = basis.map(file => read(file))
      if (inputs.some(bytes => !bytes)) throw Error("Missing projection input")
      const sha256 = createHash("sha256").update(inputs[0]!).update(inputs[1]!).digest("hex")
      const parts = new Map<string, number>()
      const entries = projectedChangeParagraphs(directory, doc).map((paragraph, index) => {
        const part = parts.get(paragraph.pid) ?? 0;parts.set(paragraph.pid, part + 1)
        return { index, artifact, sha256, basis, representation: "generated_proposal" as const, initial_pid: paragraph.pid, proposed_part_index: part, ...(paragraph.label ? { label: paragraph.label } : {}), text: paragraph.text }
      })
      return { status: "available", document, view, entries }
    }
    const current = selected.id === "notes" ? undefined : currentDocumentArtifacts(directory, selected.id as `d${number}`)
    const candidates = selected.id === "notes" ? [".dochaus/lean/drafting-notes.docx"] : current?.body ? [`.dochaus/lean/${current.body}`] : []
    let body: { artifact: string; bytes: Buffer } | undefined
    for (const artifact of candidates) {
      const bytes = read(artifact)
      if (bytes) { body = { artifact, bytes };break }
    }
    if (!body?.bytes) return { status: "unavailable", document, view, reason: "No readable current Word body exists for this planned document." }
    const artifacts = view === "current_body" || selected.id === "notes" ? [body] : [body, ...(current?.redline ? [{ artifact: `.dochaus/lean/${current.redline}`, bytes: read(`.dochaus/lean/${current.redline}`) }] : [])].filter(row => row.bytes)
    // List numbers come from the lean rows recorded with this body, by paragraph PID.
    const rows = current?.body ? readParagraphs(directory, selected.id as `d${number}`, current.applied ? "paragraphs-applied.json" : "paragraphs.json") : []
    const labels = new Map(rows.flatMap(row => row.label ? [[row.pid, row.label] as const] : []))
    const entries: CurrentDocumentEntry[] = []
    for (const row of artifacts) {
      const bytes = row.bytes!
      const identity = { artifact: row.artifact, sha256: createHash("sha256").update(bytes).digest("hex") }
      if (view === "current_body") for (const paragraph of readDocxCommentParagraphs(bytes)) {
        const label = paragraph.pids.map(pid => labels.get(pid)).find(Boolean)
        entries.push({ index: entries.length, ...identity, paragraph_index: paragraph.paragraph_index, pids: paragraph.pids, ...(label ? { label } : {}), text: paragraph.paragraph_text })
      }
      if (view === "comments") for (const comment of readDocxComments(bytes)) entries.push({ index: entries.length, ...identity, comment_id: comment.comment_id, markers: comment.markers.map(marker => ({ kind: marker.kind, paragraph_index: marker.paragraph_index, xml_paragraph_index: marker.xml_paragraph_index, pids: marker.pids })), text: comment.text })
      if (view === "deleted_text") for (const text of readDocxRevisionText(bytes).deleted_text) entries.push({ index: entries.length, ...identity, text })
    }
    return { status: "available", document, view, entries }
  } catch {
    return { status: "unavailable", document, view, reason: view === "proposed_body" ? "Pending projection could not be validated or read safely. No proposed-body search or absence conclusion was made." : "Current Word artifact could not be read safely or parsed. No search or absence conclusion was made." }
  }
}

function response(value: Record<string, unknown>, truncated = false) {
  return { title: "Current document evidence", output: JSON.stringify({ ...value, notice: UNTRUSTED_DOCUMENT_NOTICE, ...(value.view === "proposed_body" ? { representation: "Generated pending revision derived from initial paragraphs and current amendments; not actual Word, adopted wording, or source/instruction authority. Initial PIDs and part indices are projection coordinates, not final native Word PIDs." } : {}) }), metadata: { truncated } }
}

/** Consecutive entries include full text, not JSON-array child descriptors. Long entries page without losing characters. */
export function readCurrentDocument(directory: string, args: { document: string; view: CurrentDocumentView; offset?: number; text_offset?: number; limit?: number; count?: number }) {
  const current = currentDocumentView(directory, args.document, args.view)
  if (current.status === "unavailable") return response(current)
  const offset = args.offset ?? 0, textOffset = args.text_offset ?? 0
  if (offset > current.entries.length || (offset === current.entries.length && textOffset)) return response({ status: "unavailable", reason: "Entry offset is outside the selected current view." })
  const entries: Record<string, unknown>[] = []
  let remaining = args.limit ?? 4000, index = offset, cursor = textOffset
  while (index < current.entries.length && entries.length < (args.count ?? 10) && remaining > 0) {
    const entry = current.entries[index]!, points = Array.from(entry.text)
    if (cursor > points.length) return response({ status: "unavailable", reason: "Text offset is outside this entry. Follow the returned next_text_offset." })
    const take = Math.min(remaining, points.length - cursor)
    const page = { ...entry, text: points.slice(cursor, cursor + take).join(""), text_offset: cursor, total_text_characters: points.length }
    // Bound escaped output too, retaining a resumable page even for unusual Unicode/JSON text.
    let length = take
    while (length > 0 && JSON.stringify({ entries: [...entries, page] }).length > 5000) { length = Math.floor(length / 2); page.text = points.slice(cursor, cursor + length).join("") }
    if (!length && take) {
      if (!entries.length) return response({ status: "unavailable", reason: "This entry identity exceeds the bounded response; use its native record pointer." })
      break
    }
    entries.push(page);remaining -= length;cursor += length
    if (cursor < points.length) break
    index++;cursor = 0
  }
  const more = index < current.entries.length
  return response({ status: "read", document: args.document, view: args.view, total_entries: current.entries.length, offset, entries, next_offset: more ? index : null, next_text_offset: more ? cursor : null, scope: "Only this named artifact view; body, comments and deleted text are distinct. Continue with both returned offsets for complete consecutive text." }, more)
}

export function searchCurrentDocument(directory: string, args: { document: string; view: CurrentDocumentView; query: string; k?: number }) {
  const current = currentDocumentView(directory, args.document, args.view)
  if (current.status === "unavailable") return response(current)
  const terms = [...new Set(args.query.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [])]
  const ranked = current.entries.map(entry => ({ entry, score: terms.filter(term => `${entry.label ?? ""} ${entry.text}`.toLowerCase().includes(term)).length })).filter(row => row.score > 0).sort((a, b) => b.score - a.score || a.entry.index - b.entry.index)
  const hits = ranked.slice(0, args.k ?? 5).map(({ entry }) => ({ ...entry, text: Array.from(entry.text).slice(0, 180).join(""), text_complete: Array.from(entry.text).length <= 180, read_offset: entry.index }))
  while (JSON.stringify(hits).length > 4500) hits.pop()
  return response({ status: "searched", document: args.document, view: args.view, method: "lexical", total_entries: current.entries.length, matched_entries: ranked.length, hits, scope: "Lexical matches in this exact current view only. No hits does not prove absence. Read consecutive entries around a hit using read-document with the same view and offset; related documents require their own exact identity." })
}
