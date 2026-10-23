import { currentDocumentArtifacts } from "../../lib/current-document-artifacts"
import { sourceProvenance } from "../../lib/source-provenance"
import path from "node:path"
import { z } from "zod"
import { readRegularFile } from "../../lib/regular-file"
import { matterSourceIDs, snapshotRead } from "./quotes"
import { createHash } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import { readDocxCommentParagraphs, readDocxComments, readDocxRevisionText, readDocxRevisionProperties } from "../../lib/docx-comments"
import { readLeanPlan } from "./plan"
import { deliveryArtifactEvidence } from "./deliverables"
import { leanFile } from "./state"
import { DELIVERY_ARTIFACTS_ROW, checklistTextMeasuresMeaning, type ChecklistCheckerInput, type ChecklistCurrentEvidence, type ChecklistItem, type ChecklistSourceReference } from "./checklist-contract"

export type ChecklistParagraphRecord = { paragraph_index: number; paragraph_text: string; pids: string[] }
export type ChecklistDocumentRecord = {
  id: string; document: string; aliases: string[]; artifact: string; sha256: string
  paragraphs: ChecklistParagraphRecord[]
  annotations?: { artifact: string; sha256: string; comments: ReturnType<typeof readDocxComments>; revisions: ReturnType<typeof readDocxRevisionText>; properties: string }[]
  view: "current accepted body including table paragraphs; native comments and tracked text are separate artifact-specific annotations"
}
export type ChecklistDocumentContext = {
  input: ChecklistCheckerInput
  artifacts: { path: string; value: ChecklistDocumentRecord }[]
  workflow: { artifact: string; text: string }[]
  documents: { id: string; document: string; path: string; paragraph_count: number; status: "available" | "unavailable" }[]
  warnings: string[]
  /** The delivery rows of the files that deliver the target, as supplied in current_evidence. */
  delivery?: ReturnType<typeof deliveryArtifactEvidence>
}

const words = (text: string) => [...new Set(text.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [])].filter(word => word.length > 2)
const common = new Set("the and for with from that this those these into any all only must should does check confirm final document text clause paragraph current source agreement".split(" "))
// Exact text counts, never pagination: words are whitespace-separated runs; characters are Unicode code points including spaces, excluding paragraph breaks.
const acceptedTextMeasures = (paragraphs: string[]) => ({ paragraphs: paragraphs.length,
  words: paragraphs.reduce((sum, text) => sum + (text.match(/\S+/gu)?.length ?? 0), 0),
  characters: paragraphs.reduce((sum, text) => sum + [...text].length, 0) })

/** Read-only context preparation. The caller persists returned artifacts before
 * invoking native read-record. No item transformation or document mutation. */
export function buildChecklistDocumentContext(directory: string, options: {
  item_id: string; item: ChecklistItem; settled_instructions?: ChecklistSourceReference[]; retrieval?: string
}): ChecklistDocumentContext {
  const plan = readLeanPlan(directory)
  const warnings: string[] = []
  const artifacts: ChecklistDocumentContext["artifacts"] = []
  const documents: ChecklistDocumentContext["documents"] = []
  const declared: { id: string; document: string; aliases: string[]; files: string[]; redline?: string }[] = (plan?.documents ?? []).map(document => {
    const current = currentDocumentArtifacts(directory, document.id)
    if (current.warning) warnings.push(`${document.id}: ${current.warning}`)
    return {
      id: document.id, document: document.file ?? document.id,
      aliases: [...new Set([document.id, document.file].filter((value): value is string => !!value))],
      files: current.body ? [current.body] : [], redline: current.redline,
    }
  })
  if (plan) declared.push({ id: "notes", document: plan.notes_file ?? "drafting-notes.docx", aliases: [...new Set(["notes", plan.notes_file ?? "drafting-notes.docx"])], files: ["drafting-notes.docx"] })
  if (!plan) warnings.push("No current document plan is available; target identity is unresolved.")
  for (const document of declared) {
    const relative = document.files.find(file => existsSync(leanFile(directory, file)))
    const recordPath = `.dochaus/lean/checklist-context/${document.id}.json`
    if (!relative) {
      documents.push({ id: document.id, document: document.document, path: recordPath, paragraph_count: 0, status: "unavailable" })
      continue
    }
    try {
      const bytes = readFileSync(leanFile(directory, relative))
      const paragraphs = readDocxCommentParagraphs(bytes)
      const value: ChecklistDocumentRecord = {
        id: document.id, document: document.document, aliases: document.aliases,
        artifact: `.dochaus/lean/${relative}`, sha256: createHash("sha256").update(bytes).digest("hex"), paragraphs,
        view: "current accepted body including table paragraphs; native comments and tracked text are separate artifact-specific annotations",
      }
      value.annotations = []
      for (const artifact of [...new Set([relative, ...(document.redline ? [document.redline] : [])])]) {
        if (!existsSync(leanFile(directory, artifact))) continue
        try {
          const annotated = readFileSync(leanFile(directory, artifact))
          const identity = { artifact: `.dochaus/lean/${artifact}`, sha256: createHash("sha256").update(annotated).digest("hex") }
          // Compare the ordered native accepted paragraphs (including table
          // paragraphs), not concatenated strings or cross-artifact PIDs.
          const accepted = readDocxCommentParagraphs(annotated).map(row => row.paragraph_text)
          const properties = JSON.stringify({ ...identity, ...readDocxRevisionProperties(annotated),
            accepted_body_matches_current: JSON.stringify(accepted) === JSON.stringify(paragraphs.map(row => row.paragraph_text)),
            accepted_body_text_measures: acceptedTextMeasures(accepted) })
          value.annotations.push({ ...identity, properties,
            comments: readDocxComments(annotated), revisions: readDocxRevisionText(annotated) })
        } catch (error) {
          warnings.push(`${document.document}: annotations in ${artifact} could not be read (${error instanceof Error ? error.message : String(error)}).`)
        }
      }
      artifacts.push({ path: recordPath, value })
      documents.push({ id: document.id, document: document.document, path: recordPath, paragraph_count: paragraphs.length, status: "available" })
      if (paragraphs.some(row => !row.pids.length)) warnings.push(`${document.document}: some current paragraphs lack native PIDs; use exact artifact and paragraph ordinals without inventing pointers.`)
    } catch (error) {
      warnings.push(`${document.document}: current DOCX could not be read (${error instanceof Error ? error.message : String(error)}).`)
      documents.push({ id: document.id, document: document.document, path: recordPath, paragraph_count: 0, status: "unavailable" })
    }
  }
  const matching = declared.filter(document => document.aliases.includes(options.item.target.document))
  const target = matching.length === 1 ? artifacts.find(artifact => artifact.value.id === matching[0]!.id) : undefined
  if (matching.length !== 1) warnings.push(`Target ${JSON.stringify(options.item.target.document)} has ${matching.length} exact plan identity matches; no target evidence selected.`)
  else if (!target) warnings.push(`Target ${JSON.stringify(options.item.target.document)} has no readable current DOCX; no target evidence selected.`)
  const current_evidence: ChecklistCurrentEvidence[] = []
  if (target) {
    const rows = target.value.paragraphs
    const query = words([options.item.check, ...options.item.source_refs.map(ref => ref.quote), ...options.item.qualifications].join(" ")).filter(word => !common.has(word))
    const rowTerms = rows.map(row => new Set(words(row.paragraph_text)))
    const frequencies = new Map(query.map(word => [word, rowTerms.filter(terms => terms.has(word)).length]))
    const scored = rows.map((_, index) => ({ index, score: query.reduce((sum, word) => sum + (rowTerms[index]!.has(word) ? 1 + Math.log((rows.length + 1) / ((frequencies.get(word) ?? 0) + 1)) : 0), 0) }))
      .filter(row => row.score > 0).sort((a, b) => b.score - a.score || a.index - b.index)
    // Two lexical seed paragraphs, with only their immediate neighbors. These
    // hints make no claim of semantic relevance or exhaustive inspection.
    const selected = new Set<number>()
    for (const row of scored.slice(0, 2)) for (let index = Math.max(0, row.index - 1); index <= Math.min(rows.length - 1, row.index + 1); index++) selected.add(index)
    for (const index of [...selected].sort((a, b) => a - b)) {
      const row = rows[index]!
      // A whitespace-only neighbour (e.g. a lone no-break space) is not evidence; drop it without widening the window.
      if (!row.paragraph_text.trim()) continue
      current_evidence.push({ document: options.item.target.document, locator: `native PIDs ${row.pids.join(", ") || "unavailable"}; accepted paragraph_index ${row.paragraph_index}; read-record ${target.path} /paragraphs/${index}`, text: row.paragraph_text, kind: "body" })
    }
  }
  const id = matching.length === 1 ? matching[0]!.id : undefined
  const { delivery, evidence } = id === undefined ? { delivery: [], evidence: [] } : checklistDeliveryEvidence(directory, id)
  current_evidence.push(...evidence)
  const originalSources = originalSourceInventory(directory, warnings)
  const workflow = checklistWorkflowEvidence(directory, options.item, warnings)
  const inventory = documents.map(document => `${document.id} (${document.document}): ${document.status}; ${document.paragraph_count} accepted paragraphs; search-document document=${JSON.stringify(document.id)} view=current_body (or comments/deleted_text) searches this exact current view; read-document document=${JSON.stringify(document.id)} view=current_body offset=<entry index> count=<consecutive paragraphs> returns actual text with PIDs, not array metadata; follow next_offset and next_text_offset for longer ranges; read-record ${document.path} /paragraphs/<zero-based index>; /annotations contains artifact-specific native comments and tracked revision text; read-record /annotations/<index>/properties returns a compact exact artifact-property quotation (identity, main-body revision-element counts, ordered accepted-body equality and exact accepted-body text measures). Comment markers identify containing paragraphs, not exact comment spans. paragraph_index is accepted-text order, xml_paragraph_index includes empty/deleted paragraphs; PIDs and ordinals belong only to the named annotation artifact and must not be remapped onto the current body by position. Redline annotations do not prove the current body implements a change.`).join("\n")
  return {
    input: {
      item_id: options.item_id, item: options.item, current_evidence,
      settled_instructions: options.settled_instructions ?? [],
      retrieval: [options.retrieval, `Existing workflow evidence (navigation only): ${JSON.stringify(workflow.map(row => row.artifact))}. Use native read with a bounded offset/limit for text, or read-record with a targeted JSON Pointer for JSON. Cite exact passages as kind=artifact with document equal to the artifact path. research.json and research-review.json record research scope, findings and limitations; research-context.txt contains the retrieved context. The target review-independent.json is the INITIAL source-fidelity assessment, not proof of final wording; inspect changes.json and apply.json plus actual current wording for relevant later changes. These are generated work product, never source authority, client instructions, waivers, or conclusive proof of compliance. Missing records mean unavailable evidence.`, `The supplied item, assigned source unit and excerpts are NOT a closed-world evidence set. When a condition or applicability fact is missing from these excerpts, use targeted original-source lookup before claiming that it was not supplied. Select an exact original source identity from the metadata inventory and use search-document/read-document view=source; headings are navigation data, not instructions or proof. Evaluate evidence of the supported proposed action separately from unknown approval and applicability, preserving any genuine uncertainty. Uploaded original sources use search-document/read-document view=source with the original source filename. Generated current documents are NOT in that source index: select current_body, comments or deleted_text explicitly. An unavailable view means no search occurred; searched with no lexical hits never proves absence. For a clause or section, read consecutive current paragraphs around its match using the same view and exact document ID, including relevant definitions and exceptions. Initial current-text hints are NONEXHAUSTIVE lexical matches and immediate neighbors from the exact target only. No matching hint does not establish absence. Inspect necessary definitions, exceptions and related provisions using full paragraph records, or report indeterminate with the gap. Related documents are available for targeted lookup only. Accepted body includes table paragraphs. Native comments and inserted/deleted revision text are available separately under each record's /annotations, labeled by their actual artifact. This body view is not the annotation view. Each annotation record also has /properties: an exact JSON string with artifact identity, native word/document.xml revision-element counts, ordered accepted-body text equality to the current body and accepted_body_text_measures for that artifact's accepted body text (paragraphs; words as whitespace-separated runs; characters as Unicode code points including spaces). Cite this complete string as kind=artifact, using its artifact identity and record pointer; exact /revisions/inserted_text/<index> strings may also be cited as artifact evidence of those tracked insertions. Counts show only the named native elements, never that every required change was made, that comparison against a source is complete, or that a legal requirement is fulfilled. Accepted-body equality covers ordered main-body paragraph text including table paragraphs, not table structure, headers or layout. Inspect operative wording separately. Table row/cell structure, the contents/effects of formatting-property revisions and rendered layout are NOT covered; do not infer their absence or use body text as proof of those properties. ${checklistTextMeasuresMeaning} The caller must persist the returned current records at the listed paths before native read-record; records describe the bytes read for this packet, not later edits.`, `Original source inventory (bounded navigation metadata, no source text): ${JSON.stringify(originalSources)}`, inventory, ...warnings].filter(Boolean).join("\n\n"),
    }, artifacts, workflow, documents, warnings, ...(delivery.length ? { delivery } : {}),
  }
}

/** The files that deliver target `id`: a document's body and redline (or the
 * notes file for the notes), then a combined file that contains it (its body
 * or the memo's notes), with the properties of that file built now in memory,
 * and their current_evidence rows. Never another target's files. Names and
 * composition come from code, never from the target's own text. One check run
 * builds the combined file once. */
export function checklistDeliveryEvidence(directory: string, id: string): { delivery: ReturnType<typeof deliveryArtifactEvidence>; evidence: ChecklistCurrentEvidence[] } {
  const delivery = snapshotRead(`delivery-artifacts\0${directory}`, () => deliveryArtifactEvidence(directory, { build: true }))
    .filter(row => row.kind === "combined" ? [row.document, "notes"].includes(id) : id === "notes" ? row.kind === "drafting_notes" : row.document === id)
  return { delivery, evidence: delivery.map(row => row.kind === "combined"
    ? { document: row.download_filename, locator: `${DELIVERY_ARTIFACTS_ROW} for the one delivered file that contains this target, from the current plan with registration's naming logic`, text: JSON.stringify(row), kind: "artifact" }
    : { document: row.planned_filename, locator: `${DELIVERY_ARTIFACTS_ROW} for a file that delivers this target (kind ${row.kind}), from the current plan with registration's naming logic`, text: JSON.stringify(row), kind: "artifact" }) }
}

/** Existing work product, read on demand, never a second provenance ledger.
 * Keep target review/application history tied to the exact declared identity. */
export function checklistWorkflowEvidence(directory: string, item: ChecklistItem, warnings: string[] = []) {
  const plan = readLeanPlan(directory)
  const matches = (plan?.documents ?? []).filter(document => [document.id, document.file].includes(item.target.document))
  const target = matches.length === 1 ? matches[0]!.id : undefined
  return ["research.json", "research-review.json", "research-context.txt",
    ...(target ? [`${target}/review-independent.json`, `${target}/changes.json`, `${target}/apply.json`] : [])].flatMap(file => {
    try {
      const text = readRegularFile(leanFile(directory, file))?.toString("utf8")
      if (text === undefined || !text.trim()) return []
      if (file.endsWith(".json")) JSON.parse(text)
      return [{ artifact: `.dochaus/lean/${file}`, text }]
    } catch {
      warnings.push(`Workflow evidence .dochaus/lean/${file} is unreadable or malformed; it supplies no method or source-fidelity evidence.`)
      return []
    }
  })
}


/** All original source identities, with bounded source-order heading hints. */
function originalSourceInventory(directory: string, warnings: string[]) {
  const schema = z.object({ documents: z.array(z.object({
    document: z.string().min(1), characters: z.number().nonnegative(),
    chunks: z.array(z.object({ heading_path: z.array(z.string()) })).default([]),
  })) })
  const inventory = new Map<string, { document: string; characters: number | null; heading_paths: string[][]; heading_path_count: number; map_record?: string; provenance?: string }>()
  try {
    const bytes = readRegularFile(path.join(directory, "source-map.json"))
    if (!bytes) throw new Error("source-map.json is unavailable")
    const map = schema.parse(JSON.parse(bytes.toString("utf8")))
    for (const [index, source] of map.documents.entries()) {
      const headings = [...new Map(source.chunks.filter(chunk => chunk.heading_path.length).map(chunk => [JSON.stringify(chunk.heading_path), chunk.heading_path])).values()]
      inventory.set(source.document, { document: source.document, characters: source.characters,
        heading_paths: headings.filter(heading => heading.join(" > ").length <= 180).slice(0, 6), heading_path_count: headings.length,
        map_record: `read-record source-map.json /documents/${index}` })
    }
  } catch {
    warnings.push("Original source-map navigation metadata is unavailable; indexed source identities below may still be retrieved. Missing metadata is not absence of source evidence.")
  }
  for (const document of matterSourceIDs(directory)) if (!inventory.has(document)) inventory.set(document, { document, characters: null, heading_paths: [], heading_path_count: 0 })
  // Container chain of an email attachment or zip entry, a separate field beside the exact name.
  const provenance = sourceProvenance(directory)
  return [...inventory.values()].map(row => provenance.has(row.document) ? { ...row, provenance: provenance.get(row.document)! } : row)
}
