import { z } from "zod"
import { checklistTargetAssignments } from "./checklist-target-context"
import { checklistItemSchema, type ChecklistSourceReference } from "./checklist-contract"
import { checklistProducerInputs } from "./checklist-source-context"
import { readLeanJson } from "./state"
import { normalizeWithMap } from "../../lib/quote-verify"
import { readRegularFile } from "../../lib/regular-file"
import path from "node:path"
import { realpathSync } from "node:fs"

const chunkSchema = z.object({ id: z.string().min(1), heading_path: z.array(z.string()), start_char: z.number().int().nonnegative(), end_char: z.number().int().nonnegative() })
const mapSchema = z.object({ documents: z.array(z.object({ document: z.string().min(1), characters: z.number().int().nonnegative(), chunks: z.array(chunkSchema).min(1) })) })
const extractionSchema = z.object({ sources: z.array(z.object({ items: z.array(z.object({ id: z.string(), item: checklistItemSchema })) })) })
export type ChecklistRefinementScope = "source_collection" | "cited_sources" | "target_body"
export type ChecklistRefinementUnit = { id: string; passages: ChecklistSourceReference[]; role: string }
type Range = { id: string; start: number; end: number; headings: string[] }
type Source = { document: string; text: string; ranges: Range[] }
const MAX_UNIT_CHARACTERS = 4500

/** Finite explicit scopes retain original-source coordinates or actual accepted
 * target ranges. Scope preparation creates no source authority or legal finding. */
export function buildChecklistRefinementContext(directory: string, options: { parent_id: string; scope: ChecklistRefinementScope }) {
  const saved = extractionSchema.parse(readLeanJson(directory, "checklist/extraction.json"))
  const matches = saved.sources.flatMap(source => source.items).filter(record => record.id === options.parent_id)
  if (matches.length !== 1) throw Error("Refinement parent must resolve to one preserved extraction item.")
  const parent = matches[0]!.item
  const producer = checklistProducerInputs(directory)
  const request = producer.find(job => job.key === "request")!.input
  const instructions = [...request.passages, ...request.instructions]
  if (options.scope === "target_body") return {
    parent_id: options.parent_id, scope: options.scope,
    units: checklistTargetAssignments(directory, parent).assignments.map(unit => ({ ...unit, role: "Bounded current comparison under the preserved parent.", passages: parent.source_refs })),
    instructions, targets: request.targets, retrieval: "Compare only the explicit current assignment. Retrieve original support and material cross-boundary conditions through exact native source/current views; do not infer coverage from prior reviewer status.",
  }
  const mapped = mapSchema.parse(readMap(directory))
  const sources: Source[] = producer.filter(job => job.key !== "request").map(job => {
    const passage = job.input.passages[0]!
    if (passage.document.startsWith("playbook:")) return { document: passage.document, text: passage.quote, ranges: splitRange(passage.document, passage.quote, 0, passage.quote.length, []) }
    const metadata = mapped.documents.filter(row => row.document === passage.document)
    if (metadata.length !== 1) throw Error(`Original source ${JSON.stringify(passage.document)} has no unique source-map entry.`)
    const document = metadata[0]!, text = passage.quote.slice(0, document.characters)
    if (text.length !== document.characters || (passage.quote !== text && passage.quote !== `${text}\n`)) throw Error(`Canonical text length disagrees with source-map for ${JSON.stringify(passage.document)}.`)
    let covered = 0
    const ranges = document.chunks.flatMap(chunk => {
      if (chunk.start_char > covered || chunk.end_char <= chunk.start_char || chunk.end_char > text.length) throw Error(`Source-map has an unavailable or gapped canonical range for ${JSON.stringify(passage.document)}.`)
      covered = Math.max(covered, chunk.end_char)
      return chunk.end_char - chunk.start_char <= MAX_UNIT_CHARACTERS
        ? [{ id: chunk.id, start: chunk.start_char, end: chunk.end_char, headings: chunk.heading_path }]
        : splitRange(chunk.id, text, chunk.start_char, chunk.end_char, chunk.heading_path)
    })
    if (covered !== text.length) throw Error(`Source-map does not cover the complete canonical source ${JSON.stringify(passage.document)}.`)
    return { document: passage.document, text, ranges }
  })
  const unitIds = sources.flatMap(source => source.ranges.map(range => range.id))
  if (new Set(unitIds).size !== unitIds.length) throw Error("Original source-map unit IDs are not unique.")
  // Request/update/answer authority stays distinct from original source units.
  // A locally cited instruction can itself be the assigned evidence, with its exact locator.
  const selected = new Map<string, { source: Source; range: Range }>()
  if (options.scope === "source_collection") for (const source of sources) for (const range of source.ranges) selected.set(range.id, { source, range })
  if (options.scope === "cited_sources") for (const ref of parent.source_refs) {
    const source = sources.find(row => row.document === ref.document)
    if (!source) {
      const exact = instructions.find(row => row.document === ref.document && quoteRanges(row.quote, ref.quote).length)
      if (!exact) throw Error(`Cited source ${JSON.stringify(ref.document)} is unavailable or its quote is not canonical.`)
      const instructionSource: Source = { document: exact.document, text: exact.quote, ranges: splitRange(exact.locator, exact.quote, 0, exact.quote.length, []) }
      for (const range of instructionSource.ranges) selected.set(range.id, { source: instructionSource, range })
      continue
    }
    const spans = quoteRanges(source.text, ref.quote)
    if (!spans.length) throw Error(`Cited quote is not found in canonical source ${JSON.stringify(ref.document)}.`)
    for (const range of source.ranges) if (spans.some(span => span.start < range.end && span.end > range.start)) selected.set(range.id, { source, range })
  }
  if (!selected.size) throw Error("Refinement scope has no available original source units.")
  const units: ChecklistRefinementUnit[] = [...selected.values()].map(({ source, range }) => ({
    id: range.id,
    passages: [{ document: source.document, locator: `canonical UTF-16 chars ${range.start}-${range.end}${range.headings.length ? `; headings: ${range.headings.join(" > ")}` : ""}`, quote: source.text.slice(range.start, range.end) }, ...boundaries(source, range)],
    role: `Assigned original source unit from ${JSON.stringify(source.document)}${range.headings.length ? ` at headings ${JSON.stringify(range.headings)}` : ""}. The first passage is the assigned unit; adjacent-boundary passages preserve context and exceptions, not independent additional assignments. Source wording determines authority; headings alone do not establish precedence.`,
  }))
  return { parent_id: options.parent_id, scope: options.scope, units, instructions, targets: request.targets, retrieval: "This job refines only its assigned original source unit under the preserved parent and qualifications. Original request and actual included human instructions/answers are separate authority, not inferred approvals. Use targeted search-document/read-document view=source with the exact source name to resolve a definition, exception or cross-reference outside the supplied boundaries. Do not inspect generated documents, prior findings or evaluator material. Source-unit coverage is finite but does not itself establish current compliance." }
}

function readMap(directory: string) {
  const root = realpathSync(directory), file = path.join(directory, "source-map.json")
  const bytes = readRegularFile(file)
  if (!bytes || path.relative(root, realpathSync(path.dirname(file))).startsWith("..")) throw Error("Original source-map is unavailable.")
  return JSON.parse(bytes.toString("utf8"))
}

function quoteRanges(text: string, quote: string) {
  const haystack = normalizeWithMap(text, { unescapeMarkdown: true }), needle = normalizeWithMap(quote).text
  if (!needle) return []
  const spans: { start: number; end: number }[] = []
  for (let index = haystack.text.indexOf(needle); index >= 0; index = haystack.text.indexOf(needle, index + 1)) spans.push({ start: haystack.map[index]!, end: haystack.map[index + needle.length - 1]! + 1 })
  return spans
}

function splitRange(id: string, text: string, start: number, end: number, inherited: string[]): Range[] {
  const ranges: Range[] = []
  let cursor = start
  while (cursor < end) {
    const maximum = Math.min(end, cursor + MAX_UNIT_CHARACTERS)
    let stop = maximum
    if (maximum < end) {
      const paragraph = text.lastIndexOf("\n\n", maximum - 2)
      if (paragraph > cursor) stop = paragraph + 2
      else {
        const spaces = [...text.slice(cursor, maximum).matchAll(/\s+/g)]
        if (!spaces.length) throw Error("An original source contains an indivisible token larger than a bounded source unit.")
        const last = spaces.at(-1)!;stop = cursor + last.index! + last[0].length
      }
    }
    if (stop <= cursor) throw Error("Original source could not be divided into bounded paragraph units.")
    const headings = [...text.slice(start, cursor).matchAll(/^#{1,6}\s+(.+)$/gm)].map(match => match[1]!).slice(-4)
    ranges.push({ id: `${id}@${cursor}:${stop}`, start: cursor, end: stop, headings: headings.length ? headings : inherited });cursor = stop
  }
  return ranges
}

function boundaries(source: Source, range: Range): ChecklistSourceReference[] {
  let before = Math.max(0, range.start - 400), after = Math.min(source.text.length, range.end + 400)
  while (before < range.start && before > 0 && !/\s/.test(source.text[before - 1]!)) before++
  while (after > range.end && after < source.text.length && !/\s/.test(source.text[after]!)) after--
  return [{ start: before, end: range.start }, { start: range.end, end: after }].filter(row => row.end > row.start).map(row => ({ document: source.document, locator: `adjacent boundary context; canonical UTF-16 chars ${row.start}-${row.end}`, quote: source.text.slice(row.start, row.end) }))
}
