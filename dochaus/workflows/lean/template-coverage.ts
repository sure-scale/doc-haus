import type { Amendment } from "../../lib/drafting-amendment"
import type { ParagraphRow } from "./checks"
import { plannedDocument, type DocID, type PlannedPart } from "./plan"
import { excerpt, openItem } from "./quotes"
import type { OpenItem } from "./state"
import { expectedWords, plannedContract, type TemplateContract } from "./template-contract"
import { countWords, headingUnits, tokens, type TemplateUnit } from "./template-structure"

// Template coverage: does a drafted document carry what its coverage contract
// (template-contract.ts) puts in it? Each top-level template part the contract
// gives to the document's own parts is measured in the body. A paragraph headed
// with the template part's wording (a short heading line or a numbered clause
// line) opens its own region, running to the next such heading of a top-level
// part or the next outlined part. A template part with no heading of its own
// shares the text of the outlined parts the contract gives it with the other
// such parts there, in proportion to their template words, so a full clause
// under a renamed heading counts. Text inside square brackets (a placeholder or
// an open-item marker) carries nothing. A part is missing when its region
// carries no text, and thin when it carries less than a quarter of its template
// words. A gap in the rendered draft is an open item of the draft step, which
// the change cycle receives as a required repair; the final body records each
// gap that remains as an open item for the drafting notes, and a part an
// applied amendment took out, or the outline left out, as a recorded removal.
// Nothing here judges the text.

/** A body to measure: its paragraphs, the outlined part each belongs to, and
 * the paragraphs that are outlined headings (the title and each part's). */
export type Body = { paragraphs: string[]; partOf: (number | undefined)[]; headings: Set<number> }
export type UnitMeasure = { unit: TemplateUnit; expected: number; drafted: number; region: number[]; gap?: "missing" | "thin" }

/** Parts shorter than this in the template are not measured for length. */
const MEASURED_PART_WORDS = 40
/** A part drafted in less than 1/THIN_RATIO of its template words is thin. */
const THIN_RATIO = 4
const MAX_HEADING_WORDS = 12

/** How the body carries each top-level template part the contract gives to
 * the document's own parts. */
export function measureContract(contract: TemplateContract, body: Body): UnitMeasure[] {
  const top = new Set(contract.top.map(unit => unit.index))
  const heads = body.paragraphs.map((text, at) => !body.headings.has(at) && headed(text) ? headingUnits(text, contract.units) : new Set<number>())
  const cut = (at: number) => body.headings.has(at) || body.partOf[at] !== body.partOf[at - 1] || [...heads[at]!].some(index => top.has(index))
  const regionFrom = (start: number) => {
    const end = body.paragraphs.findIndex((_, at) => at > start && cut(at))
    return Array.from({ length: (end < 0 ? body.paragraphs.length : end) - start }, (_, offset) => start + offset)
  }
  const words = (region: number[]) => region.reduce((sum, at, offset) =>
    sum + (offset === 0 && countWords(body.paragraphs[at]!) <= MAX_HEADING_WORDS && headed(body.paragraphs[at]!) ? 0 : carriedWords(body.paragraphs[at]!)), 0)
  // Each top-level part's own region: of its headed paragraphs, the longest
  // region, preferring those in the outlined parts the contract gives it.
  const own = new Map(contract.top.flatMap(unit => {
    const place = contract.places.get(unit.index)
    const starts = heads.flatMap((named, at) => named.has(unit.index) ? [at] : [])
    const inPlace = starts.filter(at => place?.kind === "parts" && place.parts.includes(body.partOf[at]!))
    const region = (inPlace.length ? inPlace : starts).map(regionFrom).sort((a, b) => words(b) - words(a))[0]
    return region ? [[unit.index, region] as const] : []
  }))
  const owned = new Set([...own.values()].flat())
  const measured = contract.top.flatMap(unit => {
    const place = contract.places.get(unit.index)
    return place?.kind === "parts" ? [{ unit, parts: place.parts, expected: expectedWords(contract, unit) }] : []
  })
  const pooled = Map.groupBy(measured.filter(row => !own.has(row.unit.index)), row => row.parts.join(","))
  return measured.map((row): UnitMeasure => {
    const region = own.get(row.unit.index)
    if (region) return settle(row.unit, row.expected, words(region), region)
    const share = pooled.get(row.parts.join(","))!.reduce((sum, other) => sum + other.expected, 0)
    const pool = body.paragraphs.flatMap((_, at) => row.parts.includes(body.partOf[at]!) && !body.headings.has(at) && !owned.has(at) ? [at] : [])
    const total = pool.reduce((sum, at) => sum + carriedWords(body.paragraphs[at]!), 0)
    return settle(row.unit, row.expected, share ? Math.round(total * row.expected / share) : total, pool, total)
  })
}

export function contractGaps(contract: TemplateContract, body: Body) {
  return measureContract(contract, body).filter(row => row.gap)
}

/** The body of a rendered draft, or of its final text: the outlined headings
 * are found in the draft's paragraphs in order, and a final paragraph belongs
 * to the part its paragraph pointer had in the draft (a new one, to the part of
 * the paragraph before it). The document's title line, which the draft opens
 * with, is a heading too, so a title that holds a part's name heads no part. */
export function rowsBody(outline: { title: string; parts: PlannedPart[] }, draft: ParagraphRow[], final?: ParagraphRow[]): Body {
  if (draft.length && draft.every(row => row.draft_part !== undefined)) {
    const byPid = new Map(draft.map(row => [row.pid, row.draft_part ?? undefined]))
    const headingPids = new Set(draft.filter(row => row.draft_heading).map(row => row.pid))
    const rows = final ?? draft
    return {
      paragraphs: rows.map(row => row.text),
      partOf: rows.reduce<(number | undefined)[]>((acc, row) => [...acc, byPid.has(row.pid) ? byPid.get(row.pid) : acc.at(-1)], []),
      headings: new Set(rows.flatMap((row, at) => headingPids.has(row.pid) ? [at] : [])),
    }
  }
  const starts = outline.parts.reduce<{ at: number[]; from: number }>((acc, part) => {
    const want = tokens(part.heading).join(" ")
    const found = want ? draft.findIndex((row, at) => at >= acc.from && tokens(row.text).join(" ") === want) : -1
    return found < 0 ? { at: [...acc.at, -1], from: acc.from } : { at: [...acc.at, found], from: found + 1 }
  }, { at: [], from: 0 }).at
  const first = starts.find(at => at >= 0) ?? draft.length
  const title = draft.findIndex((row, at) => at < first && tokens(row.text).join(" ") === tokens(outline.title).join(" "))
  const draftPart = draft.map((_, at) => starts.reduce<number | undefined>((part, start, index) => start >= 0 && start <= at ? index : part, undefined))
  const headingPids = new Set([title, ...starts].filter(at => at >= 0).map(at => draft[at]!.pid))
  const byPid = new Map(draft.map((row, at) => [row.pid, draftPart[at]]))
  const rows = final ?? draft
  return {
    paragraphs: rows.map(row => row.text),
    partOf: rows.reduce<(number | undefined)[]>((acc, row) => [...acc, byPid.has(row.pid) ? byPid.get(row.pid) : acc.at(-1)], []),
    headings: new Set(rows.flatMap((row, at) => headingPids.has(row.pid) ? [at] : [])),
  }
}

/** Open items of the draft step for the template parts its rendered body
 * does not carry as its contract says. The draft is not rewritten for them:
 * the change cycle receives them as required repairs, and the apply step
 * records what the final body still lacks (templateItems). */
export function draftTemplateItems(directory: string, document: DocID, rows: ParagraphRow[]): OpenItem[] {
  const contract = plannedContract(directory, document)
  const outline = plannedDocument(directory, document)
  if (!contract || !outline) return []
  return contractGaps(contract, rowsBody(outline, rows)).map(gap => openItem("draft", "omission", gap.gap === "missing"
    ? `The template part "${gap.unit.label}" of "${contract.source}" is not in the draft.`
    : `The template part "${gap.unit.label}" of "${contract.source}" is much shorter in the draft than in the template (about ${gap.expected} words).`,
  { template_part: gap.unit.label, source_id: contract.source }))
}

/** Open items for the template parts the final body does not carry as its
 * contract says: a part the outline leaves out with a reason, or that an
 * applied amendment took out of the draft, is a recorded removal with its
 * reason; a part left out with no reason, and any other gap, is an omission. */
export function templateItems(
  directory: string, document: DocID, draft: ParagraphRow[], final: ParagraphRow[], amendments: Amendment[], applied: Set<string>,
): OpenItem[] {
  const contract = plannedContract(directory, document)
  const outline = plannedDocument(directory, document)
  if (!contract || !outline) return []
  const ref = (unit: TemplateUnit) => ({ template_part: unit.label, source_id: contract.source })
  const leftOut = [...contract.places].flatMap(([index, place]) => place.kind !== "left_out" ? []
    : place.reason.trim()
      ? [openItem("apply", "template_removal",
        `The template part "${contract.units[index]!.label}" of "${contract.source}" is left out on purpose, as the outline records: ${excerpt(place.reason)}`,
        ref(contract.units[index]!))]
      : [openItem("apply", "omission",
        `The template part "${contract.units[index]!.label}" of "${contract.source}" is left out by the outline, which gives no reason.`,
        ref(contract.units[index]!))])
  const drafted = new Map(measureContract(contract, rowsBody(outline, draft)).map(row => [row.unit.index, row]))
  const gaps = contractGaps(contract, rowsBody(outline, draft, final)).map(gap => {
    const before = drafted.get(gap.unit.index)
    const pids = new Set(before && !before.gap ? before.region.map(at => draft[at]!.pid) : [])
    const cut = amendments.filter(amendment => applied.has(amendment.id) && pids.has(amendment.pid))
    if (cut.length)
      return openItem("apply", "template_removal",
        `The template part "${gap.unit.label}" of "${contract.source}" is ${gap.gap === "missing" ? "taken out" : "cut down"} by ` +
        `${cut.length > 1 ? "amendments" : "amendment"} ${cut.map(amendment => amendment.id).join(", ")}. ` +
        `Reason given: ${cut.map(amendment => `${amendment.id}: ${excerpt(amendment.reason)}`).join("; ")}`,
        ref(gap.unit))
    return openItem("apply", "omission", gap.gap === "missing"
      ? `The template part "${gap.unit.label}" of "${contract.source}" is not in the final text, and the outline does not record it as left out.`
      : `The template part "${gap.unit.label}" of "${contract.source}" is much shorter in the final text than in the template, and the outline does not record what is left out.`,
    ref(gap.unit))
  })
  return [...leftOut, ...gaps]
}

function settle(unit: TemplateUnit, expected: number, drafted: number, region: number[], carried = drafted): UnitMeasure {
  const gap = !carried ? "missing" : expected >= MEASURED_PART_WORDS && drafted * THIN_RATIO < expected ? "thin" : undefined
  return { unit, expected, drafted, region, ...(gap ? { gap } : {}) }
}

/** A paragraph that can head a part: a short line with no closing punctuation,
 * or one that opens with clause numbering. A sentence that mentions a part's
 * name in passing heads nothing. */
function headed(text: string) {
  const line = text.trim()
  return (countWords(line) <= MAX_HEADING_WORDS && !/[.;:,]$/.test(line)) || /^\(?\d+(?:\.\d+)*[.)]?\s+\S/.test(line)
}

/** Words of a paragraph outside square brackets: a placeholder or an
 * open-item marker carries none of the template's text. */
function carriedWords(text: string) {
  return countWords(text.replace(/\[[^\]]*\]/g, " "))
}
