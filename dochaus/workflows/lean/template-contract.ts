import { readLeanPlan, type DocID, type LeftOutPart, type PlannedDocument } from "./plan"
import { matterSourceIDs } from "./quotes"
import { headingUnits, parseTemplateUnits, phraseNames, sourceText, type TemplateUnit } from "./template-structure"

// The coverage contract of a document drafted from a template: for each
// top-level part of the template, where the outline puts it. The document's own
// outline holds a part strongly in the outlined parts whose "## " heading names
// it, and weakly in the parts whose scope line names it; its title places no
// part. A part goes to the document's own strong place; else to another
// planned document drafted from the same template that holds it strongly; else
// to the document's own weak place; else to another such document that holds
// it weakly. So a part placed in another document is always in that document's own
// parts. A "Left out:" line leaves a part (or a part inside one) out on purpose,
// with its reason. A part the outline does not account for, or leaves out while
// a heading also names it, is retry feedback at outline admission; when
// admission ends with one still unaccounted, it is held by the part that holds
// the template part before it, so the draft is still measured for it. The
// contract is derived from the admitted plan and the template's indexed text
// only, so every step reads the same contract. Nothing here knows what a part
// is called.

export type DocumentTemplate = { source: string; units: TemplateUnit[] }

export type TemplatePlace =
  | { kind: "parts"; parts: number[]; assumed?: true }
  | { kind: "document"; title: string }
  | { kind: "left_out"; line: string; reason: string }

export type TemplateContract = DocumentTemplate & {
  /** The template's top-level parts, in order. */
  top: TemplateUnit[]
  /** By unit index: every top-level part, and any inner part left out. */
  places: Map<number, TemplatePlace>
  /** Top-level parts the outline does not account for. */
  unaccounted: TemplateUnit[]
  /** "Left out:" lines that name no part of the template. */
  strays: LeftOutPart[]
  /** Parts a "Left out:" line leaves out that a heading of the plan also
   * names: by the outlined part's heading, or by the title of the other
   * document whose heading names it. */
  contradicted: { unit: TemplateUnit; line: string; by: string }[]
}

type OutlinedDocument = Omit<PlannedDocument, "id">

const MAX_LISTED = 12

type TemplateEntry = { template?: string; comparison?: string; mode?: "new" | "revise" }

/** The name a planned document's template is drawn from: a revision's
 * original, else its "Template:" line. */
export function templateSource(entry: TemplateEntry | undefined) {
  return entry?.mode === "revise" ? entry.comparison ?? entry.template : entry?.template
}

/** The template a planned document is drafted from: the matter document its
 * "Template:" line names, when that document has headed parts. */
export function documentTemplate(directory: string, entry: TemplateEntry | undefined): DocumentTemplate | undefined {
  const source = templateSource(entry)
  if (!source || !matterSourceIDs(directory).includes(source)) return
  const units = parseTemplateUnits(sourceText(directory, source))
  return topUnits(units).length ? { source, units } : undefined
}

/** The contract of the document at `at` among the outlined documents. */
export function templateContract(directory: string, documents: OutlinedDocument[], at: number): TemplateContract | undefined {
  const document = documents[at]
  const template = documentTemplate(directory, document)
  if (!document || !template) return
  const units = template.units
  const top = topUnits(units)
  const lines = (document.left_out ?? []).map(line => ({ line, named: headingUnits(line.part, units) }))
  const leftOut = new Map(lines.flatMap(row => [...row.named].map(index => [index, row.line] as const)))
  const siblings = document.mode === "revise" ? [] : documents.filter((other, index) => index !== at && other.template === template.source)
  const elsewhere = (unit: TemplateUnit, strong: boolean) => siblings.find(other => {
    const place = ownPlace(other, unit, units)
    return place?.strong === strong && !(other.left_out ?? []).some(line => headingUnits(line.part, units).has(unit.index))
  })
  const placed = top.reduce<{ places: [number, TemplatePlace][]; unaccounted: TemplateUnit[]; last: number[] }>((acc, unit) => {
    const line = leftOut.get(unit.index)
    if (line) return { ...acc, places: [...acc.places, [unit.index, { kind: "left_out", line: line.part, reason: line.reason }]] }
    const own = ownPlace(document, unit, units)
    const strong = own?.strong ? undefined : elsewhere(unit, true)
    const weak = own || strong ? undefined : elsewhere(unit, false)
    const place: TemplatePlace | undefined = own?.strong ? { kind: "parts", parts: own.parts }
      : strong ? { kind: "document", title: strong.title }
      : own ? { kind: "parts", parts: own.parts }
      : weak ? { kind: "document", title: weak.title }
      : undefined
    if (!place) return { places: [...acc.places, [unit.index, { kind: "parts", parts: acc.last, assumed: true }]], unaccounted: [...acc.unaccounted, unit], last: acc.last }
    return { ...acc, places: [...acc.places, [unit.index, place]], last: place.kind === "parts" ? [place.parts.at(-1)!] : acc.last }
  }, { places: [], unaccounted: [], last: [0] })
  const inner = units.filter(unit => !top.includes(unit) && leftOut.has(unit.index))
    .map(unit => [unit.index, { kind: "left_out", line: leftOut.get(unit.index)!.part, reason: leftOut.get(unit.index)!.reason }] as [number, TemplatePlace])
  const contradicted = units.flatMap(unit => {
    const line = leftOut.get(unit.index)
    const by = line && (document.parts.find(part => headingUnits(part.heading, units).has(unit.index))?.heading ??
      siblings.find(other => ownPlace(other, unit, units)?.strong)?.title)
    return line && by ? [{ unit, line: line.part, by }] : []
  })
  return {
    ...template, top,
    places: new Map([...placed.places, ...inner]),
    unaccounted: placed.unaccounted,
    strays: lines.filter(row => !row.named.size).map(row => row.line),
    contradicted,
  }
}

/** Where a document's own outline puts a template part: the parts whose
 * heading names it (a strong place); else the parts whose scope line names it
 * (a weak place). The title places no part: a title that rewords the
 * template's own title, or one naming a part the document then drafts with
 * other text, would pool the part over the whole document and hide a part
 * drafted nowhere. Unplaced, the part is outline feedback and then measured. */
function ownPlace(document: OutlinedDocument, unit: TemplateUnit, units: TemplateUnit[]) {
  const naming = (names: (part: PlannedDocument["parts"][number]) => boolean) => document.parts.flatMap((part, index) => names(part) ? [index] : [])
  const headed = naming(part => headingUnits(part.heading, units).has(unit.index))
  if (headed.length) return { parts: headed, strong: true }
  const scoped = naming(part => phraseNames(`${part.heading} ${part.scope}`, unit))
  return scoped.length ? { parts: scoped, strong: false } : undefined
}

/** The contract of a planned document, from the saved plan. */
export function plannedContract(directory: string, document: DocID) {
  const documents = readLeanPlan(directory)?.documents ?? []
  const at = documents.findIndex(row => row.id === document)
  return at < 0 ? undefined : templateContract(directory, documents, at)
}

/** Retry feedback for the outline: each "Template:" line names a matter
 * document, each "Left out:" line names a part of the template and gives a
 * reason, and every top-level part of the template is accounted for. */
export function outlineCoverageIssues(directory: string, documents: OutlinedDocument[]) {
  const sources = matterSourceIDs(directory)
  return documents.flatMap((document, at) => {
    const invalid = document.template && !sources.includes(document.template)
      ? [`The "Template:" line of "${document.title}" names "${document.template}", which is not a matter document. Use the exact name of the matter document it follows, or leave the line out.`]
      : []
    const contract = templateContract(directory, documents, at)
    if (!contract) return invalid
    const reasonless = [...contract.places.values()].filter(place => place.kind === "left_out" && !place.reason)
    return [
      ...invalid,
      ...contract.strays.map(line =>
        `The "Left out:" line "${line.part}" of "${document.title}" names no part of its template "${contract.source}". Name the template part by its heading as the template writes it.`),
      ...[...new Set(reasonless.map(place => place.kind === "left_out" ? place.line : ""))].map(line =>
        `The "Left out:" line "${line}" of "${document.title}" gives no reason. Say after the part's name why it is left out.`),
      ...contract.contradicted.map(row =>
        `The "Left out:" line "${row.line}" of "${document.title}" leaves out the template part "${row.unit.label}", which "${row.by}" also names. ` +
        "Keep the \"Left out:\" line only if no part carries it; otherwise drop that line."),
      ...(contract.unaccounted.length
        ? [`The outline of "${document.title}" does not account for these parts of its template "${contract.source}": ${quoted(contract.unaccounted.slice(0, MAX_LISTED))}` +
          `${contract.unaccounted.length > MAX_LISTED ? ` and ${contract.unaccounted.length - MAX_LISTED} more` : ""}. ` +
          "Name each in the \"## \" heading or scope line of the part that carries it, or, for a part left out on purpose, " +
          "add a line \"Left out: \"<template part>\" – <reason>\" under the title."]
        : []),
    ]
  })
}

/** For each outlined part, the template parts it carries and the parts inside
 * them the outline leaves out. */
export function partCarries(contract: TemplateContract, parts: number) {
  return Array.from({ length: parts }, (_, index) => {
    const carried = contract.top.filter(unit => {
      const place = contract.places.get(unit.index)
      return place?.kind === "parts" && place.parts.includes(index)
    })
    const left = contract.units.filter(unit => !contract.top.includes(unit) && contract.places.get(unit.index)?.kind === "left_out" &&
      carried.some(top => ancestors(contract.units, unit).includes(top)))
    return { carried, left }
  })
}

/** The words a carried part is expected to hold: its template words less those
 * of the parts inside it that the outline leaves out. */
export function expectedWords(contract: TemplateContract, unit: TemplateUnit) {
  const left = contract.units.filter(inner => contract.places.get(inner.index)?.kind === "left_out" && ancestors(contract.units, inner).includes(unit) &&
    !ancestors(contract.units, inner).some(outer => outer !== unit && contract.places.get(outer.index)?.kind === "left_out"))
  return Math.max(0, unit.words - left.reduce((sum, inner) => sum + inner.words, 0))
}

export function topUnits(units: TemplateUnit[]) {
  return units.filter(unit => !unit.title && (unit.parent === undefined || units[unit.parent]!.title))
}

export function ancestors(units: TemplateUnit[], unit: TemplateUnit): TemplateUnit[] {
  const parent = unit.parent === undefined ? undefined : units[unit.parent]
  return parent ? [parent, ...ancestors(units, parent)] : []
}

export function quoted(units: TemplateUnit[]) {
  return units.map(unit => `"${unit.label}"`).join("; ")
}
