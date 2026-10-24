import { DOCUMENT_FORMS } from "../../lib/page-furniture"
import { finalBodyParagraphs, shownText, type ParagraphRow } from "./checks"
import { nameKey, namedIn, readLeanPlan, type DocID } from "./plan"
import { excerpt, openItem } from "./quotes"
import type { OpenItem } from "./state"
import { sourceText } from "./template-structure"

// Cross-references and defined terms of a final body, read by code after a
// document's changes are applied. No model is asked and nothing is rewritten:
// each problem is an open item for the drafting notes, and a correction goes
// through the normal change path. Few, precise findings: where the text does
// not settle a point, nothing is reported.
//
// References. A document's clause numbering is read from the numbers printed
// at the start of its paragraphs ("9.1 …", "9 Liability", "Clause 9 – …")
// before its first schedule, annex, appendix, exhibit or attachment heading.
// An unnumbered heading among the clauses stands for a clause whose number is
// not printed, so a gap in the printed numbers that unnumbered headings
// account for is not a missing clause. Reported: a reference to a clause,
// sub-clause or attachment the final text does not have; a reference carried
// over from the text before the changes (the original being revised, and this
// run's first draft) whose clause was renumbered or removed under it, so that
// it now points at another clause; and a number printed twice that something
// refers to. A reference "of" another named instrument, or placed after a
// pack document's title, is resolved in that pack document when one has that
// title and is otherwise not this document's to check, nor is a statutory
// citation. Text in square brackets (a placeholder or an open-item marker)
// refers to nothing. A paragraph numbered by the document's clause list (Word
// list numbering) is read with the label the list gives it, as a reader sees
// it; an ordinary numbered list inside a clause numbers no clauses, so its
// labels are not read. A document with no clause numbers has its clause
// references left unchecked, and so has advice or a letter, which cites the
// clauses of the documents it discusses. An original read from its source
// text carries the list labels ingest's extraction writes into it as text. A
// document planned without a known form is not checked at all.
//
// Defined terms. A term is defined by a quoted or "the"-introduced parenthetical
// ("… Limited (the "Customer")", "(the Supplier)") or by a "means" definition.
// Reported: a term defined twice in ways that cannot both hold; used in a mixed
// case ("Confidential information") mid-sentence; used after the changes
// removed its definition; and defined differently by a later pack document than
// by an earlier one (another company, or other figures). In a document that
// works by defined terms (at least MIN_DEFINITIONS of them) and is not advice or
// a letter, also: a term defined but not used (or used only in lower case), and
// a capitalised phrase used repeatedly as a defined term ("the Effective
// Date") that the pack defines nowhere. Plural and possessive forms are uses of
// the term.

/** One paragraph of a document's text; the final body's carry a pointer. */
export type Para = { pid?: string; text: string }
/** Another document of the pack, by its title; `earlier` documents are compared for definitions. */
export type PackDocument = { title: string; paras: Para[]; earlier: boolean }
export type ConsistencyFinding = { pid?: string; detail: string }

/** Open items for the cross-references and defined terms of one document's
 * final body. The baselines are the original it revises, when the outline
 * names one, and its drafted body; the other pack documents are those whose
 * final bodies exist. */
export function consistencyItems(directory: string, document: DocID, draft: ParagraphRow[], final: ParagraphRow[]): OpenItem[] {
  const plan = readLeanPlan(directory)
  const position = plan?.documents.findIndex(row => row.id === document) ?? -1
  const entry = plan?.documents[position]
  if (!plan || !entry) return []
  const original = entry.comparison ? sourceLines(sourceText(directory, entry.comparison)) : []
  const shown = (rows: (ParagraphRow & { clause?: true })[]) => rows.map(row => ({ pid: row.pid, text: row.clause ? shownText(row) : row.text }))
  const others = plan.documents.flatMap((row, at): PackDocument[] => {
    if (row.id === document) return []
    const paras = shown(finalBodyParagraphs(directory, row.id))
    return paras.length ? [{ title: row.title, paras, earlier: at < position }] : []
  })
  return consistencyFindings({ title: entry.title, form: entry.form, final: shown(final), baselines: [original, shown(draft)].filter(rows => rows.length), others })
    .map(finding => openItem("apply", "consistency", finding.detail, finding.pid ? { pid: finding.pid } : {}))
}

export function consistencyFindings(input: { title: string; form?: string; final: Para[]; baselines: Para[][]; others: PackDocument[] }): ConsistencyFinding[] {
  // A document planned without a known form (a legacy plan, or an outline's
  // last attempt) may be a letter or advice as easily as an agreement; it is
  // not checked rather than checked as the wrong kind of document.
  if (!DOCUMENT_FORMS.some(form => form === input.form)) return []
  const final = outline(input.final)
  const baselines = input.baselines.map(rows => outline(rows))
  const pack = input.others.map(other => ({ ...other, outline: outline(other.paras) }))
  // Advice and letters point at clauses of the documents they discuss and name
  // things without defining them; only references to a named pack document,
  // and definitions they do make, are theirs to check.
  const advisory = input.form === "advice" || input.form === "letter" || input.form === "advice letter"
  const findings = [...referenceFindings(input.title, final, baselines, pack, advisory), ...termFindings(input.title, final, baselines, pack, advisory)]
  return [...new Map(findings.map(finding => [`${finding.pid ?? ""}\0${finding.detail}`, finding])).values()]
}

function sourceLines(text: string): Para[] {
  return text.split(/\r?\n/).map(line => line.trim()).filter(Boolean).map(text => ({ text }))
}

// ---------------------------------------------------------------- structure

type Kind = "clause" | "section" | "article" | "schedule" | "annex" | "appendix" | "exhibit" | "attachment"
const CLAUSE_KINDS = new Set<Kind>(["clause", "section", "article"])

type Outline = {
  paras: Para[]
  /** Visible text of each paragraph: square-bracketed text blanked. */
  texts: string[]
  /** Paragraphs before this index are the clauses; attachments follow. */
  end: number
  labelAt: (string | undefined)[]
  printed: Map<string, number>
  duplicates: { label: string; at: number }[]
  /** Top-level clause numbers: printed, or standing behind unnumbered headings. */
  tops: Set<string>
  numbered: boolean
  /** Keywords a clause heading is written with ("section" for "Section 4 – …"). */
  keywords: Set<string>
  /** Attachment headings by "kind label" ("schedule 1"), and the one each paragraph is under. */
  attachments: Map<string, number>
  attachmentAt: (string | undefined)[]
  heading: Map<string, string>
  content: Map<string, string>
}

const MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December"
const DATE_START = new RegExp(`^\\d{1,2}\\s+(?:${MONTHS})\\b`)
const NUMBERED = /^(\d{1,3}(?:\.\d{1,3}){0,3})\.?\s+(?=[\p{Lu}(\[“"‘'])/u
const KEYED = /^(clause|section|article)\s+(\d{1,3}(?:\.\d{1,3}){0,3})\.?(.*)$/i
const ATTACHMENT_HEAD = /^(schedule|annex|appendix|exhibit|attachment)\s+(\d{1,3}|[A-Z]|[IVX]{1,5})(?![\w])(.*)$/i
const MAX_HEADING_WORDS = 12

function outline(paras: Para[]): Outline {
  const texts = paras.map(para => visible(para.text))
  const heads = texts.map(attachmentHead)
  const first = heads.findIndex(Boolean)
  const end = first < 0 ? paras.length : first
  // Labels are read as printed: "4.2 [Amount]" is clause 4.2 with a placeholder.
  const labels = paras.map((para, at) => at < end ? clauseLabel(para.text.trim()) : undefined)
  const labelAt = labels.map(row => row?.label)
  const printed = new Map<string, number>()
  const duplicates: Outline["duplicates"] = []
  labelAt.forEach((label, at) => {
    if (!label) return
    if (printed.has(label)) duplicates.push({ label, at })
    else printed.set(label, at)
  })
  const order = [...new Set(labelAt.flatMap(label => label ? [top(label)] : []))]
  const numbered = printed.size >= 3 && order.length >= 2 && order.every((value, at) => at === 0 || Number(value) > Number(order[at - 1]))
  const firstAt = (number: string) => labelAt.findIndex(label => label !== undefined && top(label) === number)
  const lastAt = (number: string) => labelAt.findLastIndex(label => label !== undefined && top(label) === number)
  const plain = (at: number) => !labelAt[at] && !heads[at] && unnumberedHeading(texts[at]!)
  const headingsBetween = (from: number, to: number) => texts.reduce((sum, _, at) => sum + (at > from && at < to && plain(at) ? 1 : 0), 0)
  // A clause whose first printed number is a sub-clause's ("9.1") has its
  // heading, unnumbered, above it.
  const own = (number: string) => labelAt[firstAt(number)]!.includes(".") ? 1 : 0
  const tops = new Set(order)
  const presume = (from: number, to: number, available: number) => {
    if (to >= from && available >= to - from + 1) for (let number = from; number <= to; number += 1) tops.add(String(number))
  }
  if (numbered) {
    const numbers = order.map(Number)
    presume(1, numbers[0]! - 1, headingsBetween(-1, firstAt(order[0]!)) - own(order[0]!))
    numbers.slice(1).forEach((number, at) =>
      presume(numbers[at]! + 1, number - 1, headingsBetween(lastAt(order[at]!), firstAt(order[at + 1]!)) - own(order[at + 1]!)))
    const trailing = headingsBetween(lastAt(order.at(-1)!), end)
    presume(numbers.at(-1)! + 1, numbers.at(-1)! + trailing, trailing)
  }
  const heading = new Map<string, string>()
  order.forEach((number, at) => {
    const start = firstAt(number)
    if (labelAt[start] === number) return heading.set(number, labels[start]!.rest)
    const floor = at ? lastAt(order[at - 1]!) : -1
    for (let before = start - 1; before > floor; before -= 1) if (plain(before)) return heading.set(number, texts[before]!)
  })
  const content = new Map<string, string>()
  labels.forEach(row => row && row.label.includes(".") && !content.has(row.label) && content.set(row.label, row.rest))
  const attachments = new Map<string, number>()
  heads.forEach((head, at) => head && !attachments.has(head) && attachments.set(head, at))
  const attachmentAt = heads.reduce<(string | undefined)[]>((acc, head) => [...acc, head ?? acc.at(-1)], [])
  const keywords = new Set(paras.slice(0, end).flatMap(para => {
    const text = para.text.trim()
    const keyed = KEYED.exec(text)
    return keyed && clauseLabel(text) ? [keyed[1]!.toLowerCase()] : []
  }))
  return { paras, texts, end, labelAt, printed, duplicates, tops, numbered, keywords, attachments, attachmentAt, heading, content }
}

function visible(text: string) {
  return text.replace(/\[[^\]]*\]/g, match => " ".repeat(match.length)).trim()
}

function top(label: string) {
  return label.split(".")[0]!
}

/** The clause number printed at the start of a paragraph, and the text after it. */
function clauseLabel(text: string): { label: string; rest: string } | undefined {
  const bare = NUMBERED.exec(text)
  if (bare && !DATE_START.test(text)) return { label: bare[1]!, rest: text.slice(bare[0].length).trim() }
  const keyed = KEYED.exec(text)
  if (keyed && headed(keyed[3]!, text)) return { label: keyed[2]!, rest: keyed[3]!.replace(/^\s*[-–—:.]?\s*/, "").trim() }
}

/** "schedule 1" for a schedule's (annex's …) heading paragraph. */
function attachmentHead(text: string) {
  const head = ATTACHMENT_HEAD.exec(text)
  return head && headed(head[3]!, text) ? `${head[1]!.toLowerCase()} ${head[2]!.toUpperCase()}` : undefined
}

/** A keyword label heads its paragraph when nothing follows it, a separator
 * does, or a short title does: "Clause 9 applies …" is a reference, not a heading. */
function headed(rest: string, text: string) {
  return !rest.trim() || /^\s*[-–—:.]/.test(rest) || (/^\s+\p{Lu}/u.test(rest) && headingLine(text))
}

function headingLine(text: string) {
  return words(text) <= MAX_HEADING_WORDS && !/[.;:,!?]$/.test(text.trim())
}

function unnumberedHeading(text: string) {
  return /^\p{Lu}/u.test(text) && headingLine(text)
}

function words(text: string) {
  return text.split(/\s+/).filter(Boolean).length
}

function key(text: string) {
  return text.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim()
}

// --------------------------------------------------------------- references

type Reference = { kind: Kind; labels: string[]; at: number; index: number; target?: string }

const REFERENCE_WORD = /\b(clauses?|sub-?clauses?|sections?|articles?|schedules?|annex(?:es)?|appendix|appendices|exhibits?|attachments?)\s+/gi
const CLAUSE_NUMBER = /^(\d{1,3}(?:\.\d{1,3}){0,3})(?![\d])((?:\s?\([a-zA-Z0-9]{1,5}\))*)/
const ATTACHMENT_NUMBER = /^(\d{1,3}|[A-Z]|[IVX]{1,5})(?![\w])/
const CHAIN = /^(?:\s*,\s*(?:and\s+|or\s+)?|\s+(?:and|or|to|through|until)\s+|\s*[-–—]\s*)((?:clauses?|sub-?clauses?|sections?|articles?|schedules?|annex(?:es)?|appendix|appendices|exhibits?|attachments?)\s+)?/i
const LINKING_WORDS = new Set(["and", "or", "to", "through", "until", "inclusive", "above", "below", "of", "shall", "will", "apply", "applies", "is", "are", "has", "have", "may", "must", "which", "that", "as", "in", "on", "for", "by", "with", "where", "if", "does", "do", "respectively", "nor", "but", "not", "unless", "hereof", "hereto", "herein"])
const OTHER_INSTRUMENT = /^\s*(of|to|in|under)\s+(?:the\s+)?(\p{Lu}[\p{L}\p{N}’'&.-]*(?:\s+(?:(?:of|and|for|the|on|in|to|&)\s+)?(?:\p{Lu}[\p{L}\p{N}’'&.-]*|\d{4}))*)/u
/** A caption after a cited number: "clause 4 (Payment)". */
const CAPTION = /^\s*\([^()]{2,60}\)/
/** An abbreviated statute right after the number: "section 5 CA 2006", "Article 6 GDPR". */
const ACRONYM_AFTER = /^\s*\p{Lu}{2,}[\p{Lu}\d]*\b/u
const AUTHORITY_BEFORE = /(?:\b(?:Act|Code|Regulations?|Rules?|Directive|Convention|Treaty|Order|Law|Statute|Constitution|GDPR|U\.?S\.?C\.?|C\.?F\.?R\.?)(?:\s+\d{4})?|§)[\s,]*$/
const TITLE_CONNECTORS = new Set(["of", "and", "for", "the", "on", "in", "to"])

function kindOf(word: string): Kind {
  const lower = word.toLowerCase().replace(/^sub-?/, "")
  if (lower.startsWith("clause")) return "clause"
  if (lower.startsWith("section")) return "section"
  if (lower.startsWith("article")) return "article"
  if (lower.startsWith("schedule")) return "schedule"
  if (lower.startsWith("annex")) return "annex"
  if (lower.startsWith("append")) return "appendix"
  if (lower.startsWith("exhibit")) return "exhibit"
  return "attachment"
}

/** "SPA" for "Share Purchase Agreement". */
function initials(title: string) {
  return key(title).split(" ").filter(word => word && !TITLE_CONNECTORS.has(word)).map(word => word[0]).join("").toUpperCase()
}

/** Whether a name cited for a reference is this title: its words, or its initials. */
function names(title: string, name: string) {
  const own = new Set(key(title).split(" "))
  return key(name).split(" ").every(word => own.has(word) || TITLE_CONNECTORS.has(word)) || (name === name.toUpperCase() && name.length >= 2 && initials(title) === name)
}

/** Every reference in a document's visible text. A reference "of"/"to"/"in"/
 * "under" another named instrument, or placed just after a pack document's
 * title, carries that name as its target; one to statute is left out, and so
 * is any reference to a number once cited "of" another instrument. A label
 * heading its own paragraph refers to nothing. */
function references(document: Outline, title: string, titles: string[]): Reference[] {
  const external = new Set<string>()
  const caption = (text: string) => text.replace(CAPTION, "")
  const found = document.texts.flatMap((text, at) => [...text.matchAll(REFERENCE_WORD)].flatMap((match): (Reference & { external?: boolean })[] => {
    if (match.index === 0 && (clauseLabel(text) || attachmentHead(text))) return []
    const kind = kindOf(match[1]!)
    const number = CLAUSE_KINDS.has(kind) ? CLAUSE_NUMBER : ATTACHMENT_NUMBER
    let rest = text.slice(match.index + match[0].length)
    const first = number.exec(rest)
    if (!first) return []
    const labels = [CLAUSE_KINDS.has(kind) ? first[1]! : first[1]!.toUpperCase()]
    rest = rest.slice(first[0].length)
    const bare = rest
    rest = caption(rest)
    for (;;) {
      const link = CHAIN.exec(rest)
      if (!link || (link[1] && kindOf(link[1].trim()) !== kind)) break
      const next = number.exec(rest.slice(link[0].length))
      if (!next) break
      const after = caption(rest.slice(link[0].length + next[0].length))
      const word = /^\s+([a-z]+)\b/.exec(after)?.[1]
      if (word && !LINKING_WORDS.has(word)) break
      labels.push(CLAUSE_KINDS.has(kind) ? next[1]! : next[1]!.toUpperCase())
      rest = after
    }
    const before = text.slice(Math.max(0, match.index - 60), match.index)
    const reference: Reference = { kind, labels, at, index: match.index }
    if (AUTHORITY_BEFORE.test(before) || (labels.length === 1 && ACRONYM_AFTER.test(bare))) return [{ ...reference, external: true }]
    const lead = key(before).replace(/\s+(?:under|in|pursuant to|at|see)$/, "")
    const preceding = titles.find(name => words(key(name)) >= 2 && (lead === key(name) || lead.endsWith(` ${key(name)}`)))
    if (preceding) return [{ ...reference, target: preceding }]
    const instrument = OTHER_INSTRUMENT.exec(rest)
    if (instrument && !names(title, instrument[2]!)) return [{ ...reference, target: instrument[2]!, external: instrument[1]!.toLowerCase() === "of" }]
    return [reference]
  }))
  // A number cited as statute's, or as another instrument's, is not this
  // document's wherever it recurs.
  found.forEach(row => row.external && row.labels.forEach(label => external.add(`${row.kind} ${label}`)))
  return found.filter(row => row.target || !row.external).map(({ external: _external, ...row }) => row.target ? row : { ...row, labels: row.labels.filter(label => !external.has(`${row.kind} ${label}`)) })
    .filter(row => row.labels.length)
}

type State = "present" | "missing" | "unknown"

/** Whether a numbered document has a clause: its top-level number must be
 * there; each deeper number must be printed where its siblings are. */
function clauseState(document: Outline, label: string): State {
  if (!document.numbered) return "unknown"
  const parts = label.split(".")
  if (!document.tops.has(parts[0]!)) return "missing"
  for (let depth = 2; depth <= parts.length; depth += 1) {
    const prefix = parts.slice(0, depth).join(".")
    if (document.printed.has(prefix)) continue
    const parent = `${parts.slice(0, depth - 1).join(".")}.`
    return [...document.printed.keys()].some(row => row.startsWith(parent) && row.split(".").length === depth) ? "missing" : "unknown"
  }
  return "present"
}

/** Where the clause numbered `label` before the changes is in the final text:
 * a top-level clause by its heading, a sub-clause by its text or else by its
 * clause's new number. Undefined when it is where it was or cannot be found. */
function moved(before: Outline, after: Outline, label: string): string | undefined {
  const parts = label.split(".")
  const unique = (map: Map<string, string>, value: string) => {
    const hits = [...map].filter(([, text]) => key(text) === key(value)).map(([row]) => row)
    return hits.length === 1 ? hits[0] : undefined
  }
  if (parts.length === 1) {
    const heading = before.heading.get(label)
    const found = heading ? unique(after.heading, heading) : undefined
    return found && found !== label ? found : undefined
  }
  const text = before.content.get(label)
  const found = text ? unique(after.content, text) : undefined
  if (found) return found !== label ? found : undefined
  const parent = moved(before, after, parts[0]!)
  const guess = parent ? [parent, ...parts.slice(1)].join(".") : undefined
  return guess && after.printed.has(guess) ? guess : undefined
}

function referenceFindings(title: string, final: Outline, baselines: Outline[], pack: (PackDocument & { outline: Outline })[], advisory: boolean): ConsistencyFinding[] {
  const refs = references(final, title, pack.map(row => row.title))
  const kinds = new Set<Kind>(["clause"])
  for (const kind of ["section", "article"] as const)
    if (final.keywords.has(kind) || final.texts.some(text => new RegExp(`\\b(?:this|these) ${kind}s?\\b`, "i").test(text)) || !refs.some(row => row.kind === "clause"))
      kinds.add(kind)
  // A number printed twice matters where something refers to it, unless one of
  // the two is a heading and the other the text under it ("6.2 Seller's
  // obligations" over "6.2 At Completion …").
  const referenced = new Set(refs.flatMap(row => !row.target && CLAUSE_KINDS.has(row.kind) ? row.labels : []))
  const ambiguous = final.duplicates.filter(row => referenced.has(row.label) &&
    headingLine(final.texts[row.at]!) === headingLine(final.texts[final.printed.get(row.label)!]!))
  const findings: ConsistencyFinding[] = advisory ? [] : ambiguous.map(row => ({
    pid: final.paras[row.at]!.pid,
    detail: `Two paragraphs of the final text are numbered ${row.label} (${where(final, final.printed.get(row.label)!)} and the paragraph beginning "${excerpt(final.paras[row.at]!.text, 60)}"); a reference to ${row.label} is ambiguous.`,
  }))
  for (const ref of refs) for (const label of ref.labels) {
    const pid = final.paras[ref.at]!.pid
    const from = capitalize(where(final, ref.at))
    if (ref.target) {
      const other = pack.find(row => names(row.title, ref.target!) && names(ref.target!, row.title) || (ref.target!.length >= 2 && initials(row.title) === ref.target))
      if (other && CLAUSE_KINDS.has(ref.kind) && clauseState(other.outline, label) === "missing")
        findings.push({ pid, detail: `${from} refers to ${ref.kind} ${label} of "${other.title}", but no ${ref.kind} in that document's final text is numbered ${label}.` })
      continue
    }
    if (advisory) continue
    if (!CLAUSE_KINDS.has(ref.kind)) {
      const name = `${capitalize(ref.kind)} ${label}`
      const keyName = `${ref.kind} ${label}`
      const before = baselines.some(row => row.attachments.has(keyName))
      if (final.attachments.has(keyName) || (!before && !final.attachmentAt.some(head => head?.startsWith(`${ref.kind} `)))) continue
      if (pack.some(row => key(row.title).startsWith(key(name)))) continue
      findings.push({ pid, detail: `${from} refers to ${name}, but the final text has no ${name}.${before ? ` ${name} was in the text before the changes.` : ""}` })
      continue
    }
    if (!kinds.has(ref.kind)) continue
    const state = clauseState(final, label)
    if (state === "missing") {
      const earlier = baselines.find(row => row.printed.has(label))
      const now = earlier ? moved(earlier, final, label) : undefined
      findings.push({ pid, detail: `${from} refers to ${ref.kind} ${label}, but no ${ref.kind} in the final text is numbered ${label}.` +
        (now ? ` What was ${ref.kind} ${label} before the changes is now ${ref.kind} ${now}.` : earlier ? ` There was a ${ref.kind} ${label} before the changes.` : "") })
      continue
    }
    if (state !== "present") continue
    for (const before of baselines) {
      if (!before.numbered || !carried(before, final.texts[ref.at]!, ref.kind, label)) continue
      const now = moved(before, final, label)
      if (now) {
        findings.push({ pid, detail: `${from} refers to ${ref.kind} ${label}. Before the changes that reference was to ${describe(before, label)}, which is now ${ref.kind} ${now}; ` +
          `${ref.kind} ${label} is now ${describe(final, label)}. Check that the reference points where it should.` })
        break
      }
      const was = replacedBy(before, final, label)
      if (was) {
        findings.push({ pid, detail: `${from} refers to ${ref.kind} ${label}. Before the changes that reference was to ${describe(before, label)}, which is no longer in the text; ` +
          `${ref.kind} ${label} is now ${describe(final, label)}, which was ${ref.kind} ${was} before. Check that the reference points where it should.` })
        break
      }
    }
  }
  return findings
}

/** The clause the final text numbers `label` in place of one the changes
 * removed: the old clause's heading is gone, and the clause now under that
 * number had another number before. Top-level clauses only, by heading. */
function replacedBy(before: Outline, after: Outline, label: string) {
  const [old, now] = [before.heading.get(label), after.heading.get(label)]
  if (label.includes(".") || !old || !now || key(old) === key(now) || [...after.heading.values()].some(text => key(text) === key(old))) return
  const hits = [...before.heading].filter(([, text]) => key(text) === key(now)).map(([row]) => row)
  return hits.length === 1 && hits[0] !== label ? hits[0] : undefined
}

/** Whether a reference was carried over from the text before the changes: the
 * paragraph most like this one there makes the same reference. */
function carried(before: Outline, text: string, kind: Kind, label: string) {
  const mine = new Set(key(text).split(" "))
  const best = before.texts.reduce<{ score: number; text?: string }>((acc, other) => {
    const theirs = new Set(key(other).split(" "))
    const shared = [...mine].filter(word => theirs.has(word)).length
    const score = shared / (mine.size + theirs.size - shared || 1)
    return score > acc.score ? { score, text: other } : acc
  }, { score: 0 })
  return best.score >= 0.5 && new RegExp(`\\b${kind}s?\\s+(?:[\\d.]+(?:\\s*(?:,|and|or|to)\\s*)?)*${label.replace(/\./g, "\\.")}(?![\\d.]*\\d)`, "i").test(best.text ?? "")
}

function describe(document: Outline, label: string) {
  const text = label.includes(".") ? document.content.get(label) : document.heading.get(label)
  return text ? `"${excerpt(text, 60)}"` : `${label}`
}

/** Where a paragraph is, as a lawyer would point to it: its clause or the one
 * it is under, the attachment it is in, or its opening words. */
function where(document: Outline, at: number) {
  if (at >= document.end) {
    const head = document.attachmentAt[at]
    if (head) return head.replace(/^\w/, letter => letter.toUpperCase())
  }
  for (let row = Math.min(at, document.end - 1); row >= 0; row -= 1) if (document.labelAt[row]) return `clause ${document.labelAt[row]}`
  return `the paragraph beginning "${excerpt(document.paras[at]!.text, 60)}"`
}

function capitalize(text: string) {
  return text.replace(/^\p{Ll}/u, letter => letter.toUpperCase())
}

// ------------------------------------------------------------ defined terms

/** A definition: by "means" (with its body) or by a parenthetical (with the
 * company it follows, when one is named before it in the same clause). */
type Definition = { term: string; at: number; body?: string; party?: string; alias?: true }

const CONNECTOR = "of|and|for|the|to|in|on|by|with|a|an"
const TERM_WORD = "[\\p{Lu}\\d][\\p{L}\\p{N}’'&-]*"
const TERM = `${TERM_WORD}(?:\\s+(?:(?:${CONNECTOR})\\s+)*${TERM_WORD}){0,5}`
const TERM_SHAPE = new RegExp(`^${TERM}$`, "u")
const QUOTED = /["“]([^"“”\n]{1,60})["”]|‘([^‘’\n]{1,60})’|'([^'\n]{1,60})'/g
const PARENTHETICAL = /\(([^()]{1,400})\)/g
const UNQUOTED_PARENTHETICAL = new RegExp(`^\\s*(?:(?:together|collectively|jointly|each|individually|hereinafter|hereafter)\\b[\\s,]*(?:called\\s+|referred\\s+to\\s+as\\s+)?)?the\\s+(${TERM})\\s*(?:$|,\\s*(?:being|which|that|as|meaning)\\b)`, "u")
const MEANS = "means|shall mean|has the meaning|have the meaning|includes|shall include|refers to"
const QUOTED_MEANS = new RegExp(`(?:["“]([^"“”\\n]{1,60})["”]|‘([^‘’\\n]{1,60})’)\\s*:?\\s+(?:${MEANS}|shall be|will be)\\b`, "gu")
const UNQUOTED_MEANS = new RegExp(`(?:^|[.;:]\\s+|,?\\s+(?:and|or)\\s+)(?:\\(?[a-z0-9]{1,4}[.)]\\s+)?(?:[Tt]he\\s+)?(${TERM}(?:\\s+or\\s+${TERM})?)\\s+(?:means|shall mean|has the meaning)\\b`, "gu")
const NOT_A_TERM_START = new Set(["The", "This", "That", "These", "Those", "It", "Each", "Any", "Such", "A", "An", "No", "Nothing", "Clause", "Section", "Article", "Schedule", "Annex", "Appendix", "Exhibit", "Part", "Paragraph", "In", "For", "If", "Where"])
const DEFINED_ELSEWHERE = /\b(?:has|have) the (?:same )?meanings? (?:given|ascribed|assigned) (?:to (?:it|them|that term) )?in\b|\bas defined in (?:the|that)\b|\bmeanings? given (?:to (?:them|it) )?in the\b/i
const EXPLAINED = new RegExp(`\\b[Tt]he\\s+(${TERM}),\\s+being\\b|\\breferences?\\s+to\\s+(?:the\\s+|an?\\s+)?(${TERM})\\s+(?:are|is|shall be)\\s+(?:to|references to)\\b`, "gu")
const PARTY = /(?:\p{Lu}[\p{L}\p{N}’'&.-]*\s+(?:(?:of|and|&|for|the)\s+)?){1,6}(?:Limited|LIMITED|Ltd|LTD|PLC|plc|LLP|LLC|LP|Inc|INC|Incorporated|INCORPORATED|Corporation|CORPORATION|Corp|CORP|GmbH|AG|SA|BV|NV|Pty|PTY)\b\.?/gu

function isTerm(text: string) {
  const term = text.trim()
  return term.length <= 60 && /^\p{Lu}/u.test(term) && TERM_SHAPE.test(term) && !NOT_A_TERM_START.has(term.split(/\s+/)[0]!)
}

function definitions(document: Outline): Definition[] {
  return document.texts.flatMap((text, at) => {
    const found: (Definition & { index: number })[] = []
    let from = 0
    for (const match of text.matchAll(PARENTHETICAL)) {
      const inner = match[1]!
      const quoted = [...inner.matchAll(QUOTED)].map(row => (row[1] ?? row[2] ?? row[3])!.trim()).filter(isTerm)
      const terms = quoted.length ? quoted : [UNQUOTED_PARENTHETICAL.exec(inner)?.[1]].filter((term): term is string => Boolean(term && isTerm(term)))
      if (!terms.length) continue
      const party = definedParty(text.slice(from, match.index))
      from = match.index + match[0].length
      terms.forEach(term => found.push({ term, at, index: match.index, ...(party ? { party } : {}) }))
    }
    // "“Price” or “Consideration” means …" defines both names; each is an alias.
    const means = [...text.matchAll(QUOTED_MEANS)].map(row => {
      const other = /["“]([^"“”\n]{1,60})["”]\s+or\s+$/u.exec(text.slice(0, row.index))?.[1]
      return { terms: [...other ? [other.trim()] : [], (row[1] ?? row[2])!.trim()], index: row.index, end: row.index + row[0].length }
    })
    const label = clauseLabel(text)
    const offset = label ? text.length - text.slice(text.indexOf(label.rest)).length : 0
    const body = text.slice(offset)
    means.push(...[...body.matchAll(UNQUOTED_MEANS)].map(row => ({ terms: row[1]!.split(/\s+or\s+/).map(term => term.trim()), index: offset + row.index + row[0].indexOf(row[1]!), end: offset + row.index + row[0].length })))
    const starts = means.map(row => row.index).toSorted((a, b) => a - b)
    means.forEach(row => {
      const next = starts.find(start => start > row.index) ?? text.length
      const terms = row.terms.filter(isTerm)
      const meaning = text.slice(row.end, next).trim()
      // "Borrower means Acme Ltd" names its party as a parenthetical does.
      const party = [...meaning.replace(/^the\s+/i, "").matchAll(PARTY)].find(match => match.index === 0)?.[0].trim()
      terms.forEach(term => found.push({ term, at, index: row.index, body: meaning, ...(party ? { party } : {}), ...(terms.length > 1 ? { alias: true as const } : {}) }))
    })
    return found.toSorted((a, b) => a.index - b.index).map(({ index: _index, ...definition }) => definition)
  })
}

/** The company a parenthetical definition names: the clause before it names
 * one company and no other, and either opens with it ("(1) Acme Ltd, a company
 * … (the "Buyer")", "To: Acme Ltd (the Buyer)") or ends with it ("between Acme
 * Ltd (the "Buyer")"). An agreement between two companies, or a sentence
 * about one, names no party. */
function definedParty(before: string) {
  const clause = before.replace(/\([^()]*\)/g, " ").split(";").at(-1)!
    .replace(/^[\s,.]*(?:\(?[A-Za-z0-9]{1,4}[.)]\s+)?(?:(?:to|from|by|and)\b:?\s*)?/i, "")
  const names = [...clause.matchAll(PARTY)]
  if (names.length !== 1) return
  const name = names[0]!
  return name.index === 0 || /^[\s,]*$/.test(clause.slice(name.index + name[0].length)) ? name[0].trim() : undefined
}

/** Terms a document explains without defining them: "the Effective Date,
 * being the date …", "references to the Tenant are to …". They count as
 * explained, never as definitions to compare or to find used. */
function explained(document: Outline) {
  return new Set(document.texts.flatMap(text => [...text.matchAll(EXPLAINED)].map(match => (match[1] ?? match[2])!.trim()))
    .filter(isTerm).flatMap(variants))
}

function variants(term: string) {
  const forms = new Set([term, `${term}s`, `${term}es`])
  if (/[^aeiouAEIOU]y$/.test(term)) forms.add(`${term.slice(0, -1)}ies`)
  if (/[^s]s$/.test(term) && term.length > 3) forms.add(term.slice(0, -1))
  return [...forms].toSorted((a, b) => b.length - a.length)
}

function termPattern(term: string, flags: string) {
  return new RegExp(`(?<![\\p{L}\\p{N}_])(?:${variants(term).map(escape).join("|")})(?:['’]s)?(?![\\p{L}\\p{N}_])`, flags)
}

function escape(text: string) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

/** Occurrences of a term in the visible text: as written, in lower case, and
 * in a mixed case away from a sentence's start; `body` is the first paragraph
 * using it as written that is not a heading. */
function uses(document: Outline, term: string) {
  const exact = termPattern(term, "gu")
  const any = termPattern(term, "giu")
  const rows = document.texts.map((text, at) => ({ at, text }))
  const count = (pattern: RegExp) => rows.reduce((sum, row) => sum + [...row.text.matchAll(pattern)].length, 0)
  const lower: string[] = []
  const mixed: { at: number; text: string }[] = []
  for (const row of rows) for (const match of row.text.matchAll(any)) {
    const text = match[0]
    if (new RegExp(`^${exact.source}$`, "u").test(text) || text === text.toUpperCase()) continue
    if (text === text.toLowerCase()) lower.push(text)
    else if (/\s/.test(term) && !sentenceStart(row.text.slice(0, match.index))) mixed.push({ at: row.at, text })
  }
  const using = rows.filter(row => termPattern(term, "u").test(row.text))
  return { exact: count(exact), lower, mixed, first: using[0]?.at, body: using.find(row => !headingLine(row.text))?.at }
}

function sentenceStart(before: string) {
  return /(?:^|[.!?:;]["”’)]?)\s*["“‘(]?$/.test(before.trimEnd()) || /^\s*(?:\(?[a-z0-9]{1,4}[.)]|\d+(?:\.\d+)*\.?)\s*$/i.test(before)
}

/** A document that defines fewer terms than this does not work by defined
 * terms, so its unused definitions and undefined capitalised phrases are not reported. */
const MIN_DEFINITIONS = 3

function termFindings(title: string, final: Outline, baselines: Outline[], pack: (PackDocument & { outline: Outline })[], advisory: boolean): ConsistencyFinding[] {
  const defined = definitions(final)
  const byTerm = Map.groupBy(defined, row => row.term)
  const glossary = !advisory && byTerm.size >= MIN_DEFINITIONS
  const elsewhere = pack.map(row => ({ ...row, definitions: definitions(row.outline) }))
  const definedInPack = (term: string) => elsewhere.some(row => row.definitions.some(definition => variants(term).includes(definition.term)))
  const usedInPack = (term: string) => elsewhere.some(row => uses(row.outline, term).exact > 0)
  const pid = (at: number | undefined) => at === undefined ? undefined : final.paras[at]!.pid
  const findings: ConsistencyFinding[] = []
  for (const [term, rows] of byTerm) {
    const twice = twiceDefined(rows)
    if (twice)
      findings.push({ pid: pid(twice[1].at), detail: `The term "${term}" is defined more than once: in ${where(final, twice[0].at)} and in ${where(final, twice[1].at)}.` })
    const seen = uses(final, term)
    const used = seen.exact - rows.length
    if (glossary && used <= 0 && !rows.some(row => row.alias) && !usedInPack(term))
      findings.push({ pid: pid(rows[0]!.at), detail: seen.lower.length
        ? `The term "${term}" is defined in ${where(final, rows[0]!.at)} but is written only in lower case ("${seen.lower[0]}") elsewhere, so the definition is not used as written.`
        : `The term "${term}" is defined in ${where(final, rows[0]!.at)} but is not used anywhere else in the document.` })
    // A definition may describe its term in ordinary words ("“Bank Consent” means the bank consent …").
    const mixed = seen.mixed.filter(row => !rows.some(definition => definition.at === row.at))
    if (used > 0 && mixed.length)
      findings.push({ pid: pid(mixed[0]!.at), detail: `The defined term "${term}" is written "${mixed[0]!.text}" in ${where(final, mixed[0]!.at)}; check that the defined term is used consistently.` })
  }
  // A definition the changes removed while the term is still used.
  const removed = new Set<string>()
  const described = explained(final)
  for (const before of baselines) for (const row of definitions(before)) {
    if (removed.has(row.term) || byTerm.has(row.term) || described.has(row.term) || defined.some(definition => variants(row.term).includes(definition.term)) || definedInPack(row.term)) continue
    const seen = uses(final, row.term)
    if (seen.body === undefined) continue
    removed.add(row.term)
    findings.push({ pid: pid(seen.body), detail: `The term "${row.term}" is used in the final text (first in ${where(final, seen.body)}) but is no longer defined: before the changes it was defined in ${where(before, row.at)}.` })
  }
  if (glossary && !final.texts.some(text => DEFINED_ELSEWHERE.test(text)))
    findings.push(...undefinedTerms(final, defined, definedInPack, new Set([...removed, ...described]), [title, ...pack.map(row => row.title)]))
  // A term a later pack document defines otherwise than an earlier one.
  for (const [term, rows] of byTerm) for (const other of elsewhere.filter(row => row.earlier)) {
    const theirs = other.definitions.filter(row => row.term === term)
    // Each document's own subject ("the Company" in a company's minutes) is
    // its own to name, and a document that gives the other's party another
    // term names the two companies apart on purpose.
    const renamed = (party: string, definitions: Definition[]) => definitions.some(row => row.term !== term && row.party && !partiesDiffer(row, { term, at: 0, party }))
    const differs = rows.some(mine => theirs.some(row => (mine.body !== undefined && row.body !== undefined && numbersDiffer(mine.body, row.body)) ||
      (partiesDiffer(mine, row) && !namedIn(title, mine.party!) && !namedIn(other.title, row.party!) && !renamed(row.party!, defined) && !renamed(mine.party!, other.definitions))))
    if (!differs) continue
    findings.push({ pid: pid(rows[0]!.at), detail: `The term "${term}" is defined differently in this document (${where(final, rows[0]!.at)}) and in "${other.title}" (${where(other.outline, theirs[0]!.at)}).` })
    break
  }
  return findings
}

/** Two definitions of a term that cannot both hold: two "means" definitions,
 * or two parentheticals naming different things. */
function twiceDefined(rows: Definition[]): [Definition, Definition] | undefined {
  const means = rows.filter(row => row.body !== undefined)
  if (means.length > 1) return [means[0]!, means[1]!]
  if (means.length) return
  for (const [index, first] of rows.entries()) for (const second of rows.slice(index + 1)) if (partiesDiffer(first, second)) return [first, second]
}

/** Two parenthetical definitions name different companies: neither name, as
 * compared, contains the other ("BY ACME LTD" and "Acme Limited" agree). */
function partiesDiffer(first: Definition, second: Definition) {
  if (!first.party || !second.party) return false
  const [a, b] = [nameKey(first.party), nameKey(second.party)]
  return !a.includes(b) && !b.includes(a)
}

/** Numbers cited as a document's parts, which each document numbers its own way. */
const CITED_NUMBERS = /\b(?:clauses?|sub-?clauses?|sections?|articles?|paragraphs?|recitals?|parts?|schedules?|annex(?:es)?|appendix|appendices|exhibits?|attachments?)\s+[\d.()a-zA-Z]+(?:\s*(?:,|and|or|to)\s*[\d.()a-zA-Z]+)*/gi

/** Two "means" definitions differ in substance when each states a figure the
 * other does not; wording, and detail one adds, are left to the reviewer. */
function numbersDiffer(first: string, second: string) {
  const figures = (text: string) => new Set(text.replace(CITED_NUMBERS, " ").match(/\d+(?:[.,]\d+)*/g) ?? [])
  const [a, b] = [figures(first), figures(second)]
  return [...a].some(figure => !b.has(figure)) && [...b].some(figure => !a.has(figure))
}


// Words of a capitalised phrase that make it a name of a place, body, law,
// date or part of a document rather than a term the document should define.
const NAME_WORDS = new Set([
  "Schedule", "Schedules", "Clause", "Clauses", "Section", "Sections", "Article", "Articles", "Annex", "Appendix", "Exhibit", "Part", "Paragraph",
  "Act", "Acts", "Code", "Regulation", "Regulations", "Rule", "Rules", "Directive", "Law", "Laws", "Statute", "Order", "Convention", "Treaty",
  "Court", "Courts", "Tribunal", "Commission", "Commissioner", "Authority", "Office", "Registry", "Council", "Government", "Department", "Agency",
  "Ministry", "Bank", "Union", "Kingdom", "States", "State", "Republic", "Parliament", "Crown", "Exchange", "Street", "Road", "Park", "House",
  ...MONTHS.split("|"), "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
])
const MIN_UNDEFINED_USES = 3

/** Capitalised phrases used as defined terms ("the Effective Date", "the
 * Buyer") at least MIN_UNDEFINED_USES times that the pack defines nowhere.
 * A phrase that is also used without an article mid-sentence is a name, and
 * one in a heading or the title is the document's own vocabulary. */
function undefinedTerms(final: Outline, defined: Definition[], definedInPack: (term: string) => boolean, reported: Set<string>, titles: string[]): ConsistencyFinding[] {
  const known = new Set(defined.flatMap(row => variants(row.term)))
  const heading = (text: string) => headingLine(text) || (words(text) <= MAX_HEADING_WORDS && /:$/.test(text.trim()))
  const headings = [...final.texts.filter((text, at) => heading(text) || at === 0), ...titles].map(key).join(" | ")
  const counts = new Map<string, { count: number; at: number }>()
  // Quoted passages are another text's words.
  const own = final.texts.map(text => text.replace(/["“][^"“”]*["”]/g, quote => words(quote) >= 4 ? " " : quote))
  own.forEach((text, at) => {
    if (heading(text)) return
    for (const match of text.matchAll(/\b(?:[Tt]he|[Tt]his|[Ee]ach|[Aa]ny|[Ss]uch|[Ee]ither|[Nn]either)\s+(\p{Lu}[\p{L}’'-]*(?:\s+(?:of\s+|and\s+|for\s+)?\p{Lu}[\p{L}’'-]*){0,3})(?![\p{L}\p{N}’'-])/gu)) {
      const phrase = match[1]!.replace(/['’]s$/, "")
      const current = counts.get(phrase)
      counts.set(phrase, { count: (current?.count ?? 0) + 1, at: current?.at ?? at })
    }
  })
  return [...counts].flatMap(([phrase, seen]) => {
    if (seen.count < MIN_UNDEFINED_USES || known.has(phrase) || reported.has(phrase) || definedInPack(phrase) || /['’]s\s/.test(phrase)) return []
    if (phrase === phrase.toUpperCase() || phrase.split(/\s+/).some(word => NAME_WORDS.has(word))) return []
    if ([...known].some(term => phrase.startsWith(`${term} `) || phrase.endsWith(` ${term}`))) return []
    if (headings.includes(key(phrase))) return []
    // A phrase cited as an instrument ("section 3 of the Bylaws") is that instrument's name.
    if (final.texts.some(text => new RegExp(`${CITED_NUMBERS.source}\\s+(?:of|to|in|under)\\s+(?:the\\s+)?${escape(phrase)}(?![\\p{L}\\p{N}])`, "iu").test(text))) return []
    const bare = new RegExp(`(?<!\\b(?:the|this|each|any|such|either|neither|The|This|Each|Any|Such|Either|Neither)\\s+)(?<![\\p{L}\\p{N}_])${escape(phrase)}(?![\\p{L}\\p{N}_])`, "u")
    if (final.texts.some(text => [...text.matchAll(new RegExp(bare.source, "gu"))].some(match => !sentenceStart(text.slice(0, match.index))))) return []
    return [{ pid: final.paras[seen.at]!.pid, detail: `"${phrase}" is used as a defined term (${seen.count} times, first in ${where(final, seen.at)}) but is not defined in the document.` }]
  })
}
