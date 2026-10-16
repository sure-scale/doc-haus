import type { DocxSession } from "docxodus"
import { visibleLines, type VisibleLine } from "./markdown-visible-text"

// Docxodus's insertParagraph follows markdown soft-wrap: consecutive non-blank
// lines merge into ONE block whose text carries literal newlines, and Word does
// not render a newline inside a run as a break — a heading line swallows its
// body, and "(a)"/"(b)" sub-clause lines collapse into one run-on paragraph.
// Agent-authored bodies use a newline to mean a new paragraph, so make every
// line its own block. Markdown lists are unaffected: docxodus splits list items
// into their own paragraphs whether or not blank lines separate them.
export function splitBlocks(markdown: string) {
  return markdown
    .split("\n")
    .filter((line) => !isAlignmentWrapper(line))
    .filter((line) => line.trim())
    .join("\n\n")
}

function isAlignmentWrapper(line: string) {
  return /^<\/?center>$/i.test(line.trim())
}

// A template defect multiplies into every document later drafted from it, so
// agent-authored template bodies are linted before creation. These are the
// defects no model is needed for: a duplicate placeholder collapses two
// different terms into one fill value, a bare [___] cannot be filled by a
// descriptive name, a heading-number gap ships in every draft, and a duplicate
// optional-clause short-name makes omit ambiguous.
export function lintTemplateBody(markdown: string) {
  const brackets = markdown.match(/\[[^\]\n]+\]/g) ?? []
  const markers = brackets.filter((b) => b.startsWith("[optional:"))
  const fills = brackets.filter((b) => !b.startsWith("[optional:"))
  const dupes = (list: string[]) => [...new Set(list.filter((b, i) => list.indexOf(b) !== i))]
  const numbers = markdown
    .split("\n")
    .map((line) => line.match(/^ {0,3}#{2,6} (\d+)\./)?.[1])
    .filter((n): n is string => n !== undefined)
  const numberingBreak = numbers.findIndex((n, i) => n !== `${i + 1}`)
  return [
    ...dupes(fills).map((b) => `placeholder ${b} appears more than once — every placeholder must be unique`),
    ...fills.filter((b) => /^\[_+\]$/.test(b)).map((b) => `bare ${b} — use a descriptive [insert ...] placeholder`),
    ...dupes(markers).map((b) => `optional-clause marker ${b} appears more than once — short-names must be unique`),
    ...(numberingBreak === -1
      ? []
      : [`heading numbering breaks at "${numbers[numberingBreak]}." — expected ${numberingBreak + 1}`]),
  ]
}

// Deterministic consistency lint over a COMPOSED draft's projected markdown —
// the defects no model is needed for: a defined term never used, a term used
// like a defined term but never defined, a cross-reference to a section number
// no heading carries (renumberHeadings makes stale cross-references a live
// risk), and date/number formats that break the firm's preferences. Advisory:
// draft-document reports the problems so the agent fixes them by redline in
// its read-back; it never blocks the draft. Forward references (a term used in
// the recitals, defined in a later section) are standard drafting and are NOT
// flagged. Wording with legal effect — statutory notices, conspicuousness,
// jurisdiction-specific phrasing — stays with legal review and the
// jurisdiction packs, never here.
export type DraftStyle = {
  dateFormat: "month-day-year" | "day-month-year" | "iso"
  numberStyle: "words-and-numerals" | "numerals"
}

export type DraftMarkdownNamespaceIssue = {
  code: "unused_term" | "undefined_term" | "broken_cross_reference"
  line: number
  message: string
}

export function inspectDraftMarkdownNamespace(
  markdown: string,
  options: {
    externalSectionNumbers?: readonly string[]
    documentTitle?: string
  } = {},
): DraftMarkdownNamespaceIssue[] {
  // A Word projection escapes punctuation that has no visible form. Terms and
  // references are read on each line's visible text; heading classification
  // stays on the source line, so an escaped "\#" paragraph is still a
  // paragraph. The projection keeps every line, so issue lines stay exact.
  const lines = visibleLines(markdown)
  const visible = lines.map((line) => line.visible).join("\n")
  const definitions = definedTerms(visible)
  const terms = definitions.map((definition) => definition.term)
  return [
    ...(terms.length >= 3
      ? [
          ...unusedTerms(visible, definitions).map((issue) => ({
            code: "unused_term" as const,
            ...issue,
          })),
          ...undefinedTerms(lines, terms, options.documentTitle).map((issue) => ({
            code: "undefined_term" as const,
            ...issue,
          })),
        ]
      : []),
    ...brokenCrossReferences(lines, visible, options.externalSectionNumbers).map((issue) => ({
      code: "broken_cross_reference" as const,
      ...issue,
    })),
  ]
}

export function lintDraftMarkdown(markdown: string, style: DraftStyle) {
  return [
    ...inspectDraftMarkdownNamespace(markdown).map((issue) => issue.message),
    ...lintDraftStyle(markdown, style),
    ...(visibleLines(markdown).some((line) => /\[_+\]/.test(line.visible))
      ? ["bare [___] left in the document — fill it or use a descriptive [insert ...] placeholder"]
      : []),
  ]
}

// Presentation checks are advisory and can be supplied independently of legal
// namespace checks to the wording reviewer and final drafting notes.
export function lintDraftStyle(markdown: string, style: DraftStyle) {
  return [...numberStyleProblems(markdown, style.numberStyle), ...dateFormatProblems(markdown, style.dateFormat)]
}

// A quoted capitalized phrase — (the "Receiving Party") — is how contracts
// define terms; quoting is rare enough outside definitions to treat every such
// phrase as one.
function definedTerms(markdown: string) {
  const matches = [
    ...markdown.matchAll(/["“]([A-Z][\w'-]*(?: (?:[A-Z][\w'-]*|of|and|or|to|in|for|the)){0,5})["”]/g),
    ...markdown.matchAll(/['‘]([A-Z][\w-]*(?:['’][A-Za-z][\w-]*)?(?: (?:[A-Z][\w-]*(?:['’][A-Za-z][\w-]*)?|of|and|or|to|in|for|the)){0,5})['’]/g),
    ...markdown.matchAll(/\(([A-Z][A-Za-z0-9'-]*(?: (?:[A-Z][A-Za-z0-9'-]*|of|and|or|to|in|for|the)){0,5})\)/g),
  ].toSorted((left, right) => (left.index ?? 0) - (right.index ?? 0))
  return matches.flatMap((match) => {
    const term = match[1]
    if (!term || match.index === undefined) return []
    return [{ term, line: lineAt(markdown, match.index) }]
  }).filter(
    (definition, index, definitions) =>
      definitions.findIndex((entry) => entry.term === definition.term) === index,
  )
}

function unusedTerms(
  markdown: string,
  definitions: Array<{ term: string; line: number }>,
) {
  return definitions
    .filter((definition) =>
      !variants(definition.term).some((variant) =>
        new RegExp(
          `(?<!["“])\\b${escapeRegExp(variant)}\\b(?!["”])`,
        ).test(markdown)
      )
    )
    .map((definition) => ({
      line: definition.line,
      message: `"${definition.term}" is defined but never used`,
    }))
}

function variants(term: string) {
  if (term.endsWith("y")) return [term, term.slice(0, -1) + "ies"]
  if (term.endsWith("s")) return [term, term.slice(0, -1)]
  return [term, term + "s"]
}

// Statute and place names and job titles ride determiners the same way
// defined terms do, so phrases carrying citation or office vocabulary are
// never flagged as undefined terms.
const TERM_STOPWORDS = new Set([
  "Section", "Sections", "Article", "Articles", "Exhibit", "Exhibits", "Schedule", "Schedules", "Appendix",
  "Code", "Act", "Law", "Laws", "Statute", "Statutes", "Regulation", "Regulations",
  "State", "States", "Commonwealth", "United", "Court", "County", "City",
  "Department", "Agency", "Service", "Commission", "Authority", "Administration", "Bureau", "Office",
  "Officer", "Officers", "President", "Chairman", "Secretary", "Treasurer",
  "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December",
])

// "the Equity Plan" twice with no ("Equity Plan") definition anywhere is a
// drafting defect; a single occurrence is usually just a proper noun, so one
// mention never flags. The capture runs through lowercase connectors so "the
// Change of Control" reads as the defined "Change of Control", not an
// undefined "Change". Phrases that appear inside a heading are section-title
// words ("Confidentiality"), which read naturally without a definition.
function undefinedTerms(lines: readonly VisibleLine[], terms: string[], documentTitle?: string) {
  const defined = new Set(terms.flatMap(variants))
  const headings = [
    ...lines.filter((line) => line.raw.trimStart().startsWith("#")).map((line) => line.visible),
    ...(documentTitle ? [documentTitle] : []),
  ]
  const counts = new Map<string, { count: number; line: number }>()
  lines.forEach((line, index) => {
    if (line.raw.trimStart().startsWith("#")) return
    const usages = line.visible.matchAll(
      /\b(?:[Tt]he|[Tt]his|[Ss]uch|[Ee]ach|[Aa]ny|[Ee]ither|[Nn]either|[Nn]o|[Aa]n?|[Ii]ts) ((?:[A-Z][\w'-]*)(?:(?: (?:of|and|or|to|in|for|the))* [A-Z][\w'-]*){0,7})/g,
    )
    for (const usage of usages) {
      const phrase = usage[1]!.replace(/['’]s$/, "")
      const current = counts.get(phrase)
      counts.set(phrase, {
        count: (current?.count ?? 0) + 1,
        line: current?.line ?? index + 1,
      })
    }
  })
  return [...counts]
    .filter(
      ([phrase, usage]) =>
        usage.count >= 2 &&
        !defined.has(phrase) &&
        ![...defined].some((d) =>
          phrase.startsWith(d + " ") ||
          phrase.startsWith(d + "'") ||
          phrase.endsWith(" " + d)
        ) &&
        !/^[A-Z][A-Z0-9.-]*$/.test(phrase) &&
        !phrase.split(/[ -]/).some((word) => TERM_STOPWORDS.has(word)) &&
        !headings.some((h) => h.includes(phrase)),
    )
    .map(([phrase, usage]) => ({
      line: usage.line,
      message: `"${phrase}" reads like a defined term but is never defined`,
    }))
}

function brokenCrossReferences(
  lines: readonly VisibleLine[],
  markdown: string,
  externalSectionNumbers: readonly string[] = [],
) {
  const headings = new Set([
    ...headingSectionNumbers(lines),
    ...externalSectionNumbers,
  ])
  if (headings.size === 0) return []
  // Statutory citations ("Labor Code Section 2870", "Section 7 of the Act")
  // reference external law, not this document's headings — and once a number
  // is established as statutory anywhere, a bare back-reference to it later
  // ("the provisions of Section 2870") is statutory too.
  const authorityAcronyms = new Set(
    [...markdown.matchAll(/\b([A-Z][A-Z0-9.-]{1,11})[,§ ]+[Ss]ections? \d+/g)]
      .map((match) => match[1]!),
  )
  const authorityName =
    /(?:Code|Act|Law|Laws|Statute|Statutes|Regulation|Regulations|Rules?|U\.?S\.?C\.?|C\.?F\.?R\.?)$/
  const statutory = new Set(
    [
      ...markdown.matchAll(/(?:Code|Act|Law|Laws|Statute|Statutes|Regulation|Regulations|Rules?|U\.?S\.?C\.?|C\.?F\.?R\.?|[A-Z][A-Z0-9.-]{1,11})[,§ ]+[Ss]ections? (\d+)/g),
      ...markdown.matchAll(/[Ss]ections? (\d+)(?:\([^)]+\))* of (?:the )?(?:[^.;\n]{0,100} )?(?:Code|Act|Law|Laws|Statute|Statutes|Regulation|Regulations|Rules?|[A-Z][A-Z0-9.-]{1,11})\b/g),
    ].map((match) => match[1]),
  )
  const problems = new Map<string, number>()
  for (const m of markdown.matchAll(/\b[Ss]ections? (\d+)(?!\w)(?!\.\d)/g)) {
    const before = markdown.slice(Math.max(0, m.index - 40), m.index)
    let rest = markdown.slice(m.index + m[0].length)
    const cited = [m[1]!]
    let chain
    while ((chain = rest.match(/^(?:,? (?:and|or|through) |, )(\d+)(?!\w)(?!\.\d)/))) {
      cited.push(chain[1]!)
      rest = rest.slice(chain[0].length)
    }
    const clauseStart = Math.max(
      markdown.lastIndexOf("\n", m.index - 1),
      markdown.lastIndexOf(";", m.index - 1),
      markdown.lastIndexOf(".", m.index - 1),
    )
    const clauseBefore = markdown.slice(clauseStart + 1, m.index)
    const externalAlias = [...authorityAcronyms].some((alias) =>
      new RegExp(`\\b${escapeRegExp(alias)}\\b`).test(clauseBefore)
    )
    if (authorityName.test(before.trimEnd()) || externalAlias) continue
    if (
      /^(?:\([^)]+\))* of (?!this\b)(?:the )?(?:[^.;\n]{0,100} )?(?:Code|Act|Law|Laws|Statute|Statutes|Regulation|Regulations|Rules?|[A-Z][A-Z0-9.-]{1,11})\b/.test(
        rest,
      )
    ) continue
    cited
      .filter((n) => !headings.has(n) && !statutory.has(n))
      .forEach((n) => {
        const message =
          `cross-reference to Section ${n}, but no heading is numbered ${n}`
        problems.set(
          message,
          problems.get(message) ?? lineAt(markdown, m.index),
        )
      })
  }
  return [...problems].map(([message, line]) => ({ line, message }))
}

export function draftMarkdownHeadingSectionNumbers(markdown: string) {
  return headingSectionNumbers(visibleLines(markdown))
}

// Structure (a bold section label, an ATX heading) is classified on the source
// line; the section numbers are read from its visible text.
function headingSectionNumbers(lines: readonly VisibleLine[]) {
  return lines
    .flatMap((source) => {
      const label = /^ {0,3}\*\*[Ss]ection (\d+)([^*\n]*)\*\*([^\n]*)$/.exec(source.raw)
      if (label && (!label[3]!.trim() || /\p{L}/u.test(label[2]!))) return [label[1]!]
      if (!/^ {0,3}#{1,6} /.test(source.raw)) return []
      const line = source.visible
      const numbers = [line.match(/^ {0,3}#{1,6} (\d+)\./)?.[1]]
      for (const reference of line.matchAll(/\b[Ss]ections? (\d+)/g)) {
        numbers.push(reference[1])
        let rest = line.slice((reference.index ?? 0) + reference[0].length)
        let chain
        while ((chain = rest.match(/^(?:,? (?:and|or|through) |, )(\d+)/))) {
          numbers.push(chain[1])
          rest = rest.slice(chain[0].length)
        }
      }
      return numbers.filter((number): number is string => Boolean(number))
    })
}

function lineAt(markdown: string, index: number) {
  return markdown.slice(0, index).split("\n").length
}

const NUMBER_WORDS =
  "zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred"

function numberStyleProblems(markdown: string, pref: DraftStyle["numberStyle"]) {
  if (pref === "words-and-numerals") {
    const bare = [...markdown.matchAll(/(\()?\b(\d{1,4})(\))?( (?:calendar |business )?(?:day|week|month|year)s?)\b/g)]
      .filter((m) => !(m[1] && m[3]))
      .map((m) => `"${m[2]}${m[4]}"`)
    if (bare.length === 0) return []
    return [`durations must be words then numerals ("thirty (30) days") per firm preferences: ${[...new Set(bare)].join(", ")}`]
  }
  const wordy = [
    ...markdown.matchAll(new RegExp(`\\b(?:${NUMBER_WORDS})(?:[ -](?:${NUMBER_WORDS}))?\\s*\\(\\d{1,4}\\)`, "gi")),
  ].map((m) => `"${m[0]}"`)
  if (wordy.length === 0) return []
  return [`numbers must be plain numerals ("30 days") per firm preferences: ${[...new Set(wordy)].join(", ")}`]
}

const MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December"

function dateFormatProblems(markdown: string, pref: DraftStyle["dateFormat"]) {
  const formats = {
    "month-day-year": { example: '"June 11, 2026"', pattern: new RegExp(`\\b(?:${MONTHS}) \\d{1,2}, \\d{4}\\b`, "g") },
    "day-month-year": { example: '"11 June 2026"', pattern: new RegExp(`\\b\\d{1,2} (?:${MONTHS}) \\d{4}\\b`, "g") },
    iso: { example: '"2026-06-11"', pattern: /\b\d{4}-\d{2}-\d{2}\b/g },
  }
  const offenders = Object.entries(formats)
    .filter(([key]) => key !== pref)
    .flatMap(([, format]) => [...markdown.matchAll(format.pattern)].map((m) => `"${m[0]}"`))
  if (offenders.length === 0) return []
  return [`dates must be written ${formats[pref].example} per firm preferences: ${[...new Set(offenders)].join(", ")}`]
}

function escapeRegExp(text: string) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

// Templates number their clause headings in the heading text itself ("9.
// Non-Solicitation"), so deleting an optional clause leaves a hole in the
// sequence ("8." followed by "10."). Renumber every such heading to its
// position in document order. Body cross-references by section number are not
// rewritten.
export function renumberHeadings(session: DocxSession) {
  session
    .findByKind("h")
    .filter((h) => /^\d+\. /.test(h.textPreview))
    .forEach((h, i) => {
      const current = h.textPreview.match(/^(\d+)\./)![1]
      if (current !== `${i + 1}`) session.replaceTextRange(h.id, `${current}.`, `${i + 1}.`, { maxReplacements: 1 })
    })
}
