import { visibleTextMatches } from "./markdown-visible-text"
import { INSTRUCTION_COMMENT_PATTERN, PLACEHOLDER_PATTERN } from "./marker-patterns"

export type RenderPlaceholderResolution = {
  placeholder: string
  replacement: string
  action: "internal_cross_reference" | "open_item"
  basis: string
}

export type RenderPlaceholderResolutionResult = {
  changed: boolean
  text: string
  resolutions: RenderPlaceholderResolution[]
  remaining: string[]
}

export type DraftingPlaceholderLocation = {
  // The exact raw source slice; line and column are its raw coordinates.
  placeholder: string
  line: number
  column: number
  excerpt: string
  // What the placeholder shows once Markdown escapes and markup are read.
  visible?: string
}

type Heading = {
  unit: string
  number: string
  title: string
  label: string
  tokens: string[]
}

const STRUCTURAL_REFERENCE_PLACEHOLDER_PATTERN =
  /\b(?:article|section|clause|paragraph|subsection|schedule|exhibit|appendix)\s+\[(?:[A-Z][A-Z0-9.-]*|[IVXLC]+|[A-Za-z][^[\]\n]{1,80})\]/gi

const STRUCTURAL_REFERENCE_WORDS = new Set([
  "a",
  "an",
  "and",
  "appendix",
  "article",
  "clause",
  "cross",
  "exhibit",
  "for",
  "from",
  "in",
  "number",
  "of",
  "paragraph",
  "provision",
  "reference",
  "schedule",
  "section",
  "subsection",
  "the",
  "this",
  "to",
  "under",
])

const REFERENCE_UNITS = ["article", "section", "clause", "paragraph", "subsection", "schedule", "exhibit", "appendix"]

export function resolveRenderPlaceholders(text: string): RenderPlaceholderResolutionResult {
  const headings = extractNumberedHeadings(text)
  // Resolve only internal references that the draft's own headings prove.
  // Other placeholders stay byte-for-byte visible so the readiness gate can
  // route them to the finalizer or stop; deterministic cleanup must never turn
  // missing facts or law into complete-looking prose.
  const matches = [...text.matchAll(PLACEHOLDER_PATTERN)]

  const resolutions: RenderPlaceholderResolution[] = []
  const replaced = matches.reduce(
    (state, match) => {
      const start = match.index ?? state.cursor
      const resolution = resolvePlaceholder(match[0], text, start, headings)
      resolutions.push(resolution)
      return {
        cursor: start + match[0].length,
        text: state.text + text.slice(state.cursor, start) + resolution.replacement,
      }
    },
    { cursor: 0, text: "" },
  )
  const next = replaced.text + text.slice(replaced.cursor)

  return {
    changed: next !== text,
    text: next,
    resolutions,
    remaining: draftingPlaceholders(next),
  }
}

export function draftingPlaceholders(text: string) {
  const matches = [
    ...text.matchAll(PLACEHOLDER_PATTERN),
    ...text.matchAll(INSTRUCTION_COMMENT_PATTERN),
    ...text.matchAll(STRUCTURAL_REFERENCE_PLACEHOLDER_PATTERN),
  ]
  return [...new Set(matches.map((match) => match[0].replace(/\s+/g, " ").trim()))]
}

export function stripDraftingPlaceholders(text: string) {
  return text
    .replace(PLACEHOLDER_PATTERN, "")
    .replace(INSTRUCTION_COMMENT_PATTERN, "")
    .replace(STRUCTURAL_REFERENCE_PLACEHOLDER_PATTERN, "")
}

// Placeholders are found on visible text and bound to raw bytes, so a Word
// projection's escaped `\[\_\_\_\]` is found as the `[___]` it shows and its
// location covers the whole escaped slice the finalizer must edit.
export function draftingPlaceholderLocations(text: string): DraftingPlaceholderLocation[] {
  const lines = text.split(/\r?\n/)
  return visibleTextMatches(text, [
    PLACEHOLDER_PATTERN,
    INSTRUCTION_COMMENT_PATTERN,
    STRUCTURAL_REFERENCE_PLACEHOLDER_PATTERN,
  ]).map((match) => ({
    placeholder: match.raw,
    line: match.line,
    column: match.column,
    excerpt: (lines[match.line - 1] ?? "").replace(/\s+/g, " ").trim(),
    visible: match.visible,
  }))
}

export function appendRenderPlaceholderReport(report: string, source: string, result: RenderPlaceholderResolutionResult) {
  if (!result.resolutions.length) return report
  const rows = result.resolutions
    .map(
      (resolution) =>
        `| ${escapeCell(resolution.placeholder)} | ${resolution.action} | ${escapeCell(resolution.replacement)} | ${escapeCell(resolution.basis)} |`,
    )
    .join("\n")
  return `${report.replace(/\s*$/, "")}\n\n## Render Placeholder Readiness\n\nSource: \`${source}\`\n\n| Placeholder | Action | Replacement | Basis |\n|---|---|---|---|\n${rows}\n`
}

function resolvePlaceholder(
  placeholder: string,
  text: string,
  offset: number,
  headings: Heading[],
): RenderPlaceholderResolution {
  const label = placeholderLabel(placeholder)
  const heading = isInternalCrossReference(label) ? matchingHeading(label, headings) : undefined
  if (heading) {
    const replacement = referenceReplacement(label, text.slice(Math.max(0, offset - 60), offset), heading)
    return {
      placeholder,
      replacement,
      action: "internal_cross_reference",
      basis: `matched numbered heading "${heading.label}" in the assembled body`,
    }
  }
  return {
    placeholder,
    replacement: placeholder,
    action: "open_item",
    basis: "left unresolved because no reliable internal heading match or source value is available",
  }
}

function extractNumberedHeadings(text: string): Heading[] {
  return text
    .split(/\r?\n/)
    .map((line) => {
      const match = line.match(/^#{1,6}\s+(?:(Article|Section|Clause|Schedule|Exhibit|Appendix)\s+)?([A-ZIVXLC]+|\d+(?:\.\d+)*)(?:[.)])?\s+(.+?)\s*$/i)
      if (!match) return
      const rawTitle = match[3]
      const number = match[2]
      if (!rawTitle || !number) return
      const title = rawTitle.replace(/\s+\[[^\]]+\]/g, "").trim()
      if (!title) return
      const unit = normalizedUnit(match[1] ?? "Section")
      return {
        unit,
        number,
        title,
        label: `${unit} ${number}. ${title}`,
        tokens: meaningfulTokens(title),
      }
    })
    .filter((heading): heading is Heading => Boolean(heading))
}

function matchingHeading(label: string, headings: Heading[]) {
  const wanted = meaningfulTokens(label)
  if (!wanted.length) return
  const scored = headings
    .map((heading) => ({ heading, score: headingScore(wanted, heading.tokens) }))
    .filter((candidate) => candidate.score > 0)
    .sort((a, b) => b.score - a.score)
  const best = scored[0]
  if (!best) return
  const next = scored[1]
  if (next && next.score === best.score) return
  return best.heading
}

function headingScore(wanted: string[], title: string[]) {
  const common = wanted.filter((token) => title.includes(token))
  if (common.length !== wanted.length) return 0
  const exact = wanted.join(" ") === title.join(" ")
  const phrase = title.join(" ").includes(wanted.join(" "))
  return common.length + (exact ? 10 : 0) + (phrase ? 4 : 0)
}

function isInternalCrossReference(label: string) {
  const lower = label.toLowerCase()
  return REFERENCE_UNITS.some((unit) => new RegExp(`\\b${unit}\\b`).test(lower))
}

function referenceReplacement(label: string, before: string, heading: Heading) {
  const unit = requestedUnit(label) ?? heading.unit
  if (new RegExp(`\\b${unit}s?\\s*$`, "i").test(before)) return heading.number
  if (/\bnumber\b/i.test(label) && !/\breference\b/i.test(label)) return heading.number
  return `${unit} ${heading.number}`
}

function requestedUnit(label: string) {
  const lower = label.toLowerCase()
  return REFERENCE_UNITS.find((unit) => new RegExp(`\\b${unit}\\b`).test(lower))?.replace(/^\w/, (char) =>
    char.toUpperCase(),
  )
}

function placeholderLabel(placeholder: string) {
  const raw = placeholder
    .replace(/^\[/, "")
    .replace(/\]$/, "")
    .replace(/\s+/g, " ")
    .trim()
  if (!raw || /^_+$/.test(raw)) return "unresolved value"
  return raw
    .replace(/^(?:insert|placeholder|tbd|to be determined|to be confirmed)\s+/i, "")
    .replace(/^(?:the|a|an)\s+/i, "")
    .trim()
}

function meaningfulTokens(value: string) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .split(/\s+/)
    .map(singularToken)
    .filter((token) => token && !STRUCTURAL_REFERENCE_WORDS.has(token))
}

function singularToken(token: string) {
  if (token.endsWith("ies") && token.length > 4) return `${token.slice(0, -3)}y`
  if (token.endsWith("s") && !token.endsWith("ss") && token.length > 3) return token.slice(0, -1)
  return token
}

function normalizedUnit(unit: string) {
  return unit.replace(/^\w/, (char) => char.toUpperCase())
}

function escapeCell(value: string) {
  return value.replace(/\|/g, "\\|").replace(/\s+/g, " ").trim()
}
