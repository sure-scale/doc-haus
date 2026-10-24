import { normalizePreferences } from "../../lib/drafting-preferences"
import { lintDraftStyle } from "../../lib/markdown"
import { readFirmPreferences } from "../../lib/preferences"
import { plannedDocument, type DocID } from "./plan"
import { matterSourceIDs } from "./quotes"
import { templateSource } from "./template-contract"
import { sourceText } from "./template-structure"

/** The paragraphs of a document the firm's style preferences govern. In a
 * revision they govern the text it writes or changes: a paragraph whose words
 * are the original's own, read from its canonical source text, keeps the
 * original's style and is not checked. Every paragraph of a new or template
 * draft is. */
export function styleScopedRows<Row extends { text: string }>(directory: string, document: DocID, rows: Row[]): Row[] {
  const entry = plannedDocument(directory, document)
  const source = entry?.mode === "revise" ? templateSource(entry) : undefined
  if (!source || !matterSourceIDs(directory).includes(source)) return rows
  const original = new Set(sourceText(directory, source).split(/\n\s*\n/).map(sameWords).filter(Boolean))
  return rows.filter(row => !original.has(sameWords(row.text)))
}

/** Words as compared with the original: runs of whitespace and the style of
 * quotation marks aside; case, digits and every other character count. */
function sameWords(text: string) {
  return text.replace(/[\u2018\u2019\u201A\u201B]/g, "'").replace(/[\u201C\u201D\u201E\u201F]/g, "\"").replace(/\s+/g, " ").trim()
}

// Derive advisory presentation feedback from the authoritative current paragraph
// records; do not persist a second, potentially stale style ledger.
export function draftingStyleWarnings(rows: { pid: string; text: string }[]) {
  const preferences = normalizePreferences(readFirmPreferences())
  return rows.flatMap(row => {
    const warnings = lintDraftStyle(row.text, preferences)
    return warnings.length ? [{ pid: row.pid, text: row.text, warnings }] : []
  })
}

export function draftingStyleWarningsBlock(rows: { pid: string; text: string }[], final = false) {
  const warnings = draftingStyleWarnings(rows)
  if (!warnings.length) return ""
  return [
    "<drafting_style_warnings>",
    "These date/number preference checks are advisory presentation feedback, not evidence of a legal defect or an unanswered substantive decision. Preserve numeric values, dates and meaning. Preserve exact source quotations and wording the lawyer expressly instructed you to retain; a warning alone does not authorise changing them. Never add an open-item marker or ask a human question solely for a style deviation.",
    final
      ? "These warnings concern the actual final body. The runner appends these retained style advisories to the drafting notes. Explain any intentional retention when supported, naming the document and relevant wording without internal paragraph identifiers; do not duplicate the advisory list. Distinguish a retained quotation or expressly preserved wording from an uncorrected style preference; do not claim the style was corrected when the final text still differs."
      : "Assess safe presentation corrections against the saved preferences, source evidence and settled instructions in the existing wording review and amendment flow. Distinguish presentation from substantive changes. Retain a deviation when source fidelity or express instructions require it; the final notes will disclose surviving warnings.",
    JSON.stringify(warnings, null, 2),
    "</drafting_style_warnings>",
  ].join("\n")
}

/** A non-blocking, human-facing appendix ensures surviving deviations remain
 * visible even when the notes writer does not mention the advisory handoff. */
export function draftingStyleAdvisoryNotes(documents: { title: string; rows: { pid: string; text: string }[] }[]) {
  const warnings = documents.flatMap(document => {
    const issues = [...new Set(draftingStyleWarnings(document.rows).flatMap(row => row.warnings))]
    return issues.length ? [`- ${document.title}: ${issues.join("; ")}.`] : []
  })
  if (!warnings.length) return ""
  return [
    "## Drafting style advisories",
    "The following date or number formats in the final documents differ from the saved drafting preferences. Exact source quotations and wording expressly preserved by instruction may intentionally retain their format. These presentation checks do not establish a substantive defect.",
    ...warnings,
  ].join("\n\n")
}
