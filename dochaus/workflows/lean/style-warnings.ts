import { normalizePreferences } from "../../lib/drafting-preferences"
import { lintDraftStyle } from "../../lib/markdown"
import { readFirmPreferences } from "../../lib/preferences"

// Derive advisory presentation feedback from the authoritative current paragraph
// records; do not persist a second, potentially stale style ledger.
export function draftingStyleWarnings(rows: { pid: string; text: string }[]) {
  const preferences = normalizePreferences(readFirmPreferences())
  return rows.flatMap(row => {
    const warnings = lintDraftStyle(row.text, preferences)
    return warnings.length ? [{ pid: row.pid, text: row.text, warnings }] : []
  })
}

/** `final` is the notes the warnings feed: plain notes get the appendix below; a
 * requested memo is delivered as written, without it. */
export function draftingStyleWarningsBlock(rows: { pid: string; text: string }[], final: false | "notes" | "memo" = false) {
  const warnings = draftingStyleWarnings(rows)
  if (!warnings.length) return ""
  return [
    "<drafting_style_warnings>",
    "These date/number preference checks are advisory presentation feedback, not evidence of a legal defect or an unanswered substantive decision. Preserve numeric values, dates and meaning. Preserve exact source quotations and wording the lawyer expressly instructed you to retain; a warning alone does not authorise changing them. Never add an open-item marker or ask a human question solely for a style deviation.",
    final
      ? `These warnings concern the actual final body.${final === "notes" ? " The runner appends these retained style advisories to the drafting notes." : ""} Explain any intentional retention when supported, naming the document and relevant wording without internal paragraph identifiers${final === "notes" ? "; do not duplicate the advisory list" : ""}. Distinguish a retained quotation or expressly preserved wording from an uncorrected style preference; do not claim the style was corrected when the final text still differs.`
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
