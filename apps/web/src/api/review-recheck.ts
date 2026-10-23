import type { GridCellData, GridRecheck, RecheckPassages, SeenPassage } from "./ingest/grid"
import type { Citation } from "./opencode/parts"
import { readAnswerText, readReviewAnswer, retrievedCitations } from "./review-answer"
import type { Part } from "./session-view"

// One review-grid cell: an extract call, then, for a cell left without a
// citation ("Not addressed", or a source label that did not resolve), exactly
// one re-check of the same document. The re-check either adopts an answer whose
// source is one of the deterministic pass's numbered candidate passages or
// leaves the cell negative with a record of what was searched. Negative is
// judged from the stored cell (no citation). For "Not addressed" the passages
// the first search returned are skipped: the model read them and found nothing,
// so the miss was retrieval. For an uncited answer they stay candidates: the
// answer most likely came from one of them and only its label failed.

type Messages = Array<{ parts: Part[] }>

export type CellJobs = {
  /** One extract session for the prompt; resolves with its assistant messages. */
  extract(prompt: string): Promise<Messages>
  /** The deterministic pass over the cell's own document (ingest). */
  passages(request: { seen: SeenPassage[]; includeSeen: boolean }): Promise<RecheckPassages>
  /** One extract session that sees only the passages in the prompt. */
  recheck(prompt: string): Promise<Messages>
}

export const NOT_ADDRESSED = "Not addressed"

/** Fill blank and stale cells, and a negative cell stored before re-checks
 * existed; never a reviewed cell or one already re-checked for this question. */
export function needsFill(cell: GridCellData | undefined, questionHash: string) {
  if (!cell) return true
  if (cell.status === "reviewed") return false
  return cell.questionHash !== questionHash || (!cell.citation && !cell.recheck)
}

export async function computeCell(input: { document: string; question: string; questionHash: string }, jobs: CellJobs): Promise<GridCellData> {
  const first = await jobs.extract(`Document: "${input.document}". Question: ${input.question}`)
  const { text, citation } = readReviewAnswer(first)
  const cell: GridCellData = { answer: text || NOT_ADDRESSED, citation: citation && stored(citation), status: "filled", questionHash: input.questionHash }
  if (cell.citation) return cell
  const found = await jobs.passages({
    seen: retrievedCitations(first).map(({ docPath, charStart, charEnd }) => ({ docPath, charStart, charEnd })),
    includeSeen: !/^not addressed\.?$/i.test(cell.answer),
  })
  if (!found.citations.length) return { ...cell, recheck: { outcome: "not_found", ...found.scope } }
  const adopted = readCandidateAnswer(await jobs.recheck(recheckPrompt(input.document, input.question, found.passages)), found.citations)
  return adopted.citation && adopted.text
    ? { ...cell, answer: adopted.text, citation: stored(adopted.citation), recheck: { outcome: "found", ...found.scope } }
    : { ...cell, recheck: { outcome: "not_found", ...found.scope } }
}

export function recheckPrompt(document: string, question: string, passages: string) {
  return `Document: "${document}". Question: ${question}\n` +
    "Re-check: the first answer for this document had no verified source. The numbered passages below are from this document " +
    "only. Answer only from them; tools are unavailable for this check. On the Source line give the number of the one passage " +
    `that supports the answer, as \`Source: [n]\`. If none of them supports an answer, reply \`${NOT_ADDRESSED}\`.\n\n${passages}`
}

// A re-check answer is adopted only through its passage number, unique within
// this re-check, so two candidates sharing a section label cannot collide.
function readCandidateAnswer(messages: Array<{ parts: Part[] }>, candidates: Citation[]) {
  const { text, source } = readAnswerText(messages)
  const number = Number(source?.match(/^\[(\d+)\]$/)?.[1])
  return { text, citation: number >= 1 ? candidates[number - 1] : undefined }
}

/** The stored re-check as the lawyer reads it: the outcome and what was searched. */
export function describeRecheck(recheck: GridRecheck) {
  const outcome = recheck.outcome === "found" ? "Found on re-check." : "Not found after re-check."
  if (!recheck.passages) return `${outcome} This document has no indexed passages to search.`
  const shown = recheck.terms.slice(0, 6).map(term => `"${term}"`).join(", ")
  const terms = recheck.terms.length > 6 ? `${shown} and ${recheck.terms.length - 6} more` : shown
  const scope = recheck.passages === 1 ? "the one passage" : `all ${recheck.passages} passages`
  const searched = `${outcome} Searched ${scope} of this document for ${terms} by exact terms`
  const k = recheck.candidates - recheck.reread
  if (recheck.reread) {
    // An uncited answer's re-check: the passages already read, then any further ones.
    const read = recheck.reread < recheck.read ? `${recheck.reread} of the ${recheck.read} passages`
      : recheck.read === 1 ? "the one passage" : `the ${recheck.read} passages`
    const further = !k ? "" : k === 1 ? " with the one further passage that mentions any of them" : ` with the ${k} further passages that mention any of them`
    const start = recheck.terms.length ? `${searched} and re-read` : `${outcome} The question has no searchable terms; re-read`
    return `${start} ${read} the first search returned${further}: ${recheck.outcome === "found" ? "the answer cites one of them" : "none supported a cited answer"}.`
  }
  if (!recheck.terms.length) return `${outcome} The question has no searchable terms, so the document was not searched again.`
  if (!k) return `${searched}: no passage${recheck.read ? ` beyond the ${recheck.read} the first search returned` : ""} mentions any of them.`
  if (recheck.outcome === "found")
    return `${searched}: the answer cites ${k === 1 ? "the one further passage that mentions" : `one of ${k} further passages that mention`} any of them.`
  return `${searched}: ${k} further ${k === 1 ? "passage mentions" : "passages mention"} any of them, but none supported a cited answer.`
}

function stored(citation: Citation): NonNullable<GridCellData["citation"]> {
  const { documentName, docPath, section, structuralPath, excerpt, charStart, charEnd } = citation
  return { documentName, docPath, section, ...(structuralPath ? { structuralPath } : {}), excerpt, charStart, charEnd }
}
