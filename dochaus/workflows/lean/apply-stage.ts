import { copyFileSync, existsSync, readFileSync, rmSync } from "node:fs"
import { comparisonBaseline } from "./comparison"
import { compareDocxWithPointers, paragraphPointers, withStableRevisionDates } from "../../lib/drafting-amendment"
import { applyAmendments } from "../../lib/drafting-amendment-apply"
import { finalBodyItems, readChangeList, readParagraphs, settleApply, type ApplyResult, type ChangeList } from "./checks"
import type { DocID } from "./plan"
import { paragraphRows, stepOutputs } from "./stages"
import { leanFile, matterRelative, readLeanJson, writeLeanBytes, writeLeanText, type LeanState, type OpenItem } from "./state"
import { draftStatusItems } from "./draft-status"
import { templateItems } from "./template-coverage"
import { consistencyItems } from "./consistency"
import { exportAppliedChangeComments } from "./review-evidence"
import { draftingAttribution, withGeneratedDocxAttribution } from "../../lib/docx-attribution"
import { numberDocxClauses } from "../../lib/docx-numbering"

// The apply step of one document. Applies the document's change list to its
// drafted body, always from body.docx and never onto an earlier apply output,
// checks the final body for unsupported content and compares the two bodies
// into the document's redline. It writes only the reviewed change list's own
// text and opens no session: overlapping changes were consolidated before
// change-review.
//
// Every amendment ends applied or as a visible open item: an amendment the
// applier could not apply (an overlap included), did not report, or lost
// because the applier itself stopped is unapplied and carries the applier's
// detail.

export type ApplyStageRun =
  | { status: "done"; outputs: string[]; openItems: OpenItem[] }
  | { status: "failed"; reason: string }

export async function runApply(directory: string, state: LeanState, document: DocID): Promise<ApplyStageRun> {
  const step = `${document}.apply` as const
  const changes = readChangeList(directory, document)
  if (!changes) return { status: "failed", reason: `${step}: ${document}/changes.json is missing or unreadable.` }
  const body = leanFile(directory, `${document}/body.docx`)
  const out = leanFile(directory, `${document}/body-applied.docx`)
  // An earlier round's output must not stand in for this round's.
  rmSync(out, { force: true })
  const result = await applyAmendments({ docxPath: body, outPath: out, amendments: changes.amendments })
    .then((value): ApplyResult => value, (error: unknown) => ({ error: message(directory, error) }))
  const settled = "error" in result ? settleApply(changes, undefined, false, result.error) : settleApply(changes, result, existsSync(out))
  if (settled.bodyUnchanged) copyFileSync(body, out)
  const identity = draftingAttribution()
  // A changed paragraph whose text now carries a clause number the body's list
  // supplies is numbered by the list, as at drafting. The later parts of a
  // split paragraph leave its list unless they open with a deeper number of it.
  const later = new Set(settled.result.applied.flatMap(row => row.result_pids?.slice(1) ?? []))
  const split = (paragraph: Element) => paragraphPointers(paragraph).some(pointer => later.has(pointer.pid))
  writeLeanBytes(directory, `${document}/body-applied.docx`, withGeneratedDocxAttribution(numberDocxClauses(readFileSync(out), { house: false, split }), identity))
  const rows = await paragraphRows(out)
  writeLeanText(directory, `${document}/paragraphs-applied.json`, `${JSON.stringify(rows, null, 2)}\n`)
  const annotated = await documentRedline(directory, document, readFileSync(out), changes, settled.result.applied, identity)
  // Tracked changes are dated when this step started applying them; a change
  // already in the previous redline keeps its date, so an unchanged redline
  // keeps its bytes and the evidence keyed on them.
  const previous = leanFile(directory, `${document}/redline.docx`)
  writeLeanBytes(directory, `${document}/redline.docx`, withStableRevisionDates(annotated.bytes, state.stages[step]?.started_at ?? new Date().toISOString(),
    existsSync(previous) ? readFileSync(previous) : undefined))
  writeLeanText(directory, `${document}/apply.json`, `${JSON.stringify({ ...settled.result, comments: annotated.comments }, null, 2)}\n`)
  return {
    status: "done", outputs: stepOutputs(step),
    openItems: [
      ...settled.openItems, ...annotated.openItems, ...finalBodyItems(directory, rows), ...draftStatusItems(directory, document, rows, "apply"),
      ...templateItems(directory, document, readParagraphs(directory, document), rows, changes.amendments, new Set(settled.result.applied.map(row => row.id))),
      ...consistencyItems(directory, document, readParagraphs(directory, document), rows),
    ],
  }
}

/** A document's redline: its comparison baseline (comparisonBaseline) compared
 * with `final`, each applied change's explanation attached as a Word comment.
 * The apply step passes its own outcome. A change the lawyer accepts into the
 * delivered body later rebuilds the delivered redline from the outcome the step
 * recorded (changes.json and apply.json). */
export async function documentRedline(directory: string, document: DocID, final: Uint8Array,
  changes: ChangeList | undefined = readChangeList(directory, document),
  applied: ApplyResult["applied"] | undefined = recordedApplied(directory, document), identity = draftingAttribution()) {
  if (!changes || !applied) throw new Error(`${document}: the change list or the apply outcome is missing or unreadable.`)
  const redline = withGeneratedDocxAttribution(await compareDocxWithPointers(comparisonBaseline(directory, document).bytes, final, { authorName: identity.author }), identity)
  return exportAppliedChangeComments(redline, final, document, changes.amendments, applied, changes.explanations, readParagraphs(directory, document), identity.author)
}

function recordedApplied(directory: string, document: DocID) {
  const applied = (readLeanJson(directory, `${document}/apply.json`) as Partial<ApplyResult> | undefined)?.applied
  return Array.isArray(applied) ? applied : undefined
}

function message(directory: string, error: unknown) {
  return matterRelative(directory, error instanceof Error ? error.message : String(error))
}
