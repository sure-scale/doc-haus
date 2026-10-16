import { copyFileSync, existsSync, readFileSync, rmSync } from "node:fs"
import { comparisonBaseline } from "./comparison"
import { compareDocxWithPointers } from "../../lib/drafting-amendment"
import { applyAmendments, createAmendmentMerge } from "../../lib/drafting-amendment-apply"
import { finalBodyItems, readChangeList, readParagraphs, settleApply, type ApplyResult } from "./checks"
import type { DocID } from "./plan"
import { LEAN_CHILD_PERMISSIONS, LEAN_STAGE_AGENTS, paragraphRows, stepOutputs, stepSessionTitle, type LeanDeps } from "./stages"
import { leanFile, matterRelative, writeLeanBytes, writeLeanText, type LeanState, type OpenItem } from "./state"
import { templateItems } from "./template-coverage"
import { exportAppliedChangeComments } from "./review-evidence"
import { draftingAttribution, withGeneratedDocxAttribution } from "../../lib/docx-attribution"

// The apply step of one document. Applies the document's change list to its
// drafted body, always from body.docx and never onto an earlier apply output,
// checks the final body for unsupported content and compares the two bodies
// into the document's redline. Paragraphs changed earlier in the batch go
// through the applier's merge callback, one `applier` child session per merge
// request.
//
// Every amendment ends applied or as a visible open item: an amendment the
// applier could not apply, did not report, or lost because the applier itself
// stopped is unapplied and carries the applier's detail. Only an engine failure
// during a merge stops the stage, as a resumable runtime error.

const MERGE_PREAMBLE = "This is a lean drafting run. No file tool is available; answer in your reply."

export type ApplyStageRun =
  | { status: "done"; outputs: string[]; openItems: OpenItem[] }
  | { status: "failed"; reason: string }

export async function runApply(directory: string, state: LeanState, deps: LeanDeps, document: DocID): Promise<ApplyStageRun> {
  const step = `${document}.apply` as const
  const changes = readChangeList(directory, document)
  if (!changes) return { status: "failed", reason: `${step}: ${document}/changes.json is missing or unreadable.` }
  const engineErrors: string[] = []
  const merge = createAmendmentMerge(prompt => deps.sessions.createSession({
    directory, parentID: state.owner_session_id, agent: LEAN_STAGE_AGENTS.merge, title: stepSessionTitle(directory, step, "merge"), model: state.model,
    permissions: LEAN_CHILD_PERMISSIONS,
  })
    .then(sessionID => deps.sessions.prompt({ directory, sessionID, agent: LEAN_STAGE_AGENTS.merge, text: `${MERGE_PREAMBLE}\n\n${prompt}` }))
    .catch((error: unknown) => {
      engineErrors.push(message(directory, error))
      throw error
    }))
  const body = leanFile(directory, `${document}/body.docx`)
  const out = leanFile(directory, `${document}/body-applied.docx`)
  // An earlier round's output must not stand in for this round's.
  rmSync(out, { force: true })
  const result = await applyAmendments({ docxPath: body, outPath: out, amendments: changes.amendments, merge })
    .then((value): ApplyResult => value, (error: unknown) => ({ error: message(directory, error) }))
  // The applier treats a thrown merge as an unusable answer; an engine failure is
  // a runtime stop, not a legal outcome to disclose.
  if (engineErrors.length) return { status: "failed", reason: `${step}: stopped by a runtime error during merge: ${engineErrors[0]}` }
  const settled = "error" in result ? settleApply(changes, undefined, false, result.error) : settleApply(changes, result, existsSync(out))
  if (settled.bodyUnchanged) copyFileSync(body, out)
  const identity = draftingAttribution()
  writeLeanBytes(directory, `${document}/body-applied.docx`, withGeneratedDocxAttribution(readFileSync(out), identity))
  const rows = await paragraphRows(out)
  writeLeanText(directory, `${document}/paragraphs-applied.json`, `${JSON.stringify(rows, null, 2)}\n`)
  const redline = withGeneratedDocxAttribution(await compareDocxWithPointers(comparisonBaseline(directory, document).bytes, readFileSync(out), { authorName: identity.author }), identity)
  const annotated = exportAppliedChangeComments(redline, readFileSync(out), document, changes.amendments, settled.result.applied, changes.explanations, readParagraphs(directory, document), identity.author)
  writeLeanBytes(directory, `${document}/redline.docx`, annotated.bytes)
  writeLeanText(directory, `${document}/apply.json`, `${JSON.stringify({ ...settled.result, comments: annotated.comments }, null, 2)}\n`)
  return {
    status: "done", outputs: stepOutputs(step),
    openItems: [
      ...settled.openItems, ...annotated.openItems, ...finalBodyItems(directory, rows),
      ...templateItems(directory, document, readParagraphs(directory, document), rows, changes.amendments, new Set(settled.result.applied.map(row => row.id))),
    ],
  }
}

function message(directory: string, error: unknown) {
  return matterRelative(directory, error instanceof Error ? error.message : String(error))
}
