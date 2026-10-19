import { exportDocxComments, readDocxCommentParagraphs, readDocxRevisionText } from "../../lib/docx-comments"
import { amendmentParagraphs, normalizeParagraphText, type Amendment } from "../../lib/drafting-amendment"
import type { AppliedAmendment } from "../../lib/drafting-amendment-apply"
import { finalBodyParagraphs, readParagraphs, readChangeList, type ParagraphRow, type ApplyResult, type CommentExport, type ChangeExplanation, type ChangeList } from "./checks"
import { plannedDocument, type DocID } from "./plan"
import { openItem } from "./quotes"
import { existsSync, readFileSync } from "node:fs"
import { leanFile, readLeanJson } from "./state"
import { draftingAttribution } from "../../lib/docx-attribution"

/** Application mechanics and evidential basis answer different questions.
 * Carry the actual passages so the memo writer and final reviewer can assess
 * their support without inferring it from how an amendment was applied. */
export function appliedChangeEvidence(directory: string, document: DocID) {
  return appliedChangeEvidenceFromRecords(
    readChangeList(directory, document),
    readLeanJson(directory, `${document}/apply.json`) as ApplyResult | undefined,
    finalBodyParagraphs(directory, document),
    readParagraphs(directory, document),
  )
}

/** The same evidence view from native records already read by the caller. */
export function appliedChangeEvidenceFromRecords(changes: ChangeList | undefined, apply: ApplyResult | undefined, current: ParagraphRow[], initial: ParagraphRow[]) {
  const amendments = new Map((changes?.amendments ?? []).map(row => [row.id, row]))
  return [...(apply?.applied ?? []).map(row => {
    const amendment = amendments.get(row.id)
    const attached = apply?.comments?.attached.find(comment => comment.id === row.id)
    const unresolved = apply?.comments?.unresolved.find(comment => comment.id === row.id)
    return {
      id: row.id,
      pid: row.pid,
      change_origin: "review-amendment",
      application_mode: row.mode,
      operation: amendment ? amendment.operation ?? "replace" : "unavailable",
      result_paragraphs: current.filter(paragraph => (row.result_pids ?? [row.pid]).includes(paragraph.pid)).map(paragraph => ({ pid: paragraph.pid, text: paragraph.text })),
      deleted_text: amendment?.operation === "delete" ? initial.find(paragraph => paragraph.pid === row.pid)?.text ?? "" : undefined,
      evidence_basis: !amendment ? "amendment-unavailable" : amendment.sources.length ? "source-passages-provided" : "no-source-passage",
      reason: amendment?.reason ?? "",
      sources: amendment?.sources ?? [],
      native_comment: attached ? { status: "attached", author: apply!.comments!.author, ...attached }
        : unresolved ? { status: "unresolved", ...unresolved } : { status: "unavailable" },
    }
  }), ...(changes?.explanations ?? []).map(row => {
    const attached = apply?.comments?.attached.find(comment => comment.id === row.id)
    const unresolved = apply?.comments?.unresolved.find(comment => comment.id === row.id)
    return { id: row.id, pid: row.pid, change_origin: "initial-draft-revision", application_mode: "initial-draft", operation: "explanation", result_paragraphs: current.filter(paragraph => paragraph.pid === row.pid).map(paragraph => ({ pid: paragraph.pid, text: paragraph.text })), deleted_text: undefined, evidence_basis: row.sources.length ? "source-passages-provided" : "no-source-passage", reason: row.reason, sources: row.sources,
      native_comment: attached ? { status: "attached", author: apply!.comments!.author, ...attached } : unresolved ? { status: "unresolved", ...unresolved } : { status: "unavailable" } }
  })]
}

/** Use the final applied paragraph, rather than an amendment's
 * proposed text. Comparison may discard paragraph bookmarks, so its accepted
 * text and paragraph position jointly locate the exported redline comment. */
export function exportAppliedChangeComments(bytes: Uint8Array, finalBytes: Uint8Array, document: DocID, amendments: Amendment[], applied: AppliedAmendment[], explanations: ChangeExplanation[] = [], initial: ParagraphRow[] = [], author = draftingAttribution().author) {
  const paragraphs = readDocxCommentParagraphs(finalBytes)
  const comments = [...applied, ...explanations].map(row => {
    const bodyAmendment = amendments.find(amendment => amendment.id === row.id)
    const amendment = bodyAmendment ?? explanations.find(explanation => explanation.id === row.id)
    const pids = "result_pids" in row && row.result_pids ? row.result_pids : [row.pid]
    const targets = pids.flatMap(pid => paragraphs.filter(paragraph => paragraph.pids.includes(pid)))
    const target = targets.length === pids.length ? targets[0] : undefined
    const deletion = amendment && "operation" in amendment && amendment.operation === "delete"
    const deleted = deletion ? initial.find(paragraph => paragraph.pid === row.pid)?.text : undefined
    const proposed = bodyAmendment ? amendmentParagraphs(bodyAmendment) : undefined
    return {
      id: row.id, author,
      text: amendment ? [
        ...(deletion ? [`Deleted paragraph: ${deleted ?? "Original text unavailable."}`] : []),
        amendment.reason,
        ...(amendment.sources.length ? amendment.sources.map(source => `Source: ${source.source_id}\nQuoted passage: ${source.quote}`)
          : ["No supporting source passage was provided; this explanation does not establish source support."]),
      ].join("\n\n") : "This applied review change has no available explanation record.",
      anchor: deleted ? { deleted_paragraph_text: deleted } : target
        ? targets.length > 1 ? { paragraphs: targets.map(paragraph => ({ paragraph_index: paragraph.paragraph_index, paragraph_text: paragraph.paragraph_text })) }
          : { paragraph_index: target.paragraph_index, paragraph_text: target.paragraph_text } : { pid: row.pid },
      unresolved_reason: !amendment ? "The applied change has no available explanation record."
        : deletion && pids.length ? "Later reconciliation retained wording; this deletion rationale does not describe the final result."
        : !deletion && !pids.length ? "Later reconciliation deleted the paragraph; this wording rationale does not describe the final result."
        : target && proposed && (targets.length !== proposed.length || targets.some((paragraph, index) =>
          normalizeParagraphText(paragraph.paragraph_text) !== normalizeParagraphText(proposed[index] ?? "")))
          ? "Final wording was reconciled after this proposal; its explanation needs review."
          : undefined,
    }
  })
  const exported = exportDocxComments(bytes, comments.filter(comment => !comment.unresolved_reason))
  const report: CommentExport = {
    author,
    attached: exported.attached.map(row => ({ ...row, text: comments.find(comment => comment.id === row.id)!.text })),
    unresolved: [
      ...comments.flatMap(comment => comment.unresolved_reason ? [{ id: comment.id, reason: comment.unresolved_reason, text: comment.text }] : []),
      ...exported.unresolved.map(row => ({ ...row, text: comments.find(comment => comment.id === row.id)!.text })),
    ],
  }
  return {
    bytes: exported.bytes, comments: report,
    openItems: report.unresolved.map(row => openItem("apply", "omission",
      `The explanation for revision ${row.id} could not be attached as a Word comment to the final redline: ${row.reason}. ` +
      `The body change remains applied. Proposed rationale for review: ${row.text}`,
      { document, amendment_id: row.id, pid: [...applied, ...explanations].find(amendment => amendment.id === row.id)!.pid })),
  }
}


/** This is a view of current native files, not a second authoritative ledger.
 * Counts describe export/application mechanics; legal sufficiency stays reviewed. */
export function nativeExportEvidence(directory: string, document: DocID) {
  const apply = readLeanJson(directory, `${document}/apply.json`) as ApplyResult | undefined
  const path = `.dochaus/lean/${document}/redline.docx`
  const file = leanFile(directory, `${document}/redline.docx`)
  const changes = readChangeList(directory, document)
  const plan = plannedDocument(directory, document)
  const outcomes = { delivery_registration: "pending", planned_file: plan?.file ?? null, planned_role: plan?.deliverable ?? "body", applied_amendment_count: apply?.applied?.length ?? null, unapplied_amendment_count: apply?.unapplied?.length ?? null, withheld_proposal_count: changes ? changes.withheld?.length ?? 0 : null }
  if (!existsSync(file)) return { path, status: "unavailable", ...outcomes }
  try {
    return { path, status: "exported", ...outcomes, ...readDocxRevisionText(readFileSync(file)) }
  } catch {
    return { path, status: "unreadable", ...outcomes }
  }
}
