import type { ChecklistFindingClass } from "./checklist-classify"
import type { ChecklistFinding, ChecklistItem } from "./checklist-contract"
import { checklistOutputState, type ChecklistOpenFinding, type ChecklistUnresolvedFinding, type ChecklistUnresolvedParent } from "./stages"
import { readChangeReview, deferredFindingItems } from "./change-pipeline"
import { renderNotesCoverage } from "../../lib/notes-coverage-report"
import { hasOpenItemMarker } from "../../lib/open-items"
import { draftingPlaceholders } from "../../lib/placeholder-resolution"
import { normalize } from "../../lib/quote-verify"
import { finalBodyParagraphs, type ApplyResult } from "./checks"
import { FIRST_DOCUMENT, readLeanPlan, type DocID } from "./plan"
import { readPackageReview } from "./package-review"
import { quoteFound, openItem } from "./quotes"
import { readRepairRound } from "./repairs"
import { withheldChecklistRepair } from "./requests"
import { isReviewStage, currentOpenItems, readLeanText, readLeanJson, readLeanState, readQuestions, type LeanState, type OpenItem, type Question } from "./state"

// Final legal gate: the only place an open item blocks delivery. It covers every
// planned document; each item is settled against the final body of the
// document it was found in. An item is
// resolved when its cause is gone from that final body, disclosed when the
// lawyer can see it in the delivered files, and otherwise open. What "visible"
// requires depends on the item:
// - a placeholder is itself visible in the body; the drafting notes must name it;
// - unsupported content left in the body (an unverified quotation or an
//   unsupported citation) needs a visible open-item marker in its paragraph and
//   the notes naming it; a mention in the notes alone does not make the body honest;
// - an unanswered question needs a marker in the paragraph it affects and the
//   notes naming it; with no affected paragraph it can only be answered;
// - an amendment that was not applied leaves the body unchanged, so the notes
//   naming it is enough; so is a question whose paragraph only that withheld
//   amendment would have marked;
// - a template part the final body leaves out with no recorded reason is
//   disclosed by the notes naming it;
// - a template part left out on purpose (by the outline, or by an applied
//   amendment with its reason) is a recorded removal, resolved by the notes
//   naming it;
// - a checklist finding left open (a requirement the focused check could not
//   establish, or one with decisions still open) stays non-blocking and is
//   disclosed here with its recorded reason and limitation, derived from
//   checklist/checks.json, which the gate already reads;
// - a completed checklist finding that remains adverse (and a refined parent
//   whose coverage is adverse) is never a pass: it is disclosed, with the class
//   its verified fields establish, ahead of every other item.
// Disclosed items make the output a review draft, one with checklist findings
// when any adverse finding is disclosed; execution-ready output has no
// disclosed items and no unanswered questions. Incomplete, failed or stale
// checklist output still blocks.

export type GateDelivery = "review_draft" | "review_draft_with_findings" | "execution_ready"

export type GateResult = {
  status: "pass" | "blocked"
  /** Set when the gate passes: what the delivered files may be used as. */
  delivery: GateDelivery | null
  blocking: OpenItem[]
  disclosed: OpenItem[]
  items: OpenItem[]
  checked_at: string
}

type FinalBody = { paragraphs: Map<string, string>; body: string }

type SettleContext = {
  directory: string
  documents: Map<DocID, FinalBody>
  notes: string
  questions: Map<string, Question>
}

export function finalGate(directory: string): GateResult {
  const planned = readLeanPlan(directory)?.documents.map(row => row.id) ?? [FIRST_DOCUMENT]
  const context: SettleContext = {
    directory,
    documents: new Map(planned.map(id => {
      const rows = finalBodyParagraphs(directory, id)
      const paragraphs = new Map(rows.map(row => [row.pid, normalize(row.text)]))
      const actual = new Map(paragraphs)
      const applied = (readLeanJson(directory, `${id}/apply.json`) as ApplyResult | undefined)?.applied ?? []
      // Preserve the affected original paragraph's lineage through a split.
      // Deletion has no surviving paragraph in which to disclose a question.
      for (const row of applied) if (row.result_pids)
        paragraphs.set(row.pid, row.result_pids.map(pid => actual.get(pid) ?? "").join("\n"))
      return [id, { paragraphs, body: normalize(rows.map(row => row.text).join("\n")) }]
    })),
    notes: readLeanText(directory, "drafting-notes.md") ?? "",
    questions: new Map(readQuestions(directory).map(question => [question.id, question])),
  }
  const review = readPackageReview(directory)
  const semantic = review?.status === "changes-required" ? review.findings.map(row => openItem("package-review", "omission", row.detail, { ...(row.document ? { document: row.document } : {}), ...(row.pid ? { pid: row.pid } : {}) })) : []
  const projected = planned.flatMap(document => {
    const checked = readChangeReview(directory, document)
    return checked?.status === "changes-required" ? checked.findings.map(row => openItem("change-review", "omission", row.detail,
      { document, ...(row.pid ? { pid: row.pid } : {}) })) : []
  })
  const leanState = readLeanState(directory)
  const checklistState = leanState?.stages["checklist-check"]
  const checked = checklistState && (checklistState.status !== "pending" || readLeanJson(directory, "checklist/extraction.json") !== undefined)
    ? checklistOutputState(directory) : undefined
  // Adverse checklist findings lead: the lawyer's list of items may be capped.
  const unresolved = checked?.valid ? [...checked.unresolved.map(row => checklistUnresolvedItem(directory, leanState, row)),
    ...checked.unresolvedParents.map(row => checklistUnresolvedParentItem(directory, row))] : []
  const items = [...unresolved, ...[...currentOpenItems(directory), ...deferredFindingItems(directory)].map(item => settle(item, context)), ...semantic, ...projected]
  if (checked && !checked.valid) items.push(openItem("checklist-check", "omission", "Effective checklist findings or source-unit coverage are incomplete, adverse or stale; see checklist/report.md.", {}))
  // Focused checks can correctly preserve a source-dependent open issue
  // without requesting a missing answer or inventing a disclosure. Such a
  // package remains a review draft, even when no generic gate item blocks it,
  // and each open finding is disclosed with its own reason and limitation.
  const checklistOpen = checklistState?.status === "done" && checked?.valid && checked.open
  if (checklistOpen) items.push(...checked.openFindings.map(row => checklistOpenItem(directory, row)))
  const blocking = items.filter(item => item.status === "open")
  const disclosed = items.filter(item => item.status === "disclosed")
  const unanswered = [...context.questions.values()].some(question => question.status !== "answered")
  return {
    status: blocking.length ? "blocked" : "pass",
    delivery: blocking.length ? null : unresolved.length ? "review_draft_with_findings" : disclosed.length || unanswered || checklistOpen ? "review_draft" : "execution_ready",
    blocking,
    disclosed,
    items,
    checked_at: new Date().toISOString(),
  }
}

/** The drafting-notes coverage report for the gate: every item with its outcome, and
 * the blocking items as gaps. */
export function notesCoverageReport(gate: GateResult) {
  return renderNotesCoverage({
    rows: gate.items.map(item => ({ id: `${item.id} (${item.status})${item.resolution ? `: ${item.resolution}` : ""}` })),
    issues: gate.blocking.map(item => `${item.id}: ${item.resolution ?? "open"}. ${item.detail}`),
  })
}

/** An open checklist finding as a disclosed item: the requirement and the
 * finding's own recorded reason, limitation and open decisions. The verdict
 * is the checker's; the gate only makes it visible. */
function checklistOpenItem(directory: string, { item_id, item, finding }: ChecklistOpenFinding): OpenItem {
  const document = checklistDocument(directory, item)
  const detail = [`Checklist requirement left open: ${sentence(item.check)}`, ...findingText(finding)].join(" ")
  return { ...openItem("checklist-check", "checklist_open", detail, { checklist_item: item_id, ...(document ? { document } : {}) }),
    status: "disclosed", resolution: "left open by the checklist check; disclosed with the review draft" }
}

/** Why an adverse finding remains, in the lawyer's words. The runner's report
 * and the gate's disclosure share it. It reads recorded state only: whether
 * the one reinspection ran, the source adjudication that declined a repair,
 * and the checklist repair round in repairs.json (the current round and each
 * document's retained repair context). */
export function checklistFindingLabel(directory: string, state: LeanState | undefined, row: { item_id: string; item: ChecklistItem; class: ChecklistFindingClass; reinspected: boolean }): string {
  switch (row.class) {
    case "omission":
      if (state && withheldChecklistRepair(directory, state, row.item_id, row.item)) return "omission; source adjudication declined the repair"
      return checklistRoundItems(directory).has(row.item_id) ? "omission not fixed by the repair round" : "omission; no repair round ran"
    case "missing_inspection":
      return row.reinspected ? "not established by the check after reinspection" : "not established by the check; not yet reinspected"
    case "source_conflict": return "conflicts with another source requirement"
    case "not_inspectable": return "cannot be established from document text"
    case "departure": return "deliberate departure"
  }
}

/** Item IDs a recorded checklist repair round submitted. */
function checklistRoundItems(directory: string) {
  const round = readRepairRound(directory)
  const items = [...round?.items ?? [], ...Object.values(round?.document_contexts ?? {}).flatMap(context => context?.items ?? [])]
  return new Set(items.flatMap(item => {
    if (item.stage !== "checklist-check") return []
    try {
      const value = JSON.parse(item.detail.slice(item.detail.indexOf("\n") + 1)) as { item_id?: unknown } | null
      return typeof value?.item_id === "string" ? [value.item_id] : []
    } catch { return [] }
  }))
}
const UNRESOLVED = "not established as satisfied; disclosed with the review draft"

/** A completed finding that remains adverse, disclosed with its class and its
 * own recorded reason and limitation. It is never a pass and never blocks. */
function checklistUnresolvedItem(directory: string, state: LeanState | undefined, row: ChecklistUnresolvedFinding): OpenItem {
  const { item_id, item, finding, class: kind, conflicting_item_ids } = row
  const document = checklistDocument(directory, item)
  const detail = [`Checklist requirement unresolved (${checklistFindingLabel(directory, state, row)}): ${sentence(item.check)}`, ...findingText(finding),
    conflicting_item_ids?.length ? `Conflicting checklist items: ${conflicting_item_ids.join(", ")}.` : ""].filter(Boolean).join(" ")
  return { ...openItem("checklist-check", "checklist_unresolved", detail, { checklist_item: item_id, checklist_class: kind, ...(document ? { document } : {}) }),
    status: "disclosed", resolution: UNRESOLVED }
}

/** A refined parent whose reconciled coverage is adverse, with the coverage issues. */
function checklistUnresolvedParentItem(directory: string, { parent_id, item, status, issues }: ChecklistUnresolvedParent): OpenItem {
  const document = checklistDocument(directory, item)
  const detail = [`Checklist requirement unresolved across its refined checks (${status === "unsatisfied" ? "not satisfied" : "could not be established"}): ${sentence(item.check)}`,
    issues.length ? `Coverage: ${issues.map(sentence).join(" ")}` : "A refined check remains adverse."].join(" ")
  return { ...openItem("checklist-check", "checklist_unresolved", detail, { checklist_item: parent_id, ...(document ? { document } : {}) }),
    status: "disclosed", resolution: UNRESOLVED }
}

function checklistDocument(directory: string, item: ChecklistItem) {
  return readLeanPlan(directory)?.documents.find(row => row.id === item.target.document || row.file === item.target.document)?.id
}

function findingText(finding: ChecklistFinding) {
  return [`Finding: ${sentence(finding.reason)}`, finding.limitation.trim() && `Limitation: ${sentence(finding.limitation)}`,
    finding.open_decisions?.length ? `Open decisions: ${sentence(finding.open_decisions.map(row => row.trim()).join("; "))}` : ""].filter(Boolean)
}

function sentence(text: string) {
  return /[.!?]$/.test(text.trim()) ? text.trim() : `${text.trim()}.`
}

function settle(item: OpenItem, context: SettleContext): OpenItem {
  const resolved = (resolution: string): OpenItem => ({ ...item, status: "resolved", resolution })
  const disclosed = (resolution: string): OpenItem => ({ ...item, status: "disclosed", resolution })
  const open = (resolution: string): OpenItem => ({ ...item, status: "open", resolution })
  const named = context.notes.includes(item.id)
  // Drafting and review findings are re-derived from the final body at apply.
  if (item.stage === "draft" || isReviewStage(item.stage)) return resolved("superseded by the check of the final body")
  if (item.kind === "research_gap") return named ? disclosed("legal research limitation named in the drafting notes") : open("the drafting notes do not disclose this legal research limitation")
  const document = item.ref.document ?? FIRST_DOCUMENT
  const final = context.documents.get(document)
  if (!final) return open("its document is no longer planned")
  if (item.kind === "unapplied_amendment")
    return named ? disclosed("not applied; named in the drafting notes") : open("the drafting notes do not name this unapplied change")
  if (item.kind === "omission")
    return named ? disclosed("omission named in the drafting notes") : open("the drafting notes do not name this omission")
  if (item.kind === "template_removal")
    return named ? resolved("template part left out or cut down on purpose, with a recorded reason; named in the drafting notes")
      : open("the drafting notes do not name this removal of a template part")
  if (item.kind === "question") {
    const question = item.ref.question_id ? context.questions.get(item.ref.question_id) : undefined
    if (question?.status === "answered") return resolved("answered by the lawyer")
    // A question asked again by a later document points into that document; the
    // item's own paragraph is in the document it was found in.
    const pid = item.ref.pid ?? ((question?.document ?? FIRST_DOCUMENT) === document ? question?.pid : undefined)
    if (named && pid && marked(final, pid)) return disclosed("unanswered; marked in the body and named in the drafting notes")
    if (named && item.ref.amendment_id)
      return disclosed("unanswered; the change that would have marked its paragraph was withheld, so the paragraph is unchanged; named in the drafting notes")
    return open(pid
      ? "unanswered; needs a visible open-item marker in the affected paragraph and the drafting notes naming it, or an answer"
      : "unanswered and tied to no paragraph; needs an answer, or a marker in the paragraph it affects")
  }
  const text = item.ref.placeholder ?? item.ref.citation ?? item.ref.quote
  if (!inBody(final, text, item.ref.pid)) return resolved("no longer in the final body")
  if (item.kind === "quote_miss" && item.ref.quote && quoteFound(context.directory, { quote: item.ref.quote, source_id: item.ref.source_id ?? "" }))
    return resolved("re-verified against the source")
  if (item.kind === "placeholder")
    return named ? disclosed("placeholder in the body; named in the drafting notes") : open("the drafting notes do not name this placeholder")
  if (named && item.ref.pid && marked(final, item.ref.pid))
    return disclosed("marked in the body and named in the drafting notes")
  return open(named
    ? "named in the notes, but its paragraph has no visible open-item marker"
    : "neither marked in its paragraph nor named in the drafting notes")
}

function inBody(final: FinalBody, text: string | undefined, pid: string | undefined) {
  if (!text) return false
  const haystack = pid ? final.paragraphs.get(pid) : final.body
  return Boolean(haystack?.includes(normalize(text)))
}

function marked(final: FinalBody, pid: string) {
  const text = final.paragraphs.get(pid)
  return Boolean(text && paragraphShowsOpenItem(text))
}

/** Whether a paragraph shows the lawyer an open item: a visible open-item marker
 * or a drafting placeholder. The changes check applies the same test to a
 * question's paragraph before admitting it, so the two never disagree. */
export function paragraphShowsOpenItem(text: string) {
  const normalized = normalize(text)
  return hasOpenItemMarker(normalized) || draftingPlaceholders(normalized).length > 0
}
