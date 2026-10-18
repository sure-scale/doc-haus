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
import { coverageCarries, coverageItems, gateScopeParties } from "./package-coverage"
import { excerpt, quoteFound, openItem } from "./quotes"
import { answeredScopeQuestion, GATE_EXCERPT, openScopeQuestion, scopeAnswers, scopeField, scopeQuestionShows } from "./package-scope"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, repairedChecklistItems } from "./repairs"
import { withheldChecklistRepair } from "./requests"
import { isReviewStage, currentOpenItems, readLeanText, readLeanJson, readLeanState, readQuestions, shortHash, type LeanState, type OpenItem, type Question } from "./state"

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
//   its verified fields establish, ahead of every other item, and ahead of the
//   drafting notes in the notes delivered with the draft;
// - source conflicts among those findings are put to the lawyer as one
//   question, which the gate asks after the walk: it pauses nothing, and each
//   conflict stays disclosed until the answer re-runs the stages it reaches;
// - an instruction (a directive item) the checklist check left open may be
//   carried by no planned document, which disclosure alone cannot settle; the
//   gate puts those no answered package-scope question has shown to the lawyer
//   as one package-scope question on the same terms, and none while another
//   stage's is still open; each refers to the question that showed it.
//   One the plan carries in full, left open for a fact its body marks as
//   missing, is disclosed and not asked about.
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
  /** The gate's questions: one on the disclosed source conflicts, one on the
   * package's scope for the instructions left open. The runner records them
   * with the run's questions, not in gate.json. Absent when the gate asks
   * nothing. */
  questions?: Question[]
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
  const conflict = checked?.valid ? checklistConflictQuestion(checked.unresolved) : undefined
  // Adverse checklist findings lead: the lawyer's list of items may be capped.
  const unresolved = checked?.valid ? [...checked.unresolved.map(row => checklistUnresolvedItem(directory, leanState, row, conflict?.id)),
    ...checked.unresolvedParents.map(row => checklistUnresolvedParentItem(directory, row))] : []
  const items = [...unresolved, ...[...currentOpenItems(directory), ...deferredFindingItems(directory), ...coverageItems(directory)].map(item => settle(item, context)), ...semantic, ...projected]
  if (checked && !checked.valid) items.push(openItem("checklist-check", "omission", "Effective checklist findings or source-unit coverage are incomplete or stale, or a check or its reinspection failed; see checklist/report.md.", {}))
  // Focused checks can correctly preserve a source-dependent open issue
  // without requesting a missing answer or inventing a disclosure. Such a
  // package remains a review draft, even when no generic gate item blocks it,
  // and each open finding is disclosed with its own reason and limitation. An
  // instruction left open is also asked about, once, with the review draft:
  // the conflict question names a conflicting one, the scope question the rest
  // the plan may leave out.
  const checklistOpen = checklistState?.status === "done" && checked?.valid && checked.open
  const paired = new Set(checked?.valid ? conflictPairs(checked.unresolved).flat() : [])
  const scoped = checklistOpen ? checked.openFindings.filter(row => instructionLeftOpen(directory, row) && !paired.has(row.item_id)) : []
  // The scope question covers the instructions no answered scope question has
  // settled (scopeAnswers). None is asked once two of the gate's own are
  // answered, or while another stage's is open. An instruction refers to an
  // earlier question only when that question settled or showed it; a plan-step
  // question shows it only when it showed every act it directs that the plan
  // step asks about, each for its own parties (gateScopeParties).
  const asked = [...context.questions.values()]
  const otherOpen = openScopeQuestion(asked, "gate")
  const quotes = (row: ChecklistOpenFinding) => row.item.source_refs.map(ref => ref.quote)
  const parties = gateScopeParties(directory)
  const shows = (question: Question, row: ChecklistOpenFinding) => scopeQuestionShows(question, quotes(row), parties(quotes(row)))
  const settled = scopeAnswers(scoped, quotes, asked.filter(answeredScopeQuestion), "gate", shows)
  const scope = otherOpen ? undefined : packageScopeQuestion(settled.ask, settled.id)
  const scopeAsks = (row: ChecklistOpenFinding) => settled.by(row)[0] ?? (otherOpen && shows(otherOpen, row) ? otherOpen : scope)
  const asks = (row: ChecklistOpenFinding) => paired.has(row.item_id) ? conflict?.id : scoped.includes(row) ? scopeAsks(row)?.id : undefined
  if (checklistOpen) items.push(...checked.openFindings.map(row => checklistOpenItem(directory, row, asks(row))))
  const questions = [conflict, scope].filter((row): row is Question => row !== undefined)
  const blocking = items.filter(item => item.status === "open")
  const disclosed = items.filter(item => item.status === "disclosed")
  // The gate's own questions ask about findings already disclosed here.
  const unanswered = [...context.questions.values()].some(row => row.status !== "answered" && row.stage !== "gate")
  return {
    status: blocking.length ? "blocked" : "pass",
    delivery: blocking.length ? null : unresolved.length ? "review_draft_with_findings" : disclosed.length || unanswered || checklistOpen ? "review_draft" : "execution_ready",
    blocking,
    disclosed,
    items,
    checked_at: new Date().toISOString(),
    ...(questions.length ? { questions } : {}),
  }
}

/** The notes delivered with a review draft that has unresolved checklist
 * findings: those findings first, each with its checklist item, the source it
 * records, its class, reason and limitation (the gate's own disclosure), then
 * the drafting notes unchanged. The drafting notes are written before the
 * checklist check, so they cannot list its findings. Undefined when there is
 * no such finding: the drafting notes are then delivered as written. */
export function deliveredNotes(directory: string, gate: GateResult): string | undefined {
  const findings = gate.items.filter(item => item.kind === "checklist_unresolved")
  if (!findings.length) return
  const checked = checklistOutputState(directory)
  const items = new Map<string, ChecklistItem>([...checked.unresolved.map(row => [row.item_id, row.item] as const),
    ...checked.unresolvedParents.map(row => [row.parent_id, row.item] as const)])
  const entries = findings.map(finding => {
    const id = finding.ref.checklist_item ?? finding.id
    const sources = (items.get(id)?.source_refs ?? []).map(ref => `${ref.document}, ${ref.locator}: "${ref.quote}"`)
    return `- ${id}: ${finding.detail}${sources.length ? ` Source: ${sources.join("; ")}` : ""}`.replace(/\s+/g, " ").trim()
  })
  return ["# Checklist findings not satisfied", "",
    "These checklist requirements were not established as satisfied. Each is disclosed with this review draft for review; none is treated as satisfied.", "",
    ...entries, "", readLeanText(directory, "drafting-notes.md") ?? ""].join("\n")
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
 * is the checker's; the gate only makes it visible. An instruction left open
 * refers to the question that names it. */
function checklistOpenItem(directory: string, { item_id, item, finding }: ChecklistOpenFinding, questionID?: string): OpenItem {
  const document = checklistDocument(directory, item)
  const detail = [`Checklist requirement left open: ${sentence(item.check)}`, ...findingText(finding)].join(" ")
  return { ...openItem("checklist-check", "checklist_open", detail, { checklist_item: item_id, ...(document ? { document } : {}),
    ...(questionID ? { question_id: questionID } : {}) }),
    status: "disclosed", resolution: "left open by the checklist check; disclosed with the review draft" }
}

/** A directive (an instruction from the request, the lawyer or the sources)
 * that the check left open rather than settled, and that the plan may leave
 * out. Whether the package carries it is the lawyer's call, not a disclosure.
 * One the plan carries is left open for a missing fact instead, which its
 * disclosure already shows: the saved plan's coverage record has its document
 * carrying, for the parties who must take it, an act for every quote it cites,
 * with none they direct left uncovered (coverageCarries), and the finding's own
 * body evidence from that document shows an open item. A bare underscore run is
 * a blank filled at execution ("Dated: ____"), not an open item; a bracketed
 * blank ("[____]") is. Either signal alone leaves it asked. */
function instructionLeftOpen(directory: string, { item, finding }: ChecklistOpenFinding) {
  if (item.basis !== "directive" || finding.status !== "open") return false
  const document = checklistDocument(directory, item)
  const unblanked = (text: string) => normalize(text).replace(/\[[^\]]*\]|_{3,}/g, run => run.startsWith("[") ? run : "")
  const marked = finding.evidence.some(row => row.kind === "body" && plannedID(directory, row.document) === document && paragraphShowsOpenItem(unblanked(row.text)))
  return !(document && marked && coverageCarries(directory, item.source_refs, document))
}

/** The gate's backstop for instructions the checklist check left open: one
 * question asking, for each, whether the package carries it and in which
 * document, or whether it is outside this request. The plan-side route for
 * instructed acts is the package-coverage check before drafting; this asks
 * only what reached the gate still open. Its `id`, from scopeAnswers, keys
 * on the verbatim instruction quotes, never on extraction-ordinal item ids, so
 * re-extraction asks the same question, and the caller leaves out the
 * instructions an answered package-scope question settled (scopeAnswers), and
 * asks none once two of the gate's own are answered or while another stage's
 * is open. */
function packageScopeQuestion(rows: ChecklistOpenFinding[], id: string): Question | undefined {
  if (!rows.length) return undefined
  const several = rows.length > 1
  return {
    id,
    stage: "gate",
    question: [`The checklist check left ${several ? "these instructions" : "this instruction"} open, so the package may not carry ${several ? "them" : "it"}:`,
      ...rows.map(({ item_id, item, finding }) => `${checklistRequirement(item_id, item)} ${finding.open_decisions?.length
        ? `Open decisions: ${sentence(finding.open_decisions.map(row => scopeField(row.trim())).join("; "))}` : "No open decision recorded; state what is missing."}`),
      `${several ? "For each, answer" : "Answer"} the open decisions above (supply any missing fact or approval they name). If the package should carry it, say in which document; otherwise say it is outside this request.`].join(" "),
    context: `Until a later check settles ${several ? "these instructions, they are" : "this instruction, it is"} disclosed with the review draft as left open; delivery is not held for ${several ? "them" : "it"}. Your answer re-runs the drafting from the start, including the document plan and the checklist check.`,
    status: "open",
  }
}

/** Why an adverse finding remains, in the lawyer's words. The runner's report
 * and the gate's disclosure share it. It reads recorded state only: whether
 * the one reinspection ran, the source adjudication that declined a repair,
 * and the checklist items this run's repair rounds submitted (repairs.json). */
export function checklistFindingLabel(directory: string, state: LeanState | undefined, row: { item_id: string; item: ChecklistItem; class: ChecklistFindingClass; reinspected: boolean }): string {
  switch (row.class) {
    case "omission":
      if (state && withheldChecklistRepair(directory, state, row.item_id, row.item)) return "omission; source adjudication declined the repair"
      if (repairedChecklistItems(readRepairRound(directory)).has(row.item_id)) return "omission not fixed by the repair round"
      // A later round may have replaced the checklist round; the cumulative
      // allowance still shows that no further round could run.
      return (readRepairRound(directory)?.substantive_attempts ?? 0) >= LEAN_MAX_SUBSTANTIVE_REPAIRS ? "omission; the repair allowance was used up" : "omission; no repair round ran"
    case "missing_inspection":
      return row.reinspected ? "not established by the check after reinspection" : "not established by the check; not yet reinspected"
    case "source_conflict": return "conflicts with another source requirement"
    case "not_inspectable": return "cannot be established from document text"
    case "departure": return "deliberate departure"
  }
}

const UNRESOLVED = "not established as satisfied; disclosed with the review draft"

/** A completed finding that remains adverse, disclosed with its class and its
 * own recorded reason and limitation. It is never a pass and never blocks. A
 * source conflict refers to the gate's question on it. */
function checklistUnresolvedItem(directory: string, state: LeanState | undefined, row: ChecklistUnresolvedFinding, questionID?: string): OpenItem {
  const { item_id, item, finding, class: kind, conflicting_item_ids } = row
  const document = checklistDocument(directory, item)
  const detail = [`Checklist requirement unresolved (${checklistFindingLabel(directory, state, row)}): ${sentence(item.check)}`, ...findingText(finding),
    conflicting_item_ids?.length ? `Conflicting checklist items: ${conflicting_item_ids.join(", ")}.` : ""].filter(Boolean).join(" ")
  return { ...openItem("checklist-check", "checklist_unresolved", detail, { checklist_item: item_id, checklist_class: kind, ...(document ? { document } : {}),
    ...(kind === "source_conflict" && questionID ? { question_id: questionID } : {}) }),
    status: "disclosed", resolution: UNRESOLVED }
}

/** One question for every source conflict among the adverse findings, so the
 * lawyer settles them together: each pair by its item ids, then each item's
 * requirement and sources. Its id is the pairs alone, so a gate re-run asks the
 * same question and an answered one is never asked again. */
function checklistConflictQuestion(rows: ChecklistUnresolvedFinding[]): Question | undefined {
  const items = new Map(rows.flatMap(row => [{ id: row.item_id, item: row.item }, ...(row.conflicting_items ?? [])]).map(row => [row.id, row.item]))
  const pairs = conflictPairs(rows)
  if (!pairs.length) return undefined
  const requirement = (id: string) => {
    const item = items.get(id)
    return item ? checklistRequirement(id, item) : id
  }
  return {
    id: `checklist-conflict-${shortHash(JSON.stringify(pairs))}`,
    stage: "gate",
    question: [`Checklist requirements taken from the sources conflict: ${pairs.map(([a, b]) => `${a} with ${b}`).join("; ")}.`,
      ...[...new Set(pairs.flat())].map(requirement),
      pairs.length > 1 ? "For each pair, which requirement governs, or how should they be reconciled?" : "Which requirement governs, or how should they be reconciled?"].join(" "),
    context: "Until you answer, each conflicting requirement is disclosed with the review draft as not satisfied; delivery is not held for it. Your answer re-runs the drafting from the start, including the checklist check.",
    status: "open",
  }
}

/** Each source conflict once, as its sorted pair of item ids, in a stable order. */
function conflictPairs(rows: ChecklistUnresolvedFinding[]) {
  return [...new Map(rows.flatMap(row => row.class === "source_conflict" ? (row.conflicting_item_ids ?? []).map(other => [row.item_id, other].toSorted()) : [])
    .map(pair => [JSON.stringify(pair), pair])).values()].toSorted((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)))
}

/** A requirement as the gate's questions list it: its id, check and sources.
 * The check is a model's words, so it goes through scopeField, as the scope
 * question's open decisions do: questionQuotes reads only the rows the gate
 * writes. The quotes are verbatim. */
function checklistRequirement(id: string, item: ChecklistItem) {
  return `${id}: ${sentence(scopeField(item.check))} Sources: ${item.source_refs.map(ref => `${ref.document} "${excerpt(ref.quote, GATE_EXCERPT)}"`).join(", ")}`
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
  return plannedID(directory, item.target.document)
}

/** The planned document a checklist record names, by its id or its file. */
function plannedID(directory: string, name: string) {
  return readLeanPlan(directory)?.documents.find(row => row.id === name || row.file === name)?.id
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
