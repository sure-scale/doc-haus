import type { ChecklistFindingClass } from "./checklist-classify"
import type { ChecklistFinding, ChecklistItem } from "./checklist-contract"
import { CHECKLIST_COVERAGE_ISSUES, type ChecklistCoverageIssue } from "./checklist-flow"
import { checklistOutputState, type ChecklistOpenFinding, type ChecklistUnresolvedFinding, type ChecklistUnresolvedParent } from "./stages"
import { readChangeReview, deferredFindingItems } from "./change-pipeline"
import { renderNotesCoverage } from "../../lib/notes-coverage-report"
import { hasOpenItemMarker } from "../../lib/open-items"
import { draftingPlaceholders } from "../../lib/placeholder-resolution"
import { normalize } from "../../lib/quote-verify"
import { CITED, finalBodyParagraphs, readerIDs, standingIDPattern, type ApplyResult } from "./checks"
import { memoMode, runItemIDs, withoutInternalIDs, withoutRefMarkers } from "./delivered-memo"
import { FIRST_DOCUMENT, readLeanPlan, type DocID } from "./plan"
import { isWithheldDeliveryItem, readPackageReview, withheldDeliveryItems } from "./package-review"
import { coverageCarries, coverageItems, gateScopeParties } from "./package-coverage"
import { excerpt, proseQuoteFound, openItem } from "./quotes"
import { answeredScopeQuestion, GATE_EXCERPT, keyedQuestionID, openScopeQuestion, scopeAnswers, scopeField, scopeQuestionShows } from "./package-scope"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, repairedChecklistItems } from "./repairs"
import { requestedFilesItems } from "./requested-files"
import { combinedFileItems } from "./deliverables"
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
//   checklist/checks.json, which the gate already reads, and in plain words in
//   the notes delivered with the draft, after any adverse finding;
// - a completed checklist finding that remains adverse (and a refined parent
//   whose coverage is adverse) is never a pass: it is disclosed, with the class
//   its verified fields establish, ahead of every other item, and in plain
//   words in the notes delivered with the draft: ahead of the drafting notes,
//   or closing a requested memo;
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
  const conflict = checked?.valid ? checklistConflictQuestion(conflictItemPairs(checked.unresolved)) : undefined
  // Adverse checklist findings lead: the lawyer's list of items may be capped.
  const unresolved = checked?.valid ? [...checked.unresolved.map(row => checklistUnresolvedItem(directory, leanState, row, conflict?.id)),
    ...checked.unresolvedParents.map(row => checklistUnresolvedParentItem(directory, row))] : []
  // A file name in the instructions the package review withheld is disclosed with the
  // draft, never blocking it: the notes were written before that review.
  const items = [...unresolved, ...[...currentOpenItems(directory), ...deferredFindingItems(directory), ...coverageItems(directory), ...requestedFilesItems(directory), ...combinedFileItems(directory)].map(item => settle(item, context)),
    ...semantic, ...projected, ...withheldDeliveryItems(directory)]
  if (checked && !checked.valid) items.push(openItem("checklist-check", "omission", "Effective checklist findings or source-unit coverage are incomplete or stale, or a check or its reinspection failed; see checklist/report.md.", {}))
  // Focused checks can correctly preserve a source-dependent open issue
  // without requesting a missing answer or inventing a disclosure. Such a
  // package remains a review draft, even when no generic gate item blocks it,
  // and each open finding is disclosed with its own reason and limitation. An
  // instruction left open is also asked about, once, with the review draft:
  // the conflict question names a conflicting one, the scope question the rest
  // the plan may leave out.
  const checklistOpen = checklistState?.status === "done" && checked?.valid && checked.open
  const paired = new Set(checked?.valid ? checked.unresolved.flatMap(row => row.class === "source_conflict" && row.conflicting_items?.length
    ? [row.item_id, ...row.conflicting_items.map(other => other.id)] : []) : [])
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

/** The notes delivered with the draft: the drafting notes as the reader gets
 * them (notesForReader), with any unresolved checklist findings and any
 * disclosed requirements the checklist check left open in plain words, and
 * any file name in the instructions the package review could not apply
 * (withheldDeliveryItems), in its item's words. The drafting notes are written
 * before the checklist check and the package review, so they cannot list
 * these. Undefined when there are neither drafting notes nor such an item.
 *
 * The findings lead the drafting notes, or close a requested memo after its
 * own title and body: the unresolved ones first, then those left open, then
 * the file names not applied, each in its own section. An unresolved one is disclosed as a requirement not
 * confirmed as met, one left open as left open; each is worded from the
 * checklist item's requirement and source quotes rather than the checker's
 * detail (findingEntries), so its lines name no internal id: no checklist
 * item, conflict list, request locator or paragraph pointer. */
export function deliveredNotes(directory: string, gate: GateResult): string | undefined {
  const drafted = readLeanText(directory, "drafting-notes.md")
  const read = drafted === undefined ? undefined : titled(directory, notesForReader(directory, drafted, gate.items))
  const unresolved = gate.items.filter(item => item.kind === "checklist_unresolved")
  // A requirement left open is listed only while the checklist output still
  // holds it open: one a later check settled is not read back from an earlier
  // gate's detail.
  const leftOpen = gate.items.filter(item => item.kind === "checklist_open" && item.status === "disclosed")
  const held = new Set(leftOpen.length ? checklistOutputState(directory).openFindings.map(row => row.item_id) : [])
  const open = leftOpen.filter(item => held.has(item.ref.checklist_item ?? item.id))
  const names = gate.items.filter(isWithheldDeliveryItem)
  if (!unresolved.length && !open.length && !names.length) return read
  const known = runItemIDs(directory, gate.items)
  // A section, with the blank line after it, only when it lists a finding.
  const section = (findings: OpenItem[], heading: string, intro: string) => findings.length
    ? [...plainWordsSection(directory, known, heading, intro, findingEntries(directory, findings, known)), ""] : []
  // A file name not applied is listed in its item's own words.
  const named = (heading: string, intro: string) => names.length
    ? [...plainWordsSection(directory, known, heading, intro, names.map(item => (hold: (text: string) => string) => `- ${hold(item.detail)}`)), ""] : []
  const body = read ?? ""
  if (!memoMode(directory)) return [...section(unresolved, "# Checklist findings not satisfied",
    "These checklist requirements were not established as satisfied. Each is disclosed with this review draft; none is treated as satisfied."),
  ...section(open, "# Checklist requirements left open",
    "These checklist requirements were left open. Each is disclosed with this review draft; none is treated as satisfied."),
  ...named("# File names not applied", "These names in your instructions could not be applied. Each is disclosed with this review draft."), body].join("\n")
  const memo = body.trimEnd()
  return [...(memo ? [memo, ""] : []),
    ...section(unresolved, "## Requirements not confirmed as met", "These requirements were not confirmed as met in this draft and need the lawyer's review."),
    ...section(open, "## Requirements left open", "These requirements were left open in this draft and need the lawyer's review."),
    ...named("## File names not applied", "These names in your instructions could not be applied in this draft.")].join("\n")
}

/** The drafting notes as the reader gets them: without their "[ref: …]"
 * markers, and with no id the notes check keeps from the reader (readerIDs,
 * with `items`). A line naming a planned document by its id names it by its
 * title (withDocumentTitles, with the notes as this memo or these notes), and
 * every other such id goes (withoutInternalIDs). A citation, footnote marker
 * or locator's number outside a quotation (CITED) is held aside from both, as
 * the check reads it, so "3(d1)" stays as written; quotations are kept as
 * withoutInternalIDs keeps them, except one holding only such an id, which goes
 * as the id does (a planned document's by its title). Notes the check passed
 * name none of these ids elsewhere, so beyond their markers and such a
 * quotation only notes saved on a last attempt, or kept on a resume, change;
 * every unchanged line keeps its exact text. */
export function notesForReader(directory: string, text: string, items: OpenItem[]) {
  const ids = readerIDs(directory, items)
  const documents = readLeanPlan(directory)?.documents ?? []
  const planned = documents.length ? new RegExp(standingIDPattern(documents.map(row => row.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"))) : undefined
  const self = memoMode(directory) ? "this memo" : "these notes"
  // Held text stands as n U+FDEE between two U+FDEF, as in plainWordsSection.
  const held: string[] = []
  const hold = (text: string) => `\uFDEF${"\uFDEE".repeat(held.push(text))}\uFDEF`
  const lines = withoutRefMarkers(text).split("\n").map(line => {
    const kept = line.split(/("[^"\n]*"|\u201c[^\u201d\n]*\u201d)/).map((part, index) => index % 2 ? part : CITED.reduce((part, mask) => part.replace(mask, hold), part)).join("")
    return planned?.test(kept) ? withDocumentTitles(kept, documents, ids, hold, self) : kept
  })
  return withoutInternalIDs(lines.join("\n"), ids).replace(/\uFDEF(\uFDEE+)\uFDEF/g, (_, count: string) => held[count.length - 1]!)
}

/** The notes as the reader gets them, under a plain title for what they are
 * when they have no heading of their own: notes saved on a last attempt, or
 * kept on a resume, may lose their only heading with its markers and ids. So
 * the delivered notes are never empty, as the gate's validity requires. */
function titled(directory: string, notes: string) {
  if (/^#{1,6}[ \t]+\S/m.test(notes)) return notes
  return `# ${memoMode(directory) ? "Memo" : "Drafting notes"}\n\n${notes.replace(/^(?:[ \t]*\r?\n)+/, "")}`
}

/** One section of the delivered notes in plain words: `heading`, the sentence
 * `intro` saying what it lists, then one line per entry. Each line names a
 * planned document by its title in place of its id, and the notes themselves
 * as this memo or these notes, then loses any `known` run id. The text an
 * entry holds aside (`hold`, such as its source quotes) and the titles put in
 * for ids are kept from both passes and put back after them, so each reaches
 * the reader as written: quotation marks inside one pair with nothing else on
 * the line. */
function plainWordsSection(directory: string, known: string[], heading: string, intro: string, entries: ((hold: (text: string) => string) => string)[]) {
  const documents = readLeanPlan(directory)?.documents ?? []
  const self = memoMode(directory) ? "this memo" : "these notes"
  return [heading, "", intro, "", ...entries.map(entry => {
    // The n-th text held aside stands as n U+FDEE between two U+FDEF:
    // noncharacters, never in document text, that neither pass reads.
    const held: string[] = []
    const hold = (text: string) => `\uFDEF${"\uFDEE".repeat(held.push(text))}\uFDEF`
    const line = entry(hold).replace(/\s+/g, " ").trim()
    return withoutInternalIDs(withDocumentTitles(line, documents, known, hold, self), known).replace(/\uFDEF(\uFDEE+)\uFDEF/g, (_, count: string) => held[count.length - 1]!)
  })]
}

/** Each unresolved or left-open checklist finding as a line for
 * plainWordsSection: its requirement, why it is unmet and its sources. The
 * request's locators (internal paths) give way to its label. A finding whose
 * item cannot be found takes its requirement from its detail, cut of its
 * conflict list, or falls back to that detail. Why it is unmet is a plain
 * status for its class, or that it was left open, then its reason. A checked
 * finding's reason is its own recorded reason, limitation and open decisions;
 * a refined parent's is one plain reason for each kind of coverage issue it
 * records, once each and in order, or one plain reason that part of it is
 * unmet when it records none. A finding that conflicts with another
 * requirement says so, without naming either. Its source quotes are held
 * aside. */
function findingEntries(directory: string, findings: OpenItem[], known: string[]) {
  const checked = checklistOutputState(directory)
  const items = new Map<string, ChecklistItem>([...[...checked.unresolved, ...checked.openFindings].map(row => [row.item_id, row.item] as const),
    ...checked.unresolvedParents.map(row => [row.parent_id, row.item] as const)])
  const conflicted = new Set(checked.unresolved.filter(row => row.conflicting_item_ids?.length).map(row => row.item_id))
  const recorded = new Map([...checked.unresolved, ...checked.openFindings].map(row => [row.item_id, row.finding]))
  const parents = new Map(checked.unresolvedParents.map(row => [row.parent_id, row]))
  return findings.map(finding => (hold: (text: string) => string) => {
    const id = finding.ref.checklist_item ?? finding.id
    const item = items.get(id)
    const sources = (item?.source_refs ?? []).map(ref => `${ref.document === "request" || ref.locator.startsWith(".dochaus/") ? ref.document
      : `${ref.document}, ${ref.locator}`}: ${hold(`"${ref.quote}"`.replace(/\s+/g, " "))}`)
    const detail = finding.detail.replace(/ Conflicting checklist items: \S+(?:, \S+)*\.$/, "")
    // Why the requirement is unmet, chosen by what the gate item is: a checked
    // finding has a class, one left open is its own kind, a refined parent has
    // neither. Each takes its reason from the checklist output the gate was
    // built from; only an unresolved item that output no longer holds is read
    // back from its detail (deliveredNotes lists one left open only while the
    // output holds it).
    const open = finding.kind === "checklist_open"
    const kind = finding.ref.checklist_class
    const own = kind || open ? recorded.get(id) : undefined
    const parent = kind || open ? undefined : parents.get(id)
    const read = own || parent || open ? undefined : readMemoFinding(detail, !kind, known)
    const status = open ? MEMO_OPEN_STATUS : memoStatus(kind, parent?.status ?? read?.coverage)
    const reasons = own ? [findingText(own).join(" ").replace(/^Finding: /, "")]
      : parent ? (parent.issue_kinds.length ? parent.issue_kinds.map(issue => MEMO_COVERAGE_REASON[issue]) : [MEMO_PARENT_REASON]) : read?.reasons ?? []
    const reason = [...new Set(reasons)].join(" ")
    const requirement = item ? sentence(item.check) : read?.requirement
    const why = [status, reason ? `Reason: ${reason}` : ""].filter(Boolean).join(" ")
    const text = requirement ? `${requirement}${why ? ` ${why}` : ""}` : detail
    // A source-conflict status already says so.
    const conflict = (conflicted.has(id) || detail !== finding.detail) && finding.ref.checklist_class !== "source_conflict" ? " It conflicts with another requirement." : ""
    return `- ${text}${conflict}${sources.length ? ` Source: ${sources.join("; ")}` : ""}`
  })
}

/** What a memo finding's class means for the reader and what the lawyer should
 * check, in place of the workflow's own label, which names its stages. */
const MEMO_STATUS: Record<ChecklistFindingClass, string> = {
  omission: "Status: not met in this draft; check whether the draft should be changed to meet it.",
  missing_inspection: "Status: could not be confirmed from the draft; check the draft against it.",
  source_conflict: "Status: conflicts with another requirement in the sources; check which one governs.",
  not_inspectable: "Status: cannot be confirmed from the document text; check it directly.",
  departure: "Status: the draft departs from it on purpose; check that the departure is intended.",
}

/** The plain status of a requirement the checklist check left open: neither
 * met nor unmet, with what is left open for the lawyer to settle. */
const MEMO_OPEN_STATUS = "Status: left open in this draft; settle what is left open and check the draft against it."

/** A memo finding's plain status: by its class, or for a refined parent by its
 * reconciled coverage; none when neither is recorded. */
function memoStatus(kind: ChecklistFindingClass | undefined, coverage: ChecklistUnresolvedParent["status"] | undefined) {
  if (kind) return MEMO_STATUS[kind]
  return coverage === "unsatisfied" ? MEMO_STATUS.omission : coverage === "indeterminate" ? MEMO_STATUS.missing_inspection : ""
}

/** Why a refined parent is unmet, in plain words for each kind of coverage
 * issue it records, saying only what that issue says, without the checker's
 * terms or the ids of the checks it names. */
const MEMO_COVERAGE_REASON: Record<ChecklistCoverageIssue, string> = {
  scope: "The checks it was broken into are incomplete or do not match it.",
  no_finding: "A check of part of it did not complete.",
  inspect: "Whether part of it was carried out, or applies, could not be confirmed from the draft.",
  adoption_evidence: "No current text or file in the draft shows that part of it was adopted.",
  deferred_property: "Part of it depends on a document-wide property that was not decided.",
}

/** Why a refined parent is unmet when it records no coverage issue: one of the
 * checks it was broken into is itself unmet or unconfirmed. */
const MEMO_PARENT_REASON = "Part of it was not confirmed as met."

/** A memo finding read back from its gate detail, for an item the checklist
 * output no longer holds: its coverage, its requirement and its reasons. The
 * detail's builder wrote the requirement first and the reason field ("Finding:"
 * for a checked finding, "Coverage:" for a refined parent) after it, so the
 * field is found from the end, and a requirement that itself names such a
 * field stays whole. */
function readMemoFinding(detail: string, parent: boolean, known: string[]) {
  const [, label, rest] = /^Checklist requirement unresolved(?: across its refined checks)? \(([^)]*)\): (.*)$/.exec(detail) ?? []
  if (rest === undefined) return undefined
  const coverage = label === "not satisfied" ? "unsatisfied" as const : label === "could not be established" ? "indeterminate" as const : undefined
  const field = parent ? " Coverage: " : " Finding: "
  const at = rest.lastIndexOf(field)
  if (at < 0) {
    const requirement = parent ? rest.replace(/ A refined check remains adverse\.$/, "") : rest
    return { coverage, requirement, reasons: requirement === rest ? [] : [MEMO_PARENT_REASON] }
  }
  const recorded = rest.slice(at + field.length)
  return { coverage, requirement: rest.slice(0, at), reasons: parent ? recorded.split(/(?<=[.!?]) +/).map(issue => memoCoverageReason(issue, known)) : [recorded] }
}

/** One coverage issue read back from a refined parent's detail: the plain
 * reason for its kind when it has the form its producer writes, else the issue
 * as written, without a leading run id that only names the check it concerns. */
function memoCoverageReason(issue: string, known: string[]) {
  const kinds = Object.keys(CHECKLIST_COVERAGE_ISSUES) as (keyof typeof CHECKLIST_COVERAGE_ISSUES)[]
  const kind = kinds.find(kind => {
    const text = CHECKLIST_COVERAGE_ISSUES[kind]
    return issue.endsWith(` ${text}`) && (kind === "deferred_property" ? /^[^\s:]+: [^\s:,]+(?:, [^\s:,]+)*$/ : /^[^\s:]+:$/).test(issue.slice(0, -text.length - 1))
  })
  if (kind) return MEMO_COVERAGE_REASON[kind]
  const [, label, rest] = /^([^\s:]+): (.+)$/.exec(issue) ?? []
  const text = label && known.includes(label.split("/")[0]!) ? rest! : issue
  return `${text.charAt(0).toUpperCase()}${text.slice(1)}`
}

/** The determiners that lead a document's title in place of its article. */
const DETERMINERS = String.raw`[Tt]he|[Aa]n?|[Tt]his|[Tt]hese|[Tt]hose|[Ee]ach|[Ee]very|[Aa]ny|[Ii]ts|[Tt]heir|[Oo]ur|[Yy]our|[Ww]hose`

/** A determiner ending the text before an id, or a possessive: the title then
 * takes no article. One modifier may stand between the determiner and the id
 * ("the inspected d1", "each delivered d1"); determinerLeads decides when that
 * word is not the id's. */
const DETERMINER = new RegExp(String.raw`\b(?:${DETERMINERS})[ \t]+(?:(?!(?:is|are|was|were|be|to|of|in|on|for|with|by|from|as|at|and|or|not|than|that|which)[ \t])[a-z-]+[ \t]+)?$|'s[ \t]+$`)

/** Whether a determiner (DETERMINER) leads the id after `before`. The one
 * modifier is not the id's when it heads a phrase of its own (PHRASE_HEAD), so
 * the id keeps its article; any other modifier is the id's ("under the
 * executed d1 is", "the inspected d1,"). */
function determinerLeads(before: string) {
  return DETERMINER.test(before) && !PHRASE_HEAD.test(before)
}

/** A modifier slot holding the head of its own phrase: after a preposition and
 * a determiner, one of the closed set of nouns that, so placed, open a clause
 * (complex subordinators: "in the event d1 and d2 are signed", "to the extent
 * d2's terms are silent", "in each case", "by the time"); or a plural noun the
 * ids stand beside ("the documents d1 and d2"), which no singular title takes
 * as a modifier. A word ending "-ss" or "-us" is not counted as plural. */
const PHRASE_HEAD = new RegExp(String.raw`(?:\b(?:[Ii]n|[Oo]n|[Aa]t|[Tt]o|[Bb]y|[Ff]or|[Ff]rom|[Ww]ith|[Oo]f|[Uu]nder|[Uu]pon)[ \t]+(?:${DETERMINERS})[ \t]+(?:event|extent|case|time|degree|moment)|\b(?:${DETERMINERS})[ \t]+[a-z-]*[a-rtv-z]s)[ \t]+$`)

/** A finding's wording with each planned document id, written by the checklist
 * for its own use, as the reader knows that document: by its title. Two ids
 * joined by an en dash ("d1–d3"), not in a longer run of joined ids, are a
 * range: every planned document from the first to the last in plan order,
 * named as one list ("the A, the B and the C") whose first title takes the
 * article a single id would. Two different ids joined by a hyphen ("d1-d3")
 * name just those two ("the A and the C"). Range ends that are reversed or the
 * same, a range over an untitled document and ids joined by "/" are each
 * named. A planned id joined by "/" or a dash to another of the run's `known`
 * ids (a change, a paragraph pointer, an untitled document) is joined to it by
 * "and" instead, so its title stands apart and the id removal names that id by
 * its kind. "notes" joined to a planned id by "/" or a dash is this memo, joined
 * to that document by "and" ("the A and this memo"). An id joined by a hyphen
 * to a word ("d1-specific") qualifies it and becomes the title alone. A
 * parenthesis holding only target ids (a planned document's, or "notes" for
 * the memo itself) labels the words before it: it is dropped when it names
 * only the memo, and otherwise names its documents ("notes" as this memo) but
 * the one whose title stands just before it; the others are then joined to
 * that title ("the A (and the B)"), not read as another name for it. Ids
 * joined by a dash there stay one label. An id between a determiner and a
 * capitalized name ("the d1 Lease") labels that name and follows it as its
 * title, unless the name already gives the title. Every other id becomes the
 * title, led by its own article or "the" unless a determiner leads it
 * (determinerLeads), and capitalized where a sentence or a field starts. A
 * period after an abbreviation, dotted ("e.g.") or a lowercase or
 * capitalized word with no vowel ("cl.", "v.", "Sch."), ends no sentence; an
 * acronym's ("GDPR.") may. A quotation, straight or curly, holding only
 * planned ids ("d2", “d1/d3”) names them; other quoted text is left to the id
 * removal that follows. Each title put in goes through `hold`, which keeps it
 * whole through that removal. "This memo" here is `self`, the name the notes
 * give themselves: "this memo" in a requested memo, otherwise "these notes". */
function withDocumentTitles(line: string, documents: { id: string; title: string }[], known: string[], hold: (title: string) => string, self: string) {
  const titles = new Map(documents.filter(row => row.title.trim()).map(row => {
    const [, article = "the", rest] = /^(?:(the|an?)[ \t]+)?(.*)$/i.exec(row.title.trim())!
    return [row.id, { article: article.toLowerCase(), rest: rest! }]
  }))
  if (!titles.size) return line
  const alternation = (rows: string[]) => rows.toSorted((a, b) => b.length - a.length).map(id => id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")
  const ids = alternation([...titles.keys()])
  // A titled id and the other run ids the id removal names, joined in one run.
  const any = alternation([...titles.keys(), ...known.filter(id => id.length > 1 && id !== "notes" && !titles.has(id))])
  const mixed = new RegExp(String.raw`(?<![\w:./-])(?:${any})(?:[/\u2013-](?:${any}))+(?![\w-]|[.:/]\w)`, "g")
  const joined = new RegExp(String.raw`${any}|[/\u2013-]`, "g")
  const label = `(?:${ids}|notes)`
  const labels = new RegExp(String.raw`[ \t]*\(${label}(?:[ \t]*,[ \t]*(?:(?:and|or)[ \t]+)?${label}|[ \t]+(?:and|or|&)[ \t]+${label}|[/\u2013-]${label})*\)`, "g")
  const memo = new RegExp(String.raw`(?<![\w:./-])${label}(?:[/\u2013-]${label})+(?![\w-]|[.:/]\w)`, "g")
  const standing = new RegExp(String.raw`(${standingIDPattern(ids)})([ \t]+[A-Z][\w'-]*(?:[ \t]+[A-Z\d][\w'-]*(?:\.\d+)*)*)?`, "g")
  const range = new RegExp(String.raw`(?<![/\u2013-])(${standingIDPattern(ids)})([\u2013-])(${standingIDPattern(ids)})(?![/\u2013-][a-z]\d)`, "g")
  const compound = new RegExp(String.raw`(?<![\w:./-])(${ids})-(?=[a-z]{2,}\b(?!-?\d))`, "g")
  const order = documents.map(row => row.id)
  // A title led by its own article or "the", or by the determiner before it;
  // its opening letter capitalized where a sentence or a field starts.
  const named = (id: string, led: boolean) => {
    const { article, rest } = titles.get(id)!
    return led ? rest : `${article} ${rest}`
  }
  const opening = (before: string, text: string) =>
    /^(?:-[ \t]+)?$|(?:(?<!\.[A-Za-z]|\b[^\W\da-zAEIOUY_]?[^\W\dA-Zaeiouy_]+)\.|[!?]|Reason:|Limitation:|Open decisions:)[ \t]+$/.test(before) ? text[0]!.toUpperCase() + text.slice(1) : text
  // Joints written "/" or with a dash between the memo and a document read "and".
  const memoJoint = (joint: string, left: string, right: string) => (left === "notes" || right === "notes") && /^[/\u2013-]$/.test(joint) ? " and " : joint
  // A quotation holding only planned ids is those ids, its marks dropped.
  const quoted = new RegExp(String.raw`^(?:${ids})(?:[/\u2013-](?:${ids}))*$`)
  const unquoted = line.replace(/"[^"\n]*"|\u201c[^\u201d\n]*\u201d/g, quote => quoted.test(quote.slice(1, -1)) ? quote.slice(1, -1) : quote)
  return unquoted.split(/("[^"\n]*"|\u201c[^\u201d\n]*\u201d)/).map((part, index) => index % 2 ? part : part
    .replace(mixed, run => {
      const parts = run.match(joined)!
      return parts.map((part, index) => index % 2 && titles.has(parts[index - 1]!) !== titles.has(parts[index + 1]!) ? " and " : part).join("")
    })
    .replace(labels, (match, offset: number, text: string) => {
      // Its labels, each with the joint before it; ids joined by a dash are
      // one, but the memo is not part of a range.
      const [, space, inside] = /^([ \t]*)\((.*)\)$/.exec(match)!
      const parts = inside!.split(/([ \t]*,[ \t]*(?:(?:and|or)[ \t]+)?|[ \t]+(?:and|or|&)[ \t]+|\/|[\u2013-](?=notes\b)|(?<=\bnotes)[\u2013-])/)
      const rows = parts.flatMap((label, index) => index % 2 ? [] : [{ joint: index ? parts[index - 1]! : "", label }])
      if (rows.every(row => row.label === "notes")) return ""
      const before = text.slice(0, offset).toLowerCase()
      const shown = rows.filter(row => !titles.has(row.label) || !before.endsWith(titles.get(row.label)!.rest.toLowerCase()))
      if (!shown.length) return ""
      // Labels left after the title's own is dropped join that title, and the
      // list keeps its last joint.
      const joints = shown.map(row => row.joint)
      if (shown.length > 1 && shown.at(-1) !== rows.at(-1)) joints[joints.length - 1] = rows.at(-1)!.joint
      const lead = shown.length === rows.length ? "" : rows.some(row => /\bor\b/.test(row.joint)) ? "or " : "and "
      return `${space}(${lead}${shown.map((row, index) => `${index ? memoJoint(joints[index]!, shown[index - 1]!.label, row.label) : ""}${row.label === "notes" ? self : row.label}`).join("")})`
    })
    .replace(memo, (match, offset: number, text: string) => {
      const parts = match.split(/([/\u2013-])/)
      if (!parts.includes("notes") || parts.every((part, index) => index % 2 || part === "notes")) return match
      const run = parts.map((part, index) => index % 2 ? memoJoint(part, parts[index - 1]!, parts[index + 1]!) : part === "notes" ? self : part).join("")
      return parts[0] === "notes" ? opening(text.slice(0, offset), run) : run
    })
    .replace(range, (match, first: string, joint: string, last: string, offset: number, text: string) => {
      const run = joint === "-" ? (first === last ? [] : [first, last]) : order.slice(order.indexOf(first), order.indexOf(last) + 1)
      if (run.length < 2 || !run.every(id => titles.has(id))) return match
      const before = text.slice(0, offset)
      const led = determinerLeads(before)
      const list = run.map((id, index) => named(id, led && !index))
      return hold(opening(before, `${list.slice(0, -1).join(", ")} and ${list.at(-1)}`))
    })
    .replace(compound, (_, id: string, offset: number, text: string) => `${hold(opening(text.slice(0, offset), titles.get(id)!.rest))}-`)
    .replace(standing, (_, id: string, name: string | undefined, offset: number, text: string) => {
      const before = text.slice(0, offset), after = text.slice(offset + id.length)
      const { article, rest } = titles.get(id)!
      const led = determinerLeads(before)
      if (led && name) return after.trimStart().toLowerCase().startsWith(rest.toLowerCase()) ? name.trimStart() : `${name.trimStart()} (${hold(`${article} ${rest}`)})`
      return `${hold(opening(before, named(id, led)))}${name ?? ""}`
    })).join("")
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
    // Besides the package's scope, it always asks for the open decisions' missing facts or approvals.
    kind: "scope",
    mixed: true,
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

/** The other checklist items an adverse finding conflicts with, as its
 * disclosure names them. */
function conflictText(ids: string[] | undefined) {
  return ids?.length ? `Conflicting checklist items: ${ids.join(", ")}.` : ""
}

/** A completed finding that remains adverse, disclosed with its class and its
 * own recorded reason and limitation. It is never a pass and never blocks. A
 * source conflict refers to the gate's question on it. */
function checklistUnresolvedItem(directory: string, state: LeanState | undefined, row: ChecklistUnresolvedFinding, questionID?: string): OpenItem {
  const { item_id, item, finding, class: kind, conflicting_item_ids } = row
  const document = checklistDocument(directory, item)
  const detail = [`Checklist requirement unresolved (${checklistFindingLabel(directory, state, row)}): ${sentence(item.check)}`, ...findingText(finding),
    conflictText(conflicting_item_ids)].filter(Boolean).join(" ")
  return { ...openItem("checklist-check", "checklist_unresolved", detail, { checklist_item: item_id, checklist_class: kind, ...(document ? { document } : {}),
    ...(kind === "source_conflict" && questionID ? { question_id: questionID } : {}) }),
    status: "disclosed", resolution: UNRESOLVED }
}

const CHECKLIST_CONFLICT_QUESTION = "checklist-conflict-"

/** One question for every source conflict among the adverse findings, so the
 * lawyer settles them together: each pair by its two requirements, each with
 * its check and sources, never an internal item id. Its id keys on each
 * requirement's source documents and verbatim quotes (conflictKey), never on
 * extraction-ordinal item ids, so a re-run or a re-extraction asks the same
 * question and an answered one is never asked again. */
function checklistConflictQuestion(pairs: ChecklistItem[][]): Question | undefined {
  if (!pairs.length) return undefined
  const several = pairs.length > 1
  return {
    id: keyedQuestionID(CHECKLIST_CONFLICT_QUESTION, pairs.map(pair => JSON.stringify(pair.map(conflictKey)))),
    stage: "gate",
    question: [several ? `Checklist requirements taken from the sources conflict, in ${pairs.length} pairs.` : "Two checklist requirements taken from the sources conflict.",
      ...pairs.map(([a, b], index) => `${several ? `Pair ${index + 1}, first` : "First"}: ${requirementText(a!)} Second: ${requirementText(b!)}`),
      several ? "For each pair, which requirement governs, or how should they be reconciled?" : "Which requirement governs, or how should they be reconciled?"].join(" "),
    context: `Until a later check no longer finds ${several ? "these conflicts, they are" : "this conflict, it is"} disclosed with the review draft; delivery is not held for ${several ? "them" : "it"}. Your answer re-runs the drafting from the start, including the checklist check.`,
    status: "open",
    kind: "conflict",
  }
}

/** A conflicting requirement as the conflict question's id knows it: its
 * source documents and normalized verbatim quotes, never its item id or its
 * check (a model's words), so re-extraction keeps it. Two requirements citing
 * the same source passages are one conflict for asking purposes. */
function conflictKey(item: ChecklistItem) {
  return JSON.stringify([...new Set(item.source_refs.map(ref => `${ref.document}\n${normalize(ref.quote)}`))].toSorted())
}

/** Each source conflict once, as its two requirements, in a stable order: by
 * their conflict keys, then their checks. An identical copy of a requirement
 * under another item id is the same requirement. */
function conflictItemPairs(rows: ChecklistUnresolvedFinding[]) {
  const order = (item: ChecklistItem) => JSON.stringify([conflictKey(item), item.check])
  const pairs = rows.flatMap(row => row.class === "source_conflict"
    ? (row.conflicting_items ?? []).map(other => [row.item, other.item].toSorted((a, b) => order(a).localeCompare(order(b)))) : [])
  return [...new Map(pairs.map(pair => [JSON.stringify(pair), pair])).values()]
    .toSorted((a, b) => JSON.stringify(a.map(order)).localeCompare(JSON.stringify(b.map(order))))
}

/** A requirement as the gate's questions list it: its id, check and sources. */
function checklistRequirement(id: string, item: ChecklistItem) {
  return `${id}: ${requirementText(item)}`
}

/** A requirement's check and sources. The check is a model's words, so it
 * goes through scopeField, as the scope question's open decisions do:
 * questionQuotes reads only the rows the gate writes. The quotes are verbatim. */
function requirementText(item: ChecklistItem) {
  return `${sentence(scopeField(item.check))} Sources: ${checklistSources(item)}`
}

function checklistSources(item: ChecklistItem) {
  return item.source_refs.map(ref => `${ref.document} "${excerpt(ref.quote, GATE_EXCERPT)}"`).join(", ")
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
  if (item.kind === "quote_miss" && item.ref.quote && proseQuoteFound(context.directory, { quote: item.ref.quote, source_id: item.ref.source_id ?? "" }))
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
