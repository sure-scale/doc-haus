import type { LeanDraftingStatus, LeanOpenItem, LeanStageStatus } from "../../api/ingest/lean"

// Pure view model for the drafting run header: plain step names, steps grouped per
// planned document, the run's state in four words, and open items grouped by
// kind. Everything is derived from the run's status; nothing names a document.

const STEP_NAMES: Record<string, string> = {
  discrepancy: "Assess source discrepancies",
  intake: "Resolve material questions",
  research: "Legal research",
  "research-review": "Check legal research",
  spine: "Build the drafting spine",
  draft: "Draft",
  review: "Review",
  "review-playbook": "Check applicable playbook",
  "review-independent": "Independent function and fidelity review",
  "review-wording": "Review problematic wording",
  "review-client-disadvantage": "Review client disadvantage",
  "change-plan": "Assess review findings",
  "change-prepare": "Prepare amendments",
  changes: "Reconcile amendments",
  "change-review": "Check amended document",
  apply: "Apply changes",
  finalize: "Drafting notes",
  "checklist-check": "Check source requirements against documents",
  "package-review": "Final semantic and package review",
  gate: "Final check",
}

const PREPARATION_STEPS = ["discrepancy", "intake", "research", "research-review", "spine"]

/** A run-level step ("finalize", "gate") or a document step ("d2.review"). */
export function stepName(step: string) {
  const stage = step.includes(".") ? step.slice(step.indexOf(".") + 1) : step
  return STEP_NAMES[stage] ?? sentence(stage)
}

/** A planned title as the lawyer should read it: no markdown code ticks. */
export function cleanTitle(title: string) {
  return title.replace(/`/g, "").replace(/\s+/g, " ").trim()
}

export type Step = { key: string; name: string; row: LeanStageStatus }
export type StepGroup = { id: string; title: string; steps: Step[] }

/** Steps in walk order, grouped by the document they draft; the run-level
 * steps (drafting notes, final check) close the list as "Finishing". */
export function stepGroups(status: LeanDraftingStatus) {
  const planned = status.plan?.documents ?? []
  return Object.entries(status.stages).reduce<StepGroup[]>((groups, [key, row]) => {
    const id = key.includes(".") ? key.slice(0, key.indexOf(".")) : PREPARATION_STEPS.includes(key) ? "preparation" : "run"
    const step = { key, name: stepName(key), row }
    const last = groups.at(-1)
    if (last?.id === id) {
      last.steps.push(step)
      return groups
    }
    return [...groups, { id, title: groupTitle(id, planned, groups.filter(group => group.id !== "preparation" && group.id !== "run").length), steps: [step] }]
  }, [])
}

function groupTitle(id: string, planned: Array<{ id: string; title: string }>, position: number) {
  if (id === "preparation") return "Preparation"
  if (id === "run") return "Finishing"
  const title = planned.find((row) => row.id === id)?.title
  return title ? cleanTitle(title) || `Document ${position + 1}` : `Document ${position + 1}`
}

/** The step the run is in or stopped at: its key, 1-based position and the total. */
export function currentStep(status: LeanDraftingStatus) {
  const keys = Object.keys(status.stages)
  const index = keys.findIndex((key) => status.stages[key]!.status !== "done")
  return index === -1 ? undefined : { key: keys[index]!, position: index + 1, total: keys.length }
}

/** Where the run is as the lawyer counts it: "Document 1 of 2 · Review", or
 * "Finishing · Final check". The total is the plan's, so it never grows mid-run;
 * `document` is the planned title when the run drafts several documents. */
export function runPosition(status: LeanDraftingStatus) {
  const step = currentStep(status)
  if (!step) return undefined
  const dot = step.key.indexOf(".")
  if (dot === -1) return { text: `${PREPARATION_STEPS.includes(step.key) ? "Preparation" : "Finishing"} · ${stepName(step.key)}` }
  const planned = status.plan?.documents ?? []
  const at = planned.findIndex((row) => row.id === step.key.slice(0, dot))
  if (at === -1 || planned.length < 2) return { text: stepName(step.key) }
  return { text: `Document ${at + 1} of ${planned.length} · ${stepName(step.key)}`, document: cleanTitle(planned[at]!.title) }
}

export type RunState = "running" | "input" | "paused" | "done"

/** The lawyer stopped this run: it is finished, asks nothing more and takes no
 * Continue or update. Its record stays; the next drafting request starts anew. */
export function leanRunStopped(status: LeanDraftingStatus) {
  return status.stopped === true
}

/** Stop is offered for an unfinished run that nothing is working on: paused,
 * held, or left running by an ingest restart. */
export function leanStopAvailable(status: LeanDraftingStatus) {
  return (status.status === "blocked" || status.status === "running") && !status.in_flight
}

/** The composer sends updates to an open run; once the run completed or was
 * stopped, a message there is a new request. */
export function leanRunTakesUpdates(status: LeanDraftingStatus) {
  return status.status !== "completed" && !leanRunStopped(status)
}

export function openQuestions(status: LeanDraftingStatus) {
  // A stopped run's questions stay in its record but are no longer asked.
  if (leanRunStopped(status)) return []
  const queued = new Set(status.queued_answer_ids ?? [])
  return status.questions.filter((question) => question.status === "open" && !queued.has(question.id))
}

/** Every stop short of completion is recoverable from the run header (Continue or an
 * answer), so it reads as paused, never as an error. */
export function runState(status: LeanDraftingStatus): RunState {
  if (status.status === "completed") return "done"
  if (status.status === "running" && status.in_flight) return "running"
  if (openQuestions(status).length) return "input"
  return "paused"
}

/** How long a step has run: to its finish, or to `now` while it is running. */
export function stageElapsed(row: LeanStageStatus, now: number) {
  if (!row.started_at) return undefined
  const end = row.finished_at ? Date.parse(row.finished_at) : row.status === "running" ? now : undefined
  return end === undefined ? undefined : Math.max(0, end - Date.parse(row.started_at))
}

/** A step past this has been working a while; the header reassures, not alarms. */
export const LONG_STEP_MS = 20 * 60 * 1000

/** Why the run stopped at this step, for the step's own row. The run records
 * the cause as "<step>: <reason>"; the final check is summarised by its count. */
export function stepCause(status: LeanDraftingStatus, key: string) {
  const row = status.stages[key]
  if (!row || (row.status !== "failed" && row.status !== "blocked")) return undefined
  if (key === "gate" && status.gate?.status === "blocked") {
    const count = status.gate.blocking.length
    return `${count} open item${count === 1 ? "" : "s"} still to resolve.`
  }
  const prefix = `${key}: `
  return status.stop_reason?.startsWith(prefix) ? status.stop_reason.slice(prefix.length) : undefined
}

const KIND_LABELS: Record<string, string> = {
  quote_miss: "Quotes not found in the sources",
  research_gap: "Legal research limitations",
  citation_unsupported: "Citations without support",
  placeholder: "Placeholders to fill",
  unapplied_amendment: "Changes not applied",
  question: "Questions for you",
  omission: "Points left out",
  template_removal: "Template text removed",
  checklist_open: "Requirements left open by the checklist check",
  checklist_unresolved: "Unresolved checklist findings",
}

/** Where a disclosed item is recorded for the lawyer. An open checklist finding
 * is disclosed by the final check itself, not by the drafting notes; an
 * unresolved one is a requirement the delivered draft does not satisfy. */
export function openItemTag(item: LeanOpenItem) {
  if (item.status !== "disclosed") return undefined
  if (item.kind === "checklist_unresolved") return "Not satisfied"
  return item.kind === "checklist_open" ? "Left open" : "In the notes"
}

// The class only. Whether a repair round or reinspection ran is in the item's
// detail, as the final check worded it.
const CHECKLIST_CLASS_LABELS: Record<string, string> = {
  omission: "left out of the draft",
  source_conflict: "conflicts with another source requirement",
  not_inspectable: "cannot be established from document text",
  missing_inspection: "not established by the check",
  departure: "deliberate departure",
}

/** Why an unresolved checklist finding stayed unresolved, in plain words. */
export function openItemClass(item: LeanOpenItem) {
  if (item.kind !== "checklist_unresolved" || !item.ref?.checklist_class) return undefined
  return CHECKLIST_CLASS_LABELS[item.ref.checklist_class] ?? sentence(item.ref.checklist_class).toLowerCase()
}

export type OpenItemGroup = { kind: string; label: string; items: LeanOpenItem[] }

/** Open items grouped by kind, in the order each kind first appears, except
 * that unresolved checklist findings always come first. */
export function groupOpenItems(items: LeanOpenItem[]) {
  const groups = items.reduce<OpenItemGroup[]>((groups, item) => {
    const group = groups.find((row) => row.kind === item.kind)
    if (group) {
      group.items.push(item)
      return groups
    }
    return [...groups, { kind: item.kind, label: KIND_LABELS[item.kind] ?? sentence(item.kind), items: [item] }]
  }, [])
  const first = (group: OpenItemGroup) => group.kind === "checklist_unresolved"
  return [...groups.filter(first), ...groups.filter((group) => !first(group))]
}

/** The first `cap` items across the groups, in display order, so the first
 * groups stay whole; each group keeps its full count as `total`. */
export function capGroups(groups: OpenItemGroup[], cap: number) {
  return groups.reduce<{ left: number; rows: Array<OpenItemGroup & { total: number }> }>(
    (acc, group) => ({
      left: Math.max(0, acc.left - group.items.length),
      rows: acc.left > 0 ? [...acc.rows, { ...group, items: group.items.slice(0, acc.left), total: group.items.length }] : acc.rows,
    }),
    { left: cap, rows: [] },
  ).rows
}

/** The items still worth the lawyer's attention: as the final check settled them
 * once it ran for this pass (blocking or disclosed), else the run's open items
 * and any checklist requirement recorded as not satisfied. */
export function attentionItems(status: LeanDraftingStatus) {
  const ran = status.stages.gate?.status === "done" || status.stages.gate?.status === "blocked"
  if (ran && Array.isArray(status.gate?.items)) return status.gate.items.filter((item) => item.status !== "resolved")
  return status.open_items.filter((item) => item.status === "open" || isUnresolvedFinding(item))
}

function isUnresolvedFinding(item: LeanOpenItem) {
  return item.kind === "checklist_unresolved" && item.status !== "resolved"
}

/** The checklist requirements the delivered draft does not satisfy, as the final
 * check settled them (or as the run recorded them before the check ran). */
export function unresolvedChecklistFindings(status: LeanDraftingStatus) {
  return (leanSettledItems(status) ?? status.open_items).filter(isUnresolvedFinding)
}

/** How many unresolved checklist findings the final check delivered the draft
 * with; undefined unless the run carries any or its delivery says so, in which
 * case 0 means the count did not reach the status. */
export function checklistFindingsCount(status: LeanDraftingStatus) {
  const count = unresolvedChecklistFindings(status).length
  return count > 0 || status.gate?.delivery === "review_draft_with_findings" ? count : undefined
}

/** "2 unresolved checklist findings"; without a count, the plain plural. */
export function unresolvedFindingsPhrase(count: number) {
  if (count === 0) return "unresolved checklist findings"
  return `${count} unresolved checklist finding${count === 1 ? "" : "s"}`
}

/** The lawyer-facing title of a review draft delivered with unresolved findings. */
export function findingsTitle(count: number) {
  return `Review draft — ${unresolvedFindingsPhrase(count)}`
}

/** "The draft does not satisfy those requirements." */
export function findingsConsequence(count: number) {
  return `The draft does not satisfy ${count === 1 ? "that requirement" : "those requirements"}.`
}

/** The open items in one sentence: unresolved checklist findings first, then
 * questions for the lawyer apart from the placeholders to fill; undefined while
 * there is nothing to report. */
export function openItemsSentence(status: LeanDraftingStatus) {
  const asks = openQuestions(status).length
  const all = attentionItems(status).filter((item) => item.kind !== "question")
  const unresolved = all.filter(isUnresolvedFinding).length
  const items = all.filter((item) => !isUnresolvedFinding(item))
  const placeholders = items.filter((item) => item.kind === "placeholder").length
  const parts = [
    unresolved ? unresolvedFindingsPhrase(unresolved) : "",
    asks ? `${asks} question${asks === 1 ? "" : "s"} for you` : "",
    placeholders ? `${placeholders} placeholder${placeholders === 1 ? "" : "s"} to fill` : "",
    items.length > placeholders ? `${items.length - placeholders} other item${items.length - placeholders === 1 ? "" : "s"} to check` : "",
  ].filter(Boolean)
  if (parts.length === 0) return status.status === "completed" ? "No open items remain." : undefined
  const list = parts.length === 1 ? parts[0] : `${parts.slice(0, -1).join(", ")} and ${parts.at(-1)}`
  return `Open items: ${list}.`
}

export type DraftedOutput = { name: string; role: "draft" | "redline" | "notes" }
export type DraftedGroup = { main: DraftedOutput; extras: DraftedOutput[] }

export const OUTPUT_ROLE: Record<DraftedOutput["role"], string> = { draft: "Draft", redline: "Redline", notes: "Notes" }

/** The Word files a completed run delivered: each drafted document with its
 * redline beneath it; the drafting notes sit under the draft when the run made
 * one document, otherwise in their own row. */
export function draftedGroups(status: LeanDraftingStatus | undefined) {
  if (status?.documents?.status !== "completed") return []
  const documents = status.documents.documents.filter((row) => row.deliverable)
  const groups = documents
    .filter((row) => row.kind === "body")
    .map((body): DraftedGroup => ({
      main: { name: body.deliverable!, role: "draft" },
      extras: documents
        .filter((row) => row.kind === "redline" && row.document === body.document)
        .map((row) => ({ name: row.deliverable!, role: "redline" })),
    }))
  const notes = documents
    .filter((row) => row.kind === "drafting_notes" || row.kind === "drafting_notes_markdown")
    .map((row): DraftedOutput => ({ name: row.deliverable!, role: "notes" }))
  if (groups.length === 1) return [{ ...groups[0]!, extras: [...groups[0]!.extras, ...notes] }]
  return [...groups, ...notes.map((main) => ({ main, extras: [] }))]
}

/** The matter's generated files as drafted groups: what the run's status names
 * first, then any other deliverable by the run's naming ("<name> - redline",
 * "<name> - drafting notes"), so files from earlier runs group the same way. */
export function deliverableGroups(names: string[], status: LeanDraftingStatus | undefined) {
  const present = new Set(names)
  const known = draftedGroups(status)
    .filter((group) => present.has(group.main.name))
    .map((group) => ({ ...group, extras: group.extras.filter((output) => present.has(output.name)) }))
  const placed = new Set(known.flatMap((group) => [group.main.name, ...group.extras.map((output) => output.name)]))
  const rest = names.filter((name) => !placed.has(name))
  const grouped = rest
    .filter((name) => outputRole(name) === "draft")
    .map(
      (name): DraftedGroup => ({
        main: { name, role: "draft" },
        extras: rest
          .filter((other) => outputRole(other) === "redline" && other.replace(REDLINE, "$1") === name)
          .map((other) => ({ name: other, role: "redline" })),
      }),
    )
  const used = new Set(grouped.flatMap((group) => [group.main.name, ...group.extras.map((output) => output.name)]))
  const loose = rest.filter((name) => !used.has(name)).map((name): DraftedOutput => ({ name, role: outputRole(name) }))
  const groups = [...known, ...grouped]
  // One draft keeps its notes beneath it, as a one-document run delivers them.
  if (groups.length === 1 && loose.every((output) => output.role === "notes"))
    return [{ ...groups[0]!, extras: [...groups[0]!.extras, ...loose] }]
  return [...groups, ...loose.map((main) => ({ main, extras: [] }))]
}

const REDLINE = / - redline(\.[a-z]+)$/i

function outputRole(name: string): DraftedOutput["role"] {
  if (REDLINE.test(name)) return "redline"
  if (/ - drafting notes\.[a-z]+$/i.test(name)) return "notes"
  return "draft"
}

function sentence(value: string) {
  const words = value.replace(/[_-]+/g, " ").trim()
  return words ? `${words[0]!.toUpperCase()}${words.slice(1)}` : value
}

/** Continue is offered when the run stopped (blocked or failed) or when its
 * state says running but no coordinator is running it (ingest restarted). */
export function leanContinueAvailable(status: LeanDraftingStatus) {
  if (leanRunStopped(status)) return false
  if (status.status === "blocked" || status.status === "failed") return true
  return status.status === "running" && !status.in_flight
}

/** The open items as the final check settled them, once the check has run for
 * the current pass; null before then, when the run's own open items are the count. */
export function leanSettledItems(status: LeanDraftingStatus) {
  const ran = status.stages.gate?.status === "done" || status.stages.gate?.status === "blocked"
  return ran && Array.isArray(status.gate?.items) ? status.gate.items : null
}

export function leanSettledSummary(items: LeanOpenItem[]) {
  const count = (state: LeanOpenItem["status"]) => items.filter((item) => item.status === state).length
  const checklist = items.filter((item) => item.status === "disclosed" && item.kind === "checklist_open").length
  const unresolved = items.filter((item) => item.status === "disclosed" && item.kind === "checklist_unresolved").length
  const parts = [
    [count("open"), "still blocking"],
    [unresolved, "not satisfied (unresolved checklist findings)"],
    [count("disclosed") - checklist - unresolved, "disclosed in the drafting notes"],
    [checklist, "left open by the checklist check"],
    [count("resolved"), "resolved"],
  ] as const
  return `Final check settled ${items.length} open item${items.length === 1 ? "" : "s"}: ${parts
    .filter(([n]) => n > 0)
    .map(([n, label]) => `${n} ${label}`)
    .join(", ")}.`
}

const DELIVERY: Record<string, string> = {
  execution_ready: "ready for signature",
  review_draft: "a review draft, not yet ready for signature",
}

/** The final check's verdict for a delivered run. A draft changed after the
 * check is not what it checked, so the verdict is not repeated for it. A draft
 * delivered with unresolved checklist findings did not pass clean. */
export function finalCheckNote(status: LeanDraftingStatus) {
  if (!status.gate?.delivery) return undefined
  if (status.changed_after_check?.documents.length) return "Changed after final check: review the draft."
  const findings = checklistFindingsCount(status)
  if (findings !== undefined)
    return `Final workflow check: a review draft with ${unresolvedFindingsPhrase(findings)}, not yet ready for signature. ${findingsConsequence(findings)}`
  return `Final workflow check passed: ${DELIVERY[status.gate.delivery] ?? status.gate.delivery.replace(/_/g, " ")}.`
}

/** Admission and execution are distinct: only explicit runner progress claims application. */
export function updateProgressLabel(update: NonNullable<LeanDraftingStatus["updates"]>[number]) {
  if (update.kind === "clarification") {
    if (update.progress === "applied") return "Clarification reviewed — current final checks completed."
    if (update.progress === "applying") return "Clarification is being inspected against the current documents."
    if (update.progress === "requires-rework") return "Clarification received — affected checks remain incomplete."
    if (update.status === "queued" || update.progress === "received") return "Clarification received — waiting for checklist inspection."
    return "Clarification included for inspection; delivery remains subject to final checks."
  }
  if (update.progress === "applied") return "Update applied — affected drafting and final checks completed."
  if (update.progress === "applying") return "Update is being applied — affected work is running again."
  if (update.progress === "requires-rework") return "Update requires rework — affected drafting or checks remain incomplete."
  if (update.status === "queued" || update.progress === "received") return "Update received — waiting for the current step to finish."
  return "Included in drafting instructions."
}
