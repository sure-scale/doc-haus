import { checkChanges, parseJsonObject } from "./change-checks"
import { readChangeList, readParagraphs, record, type ChangeList } from "./checks"
import { uniqueReplacementIssues, readChangePlan, type ChangePlan, type ChangePreparation } from "./change-pipeline"
import { checklistProducerInputs } from "./checklist-source-context"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { packageCompositionBlock, retryPrompt } from "./prompts"
import { matterSourceIDs, openItem, quoteFound } from "./quotes"
import { drainLeanRequests } from "./requests"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, stageInputs, stepOutputs, stepSessionTitle, type LeanDeps } from "./stages"
import { hashStageInputs, readLeanJson, readQuestions, writeLeanText, type LeanState, type LeanStageState } from "./state"
import type { DocID } from "./plan"
import type { Amendment } from "../../lib/drafting-amendment"

type Component = { id: string; pids: string[]; groups: ChangePlan["groups"]; prior_only: string[]; changes: ChangeList }
const MAX_CONFLICT_CONTEXT = 60_000
export type ReconciliationWork = { direct: ChangeList; conflicts: Component[] }

/** Preparation already supplies source-checked operations. Revisit only a
 * conflicting effect, a prior operation being withdrawn, or affected commentary. */
export function reconciliationWork(directory: string, document: DocID, plan: ChangePlan, preparation: ChangePreparation): ReconciliationWork {
  const rows = new Map(readParagraphs(directory, document).map(row => [row.pid, row]))
  const previous = readChangeList(directory, document)
  const valid = (row: Amendment) => rows.get(row.pid)?.sha256 === row.existing_sha256
  if (preparation.groups.length !== plan.groups.length || plan.groups.some(group => preparation.groups.filter(row => row.group === group.id).length !== 1)) throw Error("Every adjudicated group requires exactly one accepted preparation result.")
  for (const group of plan.groups) {
    const prepared = preparation.groups.find(row => row.group === group.id)!
    const checked = checkChanges(directory, JSON.stringify(prepared.changes), [...rows.values()].filter(row => group.pids.includes(row.pid)), { document })
    if (checked.issues.length) throw Error(`Prepared group ${group.id} is not current and valid: ${checked.issues.join(" ")}`)
  }
  const components: Component[] = []
  for (const group of plan.groups) {
    const connected = components.filter(component => component.pids.some(pid => group.pids.includes(pid)))
    const component: Component = { id: group.id, pids: [...new Set([...group.pids, ...connected.flatMap(row => row.pids)])], groups: [group, ...connected.flatMap(row => row.groups)], prior_only: [], changes: emptyChanges() }
    connected.forEach(row => components.splice(components.indexOf(row), 1))
    components.push(component)
  }
  const assigned = new Set(components.flatMap(component => component.pids))
  const direct = emptyChanges()
  direct.amendments = previous?.amendments.filter(row => valid(row) && !assigned.has(row.pid)) ?? []
  direct.explanations = previous?.explanations?.filter(row => rows.has(row.pid) && !assigned.has(row.pid)) ?? []
  direct.retained_initial = previous?.retained_initial?.filter(row => rows.has(row.pid) && !assigned.has(row.pid)) ?? []
  direct.withheld = previous?.withheld?.filter(row => valid(row.amendment) && !assigned.has(row.amendment.pid)) ?? []
  direct.questions = previous?.questions.filter(row => (!row.pid || rows.has(row.pid)) && readQuestions(directory).some(saved => saved.id === row.id && saved.status === "open")) ?? []
  for (const component of components) {
    const prepared = preparation.groups.filter(row => component.groups.some(group => group.id === row.group))
    if (prepared.length !== component.groups.length) throw Error("Every adjudicated group requires exactly one accepted preparation result.")
    const lists = prepared.map(row => row.changes)
    const old = previous && {
      ...previous, amendments: previous.amendments.filter(row => valid(row) && component.pids.includes(row.pid)),
      retained_initial: previous.retained_initial?.filter(row => rows.has(row.pid) && component.pids.includes(row.pid)),
      explanations: previous.explanations?.filter(row => rows.has(row.pid) && component.pids.includes(row.pid)),
      withheld: previous.withheld?.filter(row => valid(row.amendment) && component.pids.includes(row.amendment.pid)), questions: [],
    }
    component.prior_only = old?.amendments.filter(row => !lists.some(list => list.amendments.some(candidate => candidate.pid === row.pid))).map(row => row.pid) ?? []
    component.changes = combine([...(old ? [old] : []), ...lists])
    const byPID = Map.groupBy(component.changes.amendments, row => row.pid)
    const conflicting = [...byPID.values()].some(operations => new Set(operations.map(operationKey)).size > 1)
    const hasWithheld = Boolean(component.changes.withheld?.length)
    const changedExplanation = [...(component.changes.explanations ?? []), ...(component.changes.retained_initial ?? [])].some(explanation => component.changes.amendments.some(row => row.pid === explanation.pid))
    if (!conflicting && !changedExplanation && !hasWithheld && !component.prior_only.length) {
      component.changes.amendments = [...byPID.values()].map(coalesce)
      Object.assign(direct, combine([direct, component.changes]))
    }
  }
  return { direct, conflicts: components.filter(component => {
    const groups = Map.groupBy(component.changes.amendments, row => row.pid)
    return Boolean(component.prior_only.length || component.changes.withheld?.length) || [...groups.values()].some(operations => new Set(operations.map(operationKey)).size > 1) || [...(component.changes.explanations ?? []), ...(component.changes.retained_initial ?? [])].some(explanation => groups.has(explanation.pid))
  }) }
}

export function reconciliationPrompt(directory: string, document: DocID, component: Component) {
  const request = checklistProducerInputs(directory).find(row => row.key === "request")!.input
  const plan = readChangePlan(directory, document)
  const composition = packageCompositionBlock(directory, document)
  const prompt = [
    "Resolve ONE prepared dependency component into one cumulative operation per listed initial paragraph. Do not independently review the matter or rewrite other components. Prepared operations, goals and prior changes are generated proposals, not authority; verify necessary source conditions with targeted tools. Actual instructions and answers govern. Preserve every supported correction and unaffected wording; a proposal is not approval or adoption.",
    `<actual_instructions>\n${JSON.stringify([...request.passages, ...request.instructions])}\n</actual_instructions>`,
    ...(composition ? [composition] : []),
    `<component_goals>\n${JSON.stringify(component.groups)}\n</component_goals>`,
    `<shared_dependencies>\n${JSON.stringify(plan?.shared_decisions ?? [])}\n</shared_dependencies>`,
    `<initial_targets>\n${JSON.stringify(readParagraphs(directory, document).filter(row => component.pids.includes(row.pid)))}\n</initial_targets>`,
    `<competing_operations_and_explanations>\n${JSON.stringify(component.changes)}\n</competing_operations_and_explanations>`,
    `Original source inventory: ${JSON.stringify(matterSourceIDs(directory))}. Use search-document/read-document view=source with exact filename. Current documents use view=current_body, comments/deleted_text separately. Read-record .dochaus/lean/${document}/paragraphs.json supplies initial targets. Retrieve only necessary conditions or related document passages. Other components are not amendment targets.`,
    'Return {"amendments":[{"pid":"listed initial pid","operation":"replace|delete|split (omit for replace)","updated":"complete replacement text, or array for split","reason":"combined supported effect","sources":[{"source_id":"canonical source or request","quote":"exact words"}]}],"explanations":[{"pid":"listed initial pid","reason":"actual source-to-initial revision, corrected for resulting effect","sources":[]}],"questions":[{"pid":"listed initial pid","question":"indispensable unresolved human decision","context":"why needed"}],"dependencies":[]}. Also return "retained_initial":[{"pid":"listed initial pid","reason":"source-grounded reason an earlier operation is rejected and initial wording retained","sources":[{"source_id":"canonical source or request","quote":"exact words"}]}] when withdrawing an earlier operation; this is a recorded disposition, never a no-op edit. Every listed candidate PID needs a resulting operation, explicit retained_initial disposition, or an explicit dependencies entry explaining why it cannot safely be reconciled. dependencies holds a newly discovered cross-component dependency or unsupported candidate needing renewed source adjudication; name affected PIDs/components and specific issue. Such work stops publication, never silently drops an operation. Source questions retain visible [open item: ...] carriers. Unsupported drafting judgment retains a new visible confirmation marker. Do not emit no-op edits to create comments. Never turn missing approval into adoption. No file writes.',
  ].join("\n\n")
  if (prompt.length > MAX_CONFLICT_CONTEXT) throw Error(`Dependency component ${component.id} needs decomposition before model admission; no whole-matter fallback is available.`)
  return prompt
}

export async function runChangeReconciliation(directory: string, state: LeanState, deps: LeanDeps, document: DocID, progress: (patch: Partial<LeanStageState>) => void) {
  const plan = readChangePlan(directory, document)
  const preparation = readLeanJson(directory, `${document}/change-preparation.json`) as ChangePreparation | undefined
  if (!plan || !preparation?.groups) return { status: "failed" as const, reason: "Accepted change plan and prepared groups are required before reconciliation." }
  const work = reconciliationWork(directory, document, plan, preparation)
  const step = `${document}.changes` as const
  const key = hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))
  const results: ChangeList[] = []
  const failures: string[] = []
  const counts = { next: 0, calls: 0, attempts: 0, sessions: [] as string[] }
  await Promise.all(Array.from({ length: Math.min(4, work.conflicts.length) }, async () => {
    for (;;) {
      const component = work.conflicts[counts.next++]
      if (!component) return
      try {
        const initialPrompt = reconciliationPrompt(directory, document, component)
        const agent = LEAN_STAGE_AGENTS.changes
        const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, model: state.model, title: stepSessionTitle(directory, step, component.id), permissions: LEAN_CHILD_PERMISSIONS })
        counts.sessions.push(sessionID)
        let issues: string[] = []
        for (let index = 0; index <= LEAN_MAX_RETRIES; index++) {
          counts.calls++; counts.attempts = Math.max(counts.attempts, index + 1)
          progress({ session_id: counts.sessions[0], session_ids: [...counts.sessions], calls: counts.calls, attempts: counts.attempts })
          appendLeanEvent(directory, { stage: "changes", document, step: component.id, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
          const asked = performance.now()
          const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("changes", issues) : initialPrompt })
          const answered = performance.now()
          const parsed = parseJsonObject(reply)
          if (parsed.ok && Array.isArray(parsed.value.dependencies) && parsed.value.dependencies.length) { failures.push(`${component.id}: bounded dependency reconciliation required: ${JSON.stringify(parsed.value.dependencies)}`); break }
          const check = checkChanges(directory, reply, readParagraphs(directory, document).filter(row => component.pids.includes(row.pid)), { document, withholdUnverified: index === LEAN_MAX_RETRIES })
          issues = [...check.issues, ...(check.changes ? uniqueReplacementIssues(check.changes) : []), ...(parsed.ok ? retainedInitialIssues(directory, component.pids, parsed.value) : [])]
          if (check.changes) for (const pid of new Set([...component.changes.amendments, ...(component.changes.withheld ?? []).map(row => row.amendment)].map(row => row.pid))) {
            if (![...check.changes.amendments, ...(check.changes.withheld ?? []).map(row => row.amendment), ...(parsed.ok && Array.isArray(parsed.value.retained_initial) ? parsed.value.retained_initial.filter(record) : [])].some(row => row.pid === pid)) issues.push(`[${pid}] candidate was omitted; reconcile it or return a specific dependency limitation.`)
          }
          if (check.changes) for (const pid of new Set((component.changes.explanations ?? []).map(row => row.pid))) {
            if (!check.changes.explanations?.some(row => row.pid === pid)) issues.push(`[${pid}] source-to-initial explanation was omitted; preserve/correct its actual effect or disclose the dependency.`)
          }
          appendLeanEvent(directory, { stage: "changes", document, step: component.id, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
          progress({ feedback: issues })
          if (!issues.length && check.changes) { results.push({ ...check.changes, retained_initial: parsed.ok ? parsed.value.retained_initial as ChangeList["retained_initial"] : [], withheld: [...(component.changes.withheld ?? []).filter(row => !check.changes!.amendments.some(accepted => accepted.pid === row.amendment.pid) && !(parsed.ok && Array.isArray(parsed.value.retained_initial) && parsed.value.retained_initial.some(disposition => record(disposition) && disposition.pid === row.amendment.pid))), ...(check.changes.withheld ?? [])] }); break }
          if (index === LEAN_MAX_RETRIES) failures.push(`${component.id}: ${issues.join(" ")}`)
        }
      } catch (error) { failures.push(`${component.id}: ${error instanceof Error ? error.message : String(error)}`) }
    }
  }))
  drainLeanRequests(directory)
  if (key !== hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))) return { status: "restart" as const }
  if (failures.length) return { status: "failed" as const, reason: `${step}: ${failures.join(" ")}` }
  const combined = combine([work.direct, ...results])
  const order = new Map(readParagraphs(directory, document).map((row, index) => [row.pid, index]))
  combined.amendments.sort((a, b) => order.get(a.pid)! - order.get(b.pid)!)
  const knownQuestions = readQuestions(directory)
  const candidate = { ...combined, questions: combined.questions.map(({ id, ...row }) => knownQuestions.some(saved => saved.id === id && saved.question === row.question && saved.pid === row.pid) ? { id, ...row } : row) }
  const check = checkChanges(directory, JSON.stringify(candidate), readParagraphs(directory, document), { document })
  const issues = [...check.issues, ...(check.changes ? uniqueReplacementIssues(check.changes) : []), ...retainedInitialIssues(directory, [...order.keys()], combined as unknown as Record<string, unknown>)]
  if (issues.length || !check.changes) return { status: "failed" as const, reason: `${step}: combined validation: ${issues.join(" ")}` }
  check.changes.retained_initial = combined.retained_initial
  check.changes.withheld = combined.withheld?.map((row, index) => ({ ...row, amendment: { ...row.amendment, id: `w${index + 1}` } }))
  writeLeanText(directory, `${document}/changes.json`, `${JSON.stringify(check.changes, null, 2)}\n`)
  const openItems = [...check.openItems, ...(combined.withheld ?? []).map(row => openItem("changes", "unapplied_amendment", `Withheld prepared change: ${row.issues.join(" ")}`, { document, pid: row.amendment.pid }))]
  return { status: check.changes.questions.length ? "blocked" as const : "done" as const, outputs: stepOutputs(step), openItems, questions: check.changes.questions, ...(check.changes.questions.length ? { reason: "Waiting for your answers before applying reconciled amendments." } : {}) }
}

function emptyChanges(): ChangeList { return { amendments: [], explanations: [], withheld: [], questions: [], shared_paragraphs: [] } }
function combine(lists: ChangeList[]): ChangeList {
  return { retained_initial: lists.flatMap(row => row.retained_initial ?? []), amendments: lists.flatMap(row => row.amendments), explanations: lists.flatMap(row => row.explanations ?? []), withheld: lists.flatMap(row => row.withheld ?? []), questions: [...new Map(lists.flatMap(row => row.questions).map(row => [JSON.stringify([row.pid, row.question, row.context]), row])).values()], shared_paragraphs: [] }
}
function operationKey(row: Amendment) { return JSON.stringify([row.operation ?? "replace", row.updated]) }
function coalesce(rows: Amendment[]): Amendment {
  return { ...rows[0]!, reason: [...new Set(rows.map(row => row.reason))].join("\n\n"), sources: [...new Map(rows.flatMap(row => row.sources).map(row => [JSON.stringify(row), row])).values()] }
}

function retainedInitialIssues(directory: string, pids: string[], value: Record<string, unknown>) {
  if (value.retained_initial === undefined) return []
  if (!Array.isArray(value.retained_initial)) return ["retained_initial must be a list of explicit source-grounded dispositions."]
  return value.retained_initial.flatMap(row => {
    if (!record(row) || typeof row.pid !== "string" || !pids.includes(row.pid) || typeof row.reason !== "string" || !row.reason.trim() || (!Array.isArray(row.sources) || !row.sources.length)) return ["Each retained_initial disposition needs one listed pid, a source-grounded reason, and a nonempty verified sources array."]
    const issues = row.sources.flatMap(source => !record(source) || typeof source.source_id !== "string" || typeof source.quote !== "string" || !source.quote.trim() || !quoteFound(directory, source as {source_id:string;quote:string}) ? [`[${row.pid}] retained_initial supporting quote is not verified in its canonical source.`] : [])
    if (Array.isArray(value.amendments) && value.amendments.some(amendment => record(amendment) && amendment.pid === row.pid)) issues.push(`[${row.pid}] cannot both retain initial wording and apply a replacement operation.`)
    return issues
  })
}
