import { parseJsonObject } from "./change-checks"
import { checklistProducerInputs } from "./checklist-source-context"
import { acceptChangeReview, checkChangeReview, projectedChangeParagraphs, readChangePlan, type ChangeReview } from "./change-pipeline"
import { readChangeList, readParagraphs, shownText } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { readLeanPlan, type DocID } from "./plan"
import { packageCompositionBlock } from "./prompts"
import { drainLeanRequests } from "./requests"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, stageInputs, stepSessionTitle, type LeanDeps } from "./stages"
import { hashStageInputs, readQuestions, type LeanState, type LeanStageState } from "./state"

const MAX_CONTEXT = 32_000
export type FocusedChangeReviewInput = { id: string; pids: string[]; text: string }

/** The current cumulative operations, including retained operations, are the
 * coverage inventory. Group membership supplies related targets, never a reason
 * to omit an operation or to accept its source/legal effect without review. */
export function focusedChangeReviewInputs(directory: string, document: DocID): FocusedChangeReviewInput[] {
  const plan = readChangePlan(directory, document)
  const changes = readChangeList(directory, document)
  if (!plan || !changes) throw Error("Focused semantic review requires a valid current plan and cumulative changes.")
  const initial = readParagraphs(directory, document)
  const projected = projectedChangeParagraphs(directory, document)
  const initialPids = new Set(initial.map(row => row.pid))
  const operations = [...changes.amendments, ...(changes.explanations ?? []), ...(changes.retained_initial ?? []), ...(changes.withheld ?? []).map(row => row.amendment)]
  if (operations.some(row => !initialPids.has(row.pid))) throw Error("Semantic review inventory contains a noncurrent target.")
  const covered = new Set<string>()
  const units: { id: string; pids: string[]; goal: unknown }[] = plan.groups.map(group => {
    group.pids.forEach(pid => covered.add(pid))
    return { id: `group:${group.id}`, pids: group.pids, goal: group }
  })
  for (const row of operations) {
    if (covered.has(row.pid)) continue
    covered.add(row.pid)
    units.push({ id: `retained:${row.pid}`, pids: [row.pid], goal: "Review this cumulative operation even though no current preparation group owns its target; do not silently lose a prior supported correction or explanation." })
  }
  // Rejected historical ideas do not become mandatory new review work. Current
  // unresolved/recommended dispositions still need a focused disclosure check;
  // their originating document does not prevent required advice in this one.
  for (const [index, finding] of (plan.findings ?? []).entries()) {
    if (finding.disposition === "rejected") continue
    units.push({ id: `disclosure:${index}`, pids: [], goal: finding })
  }
  const answers = readQuestions(directory).map(row => ({ id: row.id, question: row.question, status: row.status, answer: row.answer }))
  const request = checklistProducerInputs(directory).find(row => row.key === "request")!.input
  const documentPlan = readLeanPlan(directory)
  const entry = documentPlan?.documents.find(row => row.id === document)
  const composition = packageCompositionBlock(directory, document)
  const build = (unit: typeof units[number], split = false): FocusedChangeReviewInput => {
    const own = new Set(unit.pids)
    const context = new Set(unit.pids)
    initial.forEach((row, index) => {
      if (!own.has(row.pid)) return
      for (const neighbor of [index - 1, index + 1]) if (initial[neighbor]) context.add(initial[neighbor]!.pid)
    })
    const text = [
      "Independently perform one focused semantic review of the current cumulative revision BEFORE apply. Prepared/adjudicated records are proposals, not authority. Review the assigned goal, actual resulting wording and source explanations, including withheld work and qualifications. Do not assume a matched quote proves its proposed effect.",
      `<assignment>${JSON.stringify({ id: unit.id, document, ...(split ? { assigned_pids: unit.pids } : {}), goal: unit.goal })}</assignment>`,
      `<document_purpose_and_role>${JSON.stringify({ title: entry?.title, purpose: entry?.purpose, mode: entry?.mode, acting_for: documentPlan?.acting_for })}</document_purpose_and_role>`,
      `<actual_instructions>\n${JSON.stringify([...request.passages, ...request.instructions])}\n</actual_instructions>`,
      ...(composition ? [composition] : []),
      `<actual_questions_and_answers>\n${JSON.stringify(answers)}\n</actual_questions_and_answers>`,
      `<shared_decisions>\n${JSON.stringify(plan.shared_decisions)}\n</shared_decisions>`,
      `<initial_target_and_neighbors>\n${initial.filter(row => context.has(row.pid)).map(row => `[${row.pid}] ${shownText(row)}`).join("\n\n")}\n</initial_target_and_neighbors>`,
      `<projected_target_and_neighbors>\n${projected.filter(row => context.has(row.pid)).map(row => `[${row.pid}] ${shownText(row)}`).join("\n\n")}\n</projected_target_and_neighbors>`,
      `<cumulative_operations>\n${JSON.stringify({ amendments: changes.amendments.filter(row => own.has(row.pid)), explanations: (changes.explanations ?? []).filter(row => own.has(row.pid)), retained_initial: (changes.retained_initial ?? []).filter(row => own.has(row.pid)), withheld: (changes.withheld ?? []).filter(row => own.has(row.amendment.pid)), questions: changes.questions })}\n</cumulative_operations>`,
      "Your explicit dependency task: inspect the conditions, exceptions, definitions and cross-references on which this assigned wording depends, including related clauses outside these neighbors. For advice about another run-owned document, retrieve its actual relevant final clause and compare each material factual assertion, proposed resolution and remaining decision against that clause; agreement with another agent's summary is insufficient. Inspect relevant original source qualifications and actual instructions/answers. Report a cross-clause or earlier-document defect with its owning initial PID when identifiable. Do not review unrelated history or silently assume dependencies are consistent.",
      `Targeted retrieval: read-record .dochaus/lean/${document}/paragraphs.json and changes.json provide initial targets and the complete cumulative operation inventory. For this document, search-document/read-document view=proposed_body exposes the complete pending revision with initial-PID and split-part coordinates, including replacements and deletions. Use that view to locate disclosure passages even when this assignment has no inline target PID. Its words are generated proposals to evaluate, not source authority. current_body is the actual pre-apply Word artifact and may differ: never infer an omission in the pending revision from that older view. Original uploads are available using search-document/read-document view=source. Earlier run-owned final bodies can be inspected with document=<id>, view=current_body and bounded offset/count; use source-map.json and .dochaus/lean/plan.json to identify them. Approved legal propositions and limitations are in research-review.json/research.json; retrieve relevant approved text before asserting law, never substitute memory. Full plans, reviews and old findings remain targeted lookup, not authority.`,
      `Each changed paragraph, proposed source-to-initial-draft explanation and retained_initial disposition is assigned somewhere. ${split ? "This assignment owns only its assigned_pids. The complete goal retains related targets owned by other review jobs: retrieve their actual proposed passages and inspect every material dependency needed to evaluate this target, including combined effects, exceptions and shared qualifications. Separate target ownership never establishes consistency or licenses omitting a dependent requirement." : "This assignment owns only its listed targets/disclosure, but must inspect their material dependencies."} Verify each retained_initial withdrawal against its canonical supporting words, stated reason and the actual retained initial/projected text; a prior amendment may be withdrawn only when that disposition is supported, not merely absent from a candidate list. Preserve unaffected rights/obligations and every still-supported correction. Check explanation accuracy and final effect, not just body wording. For a disclosure assignment, locate the actual advice passage and test whether the unresolved material issue/action is accurately carried through; a disposition alone does not fulfill a required disclosure. Distinguish optional recommendations from essential requirements. Do not reopen answered decisions without new contrary evidence. A disclosed review-draft placeholder can be appropriate; do not invent approval or treat an uninspected essential dependency as a pass.`,
      "Native redlines and comments for this revision are generated after this review at apply. Assess proposed reasons/explanations now; missing pre-apply exports are not a defect. Do not assert old artifacts prove this revision, rewrite text, or write files.",
      'Reply exactly {"status":"pass"|"changes-required","findings":[{"document":"owning current/earlier run-owned document id, omit for current","pid":"initial paragraph id when identifiable, otherwise omit","detail":"specific material defect, actual wording/evidence and required correction","target":"body|research, omit for body"}],"limitations":["remaining disclosed limits"]}. Pass requires zero material findings; changes-required requires at least one. An earlier-body correction must name that document and its own initial PID. Research targets require a genuinely new material authority proposition needing retrieval, not an ordinary wording defect. Missing material inspection must remain explicit, not be called verified.',
    ].join("\n\n")
    return { id: unit.id, pids: unit.pids, text }
  }
  return units.flatMap(unit => {
    const input = build(unit)
    if (input.text.length <= MAX_CONTEXT) return [input]
    // A preparation group can merge many dependent targets. Review each target
    // with the complete shared goal rather than duplicating every operation in
    // one packet. The goal retains all related PIDs and source qualifications;
    // each reviewer must retrieve its material dependencies before passing.
    const inputs = unit.pids.length > 1
      ? unit.pids.map(pid => build({ ...unit, id: `${unit.id}:target:${pid}`, pids: [pid] }, true))
      : [input]
    for (const row of inputs) if (row.text.length > MAX_CONTEXT)
      throw Error(`Semantic review assignment ${row.id} exceeds bounded context; decompose its related goals before model admission.`)
    return inputs
  })
}

/** One existing focused review job, shared by the complete pool and isolated
 * evidence diagnostics. It validates a finding, never publishes an aggregate. */
export async function runFocusedChangeReviewJob(directory: string, state: LeanState, deps: LeanDeps, document: DocID, input: FocusedChangeReviewInput, progress: (patch: Partial<LeanStageState>) => void) {
  if (input.text.length > MAX_CONTEXT) throw Error(`Semantic review assignment ${input.id} exceeds bounded context; decompose before model admission.`)
  const step = `${document}.change-review` as const
  const agent = LEAN_STAGE_AGENTS["change-review"]
  const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, step, input.id), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
  progress({ session_id: sessionID, session_ids: [sessionID] })
  let issues: string[] = []
  for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
    progress({ calls: attempt + 1, attempts: attempt + 1 })
    appendLeanEvent(directory, { stage: "change-review", document, step: input.id, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
    const asked = performance.now()
    const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: attempt ? `Correct only these review-contract errors, retaining your supported findings and original assignment: ${JSON.stringify(issues)}` : input.text })
    const answered = performance.now()
    const parsed = parseJsonObject(reply)
    issues = parsed.ok ? checkChangeReview(directory, document, parsed.value) : ["Return the required semantic-review JSON object."]
    appendLeanEvent(directory, { stage: "change-review", document, step: input.id, event: "checker", attempt, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
    progress({ feedback: issues })
    if (!issues.length && parsed.ok) return { session_id: sessionID, review: parsed.value as ChangeReview, issues }
  }
  return { session_id: sessionID, issues }
}

/** Strict union: every admitted job must validate and settle. A worker failure
 * or changed input prevents acceptance of any partial passing aggregate. */
export async function runFocusedChangeReview(directory: string, state: LeanState, deps: LeanDeps, document: DocID, progress: (patch: Partial<LeanStageState>) => void) {
  const step = `${document}.change-review` as const
  const inputSha = hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))
  const inputs = focusedChangeReviewInputs(directory, document)
  const counts = { next: 0, calls: 0, attempts: 0 }
  const sessions: string[] = []
  const results = new Map<string, ChangeReview>()
  const failures: string[] = []
  await Promise.all(Array.from({ length: Math.min(4, inputs.length) }, async () => {
    for (;;) {
      const input = inputs[counts.next++]
      if (!input) return
      try {
        let previousCalls = 0
        const result = await runFocusedChangeReviewJob(directory, state, deps, document, input, patch => {
          if (patch.session_id && !sessions.includes(patch.session_id)) sessions.push(patch.session_id)
          if (patch.calls !== undefined) { counts.calls += patch.calls - previousCalls;previousCalls = patch.calls }
          counts.attempts = Math.max(counts.attempts, patch.attempts ?? 0)
          progress({ ...patch, session_id: sessions[0], session_ids: [...sessions], calls: counts.calls, attempts: counts.attempts })
        })
        if (result.review) results.set(input.id, result.review)
        else failures.push(`${input.id}: ${result.issues.join(" ")}`)
      } catch (error) { failures.push(`${input.id}: ${error instanceof Error ? error.message : String(error)}`) }
    }
  }))
  drainLeanRequests(directory)
  if (inputSha !== hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))) return { status: "restart" as const }
  if (failures.length || results.size !== inputs.length) return { status: "failed" as const, reason: `${step}: focused semantic review incomplete: ${failures.join(" ")}` }
  const ordered = inputs.map(input => results.get(input.id)!)
  const findings = ordered.flatMap(row => row.findings)
  const aggregate: ChangeReview = { status: findings.length ? "changes-required" : "pass", findings, limitations: [...new Set(ordered.flatMap(row => row.limitations))] }
  const accepted = acceptChangeReview(directory, document, JSON.stringify(aggregate))
  if (accepted.issues.length) return { status: "failed" as const, reason: `${step}: ${accepted.issues.join(" ")}` }
  return { status: "done" as const, outputs: accepted.outputs, openItems: [], questions: [] }
}
