import { escapeTagOpeners } from "../../lib/untrusted"
import { parseJsonObject } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { checkNotesTargets, notesCurrentEvidence, type NotesBody, type NotesContext, type NotesTarget } from "./focused-notes-evidence"
import type { DocID } from "./plan"
import { tagged } from "./prompts"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, stepSessionTitle, type LeanDeps } from "./stages"
import type { LeanStageState, LeanState } from "./state"

export type NotesEvidenceRequest = { document: DocID; claim: string; question: string }
export type NotesDependencyObservation = { request: NotesEvidenceRequest; targets: NotesTarget[]; limitation: string }
export type NotesDependencyResult =
  | { status: "resolved"; evidence: NotesContext[]; limitation: string; observation: NotesDependencyObservation }
  | { status: "failed"; reason: string }

/** Transient claim-specific observations share canonical passages. They never
 * turn selected PIDs, an empty limitation or EOF into a global negative. */
export function notesDependencyEvidence(observations: NotesDependencyObservation[], bodies: NotesBody[]): NotesContext[] {
  if (!observations.length) return []
  const targets = new Map<DocID, Set<string>>()
  for (const observation of observations) for (const target of observation.targets) {
    const pids = targets.get(target.document) ?? new Set<string>()
    for (const pid of target.pids) pids.add(pid)
    targets.set(target.document, pids)
  }
  const context = [...targets].flatMap(([document, pids]) => {
    const body = bodies.find(body => body.document === document)
    return body ? notesCurrentEvidence(body, [...pids]) : []
  })
  const unique = [...new Map(context.map(row => [`${row.name}\0${row.text}`, row])).values()]
  return [{ name: "dependency_evidence", text: JSON.stringify({
    role: "Evidence for checking the original prose owner, not an additional explanation owner. Initial evidence is historical; current accepted evidence controls present-body claims. Every original claim, question and actual inspection limitation remains separate. The unchanged or scoped rewritten prose must still pass its ordinary checker.",
    scope: "Selected canonical passages and their material nearby qualifications only. A selection, empty limitation or EOF alone never establishes document-wide absence, an exact global count, resolution or approval. An observation's limitation is not a finding that its claim is false.",
    observations, context: unique,
  }) }]
}

/** Missing packet evidence is not a false claim. Bind each lookup to one
 * actual prose claim and a planned document, without deciding its truth. */
export function checkNotesEvidenceRequest(value: unknown, bodies: Pick<NotesBody, "document">[], prose: string): { request?: NotesEvidenceRequest; issues: string[] } {
  if (!value || typeof value !== "object" || Array.isArray(value)) return { issues: ["request: return one object with document, claim and question."] }
  const row = value as Record<string, unknown>
  if (!bodies.some(body => body.document === row.document)) return { issues: [`request.document=${JSON.stringify(row.document)}: select one planned document ID. Planned document IDs: ${bodies.map(body => body.document).join(", ") || "none available"}.`] }
  if (typeof row.claim !== "string" || !row.claim.trim() || !prose.includes(row.claim))
    return { issues: ["request.claim: copy a nonempty exact span from proposed_prose, including capitalization and punctuation. Do not paraphrase, normalize or change the prose; correct only this request field."] }
  if (typeof row.question !== "string" || !row.question.trim()) return { issues: ["request.question: ask one concrete question about the missing current-body evidence."] }
  return { request: { document: row.document as DocID, claim: row.claim, question: row.question.trim() }, issues: [] }
}

/** One native read-only selection job. Its answer locates evidence; only code
 * copies authoritative body text, and only the ordinary checker judges prose.
 * "resolved" means the lookup answered, never that the claim passed. */
export async function resolveFocusedNotesDependency(directory: string, state: LeanState, deps: LeanDeps,
  input: { id: string; text: string }, prose: string, request: NotesEvidenceRequest, bodies: NotesBody[],
  progress: (patch: Partial<LeanStageState>) => void, observations: NotesDependencyObservation[] = []): Promise<NotesDependencyResult> {
  const validated = checkNotesEvidenceRequest(request, bodies, prose)
  if (!validated.request) return { status: "failed", reason: validated.issues.join(" ") }
  const body = bodies.find(body => body.document === request.document)!
  const agent = LEAN_STAGE_AGENTS.finalize
  const step = `${input.id}:dependency-evidence`
  const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent,
    title: stepSessionTitle(directory, "finalize", step), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
  progress({ session_ids: [sessionID] })
  const navigation = {
    document: body.document, record: `.dochaus/lean/${body.document}/paragraphs-applied.json`,
    paragraphs: body.current.map(({ pid, label, draft_part, draft_heading, text }) => ({ pid, label, part: draft_part, ...(draft_heading ? { heading: text } : {}) })),
  }
  const index = JSON.stringify(navigation)
  const prior = observations.filter(observation => observation.request.document === request.document)
  const prompt = [
    tagged("current_target_lookup", "Locate current accepted evidence for this one disputed prose claim, including material conditions, definitions and qualifications. Select evidence whether it supports or contradicts the claim. This is evidence selection only: do not judge legal correctness, draft prose, resolve a fact or grant approval. Missing inspection never establishes missing content."),
    tagged("evidence_request", escapeTagOpeners(JSON.stringify(request))),
    ...(prior.length ? [tagged("prior_dependency_inspections", escapeTagOpeners(JSON.stringify(prior)))] : []),
    tagged("current_navigation", escapeTagOpeners(index.length <= 6_000 ? index : JSON.stringify({ document: navigation.document, record: navigation.record }))),
    tagged("retrieval_scope", "Use read-record or read-document view=current_body for the requested planned document. Prior claim-specific selections and limitations are navigation, not proof of completeness or a verdict about the current prose. Use them to inspect the still-missing dependency rather than blindly restarting already read ranges; reread when a material qualification requires it. Inspect only passages material to the requested claim and their qualifications. Return existing current paragraph PIDs, not quotations or inferred text. Do not write files."),
    tagged("output_contract", 'Reply only {"targets":[{"document":"requested planned id","pids":["current paragraph pid"]}],"limitation":"remaining inspection limit, or empty"}. If no current passage can be located or inspected, return no targets with the actual limitation. A selected passage does not establish document-wide absence; identify any unavailable material dependency.'),
  ].join("\n\n")
  let issues: string[] = []
  for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
    appendLeanEvent(directory, { stage: "finalize", step, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
    const asked = performance.now()
    const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: attempt
      ? `Correct only this dependency-selection reply's format or invalid targets. Keep the same claim and evidence question. Return targets/limitation JSON.\n${issues.join("\n")}` : prompt })
    const answered = performance.now()
    const parsed = parseJsonObject(reply)
    const checked = parsed.ok ? checkNotesTargets(parsed.value, [body]) : { issues: [parsed.error] }
    issues = checked.issues
    progress({ calls: attempt + 1, attempts: attempt + 1, feedback: issues })
    appendLeanEvent(directory, { stage: "finalize", step, event: "checker", attempt, issues,
      wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
    if (!checked.targets) continue
    const context = checked.targets.flatMap(target => notesCurrentEvidence(body, target.pids))
    const limitation = checked.limitation || "Only selected current passages and their nearby qualifications were inspected; this does not establish document-wide absence."
    const observation = { request: validated.request, targets: checked.targets, limitation: checked.limitation ?? "" }
    return { status: "resolved", limitation, observation, evidence: [{ name: "dependency_evidence", text: JSON.stringify({
      document: body.document, claim: request.claim, question: request.question,
      role: "Evidence for checking this claim only, not an additional explanation owner. Generated initial evidence is historical; current accepted evidence controls. The unchanged prose must still pass its ordinary checker.",
      context, limitation,
    }) }] }
  }
  return { status: "failed", reason: `The dependency lookup for ${input.id} supplied no valid selection after ${LEAN_MAX_RETRIES} retries: ${issues.join(" ")}` }
}
