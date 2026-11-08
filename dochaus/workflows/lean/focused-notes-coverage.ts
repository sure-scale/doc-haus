import { escapeTagOpeners } from "../../lib/untrusted"
import { parseJsonObject } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import type { FocusedNotesInput } from "./focused-notes"
import { checkNotesEvidenceRequest, type NotesEvidenceRequest } from "./focused-notes-dependencies"
import type { NotesBody } from "./focused-notes-evidence"
import { readLeanPlan } from "./plan"
import { tagged } from "./prompts"
import { LEAN_MAX_RETRIES, stepSessionTitle, type LeanDeps, type PermissionRule } from "./stages"
import type { LeanStageState, LeanState } from "./state"

export const NOTES_COVERAGE_AGENT = "lean-notes-check"
export const NOTES_CHECK_PERMISSIONS: PermissionRule[] = [{ action: "*", resource: "*", effect: "deny" }]
export type NotesCoverageVerdict = { status: "pass"; reason: string }
  | { status: "rewrite"; reason: string }
  | { status: "needs-evidence"; reason: string; request: NotesEvidenceRequest }
export type NotesCoverageResult = NotesCoverageVerdict | { status: "failed"; reason: string }
export type NotesCoverageReference = { documents: Pick<NotesBody, "document">[]; prose: string }
const COVERAGE_OUTPUT_CONTRACT = 'Reply only {"status":"pass"|"rewrite","reason":"concrete explanation of what this prose explains, or the owned substantive point it fails to explain"}; or {"status":"needs-evidence","reason":"why this concrete claim needs current dependency inspection","request":{"document":"one planned document id","claim":"exact words from proposed_prose","question":"one current-body dependency question"}}. Return one request object only. A rewrite reason identifies an actual missing or unsupported owned explanation so the same writer can return its complete section; missing local inspection alone is not such a finding.'

/** A valid negative verdict is an answer to the coverage question. A missing
 * or malformed verdict supplies no answer and must never admit the prose. */
export function checkNotesCoverageReply(reply: string, reference?: NotesCoverageReference): { value?: NotesCoverageVerdict; issues: string[] } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Return the notes-coverage JSON object only. ${parsed.error}`] }
  const { status, reason, request } = parsed.value
  const issues = [
    ...(status !== "pass" && status !== "rewrite" && status !== "needs-evidence" ? ['status must be "pass", "rewrite" or "needs-evidence".'] : []),
    ...(typeof reason !== "string" || !reason.trim() ? ["reason must explain how the proposed prose does or does not substantively address this assigned point."] : []),
  ]
  if (status === "needs-evidence") {
    const row = request && typeof request === "object" && !Array.isArray(request) ? request as Record<string, unknown> : undefined
    if (!row || typeof row.document !== "string" || !/^d[1-9]\d*$/.test(row.document) ||
      typeof row.claim !== "string" || !row.claim.trim() || typeof row.question !== "string" || !row.question.trim())
      issues.push("needs-evidence must give one request object with a document id, the exact prose claim, and one concrete current-body dependency question.")
    if (issues.length) return { issues }
    if (reference) {
      const checked = checkNotesEvidenceRequest(row, reference.documents, reference.prose)
      if (!checked.request) return { issues: checked.issues }
    }
    return { value: { status, reason: (reason as string).trim(), request: {
      document: row!.document as NotesEvidenceRequest["document"], claim: row!.claim as string, question: (row!.question as string).trim(),
    } }, issues }
  }
  if (request !== undefined) issues.push("Only needs-evidence may carry an inspection request; a pass or rewrite must answer the assigned coverage question.")
  if (issues.length) return { issues }
  return { value: status === "pass" ? { status: "pass", reason: (reason as string).trim() } : { status: "rewrite", reason: (reason as string).trim() }, issues }
}

/** One independent question about one proposed section, with no retrieval or
 * drafting tools. The packet is that writer's scoped assignment, never the
 * complete notes or a bundle of neighboring jobs. */
export function notesCoveragePrompt(input: Pick<FocusedNotesInput, "kind" | "text">, prose: string, documents: Pick<NotesBody, "document">[] = []) {
  return [
    tagged("question", "Does the proposed prose itself substantively explain this one owned assignment to the lawyer, preserving its material dependency qualifications and the action or decision this assignment actually owns?"),
    tagged("section_kind", input.kind),
    tagged("assigned_writer_packet", escapeTagOpeners(input.text)),
    tagged("proposed_prose", escapeTagOpeners(prose)),
    ...(documents.length ? [tagged("planned_dependency_document_ids", JSON.stringify(documents.map(row => row.document)))] : []),
    tagged("assessment_rules", [
      "Assess only the proposed_prose block against this assignment. A previous response, invisible earlier text, errored output, a promised explanation or a claim of completion cannot supply an explanation absent from this prose. Nonresponsive, irrelevant or coverage-only prose needs rewriting, even when it is grammatical and nonempty.",
      "An open-item owner must explain the actual uncertainty, its practical consequence and the required human action, rather than an invented answer or an assertion that the work is complete. A delegated-decision owner needs the assigned choice or reason it remains open, with the material explanation that assignment requires. Other kinds need the practical explanation their own assignment asks for. Judge substantive meaning, not word overlap, length, a prescribed phrase or style preference.",
      "A change explanation owns the change's actual effect and material qualifications, including a dependency that remains unresolved. Do not duplicate a separately owned human action merely because that dependency appears in the change packet. Require that action here only when this change actually owns it, its effect creates or changes the action, or the prose wrongly claims the dependent condition has been satisfied. Preserving an unresolved caveat is not a claim that the open point is settled. Dependency context does not create another explanation owner or require the full explanation of another section.",
      "Current accepted evidence controls present-tense claims about the body. Generated initial wording and proposed or withheld changes are separate baselines; a correctly labelled historical explanation is not a current contradiction. Check the actual current targets and their material qualifications before accepting a claim that a repair, condition, citation or other content is missing. A stale deferred record alone cannot establish that it remains missing. Citation presence does not establish legal correctness or human approval. A local passage, a heading at its boundary, or unavailable current inspection cannot establish document-wide absence or unconditional effect; require a scoped explanation or an honest inspection limitation instead. Preserve genuine residual open points.",
      "Match the evidence kind to the claim. Current-body claims need current-body evidence; historical research or human-approval status needs its own research, decision or approval record, or precise attribution to a supplied unresolved record. Current clauses do not establish what research was completed or whether human approval occurred. An accurately attributed unresolved record supports its recorded uncertainty, not inferred legal validity, invalidity or approval. Do not demand exhaustive current-body inspection merely to establish that historical record status. If prose exceeds its record support, require precise support or a scoped rewrite. Preserve the original assigned duty, practical consequence and required action; a limitation cannot hide a material contradiction or drop an essential exhaustive requirement.",
      "When one concrete current-body claim depends on a material passage outside the supplied packet, request that bounded inspection with needs-evidence; never treat a missing dependency packet as evidence that the prose is false. Select one planned document and copy the exact affected claim from proposed_prose. Ask only the current wording or qualification needed to assess that claim, not a whole document, new research, legal merits or a human answer. The coordinator validates the reference, supplies trusted current-body dependency evidence and rechecks this SAME prose before deciding whether any rewrite is needed. Dependency evidence supplies context, not another owned obligation. If an actual bounded inspection reports unavailable evidence, preserve that limitation; it does not establish absence, approval or completion. Do not repeat an inspection request already answered by the dependency packet.",
      "An honest explanation of an unresolved fact or unavailable authority can pass. Do not require that the writer supply the missing answer, research new law, add facts, repeat settled confirmations or obtain human release approval. Preserve an adopted delegated judgment as settled where the assignment does; confirmation is required only when its actual evidence leaves it open.",
      "Do not require reference markers, identifiers, coverage lists or another section; code attaches references after acceptance. Existing legal quotation and citation checks remain separate. You have no tools and must neither retrieve, rewrite the section nor decide the underlying legal or commercial choice.",
    ].join("\n")),
    tagged("output_contract", COVERAGE_OUTPUT_CONTRACT),
  ].join("\n\n")
}

/** Format retries stay in this verification session. A valid rewrite returns
 * to the existing writer's bounded loop; exhausted verifier format failures
 * stop the section without asking the writer to repair verifier bookkeeping. */
export async function checkFocusedNotesCoverage(directory: string, state: LeanState, deps: LeanDeps,
  input: Pick<FocusedNotesInput, "id" | "kind" | "text">, prose: string,
  progress: (patch: Partial<LeanStageState>) => void): Promise<NotesCoverageResult> {
  const step = `${input.id}:explanation-check`
  const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent: NOTES_COVERAGE_AGENT,
    title: stepSessionTitle(directory, "finalize", step), model: state.model, permissions: NOTES_CHECK_PERMISSIONS })
  progress({ session_ids: [sessionID] })
  const reference: NotesCoverageReference = { documents: (readLeanPlan(directory)?.documents ?? []).map(row => ({ document: row.id })), prose }
  let issues: string[] = []
  for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
    appendLeanEvent(directory, { stage: "finalize", step, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
    const asked = performance.now()
    const reply = await deps.sessions.prompt({ directory, sessionID, agent: NOTES_COVERAGE_AGENT,
      text: attempt ? `Correct only this explanation-check reply's format or invalid request fields. Keep the same assigned point and proposed prose.\n${issues.map(issue => `- ${issue}`).join("\n")}\n\n${COVERAGE_OUTPUT_CONTRACT}` : notesCoveragePrompt(input, prose, reference.documents) })
    const answered = performance.now()
    const checked = checkNotesCoverageReply(reply, reference)
    issues = checked.issues
    progress({ calls: attempt + 1, attempts: attempt + 1, feedback: issues })
    // The explanation goes back to the writer; timing telemetry keeps only the bounded verdict.
    appendLeanEvent(directory, { stage: "finalize", step, event: "checker", attempt, issues,
      ...(checked.value ? { outcome: checked.value.status } : {}),
      wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
    if (checked.value) return checked.value
  }
  return { status: "failed", reason: `The explanation check for ${input.id} supplied no valid verdict after ${LEAN_MAX_RETRIES} retries: ${issues.join(" ")}` }
}
