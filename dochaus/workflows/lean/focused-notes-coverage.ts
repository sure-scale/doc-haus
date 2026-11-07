import { escapeTagOpeners } from "../../lib/untrusted"
import { parseJsonObject } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import type { FocusedNotesInput } from "./focused-notes"
import { tagged } from "./prompts"
import { LEAN_MAX_RETRIES, stepSessionTitle, type LeanDeps, type PermissionRule } from "./stages"
import type { LeanStageState, LeanState } from "./state"

export const NOTES_COVERAGE_AGENT = "lean-notes-check"
export const NOTES_CHECK_PERMISSIONS: PermissionRule[] = [{ action: "*", resource: "*", effect: "deny" }]
export type NotesCoverageVerdict = { status: "pass" | "rewrite"; reason: string }
export type NotesCoverageResult = NotesCoverageVerdict | { status: "failed"; reason: string }

/** A valid negative verdict is an answer to the coverage question. A missing
 * or malformed verdict supplies no answer and must never admit the prose. */
export function checkNotesCoverageReply(reply: string): { value?: NotesCoverageVerdict; issues: string[] } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Return the notes-coverage JSON object only. ${parsed.error}`] }
  const { status, reason } = parsed.value
  const issues = [
    ...(status !== "pass" && status !== "rewrite" ? ['status must be "pass" or "rewrite".'] : []),
    ...(typeof reason !== "string" || !reason.trim() ? ["reason must explain how the proposed prose does or does not substantively address this assigned point."] : []),
  ]
  return issues.length ? { issues } : { value: { status: status as NotesCoverageVerdict["status"], reason: (reason as string).trim() }, issues }
}

/** One independent question about one proposed section, with no retrieval or
 * drafting tools. The packet is that writer's scoped assignment, never the
 * complete notes or a bundle of neighboring jobs. */
export function notesCoveragePrompt(input: Pick<FocusedNotesInput, "kind" | "text">, prose: string) {
  return [
    tagged("question", "Does the proposed prose itself substantively explain this one assigned point to the lawyer, including its required remaining uncertainty, action or decision?"),
    tagged("section_kind", input.kind),
    tagged("assigned_writer_packet", escapeTagOpeners(input.text)),
    tagged("proposed_prose", escapeTagOpeners(prose)),
    tagged("assessment_rules", [
      "Assess only the proposed_prose block against this assignment. A previous response, invisible earlier text, errored output, a promised explanation or a claim of completion cannot supply an explanation absent from this prose. Nonresponsive, irrelevant or coverage-only prose needs rewriting, even when it is grammatical and nonempty.",
      "An open point needs its actual uncertainty or limitation and the required human action, rather than an invented answer. A delegated decision needs the assigned choice or reason it remains open, with the material explanation the assignment requires. Other kinds need the practical explanation their own assignment asks for. Judge substantive meaning, not word overlap, length, a prescribed phrase or style preference.",
      "An honest explanation of an unresolved fact or unavailable authority can pass. Do not require that the writer supply the missing answer, research new law, add facts, repeat settled confirmations or obtain human release approval. Preserve an adopted delegated judgment as settled where the assignment does; confirmation is required only when its actual evidence leaves it open.",
      "Do not require reference markers, identifiers, coverage lists or another section; code attaches references after acceptance. Existing legal quotation and citation checks remain separate. You have no tools and must neither retrieve, rewrite the section nor decide the underlying legal or commercial choice.",
    ].join("\n")),
    tagged("output_contract", 'Reply only {"status":"pass"|"rewrite","reason":"concrete explanation of what this prose explains, or the assigned substantive point it fails to explain"}. A rewrite reason identifies the missing explanation so the same writer can return its complete section.'),
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
  let issues: string[] = []
  for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
    appendLeanEvent(directory, { stage: "finalize", step, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
    const asked = performance.now()
    const reply = await deps.sessions.prompt({ directory, sessionID, agent: NOTES_COVERAGE_AGENT,
      text: attempt ? `Correct only this explanation-check reply's format. Keep the same assigned point and proposed prose, and return the required status/reason JSON.\n${issues.map(issue => `- ${issue}`).join("\n")}` : notesCoveragePrompt(input, prose) })
    const answered = performance.now()
    const checked = checkNotesCoverageReply(reply)
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
