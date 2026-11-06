import { appendLeanEvent, leanElapsedMs } from "./events"
import { checkDifferences, checkGaps, acceptEvidenceReply, evidencePrompt, gapsPrompt, type Discrepancy } from "./evidence"
import { parseJsonObject } from "./checks"
import { INSTRUCTIONS_FILE } from "./instructions"
import { retryPrompt } from "./prompts"
import type { StageRun } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, stageInputs, stepSessionTitle, type LeanDeps } from "./stages"
import { hashStageInputs, readLeanJson, writeLeanText, type LeanStageState, type LeanState } from "./state"

// The discrepancy stage: two focused jobs, each in a session of its own that
// retries its own check. DIFFERENCES compares the sources' identities, roles,
// dates, figures, versions and instructions, as the stage did before it listed
// gaps. GAPS, shown the lawyer's words, the instruction list and DIFFERENCES'
// entries, lists what the work must settle that the sources leave silent or
// only one source speaks to. Code writes discrepancy.json once, {entries, gaps,
// limitations}, checked as a whole by checkEvidence. DIFFERENCES' accepted
// reply is saved under its own key, which leaves out the instruction list, so a
// resume after it, or a new instruction list, runs only GAPS.

const DIFFERENCES_FILE = "discrepancy-differences.json"

type Differences = { key: string; entries: Discrepancy[]; limitations: string[] }
type Job = "differences" | "gaps"

/** The differences job's resume key: the stage's inputs without the instruction list it is not shown. */
function differencesKey(directory: string) {
  return hashStageInputs("discrepancy", LEAN_PROMPT_VERSION, stageInputs(directory, "discrepancy").filter(part => part.name !== INSTRUCTIONS_FILE))
}

export async function runDiscrepancyStage(directory: string, state: LeanState, deps: LeanDeps, progress: (patch: Partial<LeanStageState>) => void): Promise<StageRun> {
  const agent = LEAN_STAGE_AGENTS.discrepancy
  const sessionIDs: string[] = []
  const calls = { count: 0, attempts: 0 }
  const job = async (label: Job, first: string, check: (value: unknown) => string[]): Promise<{ value?: Record<string, unknown>; issues: string[] }> => {
    const sessionID = await deps.sessions.createSession({
      directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, "discrepancy", `discrepancy: ${label}`), model: state.model, permissions: LEAN_CHILD_PERMISSIONS,
    })
    sessionIDs.push(sessionID)
    progress({ ...(sessionIDs.length === 1 ? { session_id: sessionID } : {}), session_ids: [...sessionIDs] })
    let text = first
    for (let attempt = 0; ; attempt += 1) {
      appendLeanEvent(directory, { stage: "discrepancy", event: attempt ? "retry" : "model_call", step: label, attempt, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text })
      const answered = performance.now()
      const parsed = parseJsonObject(reply)
      const issues = parsed.ok ? check(parsed.value) : [`Reply with the requested JSON object only. ${parsed.error}`]
      calls.count += 1
      calls.attempts = Math.max(calls.attempts, attempt + 1)
      progress({ attempts: calls.attempts, calls: calls.count, feedback: issues })
      appendLeanEvent(directory, { stage: "discrepancy", event: "checker", step: label, attempt, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (parsed.ok && !issues.length) return { value: parsed.value, issues }
      if (attempt >= LEAN_MAX_RETRIES) return { issues }
      text = retryPrompt("discrepancy", issues)
    }
  }
  const failed = (label: Job, issues: string[]): StageRun => ({ status: "failed", reason: `discrepancy: ${label} checker failed after ${LEAN_MAX_RETRIES} retries: ${issues.join(" ")}` })

  const key = differencesKey(directory)
  const saved = readLeanJson(directory, DIFFERENCES_FILE) as Differences | undefined
  let differences = saved?.key === key && !checkDifferences(directory, saved).length ? saved : undefined
  if (!differences) {
    const asked = await job("differences", evidencePrompt(directory, "discrepancy"), value => checkDifferences(directory, value))
    if (!asked.value) return failed("differences", asked.issues)
    // Saved only under the inputs it was asked on; a change meanwhile restarts the walk.
    if (differencesKey(directory) !== key) return { status: "restart" }
    differences = { key, entries: asked.value.entries as Discrepancy[], limitations: asked.value.limitations as string[] }
    writeLeanText(directory, DIFFERENCES_FILE, `${JSON.stringify(differences, null, 2)}\n`)
  }
  const { entries } = differences
  const gaps = await job("gaps", gapsPrompt(directory, entries), value => checkGaps(directory, value, entries))
  if (!gaps.value) return failed("gaps", gaps.issues)
  const assembled = { entries, gaps: gaps.value.gaps, limitations: [...new Set([...differences.limitations, ...gaps.value.limitations as string[]])] }
  const accepted = acceptEvidenceReply(directory, "discrepancy", JSON.stringify(assembled))
  if (accepted.issues.length) return { status: "failed", reason: `discrepancy: ${accepted.issues.join(" ")}` }
  return { status: "done", outputs: accepted.outputs, openItems: accepted.openItems }
}
