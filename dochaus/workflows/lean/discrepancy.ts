import { parallel } from "./checklist-flow"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { checkDifferences, checkGaps, checkGapVerify, acceptEvidenceReply, evidencePrompt, gapsPrompt, gapVerifyPrompt, quoteKey, restatesDifference, type Discrepancy, type Gap, type GapVerify } from "./evidence"
import { parseJsonObject } from "./checks"
import { INSTRUCTIONS_FILE } from "./instructions"
import { retryPrompt } from "./prompts"
import type { SourceQuote } from "./quotes"
import type { StageRun } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, stageInputs, stepSessionTitle, type LeanDeps } from "./stages"
import { hashStageInputs, readLeanJson, writeLeanText, type LeanStageState, type LeanState } from "./state"

// The discrepancy stage: focused jobs, each in a session of its own that
// retries its own check. DIFFERENCES compares the sources' identities, roles,
// dates, figures, versions and instructions, as the stage did before it listed
// gaps. GAPS, shown the lawyer's words and the instruction list, lists what the
// work must settle that the sources leave silent or only one source speaks to.
// The two run in parallel. Each gap then gets a VERIFY job of its own, run in
// parallel within the workflow limit: does a matter source supply it? Code
// writes discrepancy.json once (assembleDiscrepancy), checked as a whole by
// checkEvidence. Each job's accepted reply is saved under its own content key,
// so a resume, or a new instruction list (which only GAPS reads), runs only the
// jobs whose answer is missing.

const DIFFERENCES_FILE = "discrepancy-differences.json"
const GAPS_FILE = "discrepancy-gaps.json"
const VERIFY_FILE = "discrepancy-verify.json"
/** Added to a gap whose verify job never passed its check: the gap stands. */
export const UNVERIFIED_GAP = "Whether a matter source supplies this point could not be checked, so it stays a gap."

type Differences = { key: string; entries: Discrepancy[]; limitations: string[] }
type Gaps = { key: string; gaps: Gap[]; limitations: string[] }
type Verified = GapVerify & { key: string }
/** A job's accepted answer, or how the stage stops. */
type Outcome<T> = { value: T } | { stop: Extract<StageRun, { status: "failed" | "restart" }> }

/** The stage's inputs without the instruction list: all DIFFERENCES and VERIFY read. */
function sourceInputs(directory: string) {
  return stageInputs(directory, "discrepancy").filter(part => part.name !== INSTRUCTIONS_FILE)
}
const differencesKey = (directory: string) => hashStageInputs("discrepancy", LEAN_PROMPT_VERSION, sourceInputs(directory))
const gapsKey = (directory: string) => hashStageInputs("discrepancy", LEAN_PROMPT_VERSION, stageInputs(directory, "discrepancy"))
/** A verify job's key: the sources it searches and the gap it is shown. */
const verifyKey = (inputs: ReturnType<typeof sourceInputs>, gap: Gap) =>
  hashStageInputs("discrepancy", LEAN_PROMPT_VERSION, [...inputs, { name: "gap", bytes: JSON.stringify([gap.detail, gap.requirement.source_id, gap.requirement.quote]) }])

export async function runDiscrepancyStage(directory: string, state: LeanState, deps: LeanDeps, progress: (patch: Partial<LeanStageState>) => void): Promise<StageRun> {
  const agent = LEAN_STAGE_AGENTS.discrepancy
  const sessionIDs: string[] = []
  const calls = { count: 0, attempts: 0 }
  const job = async (label: string, first: string, check: (value: unknown) => string[]): Promise<{ value?: Record<string, unknown>; issues: string[] }> => {
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
  const failed = (label: string, issues: string[]) => ({ stop: { status: "failed" as const, reason: `discrepancy: ${label} checker failed after ${LEAN_MAX_RETRIES} retries: ${issues.join(" ")}` } })
  const restart = { stop: { status: "restart" as const } }
  const inputs = sourceInputs(directory)
  const keys = { differences: hashStageInputs("discrepancy", LEAN_PROMPT_VERSION, inputs), gaps: gapsKey(directory) }
  // Each answer is saved only under the inputs it was asked on; a change meanwhile restarts the walk.
  const sourcesChanged = () => differencesKey(directory) !== keys.differences

  const differences = async (): Promise<Outcome<Differences>> => {
    const saved = readLeanJson(directory, DIFFERENCES_FILE) as Differences | undefined
    if (saved?.key === keys.differences && !checkDifferences(directory, saved).length) return { value: saved }
    const asked = await job("differences", evidencePrompt(directory, "discrepancy"), value => checkDifferences(directory, value))
    if (!asked.value) return failed("differences", asked.issues)
    if (sourcesChanged()) return restart
    const value = { key: keys.differences, entries: asked.value.entries as Discrepancy[], limitations: asked.value.limitations as string[] }
    writeLeanText(directory, DIFFERENCES_FILE, `${JSON.stringify(value, null, 2)}\n`)
    return { value }
  }

  const gaps = async (): Promise<Outcome<Gaps>> => {
    const saved = readLeanJson(directory, GAPS_FILE) as Gaps | undefined
    if (saved?.key === keys.gaps && !checkGaps(directory, saved).length) return { value: saved }
    const asked = await job("gaps", gapsPrompt(directory), value => checkGaps(directory, value))
    if (!asked.value) return failed("gaps", asked.issues)
    if (gapsKey(directory) !== keys.gaps) return restart
    const value = { key: keys.gaps, gaps: asked.value.gaps as Gap[], limitations: asked.value.limitations as string[] }
    writeLeanText(directory, GAPS_FILE, `${JSON.stringify(value, null, 2)}\n`)
    return { value }
  }

  // One verify job per gap, within the workflow limit. A job that never passes
  // leaves its gap unverified (assembleDiscrepancy). An error stops new jobs,
  // waits for the running ones and is rethrown, so no job outlives the stage.
  const verifyGaps = async (candidates: Gap[]): Promise<Outcome<(GapVerify | undefined)[]>> => {
    const saved = (readLeanJson(directory, VERIFY_FILE) as { records?: Verified[] } | undefined)?.records ?? []
    const jobKeys = candidates.map(gap => verifyKey(inputs, gap))
    const records = jobKeys.map(key => saved.find(row => row.key === key && !checkGapVerify(directory, row).length))
    const save = () => writeLeanText(directory, VERIFY_FILE, `${JSON.stringify({ records: records.filter(row => row !== undefined) }, null, 2)}\n`)
    const stopped: { error?: { cause: unknown }; restart?: boolean } = {}
    await parallel(jobKeys, async (key, index) => {
      if (records[index] || stopped.error || stopped.restart) return
      try {
        const asked = await job(`verify gap ${index + 1}`, gapVerifyPrompt(directory, candidates[index]!), value => checkGapVerify(directory, value))
        if (!asked.value) return
        if (sourcesChanged()) { stopped.restart = true; return }
        records[index] = { key, supplied: asked.value.supplied as boolean, evidence: asked.value.evidence as SourceQuote[], limitations: asked.value.limitations as string[] }
        save()
      } catch (error) { stopped.error ??= { cause: error } }
    })
    if (stopped.error) throw stopped.error.cause
    return stopped.restart ? restart : { value: records }
  }

  const listedAndVerified = async (): Promise<Outcome<{ gaps: Gaps; verified: (GapVerify | undefined)[] }>> => {
    const listed = await gaps()
    if ("stop" in listed) return listed
    const verified = await verifyGaps(listed.value.gaps)
    return "stop" in verified ? verified : { value: { gaps: listed.value, verified: verified.value } }
  }

  // DIFFERENCES runs beside GAPS and its verify jobs. Both settle before the
  // stage does, so neither outlives it; an error from either is rethrown.
  const running = [differences(), listedAndVerified()] as const
  await Promise.allSettled(running)
  const [compared, listed] = [await running[0], await running[1]]
  if ("stop" in compared || "stop" in listed) {
    const stops = [compared, listed].flatMap(row => "stop" in row ? [row.stop] : [])
    const reasons = stops.flatMap(row => row.status === "failed" ? [row.reason] : [])
    return reasons.length < stops.length ? { status: "restart" } : { status: "failed", reason: reasons.join(" ") }
  }
  const assembled = assembleDiscrepancy(compared.value, listed.value.gaps, listed.value.verified)
  const accepted = acceptEvidenceReply(directory, "discrepancy", JSON.stringify(assembled))
  if (accepted.issues.length) return { status: "failed", reason: `discrepancy: ${accepted.issues.join(" ")}` }
  return { status: "done", outputs: accepted.outputs, openItems: accepted.openItems }
}

/** discrepancy.json, in order: the differences' entries; then the gaps no
 * matter source supplies, each with any passage its verify job found speaking
 * to it, less any that restates a difference (restatesDifference); then, under
 * `supplied`, the gaps a source supplies, each with the words that supply it.
 * A gap whose verify job never passed stays a gap, with UNVERIFIED_GAP among
 * its limitations. `supplied` is left out when empty, so a record with none
 * keeps its earlier form. */
export function assembleDiscrepancy(differences: { entries: Discrepancy[]; limitations: string[] }, listed: { gaps: Gap[]; limitations: string[] }, verified: (GapVerify | undefined)[]) {
  const gaps: Gap[] = []
  const supplied: Gap[] = []
  listed.gaps.forEach((gap, index) => {
    const check = verified[index]
    const limitations = [...new Set([...gap.limitations, ...(check ? check.limitations : [UNVERIFIED_GAP])])]
    if (check?.supplied) return void supplied.push({ ...gap, evidence: check.evidence, limitations })
    const evidence = [...new Map([...gap.evidence, ...check?.evidence ?? []].map(ref => [quoteKey(ref), ref])).values()]
    const row = { ...gap, evidence, limitations }
    if (!restatesDifference(row, differences.entries)) gaps.push(row)
  })
  return { entries: differences.entries, gaps, ...(supplied.length ? { supplied } : {}), limitations: [...new Set([...differences.limitations, ...listed.limitations])] }
}
