import type { Part } from "../../api/session-view"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import { basename, stepLabel } from "../chat/step-labels"
import { checklistFindingsCount, notesWording, unresolvedFindingsPhrase } from "../drafting/model"

// The deterministic half of the drafting narration: plain activity lines read
// from the stage sessions' streamed parts and from the run's stage rows. No
// model is involved, so these lines appear the moment the run does something.

export type ActivityLine = {
  /** The part id or milestone key. Event refetches re-emit every part, so a
   * line is replaced in place rather than repeated. */
  id: string
  at: number
  kind: "tool" | "thought" | "milestone"
  /** "paused" is a stop the lawyer can Continue from with nothing lost. */
  state: "running" | "done" | "error" | "paused"
  text: string
  /** Consecutive lines of one group fold into a single row (see TOOL_GROUPS). */
  group?: string
  /** What the line acted on, shown as the folded row's latest target. */
  target?: string
  /** The full text where the line shows a clipped one. */
  detail?: string
  /** A result the lawyer should notice, such as a quotation not found word for word. */
  flag?: boolean
  /** The run step the line belongs to: a stage line's own step, or the step of
   * the stage session that produced an activity line. */
  step?: string
  /** When a stage line finished; its duration runs from `at` to here. */
  until?: number
}

type ToolPart = Extract<Part, { type: "tool" }>

// The run's own working files, named as a lawyer would know them rather than by
// their file names.
const RUN_FILES: Record<string, string> = {
  "request.md": "your request",
  "plan.json": "the drafting plan",
  "draft.md": "the working draft",
  "review.md": "the review notes",
  "review-playbook.md": "the playbook review notes",
  "review-independent.md": "the independent review notes",
  "review-wording.md": "the wording review notes",
  "review-client-disadvantage.md": "the client disadvantage review notes",
  "change-plan.json": "the amendment plan",
  "change-preparation.json": "the prepared amendments",
  "changes.json": "the reconciled change list",
  "change-review.json": "the proposed amendment review",
  "drafting-notes.md": "the drafting notes",
}

// Each stage in words: what it is doing while it runs, and its finished line.
const STAGES: Record<string, { during: string; done: string }> = {
  discrepancy: { during: "assessing discrepancies in the source materials", done: "Source discrepancies assessed" },
  intake: { during: "checking material decisions before research", done: "Material decisions checked" },
  research: { during: "researching legal authority", done: "Research findings and limitations recorded" },
  "research-review": { during: "checking research against retrieved sources", done: "Research support and limitations reviewed" },
  spine: { during: "building the source-backed drafting spine", done: "Drafting spine prepared" },
  draft: { during: "drafting", done: "Draft written" },
  review: { during: "reviewing the draft", done: "Reviewed the draft" },
  "review-playbook": { during: "Checking the applicable playbook", done: "Checked the applicable playbook" },
  "review-independent": { during: "independently checking function and fidelity", done: "Independent function and fidelity review completed" },
  "review-wording": { during: "reviewing problematic wording", done: "Problematic wording review completed" },
  "review-client-disadvantage": { during: "reviewing client disadvantage", done: "Client disadvantage review completed" },
  "change-plan": { during: "assessing review findings for amendments", done: "Amendment plan prepared" },
  "change-prepare": { during: "preparing amendment groups", done: "Amendment groups prepared" },
  changes: { during: "reconciling the proposed amendments", done: "Proposed amendments reconciled" },
  "change-review": { during: "independently checking the amended document", done: "Amended document checked" },
  apply: { during: "applying the changes", done: "Changes applied" },
  finalize: { during: "writing the drafting notes", done: "Drafting notes written" },
  "checklist-check": { during: "checking source requirements against the generated documents", done: "Source requirements checked against the documents" },
  "package-review": { during: "reviewing final semantics and the document package", done: "Final semantic and package review completed" },
  gate: { during: "running the final check", done: "Final workflow check passed" },
}

/** How a run of one tool reads once folded into a single row. */
export const TOOL_GROUPS: Record<string, { running: string; done: string; flagged?: string }> = {
  read: { running: "Reading source material", done: "Read source material" },
  search: { running: "Searching matter documents", done: "Searched matter documents" },
  "web-search": { running: "Searching the web", done: "Searched the web" },
  "web-fetch": { running: "Retrieving web sources", done: "Retrieved web sources" },
  "source-fetch": { running: "Retrieving source passages", done: "Retrieved source passages" },
  quote: { running: "Checking quotations", done: "Checked quotations", flagged: "not found word for word" },
}

const SEARCH_CHARS = 60

export function clip(text: string, max: number) {
  const flat = text.replace(/\s+/g, " ").trim()
  return flat.length <= max ? flat : `${flat.slice(0, max - 1).trimEnd()}…`
}

export function fileLabel(target: string) {
  if (target.includes("/.dochaus/") || target.startsWith(".dochaus/"))
    return RUN_FILES[basename(target)] ?? "the run's working files"
  return basename(target)
}

/** One activity line for a stage session's part; undefined for parts that are
 * not activity (answer text, compaction markers, empty thoughts). */
export function partActivity(part: Part, at: number, matter: string): ActivityLine | undefined {
  if (part.type === "reasoning") {
    const text = reasoningHeadline(part.text)
    if (!text) return undefined
    return { id: part.id, at: part.time.start || at, kind: "thought", state: part.time.end ? "done" : "running", text }
  }
  if (part.type !== "tool") return undefined
  const state = part.state.status === "completed" ? "done" : part.state.status === "error" ? "error" : "running"
  const activity = toolActivity(part, state, matter)
  // A failed step stays on its own line so it is never folded out of sight.
  return { id: part.id, at, kind: "tool", state, ...activity, group: state === "error" ? undefined : activity.group }
}

/** The latest bold heading a thinking block opened, else its first sentence. */
export function reasoningHeadline(text: string) {
  const headings = [...text.matchAll(/^\s*\*\*(.+?)\*\*\s*$/gm)].map((match) => match[1]!)
  const line = headings.at(-1) ?? text.trim().split(/(?<=[.!?])\s|\n/)[0] ?? ""
  const plain = clip(line.replace(/[*_`#>]/g, ""), 110)
  return plain || undefined
}

type ToolActivity = Pick<ActivityLine, "text" | "group" | "target" | "detail" | "flag">

function toolActivity(part: ToolPart, state: ActivityLine["state"], matter: string): ToolActivity {
  const input = part.state.input
  const running = state === "running"
  const document = typeof input.document === "string" && input.document ? fileLabel(input.document) : undefined
  if (["web-search", "websearch"].includes(part.tool)) {
    const query = typeof input.query === "string" ? input.query.trim() : ""
    const target = query ? `"${clip(query, SEARCH_CHARS)}"` : undefined
    if (state === "error") return { text: "A web search did not complete" }
    if (part.state.status === "completed" && part.state.metadata?.authenticationFailed === true)
      return { text: "A web search could not authenticate", group: "web-search", target, flag: true }
    if (part.state.status === "completed" && part.state.metadata?.results === false)
      return { text: "A web search returned no results", group: "web-search", target, flag: true }
    return { text: `${running ? "Searching the web" : "Searched the web"}${target ? ` for ${target}` : ""}`, group: "web-search", target }
  }
  if (part.tool === "webfetch" || part.tool === "official-source") {
    const url = typeof input.url === "string" ? input.url : ""
    const host = URL.canParse(url) ? new URL(url).hostname : undefined
    const source = part.tool === "official-source" ? "source passages" : "a web page"
    const where = host ? ` from ${host}` : ""
    if (state === "error") return { text: `Could not retrieve ${source}${where}` }
    if (part.state.status === "completed" && part.state.metadata?.blocked === true)
      return { text: `Could not retrieve ${source}${where}`, group: part.tool === "official-source" ? "source-fetch" : "web-fetch", target: host, flag: true }
    const partial = part.state.status === "completed" && part.state.metadata?.partial === true
    return { text: `${running ? "Retrieving" : "Retrieved"} ${source}${where}${partial ? " (partial retrieval)" : ""}`, group: part.tool === "official-source" ? "source-fetch" : "web-fetch", target: host, flag: partial || undefined }
  }
  if (part.tool === "read-document" || part.tool === "read") {
    const path = [input.filePath, input.path].find((value) => typeof value === "string" && value)
    const target = document ?? (typeof path === "string" ? fileLabel(path) : undefined)
    if (state === "error") return { text: `Could not read ${target ?? "a document"}` }
    return { text: `${running ? "Reading" : "Read"} ${target ?? "a document"}`, group: "read", target }
  }
  if (part.tool === "search-document") {
    const where = document ?? "the matter documents"
    const full = typeof input.query === "string" ? input.query.trim() : ""
    if (state === "error") return { text: `A search of ${where} did not complete` }
    if (!full) return { text: running ? `Searching ${where}` : `Searched ${where}`, group: "search" }
    const query = `"${clip(full, SEARCH_CHARS)}"`
    const text = `${running ? "Searching" : "Searched"} ${where} for ${query}`
    return { text, group: "search", target: query, detail: full.length > SEARCH_CHARS ? `"${full}"` : undefined }
  }
  if (part.tool === "verify-quote") {
    const where = document ? `in ${document}` : "in the sources"
    if (running) return { text: `Checking a quotation ${where}`, group: "quote", target: document }
    if (state === "error") return { text: `Could not check a quotation ${where}` }
    const output = part.state.status === "completed" ? part.state.output : ""
    if (output.trimStart().startsWith("NOT FOUND"))
      return { text: `A quotation was not found word for word ${where}`, group: "quote", target: document, flag: true }
    return { text: `Confirmed a quotation ${where}`, group: "quote", target: document }
  }
  const label = stepLabel(part, matter)
  return { text: state === "error" ? `${label} (did not complete)` : label }
}

/** One line per stage run, keyed by its start time so a re-run reads as a new
 * line; it updates in place as the stage runs, finishes, pauses or stops. */
export function stageMilestones(status: LeanDraftingStatus): ActivityLine[] {
  const working = isWorking(status)
  return Object.entries(status.stages).flatMap(([step, row]): ActivityLine[] => {
    if (!row.started_at) return []
    const at = Date.parse(row.started_at)
    const line = { id: `${step}:stage:${row.started_at}`, at, kind: "milestone" as const, step }
    if (row.status === "running" && working)
      return [{ ...line, state: "running" as const, text: `${stepActivityLabel(status, step)}${row.attempts > 1 ? " (trying again)" : ""}` }]
    if (row.status === "running" || row.status === "pending")
      return [{ ...line, state: "paused" as const, text: `Paused while ${during(step, status.memo)}`, until: Date.parse(status.updated_at) }]
    const until = row.finished_at ? Date.parse(row.finished_at) : undefined
    // A final check that delivered unresolved checklist findings did not pass
    // clean; flagged, it stays outside the folded activity.
    const findings = step === "gate" && row.status === "done" ? checklistFindingsCount(status) : undefined
    if (findings !== undefined)
      return [{ ...line, state: "done" as const, text: `Final check: review draft with ${unresolvedFindingsPhrase(findings)}`, until, flag: true }]
    if (row.status === "done") return [{ ...line, state: "done" as const, text: words(step, status.memo).done, until }]
    if (row.status === "failed" && !stoppedByTechnicalProblem(status))
      return [{ ...line, state: "error" as const, text: `Stopped while ${during(step, status.memo)}`, until }]
    return [{ ...line, state: "paused" as const, text: pauseText(status, step, row.status), until }]
  })
}

function pauseText(status: LeanDraftingStatus, step: string, state: string) {
  if (state === "failed") return `Paused while ${during(step, status.memo)} after a technical problem`
  if (step !== "gate") return `Paused while ${during(step, status.memo)}`
  const blocking = status.gate?.blocking.length ?? status.open_items.filter((item) => item.status === "open").length
  return `Final check found ${blocking} open item${blocking === 1 ? "" : "s"} to resolve`
}

/** A stage in words; the notes step of a run whose notes are the requested
 * memo (`memo`) names that memo. */
function words(step: string, memo?: string | null) {
  const stage = step.split(".").at(-1) ?? step
  if (stage === "finalize" && memo?.trim()) {
    const notes = notesWording(memo)
    return { during: `writing ${notes.phrase}`, done: `${notes.label} written` }
  }
  return STAGES[stage] ?? { during: step, done: `Finished ${step}` }
}

function during(step: string, memo?: string | null) {
  return words(step, memo).during
}

/** A line for the run starting or stopping work between two polls: a resume
 * after a stop, or a pause because nothing is running the run any more. Stage
 * failures and blocks already have their own milestones. */
export function runTransition(prev: LeanDraftingStatus, next: LeanDraftingStatus) {
  if (isWorking(prev) === isWorking(next)) return undefined
  if (isWorking(next)) {
    const step = activeStep(next)
    return { state: "done" as const, text: step ? `Resumed, picking up with ${during(step, next.memo)}` : "Resumed" }
  }
  if (next.status !== "running") return undefined
  return {
    state: "paused" as const,
    text: "Paused: the drafting service stopped working on this run. The work so far is kept.",
  }
}

export function isWorking(status: LeanDraftingStatus) {
  return status.status === "running" && status.in_flight
}

/** A run that stopped on a runtime fault rather than on its own checks; the
 * runner records such stops in this form, and Continue resumes them. */
export function stoppedByTechnicalProblem(status: LeanDraftingStatus) {
  return status.status === "failed" && (status.stop_reason ?? "").includes("stopped by a runtime error")
}

/** The step the run is working on, if any. */
export function activeStep(status: LeanDraftingStatus) {
  return Object.entries(status.stages).find(([, row]) => row.status === "running")?.[0]
}

/** A plain label for a step while it runs, e.g. "Drafting" or "Reviewing the draft". */
export function stepActivityLabel(status: LeanDraftingStatus, step: string) {
  const text = during(step, status.memo)
  return `${text[0]!.toUpperCase()}${text.slice(1)}`
}
