import type { Part } from "../../api/session-view"
import type { ChangeHunk, CompareParams, ProposalOutcome } from "../../api/ingest/compare"
import type { Citation } from "../../api/opencode/parts"
import type { Todo } from "../TodoView"
import { basename, stepLabel } from "./step-labels"
import { isDraftingRecoveryMessage } from "./drafting-recovery"
import { templateResult, type TemplateResult } from "./template-result"

// One row in the assistant's reasoning timeline: a thinking block or a tool call.
export type Step =
  | { kind: "reasoning"; text: string; done: boolean }
  | { kind: "tool"; label: string; status: "running" | "done" | "error"; detail?: string; templateResult?: TemplateResult; children?: Step[] }

// A redline/tracked-change a turn proposed: the clause delta plus a pointer back
// to the document, so the assistant bubble can preview the red/green change and
// link into the full tracked-changes review (see RedlineView, DocumentViewer).
export type RedlineProposal = { id?: number; document: string; oldText: string; newText: string; comment?: string }

// A comparison of two Word versions the compare-versions tool ran: what to send
// back to the blackline route, the counts the turn's card reports, and the
// changes its pages listed, so the answer can draw each one it names.
export type VersionCompare = Required<CompareParams> & { changes: number; untracked: number; hunks: ListedChange[] }
// A proposal (P-id) also carries its outcome and the untracked edit it names.
export type ListedChange = Pick<ChangeHunk, "id" | "kind" | "before" | "after"> & Partial<Pick<ProposalOutcome, "outcome" | "untracked">>

// id is the server message id, carried on user turns so edit/retry can revert
// the session back to that exact message. Optimistic turns added before a
// finalize reload have no id, so their actions stay hidden until settled.
// agent is the assistant that produced the turn — the server stamps it on the
// user message, so an assistant turn inherits the agent of the user turn it
// answered. Carried so mixed-agent threads read clearly (see agentLabel).
export type Turn = {
  role: "user" | "assistant"
  text: string
  citations: Citation[]
  redlines: RedlineProposal[]
  // Documents the turn created (draft-document tool), as matter-relative names —
  // each renders as a card linking into the document viewer.
  drafts: string[]
  // Versions the turn compared (compare-versions tool), each a card that opens
  // the blackline view.
  compares?: VersionCompare[]
  // The latest completeness self-audit checklist the turn streamed via the
  // upstream todowrite tool — draft mode's term/number/cross-reference audit or
  // the reviewer's per-area present/absent/deviating walk.
  todos: Todo[]
  steps: Step[]
  id?: string
  agent?: string
  // A failed assistant turn carries the provider's error message instead of an
  // answer. Without surfacing it the turn has no text/steps and toTurns would
  // drop it, so the chat would silently show nothing — the model appears to
  // "think" then vanish. Render this so failures (bad project, quota, auth) are
  // visible and actionable.
  error?: string
  // A durable recovery user message starts a new assistant boundary and marks
  // the preceding interrupted assistant group as replaced by that recovery.
  superseded?: boolean
}

// Reduce a turn's parts to its visible text, search-document citations, and
// any redline proposals. Each text part (one per step) stays its own block, as
// in the live view (readTurn).
function contentOf(parts: Part[]) {
  const text = parts
    .filter((p): p is Extract<Part, { type: "text" }> => p.type === "text")
    .filter((p) => !p.synthetic)
    .map((p) => p.text)
    .join("\n\n")
  const citations = parts
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => p.tool === "search-document" && p.state.status === "completed")
    .flatMap((p) => ((p.state as { metadata?: { citations?: Citation[] } }).metadata?.citations ?? []))
  return { text, citations, redlines: redlinesOf(parts), drafts: draftsOf(parts), compares: comparesOf(parts), todos: todosOf(parts) }
}

// The upstream todowrite tool emits the WHOLE checklist on every call (it is a
// snapshot, not a delta), so the latest todowrite part is the current state of
// the turn's completeness audit — earlier snapshots are superseded. Read the
// todos off the most recent call's input (the live in-flight value) and fall
// back to its output (the settled value) so a streaming audit shows while the
// turn is still working.
function todosOf(parts: Part[]): Todo[] {
  const last = parts
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => p.tool === "todowrite")
    .at(-1)
  if (!last) return []
  const fromInput = "input" in last.state ? (last.state.input as { todos?: Todo[] } | undefined)?.todos : undefined
  const fromOutput =
    last.state.status === "completed" ? (last.state.metadata as { todos?: Todo[] } | undefined)?.todos : undefined
  return fromInput ?? fromOutput ?? []
}

// The draft-document tool stamps the created document's path on its metadata.
// Surface each as a document name the bubble can card and link into the viewer.
function draftsOf(parts: Part[]): string[] {
  return parts
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => p.tool === "draft-document" && p.state.status === "completed")
    .flatMap((p) => {
      const document = (p.state as { metadata?: Record<string, unknown> }).metadata?.document
      return typeof document === "string" ? [basename(document)] : []
    })
}

// The compare-versions tool stamps the comparison it ran on its metadata. Keep
// only a complete record: anything else could open the wrong blackline. Pages
// of one comparison make one entry holding every change they listed.
function comparesOf(parts: Part[]): VersionCompare[] {
  const compares = new Map<string, VersionCompare>()
  for (const p of parts) {
    if (p.type !== "tool" || p.tool !== "compare-versions" || p.state.status !== "completed") continue
    const metadata = (p.state as { metadata?: { compare?: unknown; hunks?: unknown } }).metadata
    const compare = metadata?.compare as Record<string, unknown> | undefined
    const options = (compare?.options ?? {}) as Record<string, unknown>
    const counts = (compare?.counts ?? {}) as Record<string, unknown>
    const text = (value: unknown) => typeof value === "string" && value.length > 0
    if (!compare || !text(compare.left) || !text(compare.right) || typeof counts.changes !== "number" || typeof counts.untracked !== "number") continue
    const entry: VersionCompare = {
      left: compare.left as string,
      right: compare.right as string,
      leftLabel: text(compare.leftLabel) ? (compare.leftLabel as string) : "",
      rightLabel: text(compare.rightLabel) ? (compare.rightLabel as string) : "",
      ignoreFormatting: options.ignoreFormatting !== false,
      detectMoves: options.detectMoves !== false,
      headersFooters: options.headersFooters !== false,
      changes: counts.changes,
      untracked: counts.untracked,
      hunks: [],
    }
    const key = JSON.stringify([entry.left, entry.right, entry.leftLabel, entry.rightLabel, entry.ignoreFormatting, entry.detectMoves, entry.headersFooters])
    const seen = compares.get(key) ?? entry
    compares.set(key, seen)
    for (const hunk of Array.isArray(metadata?.hunks) ? metadata.hunks : [])
      if (listedChange(hunk) && !seen.hunks.some((known) => known.id === hunk.id)) seen.hunks.push(hunk)
  }
  return [...compares.values()]
}

const CHANGE_KINDS: unknown[] = ["modified", "inserted", "deleted", "moved"] satisfies ChangeHunk["kind"][]
const OUTCOMES: unknown[] = [undefined, "pending", "countered", "accepted", "rejected"] satisfies (ProposalOutcome["outcome"] | undefined)[]
function listedChange(value: unknown): value is ListedChange {
  const hunk = value as Partial<Record<keyof ListedChange, unknown>> | null
  const optional = (text: unknown) => text === undefined || typeof text === "string"
  return typeof hunk?.id === "string" && CHANGE_KINDS.includes(hunk.kind) && optional(hunk.before) && optional(hunk.after)
    && (typeof hunk.before === "string" || typeof hunk.after === "string") && OUTCOMES.includes(hunk.outcome) && optional(hunk.untracked)
}

// The redline and tracked-changes tools both record a pending proposal and return
// its delta in their metadata: tracked-changes as find/replace, redline as the
// located clause's old text and its replacement. Surface either as a uniform
// old/new proposal pointing at its document so the bubble can preview and link it.
function redlinesOf(parts: Part[]): RedlineProposal[] {
  return parts
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => (p.tool === "redline" || p.tool === "tracked-changes") && p.state.status === "completed")
    .flatMap((p) => {
      const meta = (p.state as { metadata?: Record<string, unknown> }).metadata ?? {}
      const document = meta.document as string | undefined
      if (!document) return []
      const oldText = (p.tool === "redline" ? meta.oldText : meta.find) as string | undefined
      const newText = (p.tool === "redline" ? meta.replacement : meta.replace) as string | undefined
      return [{ id: meta.redline as number | undefined, document, oldText: oldText ?? "", newText: newText ?? "", comment: meta.comment as string | undefined }]
    })
}

// Turn the ordered parts into the visible reasoning timeline: each thinking
// block and each tool call becomes a step, in arrival order. The final answer
// text and citations are handled by contentOf — these are the intermediate work.
// `pool`, when given, is every part received this turn across the workflow and
// its subagent sessions. A task step then nests the steps of the subagent it
// spawned (linked by the child sessionId on the task part's metadata), so the
// reviewer's own thinking streams live beneath "Consulted the Reviewer" instead
// of the step sitting opaque while the child works.
function partsToSteps(parts: Part[], matter: string, pool?: Part[]): Step[] {
  return parts.flatMap((p): Step[] => {
    if (p.type === "reasoning") {
      const reasoning = visibleReasoning(p.text)
      return reasoning ? [{ kind: "reasoning", text: reasoning, done: Boolean(p.time.end) }] : []
    }
    if (p.type === "tool") {
      // The completeness self-audit gets its own card (TodoView); keep its raw
      // todowrite calls out of the step timeline so they do not read as a dozen
      // repeated "Todowrite" steps alongside the narrated work.
      if (p.tool === "todowrite") return []
      const status = toolStepStatus(p)
      const childID =
        pool && p.tool === "task"
          ? ((p.state as { metadata?: { sessionId?: string } }).metadata?.sessionId ?? undefined)
          : undefined
      const childParts = childID
        ? pool!.filter((c) => c.sessionID === childID && (c.type === "reasoning" || c.type === "tool"))
        : []
      const children = childParts.length ? partsToSteps(childParts, matter, pool) : undefined
      // The tool's own output (or error) — surfaced behind a disclosure so the
      // timeline stays a one-line narration but the underlying work it explored
      // (a search's passages, a read's content, a failure's reason) is one click
      // away. Subagent task steps narrate through their nested children instead.
      const input = "input" in p.state ? (p.state.input as Record<string, unknown> | undefined) : undefined
      const raw =
        p.state.status === "completed" ? p.state.output : p.state.status === "error" ? p.state.error : undefined
      // The python tool runs in a sandbox the lawyer never sees — surface the code
      // it executed, not just the output, so a step that returns "['main.py']" is
      // auditable. The source is what ran; the output is only its echo, useless on
      // its own. Prepend the code; keep the output beneath it when there is one.
      const code =
        p.tool === "python_run_python_code"
          ? ((input?.python_code ?? input?.code ?? input?.source) as string | undefined)
          : undefined
      const body =
        code?.trim() && raw?.trim()
          ? `${code.trim()}\n\n--- output ---\n${raw.trim()}`
          : (code?.trim() ?? raw?.trim())
      const emptyTemplates = p.tool === "list-templates" && p.state.status === "completed" && body === "[]"
      const detail = !children && body && !emptyTemplates ? body : undefined
      const formatted = p.state.status === "completed" && detail
        ? templateResult(p.tool, detail, p.state.metadata)
        : undefined
      // A read that finds no .dochaus/profile.md is the firm-profile skill's
      // expected "no profile yet" branch, not a failure — settle it as a
      // neutral step instead of an error.
      if (
        status === "error" &&
        p.tool === "read" &&
        typeof input?.filePath === "string" &&
        input.filePath.endsWith(".dochaus/profile.md")
      )
        return [{ kind: "tool", label: "No practice profile on file", status: "done", detail }]
      return [{ kind: "tool", label: emptyTemplates ? "No saved templates found" : stepLabel(p, matter), status, detail, children, ...(formatted ? { templateResult: formatted } : {}) }]
    }
    return []
  })
}

export function toolStepStatus(p: Extract<Part, { type: "tool" }>): "running" | "done" | "error" {
  if (p.state.status === "error") return "error"
  if (p.state.status !== "completed") return "running"
  if (p.tool === "task" && p.state.output.includes('<handoff_gate status="blocked">')) return "error"
  return "done"
}

// Some OpenAI-compatible thinking models emit only a Markdown separator in the
// reasoning channel. It carries no thought content, so rendering it creates an
// empty "Thought process" disclosure. Preserve all substantive Markdown and
// discard only whitespace and standalone horizontal-rule markers.
export function visibleReasoning(text: string) {
  const trimmed = text.trim()
  if (!trimmed || /^(?:-{3,}|\*{3,}|_{3,})$/.test(trimmed)) return undefined
  return trimmed
}

// A real answer always contains at least one letter or digit. When a turn
// settles with text that is only braces or punctuation (e.g. a lone "}" leaked
// from a truncated or malformed generation), it is not an answer — render a
// retry affordance instead of the garbage. Citations still show below.
export const isEmptyAnswer = (text?: string) => !text || !/[\p{L}\p{N}]/u.test(text)

// Aggregate every part received this turn into the live view. Reasoning and
// tool parts are only ever emitted by the assistant, so the timeline and
// citations read straight off them — independent of the role map. The role map
// gates only the answer text, so the user's own prompt part is never echoed
// back into the assistant bubble. Gating steps on the role would strand them as
// "Thinking..." whenever a part.updated outraces its message.updated.
export function readTurn(parts: Map<string, Part>, roles: Map<string, string>, matter: string, sessionID: string) {
  const all = [...parts.values()]
  // The answer and the top-level timeline belong to the workflow's own session;
  // subagent-session parts are pooled only to nest under their task step.
  const own = all.filter((p) => p.sessionID === sessionID)
  const text = own
    .filter((p): p is Extract<Part, { type: "text" }> => p.type === "text")
    .filter((p) => roles.get(p.messageID) === "assistant")
    .map((p) => p.text)
    .join("\n\n")
  const work = own.filter((p) => p.type === "reasoning" || p.type === "tool")
  const citations = work
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => p.tool === "search-document" && p.state.status === "completed")
    .flatMap((p) => ((p.state as { metadata?: { citations?: Citation[] } }).metadata?.citations ?? []))
  // The reviewer subagent runs its present/absent/deviating audit in its own
  // child session, so pool every session's todowrite parts (in arrival order)
  // for the audit — the latest snapshot across the turn and its subagents is the
  // live state — while text/citations stay scoped to the workflow's own session.
  const todos = todosOf(all.filter((p) => p.type === "tool" && p.tool === "todowrite"))
  return {
    text,
    citations,
    redlines: redlinesOf(work),
    drafts: draftsOf(work),
    compares: comparesOf(work),
    todos,
    steps: partsToSteps(work, matter, all),
  }
}

// Group stored messages into turns. A tool-using turn spans several consecutive
// assistant messages (one per model step); merging their parts reconstructs the
// whole reasoning timeline instead of showing only the final answer bubble.
export function toTurns(
  msgs: { info: { id: string; role: "user" | "assistant"; agent?: string; error?: MessageError }; parts: Part[] }[],
  matter: string,
  pool?: Part[],
): Turn[] {
  const groups: {
    role: "user" | "assistant"
    id: string
    parts: Part[]
    agent?: string
    error?: string
    superseded?: boolean
  }[] = []
  // An assistant turn answers the most recent user turn, so it inherits that
  // user message's agent — the server only stamps the agent on user messages.
  let lastAgent: string | undefined
  let recoveryBoundary = false
  for (const m of msgs) {
    if (m.info.role === "user" && isDraftingRecoveryMessage(m)) {
      // The recovery instruction is hidden, but it is still a real turn
      // boundary. Without this boundary, the successful assistant response is
      // merged into the preceding failed assistant group and inherits its
      // stale error.
      const interrupted = groups.at(-1)
      if (interrupted?.role === "assistant") interrupted.superseded = true
      recoveryBoundary = true
      continue
    }
    if (m.info.role === "user" && m.info.agent) lastAgent = m.info.agent
    const last = groups[groups.length - 1]
    if (!recoveryBoundary && last && last.role === "assistant" && m.info.role === "assistant") {
      last.parts.push(...m.parts)
      last.error = last.error ?? errorOf(m.info.error)
    } else
      groups.push({ role: m.info.role, id: m.info.id, parts: [...m.parts], agent: lastAgent, error: errorOf(m.info.error) })
    recoveryBoundary = false
  }
  return groups
    .map((g) => {
      // The reviewer subagent's completeness audit is recorded in its child
      // session (carried in `pool`), so for an assistant turn read the audit off
      // every pooled todowrite part — matching how the steps timeline nests the
      // subagent's work — rather than only the turn's own session.
      const todos =
        g.role === "assistant" && pool
          ? todosOf([...g.parts, ...pool.filter((p) => p.type === "tool" && p.tool === "todowrite")])
          : todosOf(g.parts)
      return {
        role: g.role,
        id: g.id,
        agent: g.agent,
        error: g.error,
        superseded: g.superseded,
        ...contentOf(g.parts),
        todos,
        steps: g.role === "assistant" ? partsToSteps(g.parts, matter, pool) : [],
      }
    })
    .filter((t) => t.text || t.citations.length || t.steps.length || t.error)
}

// The provider error the engine stamps on a failed assistant message. Shape is
// { name, data: { message, ... } } (see the APIError surfaced by the prompt
// route); fall back to the error name if no message string is present.
export type MessageError = { name?: string; message?: string; data?: { message?: string; error?: { message?: string } } }
export function errorOf(error: MessageError | undefined) {
  if (!error) return undefined
  return error.data?.message ?? error.data?.error?.message ?? error.message ?? error.name ?? "The model returned an error."
}
