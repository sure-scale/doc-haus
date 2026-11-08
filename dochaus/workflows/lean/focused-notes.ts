import { parallel } from "./checklist-flow"
import { checkDraftingNotes, finalBodyParagraphs, parseJsonObject, readChangeList, readParagraphs, shownText, type ApplyResult, type ParagraphRow } from "./checks"
import { readChangePlan, reviewDispositionText } from "./change-pipeline"
import { comparisonDescription } from "./comparison"
import { deliveryArtifactEvidence, requestedMemo } from "./deliverables"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { checkFocusedNotesCoverage } from "./focused-notes-coverage"
import { resolveFocusedNotesDependency } from "./focused-notes-dependencies"
import { assembleNotesOwnerEvidence, checkNotesInspection, checkNotesTargets, currentNotesPids, notesCurrentEvidence, notesEvidencePlan, NOTES_CURRENT_EVIDENCE_VERSION, type NotesBody } from "./focused-notes-evidence"
import { raisedFindings } from "./review-findings"
import { matterDatesBlock } from "./intake"
import { readLeanPlan, type DocID } from "./plan"
import { actingForBlock, delegatedLines, focusedNotesPrompt, tagged } from "./prompts"
import { escapeTagOpeners } from "../../lib/untrusted"
import { matterSourceBasis } from "./quotes"
import { notesRepair, readRepairRound } from "./repairs"
import { appliedChangeEvidence, nativeExportEvidence } from "./review-evidence"
import type { StageRun } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, acceptStageReply, disclosableItems, stageDraftingPreferences, stageInputs, stepSessionTitle, type LeanDeps } from "./stages"
import { answeredQuestionsText, delegatedQuestions, hashStageInputs, readLeanJson, readLeanText, writeLeanText, type LeanStageState, type LeanState, type OpenItem } from "./state"

// Actual prose outputs are saved under their content keys. They are the
// accepted sections themselves, not a second progress or coverage record.
const NOTES_SECTIONS_CONTRACT = "focused-notes-sections-2-checked-explanation"
const NOTES_OWNER_CONTRACT = "one-semantic-owner-1-inspected-dependencies"
const PASSAGE_CHARS = 4_000

export type FocusedNotesInput = {
  id: string
  kind: "package" | "document" | "passage" | "change" | "disposition" | "open-item" | "delegated"
  heading: string
  document?: DocID
  pids: string[]
  refs: string[]
  items: OpenItem[]
  delegatedIDs: string[]
  text: string
  needsCurrentLookup?: boolean
}
type Assignment = Omit<FocusedNotesInput, "text"> & {
  question: string
  context: { name: string; text: string }[]
  repairs: string[]
}
type BodyPassage = { pids: string[]; text: string; part: number | null }

/** Every final paragraph is assigned exactly once, respecting its planned
 * part. A long paragraph becomes adjacent passages with the same source PID;
 * its full owning paragraph remains dependency context for writer and checker. */
function bodyPassages(rows: ParagraphRow[]): BodyPassage[] {
  const result: BodyPassage[] = []
  for (const row of rows) {
    const text = shownText(row)
    const part = row.draft_part ?? null
    if (text.length > PASSAGE_CHARS) {
      for (let at = 0; at < text.length; at += PASSAGE_CHARS)
        result.push({ pids: [row.pid], part, text: `[${row.pid}] ${text.slice(at, at + PASSAGE_CHARS)}` })
      continue
    }
    const last = result.at(-1)
    if (last && last.part === part && last.text.length + text.length + 20 <= PASSAGE_CHARS) {
      last.pids.push(row.pid)
      last.text += `\n\n[${row.pid}] ${text}`
    } else result.push({ pids: [row.pid], part, text: `[${row.pid}] ${text}` })
  }
  return result
}

const json = (value: unknown) => JSON.stringify(value, null, 2) ?? "null"
const context = (name: string, value: unknown) => ({ name, text: typeof value === "string" ? value : json(value) })
const requiredNotesRepairs = (directory: string) => (readRepairRound(directory)?.items ?? []).filter(item => item.stage === "package-review" || item.ref.repair_target === "memo" || notesRepair(item))

/** This stage produces the notes, so their pre-render absence is not evidence
 * of a missing requested artifact and must not invalidate its own prose. */
function notesDeliveryEvidence(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(notesDeliveryEvidence)
  if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).map(([key, child]) =>
    [key, key === "file_status" && "kind" in value && value.kind === "drafting_notes" ? "prepared by this notes stage" : notesDeliveryEvidence(child)]))
  return value
}

/** Short controlling instructions belong in the assignment itself. Long
 * instruction files stay available through bounded, explicit navigation; a
 * writer is told they are current and must retrieve the relevant passages,
 * rather than being expected to discover a changed file from an opaque key. */
function humanText(name: string, text: string | undefined, file: string, format: "text" | "json" = "text") {
  if (!text?.trim()) return []
  if (text.length <= 1_600) return [context(name, text)]
  const lines = text.split("\n")
  return [context(`${name}_navigation`, [
    `Current controlling human instructions or answers: ${file}${format === "text" ? ` (${lines.length} lines)` : " (canonical JSON answers record)"}.`,
    `This context is too long to inline. Use ${format === "json" ? "read-record with bounded JSON Pointer pages" : "read in bounded line pages"}, locating the request/updates/answers relevant to this section before writing it. Explicit later instructions control. Do not assume earlier context is current. The excerpts below are navigation only and do not replace the relevant full instruction or answer; retrieve surrounding qualifications and read further if relevance is uncertain.`,
    `Opening excerpt:\n${text.slice(0, 300)}`,
    `Latest excerpt (it may start mid-passage):\n${text.slice(-900)}`,
  ].join("\n\n"))]
}

function currentHumanContext(directory: string, document?: DocID) {
  const answers = answeredQuestionsText(directory, { asked: question => !document || !question.document || question.document === document })
    .replace(/^- \[[^\]\n]+\] Q:/gm, "- Q:")
  return [
    context("notes_role", requestedMemo(directory)
      ? `These sections form the requested accompanying memo, titled ${requestedMemo(directory)!.name}. Follow the memo's controlling audience, purpose and explanation instructions; code supplies the title and section structure.`
      : "These sections form the drafting notes for the lawyer accompanying the drafted documents. Code supplies the title and section structure."),
    ...humanText("acting_for", actingForBlock(directory).split("\n").slice(1, -1).join("\n"), ".dochaus/lean/plan.json", "json"),
    ...humanText("matter_dates", matterDatesBlock(directory).split("\n").slice(1, -1).join("\n"), ".dochaus/lean/plan.json", "json"),
    ...humanText("current_request", readLeanText(directory, "request.md"), ".dochaus/lean/request.md"),
    ...humanText("current_updates", readLeanText(directory, "updated-notes.md"), ".dochaus/lean/updated-notes.md"),
    ...humanText("current_answers", answers, ".dochaus/lean/questions.json", "json"),
  ]
}

function repairOwner(inputs: Pick<FocusedNotesInput, "id" | "kind" | "items" | "delegatedIDs" | "document" | "pids">[], repair: OpenItem) {
  return inputs.find(row => row.items.some(item => item.id === repair.id))
    ?? inputs.find(row => repair.ref.question_id && row.delegatedIDs.includes(repair.ref.question_id))
    ?? inputs.find(row => row.kind === "passage" && row.document === repair.ref.document && repair.ref.pid && row.pids.includes(repair.ref.pid))
    ?? inputs.find(row => row.kind === "document" && repair.ref.document && row.document === repair.ref.document)
}

/** Select exact final target passages; unlocated records retrieve their own
 * targets rather than receiving another copy of the whole document. */
function notesBody(directory: string, document: DocID): NotesBody {
  return { document, initial: readParagraphs(directory, document), current: finalBodyParagraphs(directory, document),
    applied: (readLeanJson(directory, `${document}/apply.json`) as ApplyResult | undefined)?.applied }
}

function relevantBody(directory: string, document: DocID | undefined, pids: string[]) {
  return document ? notesCurrentEvidence(notesBody(directory, document), pids) : []
}

/** Builds independent concrete questions. Coverage comes from iterating the
 * authoritative input records in code, never from asking a writer to copy ids. */
export function focusedNotesInputs(directory: string, linkedRepairs: ReadonlyMap<string, readonly string[]> = new Map()): FocusedNotesInput[] {
  const plan = readLeanPlan(directory)
  const documents = plan?.documents ?? []
  const items = [...new Map(disclosableItems(directory).map(item => [item.id, item])).values()]
  const questions = delegatedQuestions(directory)
  const raised = raisedFindings(directory)
  const recordPids = (document: DocID, value: unknown) => {
    const text = json(value)
    const known = new Set(readParagraphs(directory, document).map(row => row.pid))
    return [...new Set([
      ...[...text.matchAll(/\bp[0-9a-f]{12}\b/g)].map(match => match[0]),
      ...raised.filter(row => row.document === document && text.includes(row.id)).flatMap(row => row.pids ?? []),
      ...items.filter(item => item.ref.document === document && text.includes(item.id)).flatMap(item => item.ref.pid ? [item.ref.pid] : []),
    ].filter(pid => known.has(pid) || finalBodyParagraphs(directory, document).some(row => row.pid === pid)))]
  }
  const assignments: Assignment[] = []
  const add = (input: Pick<Assignment, "id" | "kind" | "heading" | "question" | "context"> & Partial<Assignment>) => {
    const row: Assignment = { pids: [], refs: [], items: [], delegatedIDs: [], repairs: [], ...input }
    assignments.push(row)
    return row
  }
  add({ id: "package", kind: "package", heading: "Documents and audience",
    question: "Explain the delivered files and whom they were drafted for. This is the package overview only; individual document effects, changes, open points and delegated choices have their own sections. A pending notes rendering or registration is normal, not a missing deliverable. A generated note can fulfil an explanatory-note request when its audience, content and naming fit; a substantive client-advice function needs its own appropriate body. Only a side explicitly labelled inferred needs confirmation; a stated or matter-recorded side is settled. Identify any requested file arrangement the supplied evidence does not deliver.",
    context: [context("document_roles", documents.map(row => ({ title: row.title, purpose: row.purpose, file: row.file, deliverable: row.deliverable }))), context("delivery_artifacts", notesDeliveryEvidence(deliveryArtifactEvidence(directory))),
      context("requested_files", readLeanJson(directory, "requested-files.json") ?? "No separate requested-files record.")],
  })
  for (const document of documents) {
    const rows = finalBodyParagraphs(directory, document.id)
    const initial = readParagraphs(directory, document.id)
    const metadata = context("document", { id: document.id, title: document.title, purpose: document.purpose, mode: document.mode, file: document.file })
    const passages = bodyPassages(rows)
    const exported = nativeExportEvidence(directory, document.id)
    const { inserted_text: _inserted, deleted_text: _deleted, ...exportMetadata } = exported as typeof exported & { inserted_text?: string[]; deleted_text?: string[] }
    add({ id: `document:${document.id}`, kind: "document", document: document.id, heading: document.title,
      question: "Explain this document's requested purpose and its comparison file, using the stated comparison baseline exactly. Do not infer comparison absence from clean body text. The passage and change sections separately explain its substantive wording; do not claim uninspected body content or repeat their work. The naming selections are not proof of the requested delivery shape; inspect relevant instructions if uncertain.",
      context: [metadata, context("comparison_baseline", comparisonDescription(directory, document.id)), context("native_export", exportMetadata),
        ...(passages[0] ? [context("opening_final_passage", passages[0].text)] : [])],
    })
    passages.forEach((passage, index) => add({ id: `passage:${document.id}:${index}`, kind: "passage", document: document.id, pids: passage.pids,
      heading: `${document.parts[passage.part ?? -1]?.heading ?? document.title} — Passage ${index + 1}`,
      question: "Explain the material effect of this final passage in this document. This section owns only this passage; retrieve its material definitions, exceptions and dependencies before describing them. Preserve proposal status and disclosed limitations. Do not describe earlier proposals as applied, repeat file delivery explanations or reopen settled choices. Other passages have their own explanation jobs.",
      context: [metadata, context("planned_part", passage.part === null ? "Unattributed final passage." : document.parts[passage.part]), context("final_passage", passage.text),
        ...(rows.some(row => passage.pids.includes(row.pid) && shownText(row).length > PASSAGE_CHARS)
          ? [context("passage_dependency_scope", "The complete owning paragraph, its earlier baseline and nearby qualifications below are dependency context for this assigned range, because a condition may appear beyond its boundary. Explain only final_passage; do not reassign or repeat the whole paragraph's explanation."), ...relevantBody(directory, document.id, passage.pids)] : []),
        context("current_inspection_scope", { version: NOTES_CURRENT_EVIDENCE_VERSION, passage: index + 1, pids: passage.pids, scope: "Only this passage is assigned; other passages continue separately. A passage ending at a heading does not establish that the following content is absent. Do not turn a local boundary into a whole-document omission." })],
    }))
    for (const [index, change] of appliedChangeEvidence(directory, document.id).entries()) {
      const pids = [change.pid, ...change.result_paragraphs.map(row => row.pid)]
      add({ id: `change:${document.id}:${index}`, kind: "change", document: document.id, pids, heading: "Changes made",
        question: "Explain this one change's actual result, source basis and legal or practical effect. Distinguish source-to-initial-draft revisions from later amendments using change_origin. application_mode says only how text was applied, never whether it has source support. Assess the provided exact source passages; disclose a no-source-passage judgment honestly without inventing a body marker or reopening a settled decision. For a genuinely new unsupported legal view, retain the need for lawyer verification. Explain an unresolved native comment honestly; the applied body result still controls.",
        context: [metadata, context("change", change), context("initial_target", initial.filter(row => row.pid === change.pid).map(row => `[${row.pid}] ${shownText(row)}`).join("\n\n")), ...relevantBody(directory, document.id, pids)],
      })
    }
    const changes = readChangeList(directory, document.id)
    const apply = readLeanJson(directory, `${document.id}/apply.json`) as ApplyResult | undefined
    for (const [index, row] of [...(apply?.unapplied ?? []).map(outcome => ({ outcome, proposal: changes?.amendments.find(change => change.id === outcome.id) })), ...(changes?.withheld ?? []).map(held => ({ outcome: { reason: "withheld", issues: held.issues }, proposal: held.amendment }))].entries()) {
      const pids = row.proposal ? [row.proposal.pid] : []
      add({ id: `unapplied:${document.id}:${index}`, kind: "change", document: document.id, pids, heading: "Changes not made",
        question: "Explain this one proposed change that was not made, its material consequence and remaining action. Verify the current wording; never present the proposal as applied or source support as established by an application status.",
        context: [metadata, context("unapplied_change", row), ...relevantBody(directory, document.id, pids)],
      })
    }
    let dispositions: { findings?: unknown[]; legacy_decisions?: string[]; unadjudicated_review_history?: { stage: string; text: string }[] } = {}
    try { dispositions = JSON.parse(reviewDispositionText(directory, document.id)) } catch { /* No accepted plan has no adjudicated dispositions. */ }
    const records = [
      ...(dispositions.findings ?? []).map(value => ({ type: "adjudicated", value })),
      ...(dispositions.legacy_decisions ?? []).map(value => ({ type: "legacy navigation only", value })),
      ...(dispositions.unadjudicated_review_history ?? []).flatMap(row => {
        let value: unknown
        try { value = JSON.parse(row.text) } catch { value = row.text }
        const findings = value && typeof value === "object" && "findings" in value && Array.isArray(value.findings) ? value.findings : [value]
        return findings.filter(Boolean).map(value => ({ type: `unadjudicated ${row.stage}: navigation only`, value }))
      }),
    ]
    records.forEach((record, index) => {
      const detail = record.value && typeof record.value === "object" && "detail" in record.value ? record.value.detail : undefined
      const original = readChangePlan(directory, document.id)?.findings?.find(row => row.detail === detail)
      const pids = recordPids(document.id, { record, resolves: original?.resolves })
      add({ id: `disposition:${document.id}:${index}`, kind: "disposition", document: document.id, pids, heading: "Review decisions and recommendations",
        question: "Explain this one material review disposition, its practical consequence and remaining action in the appropriate advice. A deferred finding remains unresolved, a recommendation is optional, and a rejected suggestion is not an outstanding mandatory change. A legacy decision or unadjudicated review is navigation, not proof of resolution or rejection: inspect its current final passage and underlying sources. Omit only an immaterial rejected suggestion, saying briefly that it was not pursued; preserve material advice and qualifications. Do not assert that proposed wording was applied without inspecting the final text.",
        context: [metadata, context("review_disposition", record), ...relevantBody(directory, document.id, pids)],
      })
    })
  }
  for (const item of items) {
    const document = documents.find(row => row.id === item.ref.document)
    const question = item.ref.question_id ? questions.find(row => row.id === item.ref.question_id) : undefined
    const pids = item.ref.document ? recordPids(item.ref.document, item) : []
    add({ id: `item:${item.id}`, kind: "open-item", document: item.ref.document, pids, heading: document ? `Open point — ${document.title}` : "Open point",
      refs: [item.id], items: [item],
      question: "Explain this one open point, its consequence and exactly what the lawyer needs to do. Keep it unresolved where the record does. For a marked drafted default, state what was chosen and the required confirmation; do not convert an already settled answer into a new confirmation request. This is disclosure, not an answer or release approval. Do not copy its marker or identifiers; code attaches its exact reference.",
      context: [context("open_point", { kind: item.kind, detail: item.detail, location: item.ref }), ...(document ? [context("document", { id: document.id, title: document.title, purpose: document.purpose })] : []),
        ...(question ? [context("related_decision", delegatedLines(question).join("\n"))] : []), ...relevantBody(directory, item.ref.document, pids)],
    })
  }
  for (const question of questions) add({ id: `delegated:${question.id}`, kind: "delegated", document: question.document, pids: question.pid ? [question.pid] : [],
    heading: "Delegated decision", refs: [question.id], delegatedIDs: [question.id],
    question: "Explain this one delegated or unanswered decision: the decision taken, why, its exact evidence and alternatives not taken, and that the lawyer delegated it. An adopted judgment is settled unless contrary evidence prevents it; do not ask to confirm it again. Inspect any affected final wording and disclose a supported departure with its reason. If no decision was taken, explain why the point stays open and the required human action. Delegation supplies neither missing facts nor human release approval. Do not copy the question id or a reference marker; code attaches it.",
    context: [context("question", question.question), context("decision", delegatedLines(question).join("\n")), ...relevantBody(directory, question.document, question.pid ? [question.pid] : [])],
  })

  // Scoped corrections go to their existing owner. A run-wide memo finding
  // has no stored semantic section target; a separate linking job selects the
  // sections to replace. Adding a correction after stale prose would leave a
  // contradictory explanation, so a repair never becomes an additive section.
  for (const repair of requiredNotesRepairs(directory)) {
    const owner = repairOwner(assignments, repair)
    const targets = owner ? [owner.id] : linkedRepairs.get(repair.id) ?? []
    for (const id of targets) assignments.find(row => row.id === id)?.repairs.push(repair.detail)
  }
  return assignments.map(({ question, context, repairs, ...input }) => {
    const mapped = input.document ? currentNotesPids(notesBody(directory, input.document), input.pids) : undefined
    const needsCurrentLookup = Boolean(["change", "disposition", "open-item", "delegated"].includes(input.kind) &&
      (!input.pids.length || mapped?.missing.length))
    return { ...input, ...(needsCurrentLookup ? { needsCurrentLookup: true } : {}), text: focusedNotesPrompt({ kind: input.kind, question,
      context: [...currentHumanContext(directory, input.document), ...context], repairs }) }
  })
}

function sectionText(input: FocusedNotesInput, prose: string) {
  // Headings are presentation only. Preserve their words below code's section
  // heading, preventing a writer's title from becoming another document title.
  const text = prose.trim().replace(/^#{1,6}\s+/gm, "### ")
  return `${text}${input.refs.length ? ` ${input.refs.map(id => `[ref: ${id}]`).join(" ")}` : ""}`
}

/** The ordinary notes checker still verifies every quote, citation, supplied
 * open item, placeholder and reader-facing id. Only delegated coverage is
 * scoped here; the assembled notes pass the default whole-package check. */
export function checkFocusedNotesSection(directory: string, input: FocusedNotesInput, reply: string, final = false) {
  if (!reply.trim()) return ["This notes section is empty. Return its complete explanation in words."]
  if (!reply.split("\n").some(line => line.trim() && !/^\s*#{1,6}\s/.test(line))) return ["This section has a heading but no explanation. Return its substantive prose."]
  if (/\[ref\s*:/i.test(reply)) return ["Write this section's explanation only. The coordinator attaches its reference markers; do not supply markers or coverage bookkeeping."]
  return checkDraftingNotes(directory, `# Notes section\n\n${sectionText(input, reply)}\n`, input.items, { final, delegatedIDs: input.delegatedIDs }).issues
}

export function assembleFocusedNotes(directory: string, sections: { input: FocusedNotesInput; prose: string }[]) {
  const title = (requestedMemo(directory)?.name ?? "Drafting notes").replace(/\s+/g, " ").trim()
  return [`# ${title}`, ...sections.map(({ input, prose }) => `## ${input.heading.replace(/\s+/g, " ").trim()}\n\n${sectionText(input, prose)}`)].join("\n\n") + "\n"
}

/** Focused notes share the coordinator's single writer. Only failed sections
 * retry; accepted prose survives a failed neighbor and an unchanged resume. */
export async function runFocusedNotes(directory: string, state: LeanState, deps: LeanDeps, progress: (patch: Partial<LeanStageState>) => void): Promise<StageRun> {
  let inputs = focusedNotesInputs(directory)
  const stageKey = () => hashStageInputs("finalize", LEAN_PROMPT_VERSION, stageInputs(directory, "finalize"))
  const initial = stageKey()
  const common = [
    { name: "effective-plan", bytes: json(readLeanPlan(directory)) },
    { name: "source-content", bytes: matterSourceBasis(directory) },
    { name: "preferences", bytes: stageDraftingPreferences(directory) },
    { name: "instructions", bytes: [readLeanText(directory, "request.md"), readLeanText(directory, "updated-notes.md"), answeredQuestionsText(directory)].join("\n") },
    { name: "research", bytes: [readLeanText(directory, "research.json"), readLeanText(directory, "research-review.json")].join("\n") },
    // A section can retrieve definitions, conditions and other final passages.
    // Hash that complete dependency evidence without sending it in every prompt.
    { name: "final-dependencies", bytes: json((readLeanPlan(directory)?.documents ?? []).map(document => ({ document: document.id,
      body: finalBodyParagraphs(directory, document.id), changes: readChangeList(directory, document.id), apply: readLeanJson(directory, `${document.id}/apply.json`), dispositions: reviewDispositionText(directory, document.id),
    }))) },
  ]
  const sections = new Map<string, string>()
  const files: string[] = []
  const sessions: string[] = []
  const failures: string[] = []
  const calls = { count: 0, attempts: 0 }
  const stopped: { error?: { cause: unknown } } = {}
  const throwIfStopped = () => { if (stopped.error) throw stopped.error.cause }
  const sectionFile = (input: FocusedNotesInput, owner = false) => `notes-sections/${hashStageInputs("finalize", NOTES_SECTIONS_CONTRACT, [...common,
    ...(owner ? [{ name: "owner-contract", bytes: NOTES_OWNER_CONTRACT }] : []), { name: "assignment", bytes: json(input) }])}.md`
  // Locate only an unowned current target. The accepted prose remains keyed
  // on its unchanged assignment and the existing complete dependency basis,
  // so recovery can reuse it without repeating this transient lookup.
  const locateCurrentInput = async (input: FocusedNotesInput): Promise<FocusedNotesInput | undefined> => {
    if (!input.needsCurrentLookup) return input
    const bodies = (readLeanPlan(directory)?.documents ?? []).map(document => notesBody(directory, document.id))
    const navigation = bodies.map(body => ({ document: body.document, record: `.dochaus/lean/${body.document}/paragraphs-applied.json`,
      paragraphs: body.current.map(({ pid, label, draft_part, draft_heading, text }) => ({ pid, label, part: draft_part, ...(draft_heading ? { heading: text } : {}) })) }))
    const index = json(navigation)
    const agent = LEAN_STAGE_AGENTS.finalize
    const step = `${input.id}:current-targets`
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent,
      title: stepSessionTitle(directory, "finalize", step), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    sessions.push(sessionID)
    progress({ session_id: sessions[0], session_ids: [...sessions] })
    const prompt = [
      tagged("current_target_lookup", "Locate the actual current accepted paragraphs needed to assess this one notes point, including material qualifications, definitions and dependencies. This is evidence selection only: do not draft prose, decide legal correctness, mark a finding resolved or infer approval. Old proposals and deferred findings are navigation, not current-body evidence. Inspect the relevant current paragraphs through read-record or read-document view=current_body. Select their existing current PIDs. Inspect only material dependencies, not the whole matter."),
      tagged("assigned_point", escapeTagOpeners(input.text)),
      tagged("current_navigation", escapeTagOpeners(index.length <= 6_000 ? index : json(navigation.map(({ paragraphs: _paragraphs, ...row }) => row)))),
      tagged("output_contract", 'Reply only {"targets":[{"document":"planned id","pids":["current paragraph pid"]}],"limitation":"remaining inspection limit, or empty"}. If current passages cannot be located or inspected, return no targets with the actual limitation; this never establishes that body content is missing. Do not invent PIDs or write files.'),
    ].join("\n\n")
    let issues: string[] = []
    for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
      appendLeanEvent(directory, { stage: "finalize", step, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: attempt
        ? `Correct only this target-selection reply: ${issues.join(" ")}. Return targets/limitation JSON, retaining this one assignment.` : prompt })
      const answered = performance.now()
      const parsed = parseJsonObject(reply)
      const checked = parsed.ok ? checkNotesTargets(parsed.value, bodies) : { issues: [parsed.error] }
      issues = checked.issues
      calls.count++
      calls.attempts = Math.max(calls.attempts, attempt + 1)
      progress({ calls: calls.count, attempts: calls.attempts, feedback: issues })
      appendLeanEvent(directory, { stage: "finalize", step, event: "checker", attempt, issues,
        wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (checked.targets) {
        const evidence = checked.targets.flatMap(target => notesCurrentEvidence(bodies.find(body => body.document === target.document)!, target.pids))
        const strip = ["current_inspection_scope", "generated_initial_evidence", "current_accepted_evidence", "baseline_meaning"]
        const text = strip.reduce((text, name) => text.replace(new RegExp(`<${name}>[\\s\\S]*?</${name}>\\s*`, "g"), ""), input.text)
        return { ...input, needsCurrentLookup: false, text: [text,
          ...evidence.map(row => tagged(row.name, escapeTagOpeners(row.text))),
          tagged("current_lookup_limit", escapeTagOpeners(checked.limitation || (checked.targets.length ? "Only the selected current passages and their inlined neighbors have been inspected; do not claim document-wide absence." : "No current passage was established. Do not turn missing inspection into missing content or completed approval."))),
        ].join("\n\n") }
      }
    }
    failures.push(`${step}: current target lookup did not pass its check: ${issues.join(" ")}`)
  }
  const linked = new Map<string, string[]>()
  const unscoped = requiredNotesRepairs(directory).filter(repair => !repairOwner(inputs, repair))
  await parallel(unscoped, async repair => {
    if (stopped.error) return
    try {
      const agent = LEAN_STAGE_AGENTS.finalize
      const step = `repair-link:${repair.id}`
      const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, "finalize", step), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
      sessions.push(sessionID)
      progress({ session_id: sessions[0], session_ids: [...sessions] })
      const catalog = inputs.map((input, index) => ({ index, kind: input.kind, heading: input.heading, document: input.document, pids: input.pids,
        open_point: input.items.map(item => item.detail.slice(0, 300)),
        assignment: /<assignment>\n([\s\S]*?)\n<\/assignment>/.exec(input.text)?.[1],
        existing_section: readLeanText(directory, sectionFile(input)) ? `.dochaus/lean/${sectionFile(input)}` : null,
      }))
      let text = [
        tagged("assignment", "This is a separate repair-link job, not prose drafting. Select the existing notes sections whose explanation this one required repair must replace. Read the actual old notes at .dochaus/lean/drafting-notes.md with the read tool, and existing_section paths where available. read-record is only for JSON. Match the actual erroneous statement and every duplicate to the owning section assignments; route a missing explanation to its natural owner. Do not simply append a correction or select only a new overview while leaving the old contradiction. Retrieve only relevant evidence and source passages. The catalog describes output responsibilities, not authority."),
        tagged("required_repair", escapeTagOpeners(repair.detail)),
        tagged("section_catalog", escapeTagOpeners(json(catalog))),
        tagged("output_contract", 'Reply only {"sections":[0],"reason":"the actual erroneous or missing explanation and why these sections own its correction"}. Select every affected existing section index, at least one and no duplicates. Do not write notes, invent a section, return a coverage ledger or choose an answer for the lawyer.'),
      ].join("\n\n")
      let issues: string[] = []
      for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
        appendLeanEvent(directory, { stage: "finalize", step, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
        const asked = performance.now()
        const reply = await deps.sessions.prompt({ directory, sessionID, agent, text })
        const answered = performance.now()
        const parsed = parseJsonObject(reply)
        const selected = parsed.ok ? parsed.value.sections : undefined
        issues = !Array.isArray(selected) || !selected.length || selected.some(index => !Number.isInteger(index) || index < 0 || index >= inputs.length) || new Set(selected).size !== selected.length
          ? ["Return sections as a nonempty array of distinct existing section indices; select the owners of the actual erroneous or missing explanation."] : []
        if (!parsed.ok || typeof parsed.value.reason !== "string" || !parsed.value.reason.trim()) issues.push("Give a reason identifying the explanation and why the selected sections own its correction.")
        calls.count++
        calls.attempts = Math.max(calls.attempts, attempt + 1)
        progress({ calls: calls.count, attempts: calls.attempts, feedback: issues })
        appendLeanEvent(directory, { stage: "finalize", step, event: "checker", attempt, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
        if (!issues.length) { linked.set(repair.id, (selected as number[]).map(index => inputs[index]!.id)); return }
        text = `Correct this repair-link reply only. Return the required sections/reason JSON.\n${issues.join("\n")}`
      }
      failures.push(`${step}: ${issues.join(" ")}`)
    } catch (error) { stopped.error ??= { cause: error } }
  })
  throwIfStopped()
  if (stageKey() !== initial) return { status: "restart" }
  if (failures.length) return { status: "failed", reason: `finalize: notes repair routing failed: ${failures.join(" ")}` }
  inputs = focusedNotesInputs(directory, linked)
  await parallel(inputs, async input => {
    if (stopped.error) return
    try {
      const ownerFile = sectionFile(input, true)
      const legacyFile = sectionFile(input)
      const legacy = readLeanText(directory, legacyFile)
      // Old split parents were assembled from passage-writing children without
      // checking the original question. They are evidence of progress, not a
      // checked answer to that question. Preserve them, but use the scoped new
      // content key. Already checked unsplit owners remain reusable unchanged.
      const legacySplit = input.kind !== "passage" && Boolean(legacy && /^### (?:Current target|Deleted initial) passage/m.test(legacy))
      let file = ownerFile
      let saved = readLeanText(directory, file)
      if (!saved && !legacySplit && !notesEvidencePlan(input.text, { ownership: input.kind === "passage" ? "passage" : "one-point" }).changedScope) {
        file = legacyFile
        saved = legacy
      }
      if (saved && !checkFocusedNotesSection(directory, input, saved, true).length) {
        sections.set(input.id, saved)
        files.push(file)
        return
      }
      const located = await locateCurrentInput(input)
      if (!located) return
      const evidencePlan = notesEvidencePlan(located.text, { ownership: located.kind === "passage" ? "passage" : "one-point" })
      const selections: NonNullable<ReturnType<typeof checkNotesInspection>["value"]>[] = []
      for (const inspection of evidencePlan.inspections) {
        const step = `${input.id}:inspect:${inspection.id}`
        const agent = LEAN_STAGE_AGENTS.finalize
        const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent,
          title: stepSessionTitle(directory, "finalize", step), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
        sessions.push(sessionID)
        progress({ session_id: sessions[0], session_ids: [...sessions] })
        let issues: string[] = []
        let selected = false
        for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
          appendLeanEvent(directory, { stage: "finalize", step, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
          const asked = performance.now()
          const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: attempt
            ? `Correct only this evidence-selection reply. Keep the original relevance question and return targets/excerpts/limitation JSON, selecting exact available ranges where a complete PID is unavailable.\n${issues.join("\n")}` : inspection.text })
          const answered = performance.now()
          const checked = checkNotesInspection(reply, inspection)
          issues = checked.issues
          calls.count++
          calls.attempts = Math.max(calls.attempts, attempt + 1)
          progress({ calls: calls.count, attempts: calls.attempts, feedback: issues })
          appendLeanEvent(directory, { stage: "finalize", step, event: "checker", attempt, issues,
            wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
          if (checked.value) { selections.push(checked.value); selected = true; break }
        }
        if (!selected) { failures.push(`${step}: evidence selection supplied no valid answer: ${issues.join(" ")}`); return }
      }
      const writing = evidencePlan.changedScope ? { ...located, text: assembleNotesOwnerEvidence(evidencePlan, selections) } : located
      if (evidencePlan.changedScope || legacySplit) file = ownerFile
      const writeUnit = async (writing: FocusedNotesInput, unitFile: string): Promise<string | undefined> => {
        const input = writing
        const kept = readLeanText(directory, unitFile)
        if (kept && !checkFocusedNotesSection(directory, writing, kept, true).length) return kept.trim()
        const agent = LEAN_STAGE_AGENTS.finalize
        const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, "finalize", input.id), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
        sessions.push(sessionID)
        progress({ session_id: sessions[0], session_ids: [...sessions] })
        let feedback: string[] = []
        const dependencies: string[] = []
        for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt++) {
          appendLeanEvent(directory, { stage: "finalize", step: input.id, event: attempt ? "retry" : "model_call", attempt, session_id: sessionID })
          const asked = performance.now()
          const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: attempt ? `Correct only this section's errors, retaining its original assignment and supported explanation. Return this section only.\n${feedback.map(issue => `- ${issue}`).join("\n")}` : writing.text })
          const answered = performance.now()
          feedback = checkFocusedNotesSection(directory, writing, reply, attempt >= LEAN_MAX_RETRIES)
          calls.count++
          calls.attempts = Math.max(calls.attempts, attempt + 1)
          progress({ attempts: calls.attempts, calls: calls.count, feedback })
          appendLeanEvent(directory, { stage: "finalize", step: input.id, event: "checker", attempt, issues: feedback, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
          // Several already-ready replies can otherwise chain synchronous
          // checks through microtasks and starve native HTTP/timer callbacks.
          await new Promise<void>(resolve => setImmediate(resolve))
          if (!feedback.length) {
            const documents = readLeanPlan(directory)?.documents ?? []
            let checking = { ...writing, text: [writing.text, tagged("planned_dependency_documents", escapeTagOpeners(json(documents.map(({ id, title }) => ({ id, title }))))), ...dependencies].join("\n\n") }
            // Evidence selection is separate from rewriting: inspect one actual
            // claim, then check the same prose with trusted current passages.
            // A request or an unanswered inspection is never an accepting verdict.
            for (let round = 0; ; round++) {
              let observedCalls = 0
              const observe = (patch: Partial<LeanStageState>) => {
                for (const id of patch.session_ids ?? []) if (!sessions.includes(id)) sessions.push(id)
                if (patch.calls !== undefined) { calls.count += patch.calls - observedCalls; observedCalls = patch.calls }
                calls.attempts = Math.max(calls.attempts, patch.attempts ?? 0)
                progress({ ...patch, calls: calls.count, attempts: calls.attempts, session_id: sessions[0], session_ids: [...sessions] })
              }
              const coverage = await checkFocusedNotesCoverage(directory, state, deps, checking, reply, observe)
              if (coverage.status === "failed") { failures.push(coverage.reason); return }
              if (coverage.status === "pass") break
              if (coverage.status === "rewrite") {
                feedback = [`This prose does not substantively explain its assigned point: ${coverage.reason} Return the complete substantive section; earlier or interrupted text cannot supply missing content.`, ...dependencies]
                progress({ feedback })
                break
              }
              if (round >= LEAN_MAX_RETRIES) {
                failures.push(`${input.id}: explanation check still needs current evidence after ${LEAN_MAX_RETRIES} focused inspections: ${coverage.reason}`)
                return
              }
              observedCalls = 0
              const dependency = await resolveFocusedNotesDependency(directory, state, deps, writing, reply, coverage.request,
                documents.map(document => notesBody(directory, document.id)), observe)
              if (dependency.status === "failed") { failures.push(`${input.id}: ${dependency.reason}`); return }
              const evidence = dependency.evidence.map(row => tagged(row.name, escapeTagOpeners(row.text))).join("\n\n")
              dependencies.push(evidence)
              checking = { ...checking, text: `${checking.text}\n\n${evidence}` }
            }
          }
          if (!feedback.length) {
            // Preserve only an answer to unchanged inputs. The walk re-keys if
            // instructions or evidence changed while any section was in flight.
            if (stageKey() === initial) {
              writeLeanText(directory, unitFile, `${reply.trim()}\n`)
              if (!files.includes(unitFile)) files.push(unitFile)
            }
            return reply.trim()
          }
        }
        failures.push(`${input.id}: ${feedback.join(" ")}`)
      }
      // Inspection children select evidence only. The original point remains
      // one prose owner and must pass its own substantive coverage check before
      // it can be cached or receive its once-only reference.
      const prose = await writeUnit(writing, file)
      if (prose === undefined) return
      const issues = checkFocusedNotesSection(directory, input, prose, true)
      if (issues.length) { failures.push(`${input.id}: ${issues.join(" ")}`); return }
      if (stageKey() === initial) {
        writeLeanText(directory, file, `${prose}\n`)
        sections.set(input.id, prose)
        if (!files.includes(file)) files.push(file)
      }
    } catch (error) { stopped.error ??= { cause: error } }
  })
  throwIfStopped()
  if (stageKey() !== initial) return { status: "restart" }
  if (failures.length || sections.size !== inputs.length) return { status: "failed", reason: `finalize: focused notes incomplete after ${LEAN_MAX_RETRIES} retries: ${failures.join(" ")}` }
  const notes = assembleFocusedNotes(directory, inputs.map(input => ({ input, prose: sections.get(input.id)! })))
  const accepted = await acceptStageReply(directory, "finalize", notes, { final: true })
  if (accepted.issues.length) return { status: "failed", reason: `finalize: assembled notes failed their check: ${accepted.issues.join(" ")}` }
  return { status: "done", outputs: [...accepted.outputs, ...files], openItems: accepted.openItems, questions: accepted.questions }
}
