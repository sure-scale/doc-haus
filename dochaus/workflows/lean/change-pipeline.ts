import { draftingStyleWarningsBlock } from "./style-warnings"
import { currentDocumentArtifacts } from "../../lib/current-document-artifacts"
import { checklistProducerInputs } from "./checklist-source-context"
import { existsSync, readFileSync } from "node:fs"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { amendmentParagraphs } from "../../lib/drafting-amendment"
import { askedQuestions, checkChanges, parseJsonObject, questionIssues } from "./change-checks"
import { finalBodyParagraphs, readChangeList, readParagraphs, record, type ChangeList } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { reviewStages, earlierDocuments, readLeanPlan, type DocID } from "./plan"
import { stagePrompt, retryPrompt } from "./prompts"
import { matterSourceIDs, openItem, quoteFound, quoteMissIssue } from "./quotes"
import { changesRepairSections, documentRepairs } from "./repairs"
import { drainLeanRequests } from "./requests"
import { approvedResearch } from "./research"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, stageInputs, stepSessionTitle, type LeanDeps } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanState, readLeanText, readQuestions, writeLeanText, type LeanState, type LeanStageState, type Question } from "./state"

export type ChangeGroup = {
  id: string
  pids: string[]
  goal: string
  sources: { source_id: string; quote: string }[]
  limitations: string[]
}
export type ReviewDisposition = {
  disposition: "deferred" | "recommendation" | "rejected"
  kind: "information" | "research" | "suggestion"
  detail: string
  reason: string
  documents: DocID[]
  sources: { source_id: string; quote: string }[]
  limitations: string[]
}
export type ChangePlan = { shared_decisions: string[]; groups: ChangeGroup[]; dismissed: string[]; questions: Question[]; findings?: ReviewDisposition[] }
export type ChangePreparation = { groups: { group: string; changes: ChangeList }[] }
export type ChangeReview = { status: "pass" | "changes-required"; findings: { target?: "body" | "research"; document?: DocID; pid?: string; detail: string }[]; limitations: string[] }
const PREPARATION_WIDTH = 4

/** Adjudication supplies goals and dependencies, never replacement wording. */
export function changePlanPrompt(directory: string, document: DocID) {
  return [
    stagePrompt(directory, `${document}.change-plan`, { contextOnly: true }),
    `<draft_paragraphs>\n${readParagraphs(directory, document).map(row => `[${row.pid}] ${row.text}`).join("\n\n")}\n</draft_paragraphs>`,
    draftingStyleWarningsBlock(readParagraphs(directory, document)),
    `<findings>\n${reviewStages(directory, document).map(stage => `${stage}:\n${readLeanText(directory, `${document}/${stage}.md`) ?? "Missing review"}`).join("\n\n")}\n</findings>`,
    ...changesRepairSections(directory, document),
    ...(readChangePlan(directory, document) ? [`<previous_review_dispositions>\n${reviewDispositionText(directory, document)}\n</previous_review_dispositions>`, "Retain still-material prior dispositions in your replacement plan. Resolve or reject them only when current evidence or settled instructions supply the reason; do not lose deferred advice during a wording repair."] : []),
    ...(documentRepairs(directory, document)?.items.some(item => item.stage === "package-review" || item.stage === "change-review" || item.stage === "checklist-check") ? [
      `<current_final_body>\n${repairBodyText(directory, document, documentRepairs(directory, document)!.items.some(item => item.stage === "change-review"), documentRepairs(directory, document)!.items.some(item => item.stage === "checklist-check"))}\n</current_final_body>`,
      "Semantic findings refer to the current final body. Map them to initial draft pids; the eventual changes list is cumulative against that initial draft.",
    ] : []),
    "Adjudicate every independent review finding and required repair against actual sources, later instructions and answers. Reviewer agreement and exact quotation matching alone do not establish the suggested legal/commercial effect. Consolidate duplicates, resolve conflicts, reject unsupported or out-of-scope suggestions with a reason, and retain every supported required correction and still-supported previous amendment. Do not write replacement paragraphs yet.",
    "Before any amendment drafting, ask only indispensable unanswered decisions needed for an essential requested function. Ask each decision once, plainly, with why it matters and one affected initial draft pid. Do not repeat answered decisions, infer an answer, ask the human to supply research, or turn optional refinements into blockers. Routine missing details may remain disclosed placeholders.",
    "Keep source-supported negotiating recommendations labelled as proposals rather than agreed terms; no prior approval does not itself prevent proposing them. Human questions concern only indispensable missing information or decisions, never internal paragraph routing or artifact mechanics. Initial-draft revision explanations are authored separately by the final reconciler; do not demand no-op amendments to generate comments.",
    "Group paragraphs whose changes depend on each other into one group: definitions, conditions, exceptions, cross-references and combined effects must be prepared together. A pid belongs to at most one group. Each independent group receives its specific operative goal and exact supporting source passages; disclose unsupported drafting judgment in limitations so resulting wording carries the required new confirmation marker. Record common settled decisions for all groups. Empty groups are valid when no operative wording correction is needed; explanation-only corrections must still reach the final reconciler through the repair findings, without a no-op amendment; record non-amendment dispositions in findings. Previous supported amendments are not silently dropped during repair.",
    "Retain each material finding that does not become an amendment in findings, consolidating duplicate observations. Use deferred for unresolved information or research that must remain visible with the needed action, recommendation for an optional supported proposal, and rejected only when evidence or scope supplies a reason it need not be pursued. Do not put deferred work or advice recommendations in dismissed. Name affected planned documents; this is routing context, not a restriction on later advice about the originating document. Keep source passages separate from uncertainty in limitations; an unresolved research proposition is not verified authority. Required wording still belongs in groups; deferral alone cannot fulfill it. Indispensable decisions still belong in questions. A usable review draft can retain a disclosed gap without an additional human question.",
    'Reply with exactly one JSON object: {"shared_decisions":["settled decision and relevant dependency"],"groups":[{"id":"group label","pids":["initial draft pid"],"goal":"accepted findings, intended operative effects and what must be preserved","sources":[{"source_id":"exact matter name, request or approved authority URL","quote":"exact supporting words"}],"limitations":["source limits or judgment requiring confirmation"]}],"dismissed":[],"findings":[{"disposition":"deferred|recommendation|rejected","kind":"information|research|suggestion","detail":"material issue or proposal and action remaining","reason":"why no amendment is made and the practical consequence","documents":["affected planned document id"],"sources":[{"source_id":"canonical matter source, request or approved authority URL","quote":"exact supporting words"}],"limitations":["what remains uncertain or unsupported"]}],"questions":[{"pid":"initial draft pid","question":"one essential unresolved decision","context":"why needed","id":"existing question id only when repeating"}]}. All five arrays are required; they may be empty. No file writes.',
  ].join("\n\n")
}

export function checkChangePlan(directory: string, document: DocID, value: unknown) {
  if (!record(value) || !strings(value.shared_decisions) || !Array.isArray(value.groups) || !strings(value.dismissed) || !Array.isArray(value.questions) || (value.findings !== undefined && !Array.isArray(value.findings)))
    return ["Return shared_decisions and dismissed string arrays, groups and questions arrays, and a findings array of review dispositions."]
  const paragraphs = new Map(readParagraphs(directory, document).map(row => [row.pid, row]))
  const sources = new Set([...matterSourceIDs(directory), "request", ...approvedResearch(directory).map(row => row.url)])
  const documents = new Set(readLeanPlan(directory)?.documents.map(row => row.id) ?? [document])
  const findings = new Set<string>()
  const seen = new Set<string>()
  const pids = new Set<string>()
  return [
    ...value.groups.flatMap((group, index) => {
      const at = `groups[${index}]`
      if (!record(group) || !nonempty(group.id) || !strings(group.pids) || !group.pids.length || !nonempty(group.goal) || !Array.isArray(group.sources) || !strings(group.limitations))
        return [`${at} needs id, nonempty pids, goal, sources and limitations.`]
      const issues = seen.has(group.id) ? [`${at}: duplicate group id.`] : []
      seen.add(group.id)
      for (const pid of group.pids) {
        if (!paragraphs.has(pid)) issues.push(`${at}: [${pid}] is not an initial draft paragraph.`)
        if (pids.has(pid)) issues.push(`${at}: [${pid}] appears in more than one group; dependent changes must be prepared together.`)
        pids.add(pid)
      }
      group.sources.forEach(source => {
        if (!record(source) || !nonempty(source.source_id) || !nonempty(source.quote) || !sources.has(source.source_id)) {
          issues.push(`${at}.sources must name canonical matter sources, request or approved authority and exact nonempty quotes.`)
          return
        }
        if (!quoteFound(directory, { source_id: source.source_id, quote: source.quote }))
          issues.push(quoteMissIssue(directory, `${at}.sources`, { source_id: source.source_id, quote: source.quote }, "json"))
      })
      return issues
    }),
    ...((value.findings ?? []) as unknown[]).flatMap((finding, index) => {
      const at = `findings[${index}]`
      if (!record(finding) || !["deferred", "recommendation", "rejected"].includes(String(finding.disposition)) ||
        !["information", "research", "suggestion"].includes(String(finding.kind)) || !nonempty(finding.detail) || !nonempty(finding.reason) ||
        !strings(finding.documents) || !finding.documents.length || !Array.isArray(finding.sources) || !strings(finding.limitations))
        return [`${at} needs disposition, kind, detail, reason, affected documents, sources and limitations.`]
      const key = finding.detail.trim().toLowerCase().replace(/\s+/g, " ")
      const issues = findings.has(key) ? [`${at}: consolidate duplicate findings into one disposition with all affected documents.`] : []
      findings.add(key)
      if (finding.documents.some(id => !documents.has(id as DocID))) issues.push(`${at}: affected documents must be planned document ids.`)
      finding.sources.forEach(source => {
        if (!record(source) || !nonempty(source.source_id) || !nonempty(source.quote) || !sources.has(source.source_id)) {
          issues.push(`${at}.sources must name canonical matter sources, request or approved authority and exact nonempty quotes.`)
          return
        }
        if (!quoteFound(directory, { source_id: source.source_id, quote: source.quote }))
          issues.push(quoteMissIssue(directory, `${at}.sources`, { source_id: source.source_id, quote: source.quote }, "json"))
      })
      if (!finding.sources.length && !finding.limitations.length) issues.push(`${at}: disclose the evidence limitation when no supporting source passage is provided.`)
      return issues
    }),
    ...value.questions.flatMap((question, index) => [
      ...questionIssues(question, index, paragraphs, new Map(readQuestions(directory).map(row => [row.id, row]))),
      ...(!record(question) || !nonempty(question.pid) ? [`questions[${index}] must identify an affected initial draft pid.`] : []),
    ]),
  ]
}

export function acceptChangePlan(directory: string, document: DocID, reply: string) {
  const value = json(reply)
  const issues = checkChangePlan(directory, document, value)
  if (issues.length || !record(value)) return { issues, outputs: [], openItems: [] }
  const asked = new Map(readQuestions(directory).map(row => [row.id, row]))
  const questions = askedQuestions(value.questions as Record<string, unknown>[], asked, document)
    .filter((row): row is Question => Boolean(row)).map(row => ({ ...row, stage: "change-plan" as const }))
  const plan = { ...value, questions }
  writeLeanText(directory, `${document}/change-plan.json`, `${JSON.stringify(plan, null, 2)}\n`)
  return { issues: [], outputs: [`${document}/change-plan.json`], questions,
    openItems: questions.map(row => openItem("change-plan", "question", row.question, { question_id: row.id, pid: row.pid })) }
}

export function readChangePlan(directory: string, document: DocID): ChangePlan | undefined {
  const value = readLeanJson(directory, `${document}/change-plan.json`)
  return checkChangePlan(directory, document, value).length ? undefined : value as ChangePlan
}

/** A view of the existing adjudication record, not a second findings ledger.
 * Old records retain their unstructured decisions and review navigation until
 * the planner re-runs; never infer rejection merely from an old dismissal. */
export function reviewDispositionText(directory: string, document: DocID, priorDocument = false) {
  const plan = readChangePlan(directory, document)
  if (!plan) return "No adjudicated review dispositions available."
  const navigation = (text: string) => priorDocument ? text.replace(/\bp[0-9a-f]{12}\b/g, "earlier paragraph") : text
  return JSON.stringify({
    origin_document: document,
    // Strip authored paragraph references for other-document consumers, while
    // retaining verbatim source passages and the saved decision record.
    findings: (plan.findings ?? []).map(row => ({ ...row, detail: navigation(row.detail), reason: navigation(row.reason), limitations: row.limitations.map(navigation) })),
    ...(plan.dismissed.length ? { legacy_decisions: plan.dismissed.map(navigation) } : {}),
    ...(plan.findings === undefined ? { unadjudicated_review_history: reviewStages(directory, document).map(stage => ({ stage,
      text: (readLeanText(directory, `${document}/${stage}.md`) ?? "").replace(/\[p[0-9a-f]{12}\]/g, "(earlier paragraph)"),
    })) } : {}),
  }, null, 2)
}

/** Deferred gaps project into the existing disclosure gate directly from the
 * plan. Optional recommendations and rejected ideas are not open decisions. */
export function deferredFindingItems(directory: string) {
  return (readLeanPlan(directory)?.documents ?? []).flatMap(document =>
    (readChangePlan(directory, document.id)?.findings ?? []).filter(row => row.disposition === "deferred").map(row =>
      openItem("change-plan", row.kind === "research" ? "research_gap" : "omission", `${row.detail} ${row.reason}`, { document: document.id })))
}

export function changePreparationPrompt(directory: string, document: DocID, plan: ChangePlan, group: ChangeGroup) {
  const pids = new Set(group.pids)
  const prior = readChangeList(directory, document)
  return [
    ...(documentRepairs(directory, document)?.items.every(item => item.stage === "checklist-check") ? [
      "Prepare ONE accepted repair group using only its goals, source passages, listed initial paragraphs and related prior changes below. Original instructions and actual human answers determine scope; generated plans/research limits are not client waivers. Flagging, proposing and adopting are different actions. Preserve genuine uncertainty and do not invent approval or authority.",
      `<original_instructions_and_actual_answers>\n${JSON.stringify(checklistProducerInputs(directory).filter(row => row.key === "request").flatMap(row => [...row.input.passages, ...row.input.instructions]))}\n</original_instructions_and_actual_answers>`,
      `Canonical uploaded originals: search-document/read-document view=source with exact source identity. Current generated documents: view=current_body with planned document ID; comments/deleted_text are separate views. Unavailable means no search occurred; lexical no hits is not absence. Read consecutive paragraphs around necessary definitions/exceptions, and related documents only when needed. Initial draft mapping: read-record .dochaus/lean/${document}/paragraphs.json. Current context records: .dochaus/lean/checklist-context/${document}.json. Prior amendments: .dochaus/lean/${document}/changes.json. Inspect only necessary records.`,
    ] : [stagePrompt(directory, `${document}.change-prepare`, { contextOnly: true })]),
    `<adjudicated_dependencies_not_independent_authority>\n${plan.shared_decisions.join("\n")}\n</adjudicated_dependencies_not_independent_authority>`,
    "These generated dependencies summarize prior source adjudication. They do not themselves establish a human decision, override original source instructions or waive unresolved work. Actual source instructions and human answers govern; disclose a conflict rather than silently adopting a generated scope limit.",
    `<group_goal_and_evidence>\n${JSON.stringify(group, null, 2)}\n</group_goal_and_evidence>`,
    `<draft_paragraphs>\n${readParagraphs(directory, document).filter(row => pids.has(row.pid)).map(row => `[${row.pid}] ${row.text}`).join("\n\n")}\n</draft_paragraphs>`,
    `<previous_group_changes>\n${JSON.stringify([...(prior?.amendments ?? []), ...(prior?.withheld ?? []).map(row => row.amendment)].filter(row => pids.has(row.pid)), null, 2)}\n</previous_group_changes>`,
    "Amend the listed initial draft paragraphs, not paragraphs reconstructed from the supplied original: the original is evidence and the comparison baseline, while the listed draft is the amendment target. Use the listed group pids exactly. A scope mismatch is a preparation limitation for the reconciler, not a question for the human about internal identifiers. Ask only a genuinely indispensable human decision; never request paragraph bookkeeping from the user.",
    "Prepare this adjudicated group only. Retain unaffected wording and every supported goal; reconcile dependent paragraphs together. Supporting excerpts and the spine are retrieval guides, not authority: inspect underlying canonical passages and qualifications on demand. Other groups cannot be amended here. If the supplied scope misses a material dependency or decision, do not guess: explain the problem and ask an indispensable question when only the lawyer can decide. A final reconciler independently considers the whole body.",
    'Reply exactly {"amendments":[{"pid":"a listed group pid","updated":"full replacement paragraph without line breaks","reason":"why and intended effect","sources":[{"source_id":"canonical source or request","quote":"exact supporting words"}]}],"questions":[]}. One explicit operation per initial pid: omit operation for one replacement string; use operation delete and updated "" to remove the paragraph only with source or settled-instruction support; use operation split and updated as an array of at least two complete nonempty paragraphs to split it. Each resulting paragraph stays one paragraph without line breaks. Never encode a deletion as a blank replacement or collapse a split to its first paragraph. Unsupported deletion retains the original wording with a new confirmation marker instead. Unsupported substantive judgment requires a new visible [open item: confirm ...] marker; sources may then be empty. A newly discovered essential unanswered question may instead be included with pid/question/context and existing id when repeating, with a visible open-item amendment to its paragraph. Never assume an answer. No files or redlines.',
  ].join("\n\n")
}

/** Fresh read-only workers; only this stage writes the aggregate after every
 * worker settles. A failed group never becomes a partial accepted preparation. */
export async function runChangePreparation(directory: string, state: LeanState, deps: LeanDeps, document: DocID, progress: (patch: Partial<LeanStageState>) => void) {
  const plan = readChangePlan(directory, document)
  if (!plan) return { status: "failed" as const, reason: `${document}.change-prepare: no valid change plan.` }
  if (plan.questions.some(row => !readQuestions(directory).some(question => question.id === row.id && question.status === "answered"))) return { status: "failed" as const, reason: `${document}.change-prepare: adjudication still has unanswered decisions.` }
  const inputSha = hashStageInputs(`${document}.change-prepare`, LEAN_PROMPT_VERSION, stageInputs(directory, `${document}.change-prepare`))
  const counts = { calls: 0, attempts: 0, next: 0, sessionID: null as string | null, sessionIDs: [] as string[] }
  const results: { group: string; changes: ChangeList }[] = []
  const failures: string[] = []
  const questions: Question[] = []
  const openItems = [] as ReturnType<typeof openItem>[]
  await Promise.all(Array.from({ length: Math.min(PREPARATION_WIDTH, plan.groups.length) }, async () => {
    let failuresForGroup: string[] = []
    for (;;) {
      const group = plan.groups[counts.next++]
      if (!group) return
      try {
        const agent = LEAN_STAGE_AGENTS["change-prepare"]
        const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, `${document}.change-prepare`, `change-prepare: ${group.id}`), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
        counts.sessionIDs.push(sessionID)
        progress({ session_ids: [...counts.sessionIDs] })
        if (counts.sessionID === null) {
          counts.sessionID = sessionID
          progress({ session_id: sessionID })
        }
        for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
          counts.calls += 1
          counts.attempts = Math.max(counts.attempts, index + 1)
          progress({ calls: counts.calls, attempts: counts.attempts })
          appendLeanEvent(directory, { stage: "change-prepare", document, step: group.id, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
          const asked = performance.now()
          const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("change-prepare", failuresForGroup) : changePreparationPrompt(directory, document, plan, group) })
          const answered = performance.now()
          const check = checkChanges(directory, reply, readParagraphs(directory, document).filter(row => group.pids.includes(row.pid)), { document, withholdUnverified: index >= LEAN_MAX_RETRIES })
          failuresForGroup = [...check.issues, ...(check.changes ? uniqueReplacementIssues(check.changes) : [])]
          appendLeanEvent(directory, { stage: "change-prepare", document, step: group.id, event: "checker", attempt: index, issues: failuresForGroup, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
          progress({ feedback: failuresForGroup })
          if (!failuresForGroup.length && check.changes) {
            results.push({ group: group.id, changes: check.changes })
            questions.push(...check.changes.questions.map(row => ({ ...row, stage: "change-prepare" as const })))
            openItems.push(...check.openItems.filter(row => row.kind === "question").map(row => ({ ...row, stage: "change-prepare" as const })))
            break
          }
          if (index === LEAN_MAX_RETRIES) failures.push(`${group.id}: ${failuresForGroup.join(" ")}`)
        }
      } catch (error) { failures.push(`${group.id}: ${error instanceof Error ? error.message : String(error)}`) }
      // Worker-local retry feedback is reset before taking another group.
      failuresForGroup = []
    }
  }))
  drainLeanRequests(directory)
  if (inputSha !== hashStageInputs(`${document}.change-prepare`, LEAN_PROMPT_VERSION, stageInputs(directory, `${document}.change-prepare`)))
    return { status: "restart" as const }
  if (failures.length) return { status: "failed" as const, reason: `${document}.change-prepare: ${failures.join(" ")}` }
  const preparation: ChangePreparation = { groups: plan.groups.map(group => results.find(row => row.group === group.id)!) }
  writeLeanText(directory, `${document}/change-preparation.json`, `${JSON.stringify(preparation, null, 2)}\n`)
  const uniqueQuestions = [...new Map(questions.map(row => [row.id, row])).values()]
  return { status: uniqueQuestions.length ? "blocked" as const : "done" as const, outputs: [`${document}/change-preparation.json`], openItems, questions: uniqueQuestions,
    ...(uniqueQuestions.length ? { reason: "Waiting for your answers before reconciling amendments." } : {}) }
}

export function uniqueReplacementIssues(changes: ChangeList) {
  const repeated = [...Map.groupBy(changes.amendments, row => row.pid)].filter(([, rows]) => rows.length > 1)
  return repeated.map(([pid]) => `Consolidate all accepted changes for [${pid}] into one full replacement; the projected body must preserve every accepted effect.`)
}

export function projectedChangeParagraphs(directory: string, document: DocID) {
  const changes = readChangeList(directory, document)
  if (!changes || uniqueReplacementIssues(changes).length) throw new Error("A reconciled change list with one replacement per paragraph is required before semantic review.")
  const amendments = new Map(changes.amendments.map(row => [row.pid, amendmentParagraphs(row)]))
  const paragraphs = readParagraphs(directory, document)
  if (changes.amendments.some(row => !paragraphs.some(paragraph => paragraph.pid === row.pid))) throw new Error("Changes address an unknown initial paragraph.")
  return paragraphs.flatMap(row => (amendments.get(row.pid) ?? [row.text]).map(text => ({ ...row, text })))
}

/** A dependent document can be redrafted after an earlier-body repair. Keep
 * its saved findings as challenges, without projecting the old change list
 * onto different initial paragraphs. Strict semantic projection stays above. */
export function repairBodyText(directory: string, document: DocID, projected: boolean, actualWord = false) {
  if (actualWord) {
    const file = currentDocumentArtifacts(directory, document).body
    if (!file) throw new Error(`Checklist repair requires a readable current Word body for ${document}.`)
    const rows = readDocxCommentParagraphs(readFileSync(leanFile(directory, file)))
    const repairs = documentRepairs(directory, document)?.items ?? []
    // Other semantic repair stages retain their existing full-body context.
    if (repairs.some(item => item.stage !== "checklist-check"))
      return repairBodyText(directory, document, projected)
    const selected = new Set<number>()
    for (const repair of repairs) {
      let inspected: { pids: string[]; paragraph_text: string }[] = []
      try { inspected = JSON.parse(repair.detail.split("\n")[1]!).current_target_paragraphs ?? [] } catch { /* unlocated: targeted lookup required */ }
      rows.forEach((row, index) => {
        if (!inspected.some(evidence => evidence.paragraph_text === row.paragraph_text && evidence.pids?.some(pid => row.pids.includes(pid)))) return
        for (const neighbor of [index - 1, index, index + 1]) if (rows[neighbor]) selected.add(neighbor)
      })
    }
    return [
      `Focused actual current Word context from .dochaus/lean/${file}. This is not the whole body. For additional relevant text use read-document with document=${document}, view=current_body and bounded offset/count; search-document supports the same document/view. Current PIDs identify evidence only; amendments still target initial draft PIDs.`,
      selected.size ? [...selected].sort((a, b) => a - b).map(index => `[${rows[index]!.pids.join(", ") || "native PID unavailable"}] ${rows[index]!.paragraph_text}`).join("\n\n")
        : "No affected current paragraph was located uniquely. Retrieve the relevant current text before deciding; do not guess a paragraph or treat missing inspection as a body defect.",
    ].join("\n\n")
  }
  const paragraphs = readParagraphs(directory, document)
  const current = new Map(paragraphs.map(row => [row.pid, row.sha256]))
  const changes = readChangeList(directory, document)
  const matches = changes && !uniqueReplacementIssues(changes).length && changes.amendments.every(row => current.get(row.pid) === row.existing_sha256) &&
    (projected || readLeanState(directory)?.stages[`${document}.apply`]?.input_sha === hashStageInputs(`${document}.apply`, LEAN_PROMPT_VERSION, stageInputs(directory, `${document}.apply`)))
  const rows = matches ? projected ? projectedChangeParagraphs(directory, document) : finalBodyParagraphs(directory, document) : paragraphs
  return [!matches ? "The saved change or applied body is from an earlier draft or is not a current accepted result. The current initial draft follows. Reassess saved findings against this draft and the sources; prior paragraph pointers are not current amendment targets. Do not transfer an old amendment by guessing its location." : "",
    rows.map(row => `[${row.pid}] ${row.text}`).join("\n\n")].filter(Boolean).join("\n\n")
}

export function changeReviewPrompt(directory: string, document: DocID) {
  return [
    stagePrompt(directory, `${document}.change-review`, { contextOnly: true }),
    '<artifact_phase>before-apply</artifact_phase>',
    "This review precedes application of the current projected revision. Review the proposed explanations in reconciled_changes reason/source records, their evidence and consistency with the projected wording. Native redlines and anchored Word comments for this revision are generated by the apply step and checked against actual exported records in package review afterwards. Earlier exports may exist on disk but do not verify this revision. Absence of the current export is not a pre-apply defect or proof of missing final delivery. Do not claim it already exists or ask the read-only correction agents to create or verify an artifact before apply. The explanations array records source-to-initial-draft revisions separately from later amendments and is exported at apply without changing body text. Verify material coverage against the supplied original, and that each reason describes the actual change and its final effect. Missing or inaccurate proposed explanations remain material findings.",
    `<initial_draft>\n${readParagraphs(directory, document).map(row => `[${row.pid}] ${row.text}`).join("\n\n")}\n</initial_draft>`,
    `<projected_full_body>\n${projectedChangeParagraphs(directory, document).map(row => `[${row.pid}] ${row.text}`).join("\n\n")}\n</projected_full_body>`,
    ...earlierDocuments(directory, document).map(row => `<earlier_body document="${row.id}">\n${finalBodyParagraphs(directory, row.id).map(paragraph => `[${paragraph.pid}] ${paragraph.text}`).join("\n\n")}\n</earlier_body>`),
    `<reconciled_changes>\n${readLeanText(directory, `${document}/changes.json`)}\n</reconciled_changes>`,
    `<adjudicated_goals>\n${readLeanText(directory, `${document}/change-plan.json`)}\n</adjudicated_goals>`,
    `<questions>\n${JSON.stringify(readQuestions(directory), null, 2)}\n</questions>`,
    `<preparation_limits>\n${JSON.stringify((readLeanJson(directory, `${document}/change-preparation.json`) as ChangePreparation | undefined)?.groups?.map(row => ({ group: row.group, withheld: row.changes.withheld, questions: row.changes.questions })) ?? [], null, 2)}\n</preparation_limits>`,
    "Preparation limits describe unverified proposals, not final application outcomes. Assess whether their accepted goals are correctly fulfilled by the final projected wording; do not repeat a proposal limit as an outstanding defect after it is corrected.",
    ...changesRepairSections(directory, document).filter(section => !section.startsWith("<previous_changes>")),
    "Independently review the actual projected full resulting text BEFORE application. An amendment being suggested, accepted by another agent or quote-matched is not evidence of its correct effect. Inspect canonical supporting context on demand. Check that the operative wording and its explanation perform the accepted goals, preserve unaffected rights/obligations and settled decisions, and remain consistent across definitions, conditions, exceptions, cross-references and earlier finished documents. Identify lost supported corrections, unsupported additions or changed effects, contradictory paragraphs and inaccurate reasons/source explanations. Check explicit questions and withheld amendments: a material unresolved or silently dropped correction is not a pass merely because it is recorded as withheld or disclosed in limitations. Appropriately marked/disclosed review-draft placeholders can remain; optional expansion is not mandatory. Do not reopen answered decisions without contrary evidence or supply law from memory.",
    'Reply exactly {"status":"pass"|"changes-required","findings":[{"document":"owning document id, omit for the current body","pid":"affected paragraph id of that document when identifiable, otherwise omit","detail":"specific material defect, actual wording/evidence and required correction","target":"body|research, omit for body"}],"limitations":["remaining disclosed limits"]}. Pass requires no findings. Changes-required requires at least one material finding. A correction to an earlier_body must name that earlier document and its own paragraph rather than a paragraph in the current body; the coordinator returns it through that document’s existing cumulative amendment flow and rechecks dependent work. Only the current body or an earlier run-owned body may be targeted; never a future document or external file. Use target research only for a genuinely new material authority proposition beyond approved research that needs actual retrieval; explain why present research cannot answer it. Do not invent law or route ordinary wording defects to research. Do not rewrite paragraphs or write files; corrections return through the same bounded amendment pipeline.',
  ].join("\n\n")
}

export function checkChangeReview(directory: string, document: DocID, value: unknown) {
  if (!record(value) || !["pass", "changes-required"].includes(String(value.status)) || !Array.isArray(value.findings) || !strings(value.limitations)) return ["Return status pass/changes-required, findings array and limitations string array."]
  const documents = readLeanPlan(directory)?.documents ?? []
  const allowed = new Set(documents.slice(0, documents.findIndex(row => row.id === document) + 1).map(row => row.id))
  return [
    ...value.findings.filter(record).flatMap(finding => finding.target !== undefined && !["body", "research"].includes(String(finding.target)) ? ["Finding target must be body or research."] : []),
    ...value.findings.flatMap(finding => {
      if (!record(finding) || !nonempty(finding.detail)) return ["Each finding needs a specific material detail."]
      const owner = finding.document ?? document
      if (typeof owner !== "string" || !allowed.has(owner as DocID)) return ["Finding document must be the current or an earlier run-owned body."]
      const pids = new Set(readParagraphs(directory, owner as DocID).map(row => row.pid))
      return finding.pid !== undefined && (!nonempty(finding.pid) || !pids.has(finding.pid)) ? ["Finding pid must belong to its owning document's initial draft."] : []
    }),
    ...(value.status === "pass" && value.findings.length ? ["A pass cannot contain material correction findings."] : []),
    ...(value.status === "changes-required" && !value.findings.length ? ["Changes-required needs at least one material finding."] : []),
  ]
}

export function acceptChangeReview(directory: string, document: DocID, reply: string) {
  const value = json(reply)
  const issues = checkChangeReview(directory, document, value)
  if (issues.length) return { issues, outputs: [], openItems: [] }
  writeLeanText(directory, `${document}/change-review.json`, `${JSON.stringify(value, null, 2)}\n`)
  return { issues: [], outputs: [`${document}/change-review.json`], openItems: [] }
}

export function readChangeReview(directory: string, document: DocID): ChangeReview | undefined {
  const value = readLeanJson(directory, `${document}/change-review.json`)
  return checkChangeReview(directory, document, value).length ? undefined : value as ChangeReview
}

export function preparationOutputValid(directory: string, document: DocID) {
  const plan = readChangePlan(directory, document)
  const value = readLeanJson(directory, `${document}/change-preparation.json`)
  if (!plan || !record(value) || !Array.isArray(value.groups) || value.groups.length !== plan.groups.length) return false
  const prepared = value.groups
  return plan.groups.every(group => {
    const rows = prepared.filter(row => record(row) && row.group === group.id)
    if (rows.length !== 1 || !record(rows[0]) || !record(rows[0].changes)) return false
    const check = checkChanges(directory, JSON.stringify(rows[0].changes), readParagraphs(directory, document).filter(row => group.pids.includes(row.pid)), { document })
    return Boolean(check.changes) && !check.issues.length && !uniqueReplacementIssues(check.changes!).length
  })
}

function json(reply: string): unknown {
  const parsed = parseJsonObject(reply)
  return parsed.ok ? parsed.value : undefined
}
function nonempty(value: unknown): value is string { return typeof value === "string" && Boolean(value.trim()) }
function strings(value: unknown): value is string[] { return Array.isArray(value) && value.every(nonempty) }
