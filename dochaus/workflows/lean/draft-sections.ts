import { rmSync } from "node:fs"
import { visibleInlineText } from "../../lib/markdown-visible-text"
import { isLetter } from "../../lib/page-furniture"
import { matterClient } from "../../lib/matter-client"
import { playbookSelectionIssues } from "./playbook-review"
import { TOOL_SYNTAX_ISSUE, checkDraft, printedToolCall, readParagraphs } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { matterSourceIDs } from "./quotes"
import { memoNotesFile } from "./deliverables"
import { LEAN_COVERAGE_SCHEMA, coveragePlanHash, coveragePrompt, coverageScopeQuestion, parseCoverage, planCoverageCurrent, planCoverageIssues, type CoverageAct, type LeanCoverage } from "./package-coverage"
import { answeredScopeQuestion, openScopeQuestion } from "./package-scope"
import {
  LEAN_REQUESTED_FILES_SCHEMA, parseRequestedFiles, requestedFilesIssues, requestedFilesPlanned, requestedFilesPrompt, requestedFilesQuestion, requestedFilesRevision, requestedFilesWords, withRequestedNames, writeRequestedFiles,
  type RequestedFiles,
} from "./requested-files"
import {
  FIRST_DOCUMENT, headingText, normalizedHeading, parsePlan, plannedDocument, readLeanPlan, runPosture, withDocumentIDs, writeLeanPlan,
  type ActingFor, type DocID, type LeanOutline, type PlannedDocument, type StepID,
} from "./plan"
import { actingForBlock, packageCompositionBlock, retryPrompt, stagePrompt } from "./prompts"
import {
  LEAN_CHILD_PERMISSIONS,
  LEAN_MAX_RETRIES,
  LEAN_STAGE_AGENTS,
  acceptStageReply,
  draftOutputValid,
  outlineBasis,
  stepOutputs,
  stepSessionTitle,
  type LeanDeps,
} from "./stages"
import { answeredQuestionsText, leanFile, readLeanText, readQuestions, writeLeanText, type LeanStageState, type LeanState, type OpenItem, type Question } from "./state"
import { stageInputSha } from "./runner"
import { outlineCoverageIssues, partCarries, plannedContract } from "./template-contract"
import { draftStatusItems } from "./draft-status"
import { draftTemplateItems } from "./template-coverage"
import { originalTables, originalTablesBlock, partTableIssues, tableParts } from "./original-tables"

// Sectioned draft step. A long document can need more tokens than one model
// reply may carry, so the draft agent writes one part per reply in the same
// child session, so native history carries the parts already written. When the
// run has no plan, the first document's session starts with the outline
// of every document the request asks for, which code saves as the run's plan;
// otherwise each document's session starts from its planned entry and the final
// text of the documents before it. Each reply is checked on its own and retried with the checker's
// issues; code assembles the parts in outline order and the assembled document
// goes through the ordinary draft check, render and paragraph pointers. The
// outline comes from the model; nothing here knows what a document should hold.
// Before the plan is saved, a coverage job in its own session checks it against
// the lawyer's instructions (package-coverage.ts), its issues go back to the
// outline session for at most PLAN_COVERAGE_REVISIONS revised outlines, and the
// result is recorded in coverage.json without failing the step. When an
// instructed act is still left without a document its parties execute and no
// answered package-scope question has shown it, the step saves no plan and
// stops, blocked on one package-scope question for the lawyer about those acts
// only; the answer joins the outline basis, so the plan is drawn up again with
// it. An act an answered question showed is saved as uncovered and disclosed,
// never asked about again. While the gate's
// package-scope question is open, the step asks none of its own: it saves the
// plan, the act is disclosed, and the open question stays the one to answer. A
// coverage reply that never passed its check (unchecked) does not stop the step.
// When the outline plans a memo or more than one document and no package-scope
// question stops the step, a files job in its own session then reads which
// files the lawyer's words ask to receive (requested-files.ts). Where the plan
// does not deliver them so, the issues go back to the outline session for at
// most PLAN_FILES_REVISIONS revised outlines, from which only the deliverables
// and file names of the documents the issues concern are taken, and only while
// the documents and the memo stay as planned; a mismatch still left, or the
// memo asked for with a document planned as a redline, stops the step, with no
// plan saved, on one question to the lawyer, asked once. Once the lawyer has
// answered such a question, no revision is asked for. Code never changes the
// plan's delivery shape, but gives the plan the file names the lawyer's words
// give, and saves the files in requested-files.json beside the plan. A files
// reply that never passed its check is recorded and logged as unchecked, is
// disclosed in the drafting notes and does not stop the step.
// When a document is drafted from a matter document as its template, the
// outline is admitted only once its coverage contract accounts for every
// top-level part of the template (template-contract.ts), and each part is told
// which template parts it carries and which parts inside them the outline
// leaves out. A carried template part the rendered draft leaves out, or
// carries only a small share of (template-coverage.ts), is an open item of the
// step: the draft is never rewritten for it, the change cycle receives it as a
// required repair, and the apply step records any gap left in the final body.

export type DraftOutline = Pick<PlannedDocument, "title" | "parts" | "form" | "mode">

export type DraftStageRun =
  | { status: "done"; outputs: string[]; openItems: OpenItem[]; questions?: Question[] }
  | { status: "blocked"; outputs: string[]; openItems: OpenItem[]; questions: Question[]; reason: string }
  | { status: "failed"; reason: string }

type Checked<T> = { issues: string[]; value?: T }
type Coverage = { acts?: CoverageAct[]; issues: string[] }

/** Revised outlines the coverage job's issues may ask for before the plan is saved. */
export const PLAN_COVERAGE_REVISIONS = 2
/** Revised outlines the files job's issues may ask for before the plan is saved. */
export const PLAN_FILES_REVISIONS = 2

export async function runDraftStage(
  directory: string,
  state: LeanState,
  deps: LeanDeps,
  progress: (patch: Partial<LeanStageState>) => void,
  document: DocID,
  planned: () => void | Promise<void>,
): Promise<DraftStageRun> {
  const step: StepID = `${document}.draft`
  // A new stage attempt cannot retrieve a partial draft from an older attempt.
  writeLeanText(directory, `${document}/draft-in-progress.md`, "")
  const agent = LEAN_STAGE_AGENTS.draft
  // The coverage and files jobs' sessions join the outline's, so the step reports each of its sessions.
  const sessionIDs: string[] = []
  // Created when first prompted: a refresh that keeps the draft never prompts it.
  let drafting: Promise<string> | undefined
  const drafter = () => drafting ??= deps.sessions.createSession({
    directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, step), model: state.model, permissions: LEAN_CHILD_PERMISSIONS,
  }).then(id => { sessionIDs.unshift(id); progress({ session_id: id, session_ids: [...sessionIDs] }); return id })
  const calls = { count: 0, attempts: 0 }
  // One checked exchange: prompt, check, and up to LEAN_MAX_RETRIES retries that
  // hand the checker's issues back to the same session. The check is told when it
  // reads the last attempt (`final`), after which its issues are not retried.
  // `cause` records on the model_call event the issues of another check that the
  // call answers.
  const exchange = (session: { id: string | (() => Promise<string>); agent: string }) => {
    const ask = async <T>(label: string, text: string, check: (reply: string, final: boolean) => Checked<T>, attempt = 0, cause?: string[]): Promise<Checked<T>> => {
      const id = typeof session.id === "string" ? session.id : await session.id()
      appendLeanEvent(directory, { stage: "draft", document, event: attempt ? "retry" : "model_call", step: label, attempt, session_id: id, ...(cause ? { issues: cause } : {}) })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID: id, agent: session.agent, text })
      const answered = performance.now()
      const checked = check(reply, attempt >= LEAN_MAX_RETRIES)
      const timing = { wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) }
      calls.count += 1
      calls.attempts = Math.max(calls.attempts, attempt + 1)
      progress({ attempts: calls.attempts, calls: calls.count, feedback: checked.issues })
      appendLeanEvent(directory, { stage: "draft", document, event: "checker", step: label, attempt, issues: checked.issues, ...timing })
      if (!checked.issues.length || attempt >= LEAN_MAX_RETRIES) return checked
      return ask(label, retryPrompt("draft", checked.issues), check, attempt + 1)
    }
    return ask
  }
  const ask = exchange({ id: drafter, agent })
  const failed = (label: string, issues: string[]): DraftStageRun =>
    ({ status: "failed", reason: `${step}: ${label}: checker failed after ${LEAN_MAX_RETRIES} retries: ${issues.join(" ")}` })
  // The coverage job reads the plan in a fresh session of its own, so the
  // outline's reasoning does not steer it. Errors from the session calls,
  // including a pause or restart, propagate and leave the plan unwritten.
  const covering = async (documents: PlannedDocument[]): Promise<Coverage> => {
    const coverageAgent = LEAN_STAGE_AGENTS.coverage
    const id = await deps.sessions.createSession({
      directory, parentID: state.owner_session_id, agent: coverageAgent, title: stepSessionTitle(directory, step, "draft: coverage"), model: state.model, permissions: LEAN_CHILD_PERMISSIONS,
    })
    sessionIDs.push(id)
    progress({ ...(sessionIDs.length === 1 ? { session_id: id } : {}), session_ids: [...sessionIDs] })
    const checked = await exchange({ id, agent: coverageAgent })("coverage", coveragePrompt(directory, documents), (reply, final) => parseCoverage(reply, directory, documents, { final }))
    return checked.value ? { acts: checked.value, issues: planCoverageIssues(checked.value, documents, directory) } : { issues: checked.issues }
  }
  const recordCoverage = (documents: PlannedDocument[], actingFor: ActingFor | undefined, coverage: Coverage) => {
    const record: LeanCoverage = {
      schema: LEAN_COVERAGE_SCHEMA,
      plan: coveragePlanHash(documents, actingFor),
      status: !coverage.acts ? "unchecked" : coverage.issues.length ? "uncovered" : "covered",
      acts: coverage.acts ?? [],
      issues: coverage.issues,
    }
    writeLeanText(directory, "coverage.json", `${JSON.stringify(record, null, 2)}\n`)
    appendLeanEvent(directory, { stage: "draft", document, event: "coverage", step: "coverage", outcome: record.status, issues: record.issues })
  }
  // An instructed act still left without a document its parties execute is the
  // lawyer's call on the package's scope: the step stops before drafting with
  // one package-scope question, never answered here, about the acts no answered
  // package-scope question has settled (none is asked when all were, or once
  // two of the step's own are answered; coverageItems discloses the rest), and
  // none while another stage's is still open. An unchecked coverage reply never
  // asks. Asked about a `saved` plan, it offers no new document.
  const scopeBlock = (documents: PlannedDocument[], coverage: Coverage, saved = false): DraftStageRun | undefined => {
    if (!coverage.acts || !coverage.issues.length) return
    const questions = readQuestions(directory)
    if (openScopeQuestion(questions, "draft")) return
    const question = coverageScopeQuestion(coverage.acts, documents, questions.filter(answeredScopeQuestion), { saved })
    if (!question) return
    return { status: "blocked", outputs: ["coverage.json"], openItems: [], questions: [question],
      reason: `${step}: the document plan leaves an instructed act without a document its parties execute; drafting waits for your answer on the package's scope.` }
  }
  // The files job, like the coverage job, reads the plan in a fresh session of
  // its own, and only when the plan has a memo or more than one document to
  // put in files; undefined when it is not asked. A reply that never passed its
  // check gives no files. A name refused on one attempt is carried, by quote,
  // to the later ones (parseRequestedFiles). Errors from the session calls
  // propagate.
  const filing = async (documents: PlannedDocument[], memo: string | undefined): Promise<Checked<RequestedFiles> | undefined> => {
    if (memo === undefined && documents.length < 2) return
    const filesAgent = LEAN_STAGE_AGENTS.files
    const id = await deps.sessions.createSession({
      directory, parentID: state.owner_session_id, agent: filesAgent, title: stepSessionTitle(directory, step, "draft: files"), model: state.model, permissions: LEAN_CHILD_PERMISSIONS,
    })
    sessionIDs.push(id)
    progress({ ...(sessionIDs.length === 1 ? { session_id: id } : {}), session_ids: [...sessionIDs] })
    const refused = new Map<string, string>()
    return exchange({ id, agent: filesAgent })("files", requestedFilesPrompt(directory, documents, memo), (reply, final) => parseRequestedFiles(reply, directory, documents, memo, { final, refused }))
  }
  // Once saved, the plan is followed as it stands, including a lawyer's edits;
  // only its coverage record is refreshed when it no longer matches the plan.
  const outlined = document === FIRST_DOCUMENT && !readLeanPlan(directory)
  if (outlined) {
    const basis = outlineBasis(directory)
    // The lawyer's own words ground requested file names and the memo.
    const request = [readLeanText(directory, "request.md"), readLeanText(directory, "updated-notes.md"), answeredQuestionsText(directory)].filter(Boolean).join("\n\n")
    // A party the lawyer's words name prevails; when they name none, the plan
    // takes the client recorded on the matter.
    const client = matterClient(directory)
    // The Posture the stages so far drafted to; the plan freezes it for the rest of the run.
    const posture = runPosture(directory)
    const best: { value?: LeanOutline } = {}
    const checkOutline = (reply: string): Checked<LeanOutline> => {
      const parsed = parsePlan(reply, request)
      if (!parsed.value) return parsed
      if (!parsed.value.acting_for.stated && client) parsed.value.acting_for = { party: client, stated: true }
      const selectionIssues = playbookSelectionIssues(directory, parsed.value.documents)
      if (selectionIssues.length) { delete best.value; return { issues: selectionIssues } }
      best.value = parsed.value
      // Keep the parser's own issues (the "Acting for:" retry) alongside the
      // template coverage issues.
      return { issues: [...parsed.issues, ...outlineCoverageIssues(directory, parsed.value.documents), ...parsed.value.documents.flatMap(document => document.comparison && (!matterSourceIDs(directory).includes(document.comparison) || !/\.docx$/i.test(document.comparison)) ? [`Compare against names an unknown source: ${document.comparison}. Use the exact supplied matter document name or omit comparison.`] : [])], value: parsed.value }
    }
    const outline = await ask<LeanOutline>("outline", stagePrompt(directory, step, { outline: true }), checkOutline)
    let value = outline.value ?? best.value
    if (!value) return failed("outline", outline.issues)
    let coverage = await covering(withDocumentIDs(value.documents))
    // The coverage issues go back to the outline session as a revision request;
    // a revision that yields no usable outline keeps the last one.
    for (let revision = 1; coverage.acts && coverage.issues.length && revision <= PLAN_COVERAGE_REVISIONS; revision++) {
      delete best.value
      const revised = await ask<LeanOutline>(`outline revision ${revision}/${PLAN_COVERAGE_REVISIONS}`, retryPrompt("draft", coverage.issues), checkOutline, 0, coverage.issues)
      const next = revised.value ?? best.value
      if (!next) continue
      value = next
      coverage = await covering(withDocumentIDs(value.documents))
    }
    const covered = withDocumentIDs(value.documents)
    recordCoverage(covered, value.acting_for, coverage)
    const blocked = scopeBlock(covered, coverage)
    if (blocked) return blocked
    // The files the lawyer's words ask for are read from the covered outline.
    // Where its deliverables do not deliver them so, the issues a revision may
    // settle go back to the outline session as a revision request, unless the
    // lawyer has already been asked about this very mismatch or has answered a
    // question on the requested files: a revision never overrides an answer. A
    // revision settles delivery only: the deliverables of the documents its
    // issues concern are taken when it plans the same titles, in order, and the
    // same memo, spelled alike, and otherwise it is spent and the outline kept.
    // File names stay the outline's: code gives each file its name. Neither job
    // reads what it changes, so neither is asked again.
    const filed = await filing(covered, value.memo)
    const requested = filed?.value
    const contents = (outline: LeanOutline) => JSON.stringify([outline.documents.map(row => normalizedHeading(row.title)), outline.memo ? normalizedHeading(outline.memo) : null])
    const delivered = ({ deliverable: _deliverable, ...row }: LeanOutline["documents"][number], revised: LeanOutline["documents"][number]) =>
      ({ ...row, ...(revised.deliverable ? { deliverable: revised.deliverable } : {}) })
    let files = requested ? requestedFilesIssues(covered, value.memo, requested).issues : []
    const revising = (documents: LeanOutline["documents"]) =>
      requested && requestedFilesQuestion(directory, requested, withDocumentIDs(documents), files) ? requestedFilesRevision(directory, files) : undefined
    for (let revision = 1, revise = revising(value.documents); requested && revise && revision <= PLAN_FILES_REVISIONS; revision++, revise = revising(value.documents)) {
      delete best.value
      const revised = await ask<LeanOutline>(`outline files revision ${revision}/${PLAN_FILES_REVISIONS}`, retryPrompt("draft", revise.feedback), checkOutline, 0, revise.feedback)
      const next = revised.value ?? best.value
      if (!next || contents(next) !== contents(value)) continue
      const concerned = revise.documents
      const ids: DocID[] = withDocumentIDs(value.documents).map(row => row.id)
      value = { ...value, documents: value.documents.map((row, index) => concerned.has(ids[index]!) ? delivered(row, next.documents[index]!) : row) }
      files = requestedFilesIssues(withDocumentIDs(value.documents), value.memo, requested).issues
    }
    const documents = withDocumentIDs(value.documents)
    // A mismatch the revisions leave is the lawyer's call, asked once; once
    // asked, the plan is saved as it stands and the record discloses the files
    // it does not deliver as asked (requestedFilesItems).
    const question = requested && requestedFilesQuestion(directory, requested, documents, files)
    if (question) return { status: "blocked", outputs: ["coverage.json"], openItems: [], questions: [question],
      reason: `${step}: the document plan does not deliver the files the lawyer's words ask for; drafting waits for your answer on how they are delivered.` }
    // A memo the request asks for is the run's notes, named for the request, and
    // a file the lawyer's words name takes that name.
    const memo = value.memo ? { name: value.memo, file: memoNotesFile(value.documents, value.memo) } : undefined
    const named = requested && withRequestedNames(value.documents, memo, requested)
    if (requested && named) {
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis, words: requestedFilesWords(directory), planned: requestedFilesPlanned(named.documents, value.memo), stated: requested.stated, files: requested.files })
      appendLeanEvent(directory, { stage: "draft", document, event: "requested_files", step: "files", outcome: requested.stated ? "stated" : "not_stated", issues: files.map(row => row.issue),
        ...(named.applied.length ? { reason: named.applied.join("; ") } : {}) })
    } else if (filed) {
      // Disclosed from the record (requestedFilesItems), as an unchecked coverage record is.
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis, words: requestedFilesWords(directory), planned: requestedFilesPlanned(value.documents, value.memo), status: "unchecked", stated: false, files: [] })
      appendLeanEvent(directory, { stage: "draft", document, event: "requested_files", step: "files", outcome: "unchecked",
        reason: "The files job's reply did not pass its check, so the files the lawyer asks to receive could not be confirmed; the plan delivers its files as outlined and the drafting notes say so." })
    } else rmSync(leanFile(directory, "requested-files.json"), { force: true })
    writeLeanPlan(directory, named?.documents ?? value.documents, basis, value.acting_for, memo && { name: memo.name, file: named?.notes_file ?? memo.file }, posture)
  } else if (document === FIRST_DOCUMENT && !planCoverageCurrent(directory)) {
    const saved = readLeanPlan(directory)!
    const coverage = await covering(saved.documents)
    recordCoverage(saved.documents, saved.acting_for, coverage)
    const blocked = scopeBlock(saved.documents, coverage, true)
    if (blocked) return blocked
  }
  // A saved plan's accepted draft is final for the run (a new run starts from
  // an empty lean folder): a step re-run after a coverage refresh, after the
  // question it stopped for is answered, or after a pause, keeps it while it
  // was accepted from this very key. The step's row keeps that key (the runner
  // never clears a draft row's input_sha), so a plan entry the lawyer edited or
  // renumbered, or a new prompt version, drafts it again. Its open items are
  // derived from its saved text as at acceptance, since a pause or a stopped
  // recheck can leave the record without them.
  if (!outlined && draftOutputValid(directory, step) && state.stages[step]?.input_sha === stageInputSha(directory, step)) {
    await planned()
    return { status: "done", outputs: stepOutputs(step), openItems: [...checkDraft(directory, readLeanText(directory, `${document}/draft.md`) ?? "").openItems, ...draftTemplateItems(directory, document, readParagraphs(directory, document)), ...draftStatusItems(directory, document, readParagraphs(directory, document), "draft")] }
  }
  const plan = plannedDocument(directory, document)
  if (!plan) return { status: "failed", reason: `${step}: the run's plan has no document ${document}.` }
  // Await source-only extraction for fresh and existing plans before any body part.
  await planned()
  const key = stageInputSha(directory, step)
  const contract = plannedContract(directory, document)
  const carries = contract ? partCarries(contract, plan.parts.length) : []
  // A revised original's tables: each part is shown those its scope holds.
  const original = originalTables(directory, document)
  const scopes = original ? tableParts(directory, document, original.tables) : []
  const drafted: string[] = []
  for (const [index, part] of plan.parts.entries()) {
    const label = `part ${index + 1}/${plan.parts.length}`
    const carried = carries[index]
    const section = sectionPrompt(plan, index, contract && carried && (carried.carried.length || carried.left.length)
      ? { source: contract.source, parts: carried.carried.map(unit => unit.label), left: carried.left.map(unit => unit.label) }
      : undefined)
    const brief = sectionHandoff(directory, document)
    const own = scopes.flatMap((parts, table) => parts.includes(index) ? [table] : [])
    // The outline's session saw the run context before the plan named a memo.
    const text = [...(index > 0 ? [] : outlined ? [actingForBlock(directory), packageCompositionBlock(directory, document)] : [stagePrompt(directory, step)]), originalTablesBlock(original, own), brief, section].filter(Boolean).join("\n\n")
    // The first reply the section check passes is checked once for an
    // original table it leaves out (partTableIssues); asked for it, the part
    // keeps that reply when no later one passes, so the table check never
    // costs a part the text it had.
    let tables = original?.tables
    let kept: string | undefined
    const written = await ask(label, text, (reply, final) => {
      const checked = checkSection(reply, plan, index, final)
      if (!tables || checked.value === undefined) return checked
      const issues = partTableIssues(checked.value, tables, own)
      tables = undefined
      if (!issues.length) return checked
      if (final) return { issues, value: checked.value }
      kept = checked.value
      return { issues }
    })
    const value = written.value ?? kept
    if (!value) return failed(`${label} "${part.heading}"`, written.issues)
    drafted.push(value)
    writeLeanText(directory, `${document}/draft-in-progress.md`, assembleDraft(plan, drafted, false))
  }
  const letter = isLetter(plan.form)
  const presentation = letter || plan.parts.some(part => part.display_heading !== undefined)
    ? { source: assembleDraft(plan, drafted, false), parts: drafted.map((text, part) => ({ heading: headingText(text.split("\n")[0]!.replace(/^##\s+/, "")), display_heading: printedDisplayHeading(plan, part), part })), ...(letter ? { title: false as const } : {}) }
    : undefined
  const accepted = await acceptStageReply(directory, step, assembleDraft(plan, drafted), { presentation })
  if (accepted.issues.length) return failed("assembled document", accepted.issues)
  // A part label the Word body could not place (presentDraft) leaves the part
  // without paragraphs of its own; the draft stands, and the record says what
  // that changes.
  const placed = new Set(readParagraphs(directory, document).map(row => row.draft_part))
  const unplaced = presentation?.parts.filter(row => !placed.has(row.part)) ?? []
  if (unplaced.length) appendLeanEvent(directory, { stage: "draft", document, event: "checker", step: "presentation", issues: unplaced.map(row =>
    `The Word body could not place the part "${row.heading}": its label stays as rendered, and its paragraphs count with the part before it when the draft is checked against a template.`) })
  // Recorded with the body, so a pause or restart before the step settles keeps it.
  progress({ input_sha: key })
  return { status: "done", outputs: accepted.outputs, openItems: [...accepted.openItems, ...draftTemplateItems(directory, document, readParagraphs(directory, document)), ...draftStatusItems(directory, document, readParagraphs(directory, document), "draft")], questions: accepted.questions }
}

export function sectionPrompt(outline: DraftOutline, index: number, template?: { source: string; parts: string[]; left?: string[] }) {
  const part = outline.parts[index]!
  const names = (labels: string[]) => labels.map(name => `"${name}"`).join("; ")
  const original = outline.mode === "revise" ? "the original" : template ? `the template "${template.source}"` : undefined
  return [
    `Draft part ${index + 1} of ${outline.parts.length} of "${outline.title}": ${part.heading}`,
    ...(part.scope ? [`This part covers: ${part.scope}`] : []),
    ...(template
      ? [[
        ...(template.parts.length ? [`From the template "${template.source}", this part carries: ${names(template.parts)}. Carry their text in full, adapted to this matter.`] : []),
        ...(template.left?.length ? [`Leave out ${names(template.left)} of the template "${template.source}", as the outline records.`] : []),
      ].join(" ")]
      : []),
    "",
    "Outline:",
    ...outline.parts.map((row, at) => `${at + 1}. ${row.heading}${at === index ? "  <- this part" : ""}${row.scope ? ` — ${row.scope}` : ""}`),
    "",
    "Reply with the full text of this part only, in Markdown:",
    `- Begin with the line "## ${part.heading}". Use "### " or lower headings inside the part if it needs them.`,
    part.display_heading === undefined ? "- The renderer prints that first line as the part's heading. Do not repeat it in the body."
      : part.display_heading ? `- The renderer replaces that first line with the visible heading "${part.display_heading}". Do not repeat it in the body.`
      // A letter reads as one letter; another document's headings are in its text.
      : isLetter(outline.form) ? "- That first line is an internal part label; the renderer omits it. Begin the body with the document text itself, without a replacement heading."
      : "- That first line is an internal part label; the renderer omits it. Write the headings the document itself needs in the part's text, in its own style" +
        (original ? `, following the heading text and hierarchy of ${original} except where the outline or the instructions change them.` : "."),
    "- Do not repeat the title or write any other part. Keep defined terms, numbering and cross-references consistent with the outline, the instructions and the parts already drafted.",
    "- The source, quotation and open item rules of the first prompt still apply.",
  ].join("\n")
}

/** Refresh decisions at each section boundary without repeating the whole
 * source bundle. A stage-owned snapshot preserves exact accepted parts; retrieval stays available
 * for shared definitions and evidence outside the current part's scope. */
function sectionHandoff(directory: string, document: DocID) {
  const updates = readLeanText(directory, "updated-notes.md")?.trim()
  const answers = answeredQuestionsText(directory)
  const files = ["spine.json", "research.json", "research-review.json"].filter(name => readLeanText(directory, name))
  return [
    ...(plannedDocument(directory, document)?.mode === "revise" ? ["Revision foundation: preserve unaffected source wording and operative clauses, including their qualifications. Read the full original part before editing it. A summary or recommendation does not replace operative wording. Explain substantive omissions against the supplied source."] : []),
    ...(readLeanText(directory, `${document}/draft-in-progress.md`)?.trim() ? [`Exact accepted parts of this stage attempt are available at .dochaus/lean/${document}/draft-in-progress.md. Read them when checking shared definitions, numbering or cross-references; a conversation summary does not replace their wording.`] : []),
    "Section handoff: use the current part's scope and provided source references to retrieve the passages it needs. Preserve shared definitions, settled decisions and cross-references from the preceding parts in this session; inspect their source dependencies when needed. The other outline scopes show where related provisions belong.",
    ...(files.length ? [`Working evidence records available with read: ${files.map(name => `.dochaus/lean/${name}`).join(", ")}. These are retrieval guides, not independent authority; verify underlying matter passages and use only research approved by research-review.`] : []),
    ...(updates ? [`Current user updates, in arrival order (explicit later changes control; tentative suggestions are not settled decisions):\n${updates}`] : []),
    ...(answers ? [`Answered decisions — apply these and do not ask them again without new contrary evidence:\n${answers}`] : []),
  ].join("\n\n")
}

/** Checks one part reply. The outlined heading is the authority for the part's
 * first line: a reply that starts straight into its text gets that heading.
 * The document title, where publication prints it (not in a letter), is
 * printed once, above the first part, and a part's display heading, or its
 * label where that is printed, from its first line: a part whose text opens
 * with a line repeating what is printed above it (openingRepeats) is retried,
 * and on the last attempt (`final`) kept without that line. */
export function checkSection(reply: string, outline: DraftOutline, index: number, final = false): Checked<string> {
  const { heading, display_heading: display } = outline.parts[index]!
  if (!reply.trim()) return { issues: [`The reply was empty. Reply with the text of the part "${heading}".`] }
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const body = reply.trim().replace(/\r\n/g, "\n")
  const drafted = /^#{1,2}\s/.test(body) ? body : `## ${heading}\n\n${body}`
  const opening = openingRepeats(drafted, {
    title: index === 0 && !isLetter(outline.form) ? outline.title : undefined,
    heading: display === undefined ? heading : display || undefined,
  })
  const text = opening.text
  const [first = "", ...rest] = text.split("\n")
  const others = new Set(outline.parts.filter((_, at) => at !== index).map(part => normalizedHeading(part.heading)))
  const strays = rest.map(line => /^##\s+(.+)$/.exec(line.trim())?.[1]).filter((name): name is string => Boolean(name) && others.has(normalizedHeading(name!)))
  const issues = [
    ...(normalizedHeading(first.replace(/^##\s+/, "")) !== normalizedHeading(heading) || !/^##\s/.test(first)
      ? [`Begin the part with the line "## ${heading}".`]
      : []),
    ...(rest.some(line => /^#\s+\S/.test(line.trim()))
      ? ["The part contains a \"# \" title line. Only the document title uses \"# \"; use \"### \" or lower headings inside a part."]
      : []),
    ...(strays.length
      ? [`The reply also contains other parts (${strays.map(name => `"${headingText(name)}"`).join(", ")}). Reply with the part "${heading}" only.`]
      : []),
    ...(!strays.length && rest.some(line => /^##\s+\S/.test(line.trim()))
      ? ['The part contains a "## " heading in its body. Only its first line identifies the part; use "### " or lower headings inside it.']
      : []),
    ...(!rest.some(line => line.trim() && !/^#{1,6}\s/.test(line.trim())) ? [`The part "${heading}" has no text under its heading.`] : []),
  ]
  const repeated = opening.issues
  if (issues.length) return { issues: [...repeated, ...issues] }
  return repeated.length && !final ? { issues: repeated } : { issues: repeated, value: text }
}

/** A part's text ("## " label first) without the lines at the very start of
 * its text that repeat what is printed above them: the document title, over
 * the part that opens the document, and the heading printed from the part's
 * label (markup, case and closing punctuation aside). Each is left out once,
 * with an issue naming the line; a later line is never touched, so a title
 * restated after a cover sheet or a mention in the text stays. */
function openingRepeats(text: string, printed: { title?: string; heading?: string }) {
  const [label = "", ...lines] = text.split("\n")
  const left = { ...printed }
  const issues: string[] = []
  let rest = lines
  for (;;) {
    const opening = rest.findIndex(line => line.trim())
    const line = rest[opening]?.trim()
    if (!line) break
    const plain = plainLine(line)
    const kind = left.title && plain === plainLine(left.title) ? "title" : left.heading && plainLine(left.heading) && plain === plainLine(left.heading) ? "heading" : undefined
    if (!kind) break
    issues.push(kind === "title"
      ? `The part opens with the document title "${left.title}". The title is printed once at the top of the document: keep the "## " line first and remove the line "${line}" after it.`
      : `The part opens with its heading "${left.heading}", which is printed from its "## " line: keep the "## " line first and remove the line "${line}" after it.`)
    delete left[kind]
    rest = rest.slice(opening + 1)
  }
  if (!issues.length) return { text, issues }
  const kept = rest.join("\n").replace(/^\s*\n/, "").trimEnd()
  return { text: kept ? `${label}\n\n${kept}` : label, issues }
}

/** A line's visible words, for comparing a part's opening line with what is printed above it. */
function plainLine(line: string) {
  return visibleInlineText(line.trim().replace(/^#{1,6}\s+/, "")).replace(/[*_`]/g, "").replace(/[\s.:;,]+$/, "").replace(/\s+/g, " ").trim().toLowerCase()
}

/** The heading publication prints in place of part `index`'s label: its
 * display heading (undefined: the label itself). The title is printed once,
 * above the part that opens the document (not in a letter), so that part's
 * display heading repeating the title (markup, case, spacing and closing
 * punctuation aside) prints nothing; a later part's stands, as a title
 * restated after a cover sheet does. */
function printedDisplayHeading(outline: DraftOutline, index: number) {
  const display = outline.parts[index]?.display_heading
  return display && index === 0 && !isLetter(outline.form) && plainLine(display) === plainLine(outline.title) ? "" : display
}

/** The draft from its parts. With `presentation`, each part label gives way to
 * its display heading, and a part that prints none has its own headings moved
 * up, keeping their relative depth, until the shallowest is a second-level
 * heading under the title, as presentDraft does in the Word body; without it,
 * the labels stay as the internal boundaries the renderer and paragraph
 * attribution read. */
export function assembleDraft(outline: DraftOutline, parts: string[], presentation = true) {
  return `# ${outline.title}\n\n${parts.map(part => {
    const text = part.trim()
    const label = headingText(text.split("\n")[0]!.replace(/^##\s+/, ""))
    const at = outline.parts.findIndex(row => normalizedHeading(row.heading) === normalizedHeading(label))
    const display = at < 0 ? undefined : printedDisplayHeading(outline, at)
    if (!presentation || display === undefined) return text
    const body = text.split("\n").slice(1).join("\n").trim()
    if (display) return [`## ${display}`, body].filter(Boolean).join("\n\n")
    const levels = [...body.matchAll(/^[ \t]*(#{3,6})[ \t]/gm)].map(match => match[1]!.length)
    if (!levels.length) return body
    const shift = Math.min(...levels) - 2
    return body.replace(/^([ \t]*)(#{3,6})(?=[ \t])/gm, (_, indent: string, marks: string) => `${indent}${"#".repeat(marks.length - shift)}`)
  }).join("\n\n")}\n`
}
