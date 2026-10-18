import { rmSync } from "node:fs"
import { playbookSelectionIssues } from "./playbook-review"
import { TOOL_SYNTAX_ISSUE, printedToolCall } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { matterSourceIDs } from "./quotes"
import { memoNotesFile } from "./deliverables"
import { LEAN_COVERAGE_SCHEMA, coveragePlanHash, coveragePrompt, coverageScopeQuestion, parseCoverage, planCoverageCurrent, planCoverageIssues, type CoverageAct, type LeanCoverage } from "./package-coverage"
import { answeredScopeQuestion, openScopeQuestion } from "./package-scope"
import {
  LEAN_REQUESTED_FILES_SCHEMA, parseRequestedFiles, requestedFilesIssues, requestedFilesPlanned, requestedFilesPrompt, requestedFilesQuestion, requestedFilesRevision, withRequestedNames, writeRequestedFiles,
  type RequestedFiles,
} from "./requested-files"
import {
  FIRST_DOCUMENT, headingText, normalizedHeading, parsePlan, plannedDocument, readLeanPlan, withDocumentIDs, writeLeanPlan,
  type ActingFor, type DocID, type LeanOutline, type PlannedDocument, type PlannedPart, type StepID,
} from "./plan"
import { actingForBlock, packageCompositionBlock, retryPrompt, stagePrompt } from "./prompts"
import {
  LEAN_CHILD_PERMISSIONS,
  LEAN_MAX_RETRIES,
  LEAN_STAGE_AGENTS,
  acceptStageReply,
  draftOutputValid,
  outlineBasis,
  planBasisCurrent,
  stepOutputs,
  stepSessionTitle,
  type LeanDeps,
} from "./stages"
import { answeredQuestionsText, handedBackBlock, leanFile, readLeanText, readOpenItems, readQuestions, writeLeanText, type LeanStageState, type LeanState, type OpenItem, type Question } from "./state"
import { outlineCoverageIssues, partCarries, plannedContract, type TemplateContract } from "./template-contract"
import { contractGaps, draftBody, draftGapIssues, type DraftPart } from "./template-coverage"
import { countWords, headingUnits, tokens } from "./template-structure"

// Sectioned draft step. A long document can need more tokens than one model
// reply may carry, so the draft agent writes one part per reply in the same
// child session, so native history carries the parts already written. When the
// run has no current plan, the first document's session starts with the outline
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
// top-level part of the template (template-contract.ts); each part is told
// which template parts it carries and which parts inside them the outline
// leaves out, and, once all parts are written, the same session is asked for
// the parts that leave a carried template part out or carry only a small share
// of it (template-coverage.ts). These checks never fail the step: after the
// allowed retries the reply with the fewest gaps is kept, and the apply step
// records any gap left in the final body as an open item.

export type DraftOutline = { title: string; parts: PlannedPart[] }

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

/** The saved plan is followed (planBasisCurrent) and, unless it was written by
 * hand, its coverage record checked this very plan (planCoverageCurrent, which
 * the first document's draft step also needs to count as done on resume). */
export function planIsCurrent(directory: string) {
  return planBasisCurrent(directory) && planCoverageCurrent(directory)
}

export async function runDraftStage(
  directory: string,
  state: LeanState,
  deps: LeanDeps,
  progress: (patch: Partial<LeanStageState>) => void,
  document: DocID,
  planned: () => void | Promise<void>,
  /** The step was done from these very inputs; only its saved output was found wanting. */
  current = false,
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
  // asks.
  const scopeBlock = (documents: PlannedDocument[], coverage: Coverage): DraftStageRun | undefined => {
    if (!coverage.acts || !coverage.issues.length) return
    const questions = readQuestions(directory)
    if (openScopeQuestion(questions, "draft")) return
    const question = coverageScopeQuestion(coverage.acts, documents, questions.filter(answeredScopeQuestion))
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
  // A plan whose outline basis is current is followed as it stands, including a
  // lawyer's edits; only its coverage record is refreshed when it no longer
  // matches the plan.
  const outlined = document === FIRST_DOCUMENT && !planBasisCurrent(directory)
  if (outlined) {
    const basis = outlineBasis(directory)
    // The lawyer's own words ground requested file names and the memo; a
    // question handed back to the drafting adds none.
    const request = [readLeanText(directory, "request.md"), readLeanText(directory, "updated-notes.md"), answeredQuestionsText(directory)].filter(Boolean).join("\n\n")
    const best: { value?: LeanOutline } = {}
    const checkOutline = (reply: string): Checked<LeanOutline> => {
      const parsed = parsePlan(reply, request)
      if (!parsed.value) return parsed
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
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis, planned: requestedFilesPlanned(named.documents, value.memo), stated: requested.stated, files: requested.files })
      appendLeanEvent(directory, { stage: "draft", document, event: "requested_files", step: "files", outcome: requested.stated ? "stated" : "not_stated", issues: files.map(row => row.issue),
        ...(named.applied.length ? { reason: named.applied.join("; ") } : {}) })
    } else if (filed) {
      // Disclosed from the record (requestedFilesItems), as an unchecked coverage record is.
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis, planned: requestedFilesPlanned(value.documents, value.memo), status: "unchecked", stated: false, files: [] })
      appendLeanEvent(directory, { stage: "draft", document, event: "requested_files", step: "files", outcome: "unchecked",
        reason: "The files job's reply did not pass its check, so the files the lawyer asks to receive could not be confirmed; the plan delivers its files as outlined and the drafting notes say so." })
    } else rmSync(leanFile(directory, "requested-files.json"), { force: true })
    writeLeanPlan(directory, named?.documents ?? value.documents, basis, value.acting_for, memo && { name: memo.name, file: named?.notes_file ?? memo.file })
  } else if (document === FIRST_DOCUMENT && !planIsCurrent(directory)) {
    const saved = readLeanPlan(directory)!
    const coverage = await covering(saved.documents)
    recordCoverage(saved.documents, saved.acting_for, coverage)
    const blocked = scopeBlock(saved.documents, coverage)
    if (blocked) return blocked
    // When only the coverage record was out of date (a lawyer's edit to another
    // document's entry, say), the draft already done from these inputs is kept
    // with its open items.
    if (current && draftOutputValid(directory, step)) {
      await planned()
      return { status: "done", outputs: stepOutputs(step), openItems: readOpenItems(directory).filter(item => item.stage === "draft" && item.ref.document === document) }
    }
  }
  const plan = plannedDocument(directory, document)
  if (!plan) return { status: "failed", reason: `${step}: the run's plan has no document ${document}.` }
  // Await source-only extraction for fresh and existing plans before any body part.
  await planned()
  const contract = plannedContract(directory, document)
  const carries = contract ? partCarries(contract, plan.parts.length) : []
  const drafted: DraftPart[] = []
  const snapshot = (parts: DraftPart[]) => writeLeanText(directory, `${document}/draft-in-progress.md`, assembleDraft(plan, parts.map(part => part.text), false))
  for (const [index, part] of plan.parts.entries()) {
    const label = `part ${index + 1}/${plan.parts.length}`
    const carried = carries[index]
    const section = sectionPrompt(plan, index, contract && carried && (carried.carried.length || carried.left.length)
      ? { source: contract.source, parts: carried.carried.map(unit => unit.label), left: carried.left.map(unit => unit.label) }
      : undefined)
    const brief = sectionHandoff(directory, document)
    // The outline's session saw the run context before the plan named a memo.
    const text = [...(index > 0 ? [] : outlined ? [actingForBlock(directory), packageCompositionBlock(directory, document)] : [stagePrompt(directory, step)]), brief, section].filter(Boolean).join("\n\n")
    const written = await ask(label, text, reply => checkSection(reply, plan, index))
    if (!written.value) return failed(`${label} "${part.heading}"`, written.issues)
    drafted.push({ part: index, text: written.value })
    snapshot(drafted)
  }
  const parts = contract ? await completeTemplateParts(ask, plan, contract, drafted, snapshot) : drafted
  const presentation = plan.parts.some(part => part.display_heading !== undefined)
    ? { source: assembleDraft(plan, parts.map(part => part.text), false), parts: parts.map(part => ({ heading: headingText(part.text.split("\n")[0]!.replace(/^##\s+/, "")), display_heading: part.part === undefined ? undefined : plan.parts[part.part]?.display_heading, part: part.part })) }
    : undefined
  const accepted = await acceptStageReply(directory, step, assembleDraft(plan, parts.map(part => part.text)), { presentation })
  if (accepted.issues.length) return failed("assembled document", accepted.issues)
  return { status: "done", outputs: accepted.outputs, openItems: accepted.openItems, questions: accepted.questions }
}

export function sectionPrompt(outline: DraftOutline, index: number, template?: { source: string; parts: string[]; left?: string[] }) {
  const part = outline.parts[index]!
  const names = (labels: string[]) => labels.map(name => `"${name}"`).join("; ")
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
    ...(part.display_heading !== undefined ? [part.display_heading ? `- The renderer replaces that first line with the visible heading "${part.display_heading}". Do not repeat it in the body.` : "- That first line is an internal part label; the renderer omits it. Begin the body with the document text itself, without a replacement heading."] : []),
    "- Do not repeat the title or write any other part. Keep defined terms, numbering and cross-references consistent with the outline and the parts already drafted.",
    "- The source, quotation and open item rules of the first prompt still apply.",
  ].join("\n")
}

/** Refresh decisions at each section boundary without repeating the whole
 * source bundle. A stage-owned snapshot preserves exact accepted parts; retrieval stays available
 * for shared definitions and evidence outside the current part's scope. */
function sectionHandoff(directory: string, document: DocID) {
  const updates = readLeanText(directory, "updated-notes.md")?.trim()
  const answers = answeredQuestionsText(directory)
  const handed = handedBackBlock(directory)
  const files = ["spine.json", "research.json", "research-review.json"].filter(name => readLeanText(directory, name))
  return [
    ...(plannedDocument(directory, document)?.mode === "revise" ? ["Revision foundation: preserve unaffected source wording and operative clauses, including their qualifications. Read the full original part before editing it. A summary or recommendation does not replace operative wording. Explain substantive omissions against the supplied source."] : []),
    ...(readLeanText(directory, `${document}/draft-in-progress.md`)?.trim() ? [`Exact accepted parts of this stage attempt are available at .dochaus/lean/${document}/draft-in-progress.md. Read them when checking shared definitions, numbering or cross-references; a conversation summary does not replace their wording.`] : []),
    "Section handoff: use the current part's scope and provided source references to retrieve the passages it needs. Preserve shared definitions, settled decisions and cross-references from the preceding parts in this session; inspect their source dependencies when needed. The other outline scopes show where related provisions belong.",
    ...(files.length ? [`Working evidence records available with read: ${files.map(name => `.dochaus/lean/${name}`).join(", ")}. These are retrieval guides, not independent authority; verify underlying matter passages and use only research approved by research-review.`] : []),
    ...(updates ? [`Current user updates, in arrival order (explicit later changes control; tentative suggestions are not settled decisions):\n${updates}`] : []),
    ...(answers ? [`Answered decisions — apply these and do not ask them again without new contrary evidence:\n${answers}`] : []),
    ...(handed ? [handed] : []),
  ].join("\n\n")
}

/** Checks one part reply. The outlined heading is the authority for the part's
 * first line: a reply that starts straight into its text gets that heading. */
export function checkSection(reply: string, outline: DraftOutline, index: number): Checked<string> {
  const heading = outline.parts[index]!.heading
  if (!reply.trim()) return { issues: [`The reply was empty. Reply with the text of the part "${heading}".`] }
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const body = reply.trim().replace(/\r\n/g, "\n")
  const text = /^#{1,2}\s/.test(body) ? body : `## ${heading}\n\n${body}`
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
  return issues.length ? { issues } : { issues, value: text }
}

/** The completeness exchange: when the assembled parts leave a carried
 * template part out or carry only a small share of it, the same session is
 * asked for the revised parts (mergeParts). Of the drafts so far, the one with
 * the fewest gaps is kept (the latest of equals), and each reply is merged onto
 * it; a reply that cannot be read leaves it as it was. */
async function completeTemplateParts(
  ask: <T>(label: string, text: string, check: (reply: string) => Checked<T>) => Promise<Checked<T>>,
  outline: DraftOutline,
  contract: TemplateContract,
  drafted: DraftPart[],
  snapshot: (parts: DraftPart[]) => void,
) {
  const first = contractGaps(contract, draftBody(drafted))
  if (!first.length) return drafted
  const best = { parts: drafted, gaps: first }
  await ask("completeness", [
    `Compared with its template "${contract.source}", the draft of "${outline.title}" falls short:`,
    ...draftGapIssues(contract, first).map(issue => `- ${issue}`),
  ].join("\n"), reply => {
    const merged = mergeParts(reply, outline, best.parts, contract, new Set(best.gaps.map(gap => gap.unit.index)))
    if (!merged.value) return merged
    const gaps = contractGaps(contract, draftBody(merged.value))
    if (gaps.length <= best.gaps.length) {
      Object.assign(best, { parts: merged.value, gaps })
      snapshot(best.parts)
    }
    return { issues: draftGapIssues(contract, gaps), value: merged.value }
  })
  return best.parts
}

/** Reads revised parts ("## " heading, then text) into the parts so far. A
 * part whose heading reads as an existing part's (numbering, case and spacing
 * aside), or names a template part the draft falls short on and every template
 * part an existing part's heading names, revises that part in place under its
 * own heading; a part headed with such a template part that no existing
 * heading names is added after the last part carrying an earlier template
 * part. A heading naming only some of what an existing part's heading names,
 * any other heading, or a revised part with less than half the words it had,
 * is an issue, so no drafted text is replaced by a reply about part of it. */
export function mergeParts(reply: string, outline: DraftOutline, parts: DraftPart[], contract: TemplateContract, gapped: Set<number>): Checked<DraftPart[]> {
  if (!reply.trim()) return { issues: ["The reply was empty. Reply with the revised parts, each beginning with its \"## \" heading."] }
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const sections = reply.replace(/\r\n/g, "\n").split(/\n(?=##\s)/).map(text => text.trim()).filter(text => /^##\s+\S/.test(text))
  if (!sections.length) return { issues: ["The reply has no part. Begin each revised part with its \"## \" heading."] }
  const headingOf = (text: string) => headingText((text.split("\n")[0] ?? "").replace(/^##\s+/, ""))
  const shape = sections.flatMap(text => {
    const rest = text.split("\n").slice(1)
    return [
      ...(rest.some(line => /^#\s+\S/.test(line.trim())) ? [`The part "${headingOf(text)}" contains a "# " title line; use "### " or lower headings inside a part.`] : []),
      ...(!rest.some(line => line.trim() && !/^#{1,6}\s/.test(line.trim())) ? [`The part "${headingOf(text)}" has no text under its heading.`] : []),
    ]
  })
  if (shape.length) return { issues: shape }
  const key = (text: string) => tokens(headingOf(text)).join(" ")
  const carries = partCarries(contract, outline.parts.length).map(row => row.carried.map(unit => unit.index))
  const carriedBy = (part: DraftPart) => part.part === undefined ? [...headingUnits(headingOf(part.text), contract.units)] : carries[part.part] ?? []
  type Merged = { parts: DraftPart[]; issues: string[] }
  // A revision keeps the part's own heading line, so the outline still finds it.
  const revise = (acc: Merged, at: number, text: string): Merged => {
    const before = acc.parts[at]!.text
    if (countWords(text) * 2 < countWords(before))
      return { ...acc, issues: [...acc.issues, `The revised part "${headingOf(text)}" is much shorter than before. Reply with its complete text, keeping what it already carries.`] }
    const revised = `${before.split("\n")[0]}\n\n${text.split("\n").slice(1).join("\n").trim()}`
    return { ...acc, parts: acc.parts.map((part, index) => index === at ? { ...part, text: revised } : part) }
  }
  const merged = sections.reduce<Merged>((acc, text) => {
    const at = acc.parts.findIndex(part => key(part.text) === key(text))
    if (at >= 0) return revise(acc, at, text)
    const named = [...headingUnits(headingOf(text), contract.units)].filter(index => gapped.has(index))
    if (!named.length)
      return { ...acc, issues: [...acc.issues, `The part "${headingOf(text)}" is neither a part of the outline nor a template part the draft leaves out. ` +
        "Begin each revised part with its exact \"## \" heading from the outline, or head a new part with the template part it carries."] }
    const headed = acc.parts.findIndex(part => named.some(index => headingUnits(headingOf(part.text), contract.units).has(index)))
    if (headed >= 0) {
      const names = headingUnits(headingOf(text), contract.units)
      if ([...headingUnits(headingOf(acc.parts[headed]!.text), contract.units)].every(index => names.has(index))) return revise(acc, headed, text)
      return { ...acc, issues: [...acc.issues, `The part "${headingOf(text)}" covers only some of the part "${headingOf(acc.parts[headed]!.text)}". ` +
        "Reply with that part under its exact \"## \" heading from the outline and its complete text, keeping what it already carries."] }
    }
    const after = acc.parts.findLastIndex(part => carriedBy(part).some(index => index < Math.min(...named)))
    return { ...acc, parts: [...acc.parts.slice(0, after + 1), { text }, ...acc.parts.slice(after + 1)] }
  }, { parts, issues: [] })
  return merged.issues.length ? { issues: merged.issues } : { issues: [], value: merged.parts }
}

export function assembleDraft(outline: DraftOutline, parts: string[], presentation = true) {
  return `# ${outline.title}\n\n${parts.map(part => {
    const text = part.trim()
    const label = headingText(text.split("\n")[0]!.replace(/^##\s+/, ""))
    const display = outline.parts.find(row => normalizedHeading(row.heading) === normalizedHeading(label))?.display_heading
    if (!presentation || display === undefined) return text
    return [display ? `## ${display}` : "", text.split("\n").slice(1).join("\n").trim()].filter(Boolean).join("\n\n")
  }).join("\n\n")}\n`
}
