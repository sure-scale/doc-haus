import { isDeepStrictEqual } from "node:util"
import { normalize } from "../../lib/quote-verify"
import { parallel } from "./checklist-flow"
import { parseJsonObject, readParagraphs, record, shownText } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { intakeActingFor, matterDatesBlock } from "./intake"
import { FIRST_DOCUMENT, planActingFor, readLeanPlan, stepDocument, stepStage, type StepID } from "./plan"
import { retryPrompt, tagged } from "./prompts"
import { REQUEST_SOURCE, isAnswerSource, matterSourceIDs, quoteFound, quoteMissIssue } from "./quotes"
import { scopeQuestionOffersNewDocument, scopeQuestionTitles } from "./package-coverage"
import { matterClient } from "../../lib/matter-client"
import { nameWithProvenance, provenanceMapLine, sourceProvenance } from "../../lib/source-provenance"
import { confirmationDocument, currentInstructions, INSTRUCTIONS_FILE, type Instruction, type InstructionDocument } from "./instructions"
import { sourceText } from "./template-structure"
import { approvedResearch, researchHandoff } from "./research"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, stepSessionTitle, type LeanDeps } from "./stages"
import { answeredQuestionsText, readLeanText, readQuestions, type LeanState, type Question, type QuestionAdvice, type QuestionOption } from "./state"

// Advice on the questions a step asks, before the runner records them: one
// fresh read-only session per question, with the same-session bounded retry
// of every lean job. The advisor sets out the options and recommends one only
// with passages code verifies, or says why it cannot; it never answers. When
// its retries run out the question is recorded all the same, its advice
// unavailable with reason_code advisor_failed.

const ADVISOR_FAILED = "No checked recommendation is available for this question."

/** `questions`, each one not already answered carrying its advice. An open
 * question asked again with the same words and code-given options keeps the
 * advice it was recorded with: an option id is resolved when the answer is
 * recorded, so the card the lawyer answers must not change under them. */
export async function adviseQuestions(directory: string, state: LeanState, step: StepID, deps: LeanDeps, questions: Question[], reviewed?: ReadonlyMap<string, { advice?: QuestionAdvice; group?: string }>, reuseSavedAdvice = true): Promise<Question[]> {
  const recorded = readQuestions(directory)
  const answered = new Set(recorded.filter(question => question.status === "answered").map(question => question.id))
  const open = new Map(recorded.filter(question => question.status === "open" && question.advice).map(question => [question.id, question]))
  const kept = questions.map(question => {
    const earlier = open.get(question.id)
    if (!earlier || earlier.question !== question.question || !isDeepStrictEqual(earlier.options, question.options)) return undefined
    const same = earlier.context === question.context && earlier.kind === question.kind && earlier.document === question.document && earlier.pid === question.pid
    const advice = earlier.advice!
    const currentBasis = !("recommendation" in advice) || !basisIssues(directory, advice.recommendation.basis).issues.length
    if (reuseSavedAdvice && same && currentBasis) return advice
    // Existing option ids keep their meaning while an unanswered card changes.
    // A stale recommendation cannot be adopted as though still checked.
    return { ...(advice.options ? { options: advice.options } : {}), unavailable: { reason: "The question or its supporting inputs changed. Check the current sources before choosing an earlier option, or answer in your own words." } }
  })
  const cached = questions.map(question => reviewed?.get(JSON.stringify([question.question, question.context, question.kind, question.document, question.pid])))
  const advised = questions.map((question, index) => kept[index] ? { ...question, ...(open.get(question.id)?.group ? { group: open.get(question.id)!.group } : {}), advice: kept[index] } : cached[index]?.advice ? { ...question, ...cached[index] } : question)
  await parallel(questions.flatMap((question, index) => answered.has(question.id) || kept[index] || cached[index]?.advice ? [] : [index]), async index => {
    advised[index] = { ...questions[index]!, ...await adviseQuestion(directory, state, step, deps, questions[index]!) }
  })
  return advised
}

export async function adviseQuestion(directory: string, state: LeanState, step: StepID, deps: LeanDeps, question: Question): Promise<{ advice: QuestionAdvice; group?: string }> {
  const agent = LEAN_STAGE_AGENTS.advise
  const at = { stage: stepStage(step), document: stepDocument(step) ?? undefined, step: `advise ${question.id}` }
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, step, `advise: ${question.id}`), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("advise", issues) : advisePrompt(directory, question) })
      const answered = performance.now()
      const checked = checkAdvice(directory, question, reply, { requireRelevance: true })
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (checked.advice) return { advice: await confirmedDecision(directory, state, step, deps, question, checked.advice), ...(checked.group ? { group: checked.group } : {}) }
    }
  } catch {
    // A runtime error ends this question's advice; a pause or a settings
    // change is the runner's to act on once every advice job has settled.
  }
  return { advice: { unavailable: { reason: ADVISOR_FAILED, reason_code: "advisor_failed" } } }
}

/** `advice`, keeping the instruction it says decides the question only when a
 * focused job of its own confirms it (decidesPrompt): a question so decided is
 * settled without being asked (settleByInstructions), so the advisor's word
 * alone does not settle it. Unconfirmed, or when the job fails, the question
 * is asked with the same advice. */
async function confirmedDecision(directory: string, state: LeanState, step: StepID, deps: LeanDeps, question: Question, advice: QuestionAdvice): Promise<QuestionAdvice> {
  const recommendation = "recommendation" in advice ? advice.recommendation : undefined
  const item = recommendation?.decided_by ? currentInstructions(directory).find(row => row.id === recommendation.decided_by) : undefined
  if (!recommendation?.decided_by) return advice
  const { decided_by: _decidedBy, ...kept } = recommendation
  const undecided = { ...advice, recommendation: kept }
  const option = advice.options?.find(row => row.id === recommendation.option)
  if (!item || !option || (question.kind !== "instruction" && question.kind !== "judgment") || question.mixed) return undecided
  const agent = LEAN_STAGE_AGENTS.advise
  const at = { stage: stepStage(step), document: stepDocument(step) ?? undefined, step: `decides ${question.id}` }
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, step, `decides: ${question.id}`), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("advise", issues) : decidesPrompt(directory, question, item, option.text) })
      const answered = performance.now()
      const checked = checkDecides(directory, item, reply)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (checked.decides !== undefined) return checked.decides ? advice : undecided
    }
  } catch {
    // As for the advice itself: a runtime error leaves the question to be asked.
  }
  return undecided
}

/** One question for one instruction: whether its words answer the question,
 * leaving the lawyer nothing to choose. */
export function decidesPrompt(directory: string, question: Question, item: Instruction, option: string) {
  return [
    "This is a lean drafting run. Decide one thing: whether one of the lawyer's instructions answers a question raised while drafting, so that it is settled without asking the lawyer. A question so settled is never put to the lawyer, so settle it only when the lawyer has already made this choice.",
    tagged("question", `Kind: ${question.kind}\n${question.question}`),
    ...(question.context ? [tagged("context", question.context)] : []),
    tagged("instruction", `[${item.id}] ${item.text}\nThe lawyer's words, in ${item.source_id}: "${item.quote}"`),
    tagged("answer_if_settled", option),
    tagged("all_instructions", readLeanText(directory, INSTRUCTIONS_FILE) ?? ""),
    tagged("answers", answeredQuestionsText(directory) || "None."),
    tagged("output_contract", [
      "- The instruction decides the question only when its words make this choice: they cover the point the question asks, and the answer above is what they direct.",
      "- It does not decide the question when the choice remains the lawyer's: the instruction is a preference, or is subject to a condition, exception or qualification the question turns on; another of the lawyer's instructions or answers bears on the same point the other way; or the question asks whether the instruction can or should be followed as given. An instruction that is related, points one way or needs inference to apply does not decide it.",
      `- When it decides the question, reply {"decides":true,"quote":"..."}, with the words of the lawyer that make the choice, copied verbatim from ${item.source_id}. When it does not, reply {"decides":false,"reason":"what leaves the choice to the lawyer"}.`,
      "- Answer from the inputs above; read-only tools are available if a source passage must be checked. The run saves and checks the reply.",
    ].join("\n")),
  ].join("\n\n")
}

const DECIDES_SHAPE = 'Reply with the JSON object only: {"decides":true,"quote":"..."} or {"decides":false,"reason":"..."}.'

/** Reads the reply to decidesPrompt: true with words `item`'s source holds, or
 * false with a reason. Every other reply is retried. */
export function checkDecides(directory: string, item: Instruction, reply: string): { issues: string[]; decides?: boolean } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`${DECIDES_SHAPE} ${parsed.error}`] }
  if (typeof parsed.value.decides !== "boolean") return { issues: [DECIDES_SHAPE] }
  const { decides, quote, reason } = parsed.value
  if (!decides) return typeof reason === "string" && reason.trim() ? { issues: [], decides: false } : { issues: [`"decides" is false, but "reason" is missing. ${DECIDES_SHAPE}`] }
  if (typeof quote !== "string" || !quote.trim()) return { issues: [`"decides" is true, but "quote" is missing. ${DECIDES_SHAPE}`] }
  const words = { source_id: item.source_id, quote: quote.trim() }
  return quoteFound(directory, words) ? { issues: [], decides: true } : { issues: [quoteMissIssue(directory, "quote", words, "json")] }
}

const OPENING_CHARS = 1_500
const CLOSE_CHARS = 600

/** The advice prompt for a confirmation question (confirmationQuestions):
 * only what an authorship question needs, each kind in its own tag, so the
 * job's context stays concise. Whether to follow a document turns on who
 * wrote it, to whom and for which work, which its opening (an email's or
 * letter's heading) and close (a sign-off) show; the instruction list,
 * research, the party acted for and the matter's other sources bear on other
 * questions. The reply contract is advisePrompt's for a question with options. */
function confirmationPrompt(directory: string, question: Question, document: InstructionDocument) {
  const text = sourceText(directory, document.source_id).trim()
  const opening = text.slice(0, OPENING_CHARS)
  const close = text.length > OPENING_CHARS ? text.slice(Math.max(OPENING_CHARS, text.length - CLOSE_CHARS)) : ""
  const updates = readLeanText(directory, "updated-notes.md")
  const basis = '"basis":[{"source_id":"request or the exact document name","quote":"words copied exactly from that source","relevance":"how this passage shows who wrote the document, to whom or for which work, or what the lawyer said about following it"}]'
  return [
    `Advise the lawyer on ONE open question raised while drafting: whether to follow ${document.source_id} as the lawyer's instructions for this work. You do not answer it; the lawyer decides. Judge it from the document itself: who wrote it, to whom, and for which work. Recommend following it when the document shows that the lawyer's client or the lawyer's own firm wrote it for this work, and treating it as evidence when it shows that someone else wrote it or that it was written for other work, or when it does not show who wrote it. The request seldom names the lawyer or their firm; that alone is no reason to treat the document as evidence. Where the request, a later instruction or an answer says whether to follow this document, recommend what it says; asking to draft from the attached documents does not say that. Quote each supporting passage exactly. The document's opening and close are below; use a targeted read-document lookup of ${document.source_id} only if they do not show who wrote it, to whom or for which work. Give a short group topic in group.`,
    tagged("question", `Kind: ${question.kind ?? "not stated"}\n${question.question}`),
    ...(question.context ? [tagged("context", question.context)] : []),
    tagged("options", (question.options ?? []).map((option, index) => `${index}: ${option.text}`).join("\n")),
    tagged("document_opening", opening || "The document's text could not be read; use read-document."),
    ...(close ? [tagged("document_close", close)] : []),
    tagged("request", readLeanText(directory, "request.md") ?? ""),
    ...(updates ? [tagged("user_updates", updates)] : []),
    tagged("answers", answeredQuestionsText(directory) || "None."),
    `Reply with one JSON object only, including "group":"short topic": {"recommendation":{"option":0,"rationale":"why this option",${basis}}}, where "option" is the index of one of the options above; or, when no option is supported, {"unavailable":{"reason":"why no option can be recommended"}}. Do not return options of your own.`,
  ].join("\n\n")
}

// The outline records whom the lawyer acts for in the plan: the party the request
// names, else the matter's client, else the side intake took from the sources.
// Questions asked before the outline get the same precedence from the request,
// matter.json and intake.json directly.
function actingForAdvice(directory: string) {
  const acting = planActingFor(directory)
  if (acting.basis === "inferred") return `The lawyer acts for: ${acting.party}, as taken from the sources; the request names no side.`
  if (acting.basis !== "none") return `The lawyer acts for: ${acting.party}.`
  if (readLeanPlan(directory)) return "The request does not say whom the lawyer acts for and no side was taken; recommend what is even-handed between the parties."
  const client = matterClient(directory)
  const inferred = intakeActingFor(directory)
  return client
    ? `The lawyer acts for: ${client} (the client recorded on the matter), unless the request, instructions or answers above name another party.`
    : inferred.basis !== "none"
    ? `The lawyer acts for: ${inferred.party}${inferred.basis === "inferred" ? ", as taken from the sources" : ""}, unless the request, instructions or answers above name another party.`
    : "Whom the lawyer acts for is not recorded yet; take it from the request, instructions or answers above where they say, otherwise recommend what is even-handed between the parties."
}

export function advisePrompt(directory: string, question: Question) {
  const confirming = question.options ? confirmationDocument(directory, question.id) : undefined
  if (confirming) return confirmationPrompt(directory, question, confirming)
  const row = question.pid ? readParagraphs(directory, question.document ?? FIRST_DOCUMENT).find(row => row.pid === question.pid) : undefined
  const paragraph = row && shownText(row)
  const updates = readLeanText(directory, "updated-notes.md")
  const provenanceLine = provenanceMapLine(matterSourceIDs(directory), sourceProvenance(directory))
  const basis = '"basis":[{"source_id":"request, an exact matter source name or an approved research URL","quote":"words copied exactly from that source","relevance":"why this passage applies to this question, party, document and conditions"}]'
  return [
    "Advise the lawyer on ONE open question raised while drafting. You do not answer it; the lawyer decides. Set out the options the lawyer could realistically choose and recommend one only where the request, the lawyer's answers, the matter sources or approved research support it, quoting each supporting passage exactly. Where an explicit instruction for this work (in the request, the lawyer's later instructions or answers, or a matter source written as instructions for it) already decides the point, recommend the option that follows it as far as it can lawfully and practically be followed, with its risks stated in the rationale; a differing source, caution or even-handedness is no reason to follow less of it. An instruction does not decide whether it can validly be carried out as given: a part that the sources or research show may be invalid, unauthorised or unenforceable as instructed cannot lawfully be followed, so it is not a risk to accept and state. Where some of it cannot be followed, recommend the option that keeps the rest of it, and include such an option when you set out the options. Where no instruction decides the point, advise from the sources as above. An option on how far the drafting qualifies or hedges names the points it applies to: a document's register is not a choice between blanket modes. Use targeted search-document or read-document lookups for the evidence the question needs; do not read the whole matter. Check that every passage applies to the affected document, represented party, version and conditions, including contrary passages and playbook exceptions; quotation existence alone is not support. State that connection in each basis passage's relevance. Reuse a matching topic from existing_topics when one exists. Give a short group topic in group (for example a neutral description of the matter point) so related questions can be considered together; never combine distinct decisions or change the question's meaning.",
    tagged("existing_topics", [...new Set(readQuestions(directory).flatMap(row => row.group ? [row.group] : []))].join("\n") || "None. Reuse the same short topic for questions about the same point."),
    tagged("question", `Kind: ${question.kind ?? "not stated"}\n${question.question}`),
    ...(question.context ? [tagged("context", question.context)] : []),
    ...(paragraph ? [tagged("affected_paragraph", paragraph)] : []),
    ...(question.options ? [tagged("options", question.options.map((option, index) => `${index}: ${option.text}`).join("\n"))] : []),
    tagged("request", readLeanText(directory, "request.md") ?? ""),
    ...(updates ? [tagged("user_updates", updates)] : []),
    tagged("answers", answeredQuestionsText(directory) || "None."),
    ...(readLeanText(directory, INSTRUCTIONS_FILE) ? [tagged("instructions", readLeanText(directory, INSTRUCTIONS_FILE)!)] : []),
    tagged("acting_for", actingForAdvice(directory)),
    ...[matterDatesBlock(directory)].filter(Boolean),
    tagged("legal_research", researchHandoff(directory)),
    `Matter source names (use these exact names with search-document and read-document): ${JSON.stringify(matterSourceIDs(directory))}`,
    ...(provenanceLine ? [provenanceLine] : []),
    question.options
      ? `Reply with one JSON object only, including "group":"short topic": {"recommendation":{"option":0,"rationale":"why this option",${basis}}}, where "option" is the index of one of the options above; or, when no option is supported, {"unavailable":{"reason":"why no option can be recommended"}}. Do not return options of your own.`
      : `Reply with one JSON object only, including "group":"short topic": {"options":[{"text":"one option the lawyer could choose"}],"recommendation":{"option":"the recommended option's text","rationale":"why this option",${basis},"decided_by":"i1"}}, with 2 to 4 distinct options and "option" the text of the one recommended, copied exactly from your options; or, when no option is supported, {"options":[],"unavailable":{"reason":"why no option can be recommended"}}, with no options or 2 to 4. Give "decided_by" only when one item of <instructions> expressly decides this question, as that item's id, and recommend the option that follows it, quoting the item's own words in basis; otherwise leave it out. An item that is only related, that would need inference to apply, or whose validity as given the question asks about, does not decide the question.`,
  ].join("\n\n")
}

/** The advice a reply gives `question`, or the issues for a retry. */
export function checkAdvice(directory: string, question: Question, reply: string, options: { requireRelevance?: boolean; request?: string } = {}): { issues: string[]; advice?: QuestionAdvice; group?: string } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object. ${parsed.error}`] }
  const { options: offered, recommendation, unavailable } = parsed.value
  const group = typeof parsed.value.group === "string" ? parsed.value.group.trim().slice(0, 120) : undefined
  if ((recommendation === undefined) === (unavailable === undefined)) return { issues: ['Return exactly one of "recommendation" or "unavailable".'] }
  const own = question.options ? { issues: offered === undefined ? [] : ['The question has its own options: leave out "options" and recommend one of them by its index.'] } : adviceOptions(directory, question, offered, recommendation !== undefined)
  const choices = question.options ?? own.options
  // The advisor's options are kept beside its advice; a question's own stay on the question.
  const kept = own.options?.length ? { options: own.options } : {}
  const issues = [...own.issues]
  if (recommendation !== undefined) {
    const row: Record<string, unknown> = record(recommendation) ? recommendation : {}
    // A question's own options are recommended by index; the advisor's own by
    // their text, which an index slip cannot move onto a neighbouring option.
    const chosen = question.options
      ? Number.isInteger(row.option) ? question.options[row.option as number] : undefined
      : typeof row.option === "string" ? own.options?.find(option => optionKey(option.text) === optionKey(row.option as string)) : undefined
    if (choices && !chosen) issues.push(question.options
      ? `"recommendation.option" must be the index, from 0, of one of the ${choices.length} options.`
      : `"recommendation.option" must be the text of one of your options, copied exactly: ${choices.map(option => JSON.stringify(option.text)).join(", ")}.`)
    if (typeof row.rationale !== "string" || !row.rationale.trim()) issues.push('"recommendation.rationale" must say why that option is recommended.')
    const basis = basisIssues(directory, row.basis, options.requireRelevance, options.request)
    issues.push(...basis.issues)
    const decided = row.decided_by === undefined || question.options ? undefined : row.decided_by
    if (decided !== undefined && !currentInstructions(directory).some(item => item.id === decided))
      issues.push(`"recommendation.decided_by" must be the id of one item listed in <instructions>, or left out.`)
    if (issues.length || !chosen) return { issues }
    return { issues: [], ...(group ? { group } : {}), advice: { ...kept, recommendation: { option: chosen.id, rationale: (row.rationale as string).trim(), basis: basis.rows, ...(typeof decided === "string" ? { decided_by: decided } : {}) } } }
  }
  const reason = record(unavailable) && typeof unavailable.reason === "string" ? unavailable.reason.trim() : ""
  if (!reason) issues.push('"unavailable.reason" must say why no option can be recommended.')
  if (issues.length) return { issues }
  return { issues: [], ...(group ? { group } : {}), advice: { ...kept, unavailable: { reason } } }
}

/** The advisor's own options with code-assigned ids: 2 to 4 with a
 * recommendation; none, or 2 to 4, without one. */
function adviceOptions(directory: string, question: Question, value: unknown, recommended: boolean): { issues: string[]; options?: QuestionOption[] } {
  if (value === undefined && !recommended) return { issues: [], options: [] }
  if (!Array.isArray(value) || value.some(row => !record(row) || typeof row.text !== "string" || !row.text.trim()))
    return { issues: ['"options" must be a list of {"text":"..."} with non-empty text.'] }
  const texts = value.map(row => (row as { text: string }).text.trim())
  const issues = [
    ...(recommended ? texts.length < 2 || texts.length > 4 ? ["Give 2 to 4 options with a recommendation."] : [] : texts.length === 1 || texts.length > 4 ? ['Give no options, or 2 to 4, with "unavailable".'] : []),
    ...(new Set(texts.map(optionKey)).size < texts.length ? ["Each option must differ from the others."] : []),
    ...(question.kind === "scope" ? [...sourceCarrierIssues(directory, question, texts), ...plannedCarrierIssues(directory, question, texts)] : []),
  ]
  return issues.length ? { issues } : { issues: [], options: texts.map((text, index) => ({ id: `o${index + 1}`, text })) }
}

/** The names an option may put an act or instruction in: the planned
 * documents' titles and files and, as the plan the question asks about may not
 * be saved yet, the titles the question shows. */
function plannedCarrierNames(directory: string, question: Question) {
  return [...(readLeanPlan(directory)?.documents ?? []).flatMap(document => [document.title, ...(document.file ? [document.file] : [])]), ...scopeQuestionTitles(question)]
}

/** A package-scope option may put an act in a planned document, a new one or
 * outside this request, never in an uploaded matter source: the package writes
 * only planned documents. An option naming a source by its file name is refused
 * unless it also names a planned document, whose title may hold a source's.
 * A source's title alone is ordinary prose ("an amendment to the Supply
 * Agreement") and passes. */
function sourceCarrierIssues(directory: string, question: Question, texts: string[]) {
  const planned = plannedCarrierNames(directory, question)
  const sources = matterSourceIDs(directory)
  return texts.flatMap(text => {
    const source = sources.find(name => holdsName(text, name))
    return source && !planned.some(name => holdsName(text, name, "iu"))
      ? [`Option ${JSON.stringify(text)} puts the act in ${JSON.stringify(source)}, an uploaded matter source the package does not write: offer a planned document, a new document, or "outside this request".`] : []
  })
}

/** A package-scope question asked about a saved plan (the gate's, or the plan
 * step's recheck) asks only which planned document carries the item, or
 * whether it is outside this request; the run cannot act on anything else. An
 * option that neither names a planned document (by title or file, in any case)
 * nor says the item is outside this request is refused. One that offers a new
 * document (scopeQuestionOffersNewDocument) is checked by sourceCarrierIssues
 * alone: a new document is named in words no list holds. */
function plannedCarrierIssues(directory: string, question: Question, texts: string[]) {
  if (scopeQuestionOffersNewDocument(question)) return []
  const planned = plannedCarrierNames(directory, question)
  const named = planned.length ? ` (${planned.map(name => JSON.stringify(name)).join(", ")})` : ""
  return texts.flatMap(text => planned.some(name => holdsName(text, name, "iu")) || OUTSIDE_REQUEST.test(normalize(text)) ? []
    : [`Option ${JSON.stringify(text)} neither names a planned document nor says the item is outside this request: the run can only carry it in a planned document or leave it out. Name the planned document that carries it, by its title or file${named}, or say it is outside this request.`])
}

/** Words that put an item outside this request: "outside" and, in the same
 * clause, "request" ("outside this request", "outside the scope of the request"). */
const OUTSIDE_REQUEST = /\boutside\b[^.;:]*\brequest\b/iu

/** Whether `text` holds `name` as whole words; a source's name as written, a planned document's in any case. */
function holdsName(text: string, name: string, flags = "u") {
  const key = normalize(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
  return !!key && new RegExp(`(?<![\\p{L}\\p{N}])${key}(?![\\p{L}\\p{N}])`, flags).test(normalize(text))
}

/** An option's text as a recommendation names it: ignoring case, spacing and
 * trailing punctuation. */
const optionKey = (text: string) => normalize(text).toLowerCase().replace(/[\s.,;:!?]+$/, "")

/** At least one supporting passage, each from the request, a matter source
 * or approved research, with non-blank words that source holds. */
export function basisIssues(directory: string, value: unknown, requireRelevance = false, request?: string): { issues: string[]; rows: { source_id: string; quote: string; relevance?: string }[] } {
  if (!Array.isArray(value) || !value.length) return { issues: ['"recommendation.basis" must list at least one {"source_id","quote"} passage supporting the recommendation.'], rows: [] }
  const sources = new Set([REQUEST_SOURCE, ...matterSourceIDs(directory), ...approvedResearch(directory).map(row => row.url)])
  const rows = value.map(row => record(row) && typeof row.source_id === "string" && typeof row.quote === "string" ? { source_id: row.source_id.trim(), quote: row.quote, ...(typeof row.relevance === "string" && row.relevance.trim() ? { relevance: row.relevance.trim() } : {}) } : undefined)
  const issues = rows.flatMap((row, index) => {
    const at = `recommendation.basis[${index}]`
    // An instruction from one of the lawyer's answers is cited as that answer (instructions.ts).
    if (!row || !(sources.has(row.source_id) || isAnswerSource(row.source_id))) return [`${at}.source_id must be "request", an exact matter source name or an approved research URL.`]
    if (!row.quote.trim()) return [`${at}.quote must be non-empty words copied exactly from that source.`]
    if (requireRelevance && !row.relevance) return [`${at}.relevance must explain how this passage supports this answer in the affected matter, document and conditions.`]
    const found = request !== undefined && row.source_id === REQUEST_SOURCE ? normalize(request).includes(normalize(row.quote)) : quoteFound(directory, row)
    return found ? [] : [quoteMissIssue(directory, at, row, "json")]
  })
  return { issues, rows: rows.filter(row => row !== undefined) }
}
