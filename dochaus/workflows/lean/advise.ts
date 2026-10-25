import { isDeepStrictEqual } from "node:util"
import { normalize } from "../../lib/quote-verify"
import { parseJsonObject } from "./change-checks"
import { parallel } from "./checklist-flow"
import { readParagraphs, record, shownText } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { FIRST_DOCUMENT, planActingFor, stepDocument, stepStage, type StepID } from "./plan"
import { retryPrompt, tagged } from "./prompts"
import { REQUEST_SOURCE, matterSourceIDs, quoteFound, quoteMissIssue } from "./quotes"
import { nameWithProvenance, provenanceMapLine, sourceProvenance } from "../../lib/source-provenance"
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
      if (checked.advice) return { advice: checked.advice, ...(checked.group ? { group: checked.group } : {}) }
    }
  } catch {
    // A runtime error ends this question's advice; a pause or a settings
    // change is the runner's to act on once every advice job has settled.
  }
  return { advice: { unavailable: { reason: ADVISOR_FAILED, reason_code: "advisor_failed" } } }
}

export function advisePrompt(directory: string, question: Question) {
  const row = question.pid ? readParagraphs(directory, question.document ?? FIRST_DOCUMENT).find(row => row.pid === question.pid) : undefined
  const paragraph = row && shownText(row)
  const updates = readLeanText(directory, "updated-notes.md")
  const acting = planActingFor(directory)
  const provenanceLine = provenanceMapLine(matterSourceIDs(directory), sourceProvenance(directory))
  const basis = '"basis":[{"source_id":"request (the request, a later instruction or the lawyer\'s answers), an exact matter source name or an approved research URL","quote":"words copied exactly from that source","relevance":"why this passage applies to this question, party, document and conditions"}]'
  return [
    "Advise the lawyer on ONE open question raised while drafting. You do not answer it; the lawyer decides. Set out the options the lawyer could realistically choose and recommend one only where the request, the lawyer's answers, the matter sources or approved research support it, quoting each supporting passage exactly. Use targeted search-document or read-document lookups for the evidence the question needs; do not read the whole matter. Check that every passage applies to the affected document, represented party, version and conditions, including contrary passages and playbook exceptions; quotation existence alone is not support. State that connection in each basis passage's relevance. Reuse a matching topic from existing_topics when one exists. Give a short group topic in group (for example a neutral description of the matter point) so related questions can be considered together; never combine distinct decisions or change the question's meaning.",
    tagged("existing_topics", [...new Set(readQuestions(directory).flatMap(row => row.group ? [row.group] : []))].join("\n") || "None. Reuse the same short topic for questions about the same point."),
    tagged("question", `Kind: ${question.kind ?? "not stated"}\n${question.question}`),
    ...(question.context ? [tagged("context", question.context)] : []),
    ...(paragraph ? [tagged("affected_paragraph", paragraph)] : []),
    ...(question.options ? [tagged("options", question.options.map((option, index) => `${index}: ${option.text}`).join("\n"))] : []),
    tagged("request", readLeanText(directory, "request.md") ?? ""),
    ...(updates ? [tagged("user_updates", updates)] : []),
    tagged("answers", answeredQuestionsText(directory) || "None."),
    tagged("acting_for", acting.stated ? `The lawyer acts for: ${acting.party}.` : "The request does not say whom the lawyer acts for; recommend what is even-handed between the parties."),
    tagged("legal_research", researchHandoff(directory)),
    `Matter source names (use these exact names with search-document and read-document): ${JSON.stringify(matterSourceIDs(directory))}`,
    ...(provenanceLine ? [provenanceLine] : []),
    question.options
      ? `Reply with one JSON object only, including "group":"short topic": {"recommendation":{"option":0,"rationale":"why this option",${basis}}}, where "option" is the index of one of the options above; or, when no option is supported, {"unavailable":{"reason":"why no option can be recommended"}}. Do not return options of your own.`
      : `Reply with one JSON object only, including "group":"short topic": {"options":[{"text":"one option the lawyer could choose"}],"recommendation":{"option":0,"rationale":"why this option",${basis}}}, with 2 to 4 distinct options and "option" the 0-based index of the one recommended; or, when no option is supported, {"options":[],"unavailable":{"reason":"why no option can be recommended"}}, with no options or 2 to 4.`,
  ].join("\n\n")
}

/** The advice a reply gives `question`, or the issues for a retry. */
export function checkAdvice(directory: string, question: Question, reply: string, options: { requireRelevance?: boolean; request?: string } = {}): { issues: string[]; advice?: QuestionAdvice; group?: string } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object: ${parsed.error}.`] }
  const { options: offered, recommendation, unavailable } = parsed.value
  const group = typeof parsed.value.group === "string" ? parsed.value.group.trim().slice(0, 120) : undefined
  if ((recommendation === undefined) === (unavailable === undefined)) return { issues: ['Return exactly one of "recommendation" or "unavailable".'] }
  const own = question.options ? { issues: offered === undefined ? [] : ['The question has its own options: leave out "options" and recommend one of them by its index.'] } : adviceOptions(offered, recommendation !== undefined)
  const choices = question.options ?? own.options
  // The advisor's options are kept beside its advice; a question's own stay on the question.
  const kept = own.options?.length ? { options: own.options } : {}
  const issues = [...own.issues]
  if (recommendation !== undefined) {
    const row: Record<string, unknown> = record(recommendation) ? recommendation : {}
    if (choices && !(Number.isInteger(row.option) && (row.option as number) >= 0 && (row.option as number) < choices.length))
      issues.push(`"recommendation.option" must be the index, from 0, of one of the ${choices.length} options.`)
    if (typeof row.rationale !== "string" || !row.rationale.trim()) issues.push('"recommendation.rationale" must say why that option is recommended.')
    const basis = basisIssues(directory, row.basis, options.requireRelevance, options.request)
    issues.push(...basis.issues)
    if (issues.length || !choices) return { issues }
    return { issues: [], ...(group ? { group } : {}), advice: { ...kept, recommendation: { option: choices[row.option as number]!.id, rationale: (row.rationale as string).trim(), basis: basis.rows } } }
  }
  const reason = record(unavailable) && typeof unavailable.reason === "string" ? unavailable.reason.trim() : ""
  if (!reason) issues.push('"unavailable.reason" must say why no option can be recommended.')
  if (issues.length) return { issues }
  return { issues: [], ...(group ? { group } : {}), advice: { ...kept, unavailable: { reason } } }
}

/** The advisor's own options with code-assigned ids: 2 to 4 with a
 * recommendation; none, or 2 to 4, without one. */
function adviceOptions(value: unknown, recommended: boolean): { issues: string[]; options?: QuestionOption[] } {
  if (value === undefined && !recommended) return { issues: [], options: [] }
  if (!Array.isArray(value) || value.some(row => !record(row) || typeof row.text !== "string" || !row.text.trim()))
    return { issues: ['"options" must be a list of {"text":"..."} with non-empty text.'] }
  const texts = value.map(row => (row as { text: string }).text.trim())
  const issues = [
    ...(recommended ? texts.length < 2 || texts.length > 4 ? ["Give 2 to 4 options with a recommendation."] : [] : texts.length === 1 || texts.length > 4 ? ['Give no options, or 2 to 4, with "unavailable".'] : []),
    ...(new Set(texts.map(text => normalize(text).toLowerCase())).size < texts.length ? ["Each option must differ from the others."] : []),
  ]
  return issues.length ? { issues } : { issues: [], options: texts.map((text, index) => ({ id: `o${index + 1}`, text })) }
}

/** At least one supporting passage, each from the request, a matter source
 * or approved research, with non-blank words that source holds. */
export function basisIssues(directory: string, value: unknown, requireRelevance = false, request?: string): { issues: string[]; rows: { source_id: string; quote: string; relevance?: string }[] } {
  if (!Array.isArray(value) || !value.length) return { issues: ['"recommendation.basis" must list at least one {"source_id","quote"} passage supporting the recommendation.'], rows: [] }
  const sources = new Set([REQUEST_SOURCE, ...matterSourceIDs(directory), ...approvedResearch(directory).map(row => row.url)])
  const rows = value.map(row => record(row) && typeof row.source_id === "string" && typeof row.quote === "string" ? { source_id: row.source_id.trim(), quote: row.quote, ...(typeof row.relevance === "string" && row.relevance.trim() ? { relevance: row.relevance.trim() } : {}) } : undefined)
  const issues = rows.flatMap((row, index) => {
    const at = `recommendation.basis[${index}]`
    if (!row || !sources.has(row.source_id)) return [`${at}.source_id must be "request", an exact matter source name or an approved research URL.`]
    if (!row.quote.trim()) return [`${at}.quote must be non-empty words copied exactly from that source.`]
    if (requireRelevance && !row.relevance) return [`${at}.relevance must explain how this passage supports this answer in the affected matter, document and conditions.`]
    const found = request !== undefined && row.source_id === REQUEST_SOURCE ? normalize(request).includes(normalize(row.quote)) : quoteFound(directory, row)
    return found ? [] : [quoteMissIssue(directory, at, row, "json")]
  })
  return { issues, rows: rows.filter(row => row !== undefined) }
}
