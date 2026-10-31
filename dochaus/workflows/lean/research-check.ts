import { parseJsonObject } from "./change-checks"
import { parallel } from "./checklist-flow"
import { record } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { retryPrompt, tagged } from "./prompts"
import { excerpt, quoteFound } from "./quotes"
import { approvedResearch } from "./research"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, stepSessionTitle, type LeanDeps } from "./stages"
import { questionOptions, readLeanJson, readQuestions, shortHash, writeLeanText, type LeanState, type Question } from "./state"

// A recommendation the lawyer delegated is adopted at intake, before legal
// research. Before the spine, one focused job per adopted recommendation asks
// whether the approved research shows it would not take effect as stated and,
// if so, which offered option the research supports; code verifies the quoted
// research. The spine then drafts to that option as the drafter's judgment and
// discloses the change. The question and its answer are left as recorded, so
// nothing is asked again and no earlier step re-runs.

const FILE = "research-checks.json"

export type ResearchCheck = { question_id: string; option: string; reason: string; evidence: { source_id: string; quote: string }[] }
type Saved = { basis: string; checks: ResearchCheck[] }

function adoptedQuestions(directory: string) {
  return readQuestions(directory).filter(question => question.status === "answered" && question.answer_kind === "adopted" && questionOptions(question).length > 1)
    .toSorted((a, b) => a.id.localeCompare(b.id))
}

// The adopted questions and the approved research the checks were made against.
function checksBasis(directory: string) {
  return shortHash(JSON.stringify({ questions: adoptedQuestions(directory).map(question => [question.id, question.option_id, questionOptions(question)]), research: approvedResearch(directory).map(row => [row.url, row.quote]) }))
}

/** Runs the checks unless they were made against the current adopted
 * questions and approved research. A job that fails its check or errors keeps
 * the adopted option. */
export async function checkAdoptedAgainstResearch(directory: string, state: LeanState, deps: LeanDeps) {
  const basis = checksBasis(directory)
  const saved = readLeanJson(directory, FILE) as Saved | undefined
  if (saved?.basis === basis) return
  const questions = approvedResearch(directory).length ? adoptedQuestions(directory) : []
  const checks: ResearchCheck[] = []
  await parallel(questions, async question => {
    const check = await checkOne(directory, state, deps, question)
    if (check) checks.push(check)
  })
  writeLeanText(directory, FILE, `${JSON.stringify({ basis, checks: checks.toSorted((a, b) => a.question_id.localeCompare(b.question_id)) } satisfies Saved, null, 2)}\n`)
}

async function checkOne(directory: string, state: LeanState, deps: LeanDeps, question: Question): Promise<ResearchCheck | undefined> {
  const agent = LEAN_STAGE_AGENTS.advise
  const at = { stage: "spine" as const, step: `research-check ${question.id}` }
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, "spine", `research check: ${question.id}`), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("advise", issues) : researchCheckPrompt(directory, question) })
      const answered = performance.now()
      const checked = checkResearchCheck(directory, question, reply)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (!issues.length) return checked.check
    }
  } catch {
    // A runtime error keeps the adopted option, as a failed check does.
  }
}

/** One adopted recommendation against the approved research only. */
export function researchCheckPrompt(directory: string, question: Question) {
  const options = questionOptions(question)
  return [
    "This is a lean drafting run. Decide one thing: whether a recommendation adopted before legal research still holds now that the research is approved. The lawyer delegated this question, so the recommended option was adopted as the drafter's judgment, not the lawyer's choice.",
    tagged("question", `${question.question}${question.context ? `\n${question.context}` : ""}`),
    tagged("options", options.map(option => `${option.id}: ${option.text}`).join("\n")),
    tagged("adopted", question.option_id ?? ""),
    tagged("approved_research", JSON.stringify(approvedResearch(directory).map(({ id: _, ...row }) => row), null, 2)),
    tagged("output_contract", [
      "- The adopted option holds unless the approved research shows it would not take effect as stated: the law the research states gives a step the option relies on to someone other than who takes it under the option, or sets a requirement for that step the option does not meet. Doubt, or research that does not reach the point, is not enough: the option holds.",
      '- When it holds, reply {"holds":true}.',
      `- When it does not, reply {"holds":false,"option":"the id of the offered option the research supports","reason":"what in the research the adopted option fails and how the chosen option meets it","evidence":[{"source_id":"the approved research url","quote":"exact words of that research quote"}]}.`,
      "- Answer from the inputs above; read-only tools are available if a passage must be checked. The run saves and checks the reply.",
    ].join("\n")),
  ].join("\n\n")
}

const SHAPE = 'Reply with the JSON object only: {"holds":true} or {"holds":false,"option":"...","reason":"...","evidence":[{"source_id":"...","quote":"..."}]}.'

/** Reads the reply to researchCheckPrompt: holds, or another offered option
 * with a reason and quotes the approved research holds. Every other reply is
 * retried. */
export function checkResearchCheck(directory: string, question: Question, reply: string): { issues: string[]; check?: ResearchCheck } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok || typeof parsed.value.holds !== "boolean") return { issues: [SHAPE] }
  const { holds, option, reason, evidence } = parsed.value
  if (holds) return { issues: [] }
  const offered = questionOptions(question).filter(row => row.id !== question.option_id).map(row => row.id)
  if (typeof option !== "string" || !offered.includes(option)) return { issues: [`"option" must be the id of an offered option other than the adopted one: ${JSON.stringify(offered)}. ${SHAPE}`] }
  if (typeof reason !== "string" || !reason.trim()) return { issues: [`"holds" is false, but "reason" is missing. ${SHAPE}`] }
  if (!Array.isArray(evidence) || !evidence.length) return { issues: [`"holds" is false, but "evidence" is empty: quote the approved research that shows it. ${SHAPE}`] }
  const urls = approvedResearch(directory).map(row => row.url)
  const quotes = evidence.filter(row => record(row) && typeof row.source_id === "string" && typeof row.quote === "string" && row.quote.trim()) as { source_id: string; quote: string }[]
  if (quotes.length !== evidence.length) return { issues: [`Each evidence entry needs source_id and quote. ${SHAPE}`] }
  const issues = quotes.flatMap(row => !urls.includes(row.source_id) ? [`Evidence source_id ${JSON.stringify(row.source_id)} is not an approved research url: ${JSON.stringify(urls)}.`]
    : quoteFound(directory, row) ? [] : [`The quotation ${JSON.stringify(excerpt(row.quote))} was not found in ${row.source_id}; copy its exact words.`])
  return issues.length ? { issues } : { issues: [], check: { question_id: question.id, option, reason: reason.trim(), evidence: quotes } }
}

/** For the spine prompt: each adopted recommendation the research showed
 * would not take effect as stated, and the option it supports instead. */
export function researchChecksText(directory: string) {
  const saved = readLeanJson(directory, FILE) as Saved | undefined
  if (saved?.basis !== checksBasis(directory)) return ""
  const questions = new Map(readQuestions(directory).map(question => [question.id, question]))
  return saved.checks.flatMap(check => {
    const question = questions.get(check.question_id)
    const options = question ? questionOptions(question) : []
    const adopted = options.find(row => row.id === question?.option_id)
    const chosen = options.find(row => row.id === check.option)
    if (!adopted || !chosen) return []
    return [`- [${check.question_id}] The recommendation adopted before legal research, "${adopted.text}", would not take effect as stated: ${check.reason} (${check.evidence.map(row => `"${row.quote}" ${row.source_id}`).join("; ")}). Draft to "${chosen.text}" instead, as the drafter's judgment. Record that as a decision entry with its research evidence, and say in it that the notes must disclose the change from the adopted recommendation and why.`]
  }).join("\n")
}
