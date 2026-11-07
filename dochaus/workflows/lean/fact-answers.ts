import { matterSourceBasis, openItem, quoteFound } from "./quotes"
import { factDisposition, questionOptions, readLeanText, readQuestions, shortHash, type FactStatus, type Question } from "./state"

export { factDisposition } from "./state"

// A changed refresh contract must reconsider holds cached by the earlier job,
// while unchanged Continue requests keep the new contract's bounded result.
const FACT_CONTEXT_VERSION = 2

/** A waiting decision is reconsidered only when its relevant inputs change. */
export function factContext(directory: string) {
  return shortHash(JSON.stringify([FACT_CONTEXT_VERSION, readLeanText(directory, "request.md") ?? "", readLeanText(directory, "updated-notes.md") ?? "", matterSourceBasis(directory)]), 40)
}

/** Only source-derived resolutions need revalidation; the lawyer's original
 * factual assertion remains their assertion, never independent verification. */
export function sourceResolvedFact(question: Question) {
  return factDisposition(question) === "supplied" && !!question.fact_resolution?.basis.some(row => row.source_id !== "request" && row.source_id !== "updated-notes")
}

export function missingFactEvidence(directory: string, question: Question) {
  return sourceResolvedFact(question) && question.fact_resolution!.basis.some(row => row.source_id !== "request" && row.source_id !== "updated-notes" && !quoteFound(directory, row))
}

export function originalFactHold(directory: string, question: Question): FactStatus {
  const action = questionOptions(question).find(option => option.id === question.option_id)?.fact_action
  if (action === "wait") return "wait"
  return action === "provisional" && !laterFactInstructions(directory, question) ? "provisional" : "unknown"
}

/** A removed supporting passage cannot leave a cached source fact supplied. */
export function currentFactDisposition(directory: string, question: Question) {
  return missingFactEvidence(directory, question) ? originalFactHold(directory, question) : factDisposition(question)
}

/** The existing request chronology establishes which human updates can change
 * an answer; same-millisecond requests keep their queue's full id ordering. */
export function laterFactInstructions(directory: string, question: Question) {
  const answered = question.answered_at ? Date.parse(question.answered_at) : NaN
  if (!question.answer_request_id && !Number.isFinite(answered)) return ""
  const text = readLeanText(directory, "updated-notes.md") ?? ""
  return [...text.matchAll(/<!-- update:((\d{15})-[^>\r\n]*) -->\s*([\s\S]*?)(?=<!-- update:|$)/g)]
    .filter(match => question.answer_request_id ? match[1]! > question.answer_request_id : Number(match[2]) > answered)
    .map(match => match[3]!.trim()).join("\n\n")
}

/** Missing facts can be supplied in the lawyer's own words. These choices
 * decide only whether the run may draft while that fact remains missing. */
export function withFactOptions(question: Question): Question {
  if (question.kind !== "fact" || question.stage !== "intake" || question.options || question.advice?.options?.length) return question
  return { ...question, options: [
    { id: "o1", text: "Continue with a marked review draft using the documents provided; leave this fact unverified, show a placeholder wherever it is needed, and disclose the outstanding check in the drafting notes. This does not authorise sending, signing or serving the document.", fact_action: "provisional" },
    { id: "o2", text: "Wait for me to supply or confirm this fact before drafting further. Keep the evidence request open; do not ask again for a status update.", fact_action: "wait" },
  ] }
}

/** An external task or an unreadable reply cannot grant leave to draft. */
export function waitingFactQuestions(directory: string) {
  return readQuestions(directory).filter(question => {
    const status = currentFactDisposition(directory, question)
    return question.stage === "intake" && question.kind === "fact" && question.status === "answered" &&
      (status === undefined || status === "wait" || status === "unknown")
  })
}

/** These are ordinary uncertainty items in the existing open-item store, not
 * answers. The gate must still see the missing fact after a drafting choice. */
export function factOpenItems(directory: string) {
  return readQuestions(directory).flatMap(question => {
    if (question.stage !== "intake" || question.kind !== "fact" || question.status !== "answered") return []
    const status = currentFactDisposition(directory, question) ?? "unknown"
    if (status === "supplied") return []
    return [openItem(question.stage, "question", `Fact remains unverified: ${question.question}`, {
      question_id: question.id,
      placeholder: `[open item: ${question.question}]`,
      ...(question.document ? { document: question.document } : {}),
      ...(question.pid ? { pid: question.pid } : {}),
    })]
  })
}
