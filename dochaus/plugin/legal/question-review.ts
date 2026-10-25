import { Tool } from "@opencode/schema/tool"
import { isProjectedUserAdmission } from "../../lib/engine-v2-projection"
import { parallel } from "../../workflows/lean/checklist-flow"
import { reviewProposedQuestion, sourceFactsFeedback, type QuestionReview, type QuestionReviewContext } from "../../workflows/lean/question-review"
import type { Question } from "../../workflows/lean/state"

type NativeQuestion = { question: string; header: string; options: { label: string; description: string }[]; multiple?: boolean }

/** Native question content contains the assistant's question and advice as
 * well as the user's answer. Only structured successful answers are authority;
 * question wording remains a separate, explicitly non-authoritative context. */
export function nativeQuestionHistory(messages: readonly unknown[]) {
  const human: string[] = []
  const context: { question: string; answers: string[] }[] = []
  for (const message of messages) {
    if (!record(message) || !Array.isArray(message.parts)) continue
    if (isProjectedUserAdmission(message)) {
      human.push(...message.parts.flatMap(part => record(part) && part.type === "text" && typeof part.text === "string" ? [part.text] : []))
      continue
    }
    for (const part of message.parts) {
      if (!record(part) || part.type !== "tool" || part.tool !== "question" || !record(part.state) || part.state.status !== "completed") continue
      const state = part.state
      if (!record(state.input) || !Array.isArray(state.input.questions) || !record(state.metadata) || !Array.isArray(state.metadata.answers)) continue
      for (const [index, question] of state.input.questions.entries()) {
        const answers: unknown = state.metadata.answers[index]
        if (!isNativeQuestion(question) || !Array.isArray(answers) || !answers.every(answer => typeof answer === "string")) continue
        const actual = (answers as string[]).filter(answer => answer.trim())
        human.push(...actual)
        context.push({ question: question.question, answers: actual })
      }
    }
  }
  return { request: human.slice(-8).join("\n\n").slice(-18_000),
    questionContext: JSON.stringify(context.slice(-8)).replaceAll("<", "\\u003c").replaceAll(">", "\\u003e").replaceAll("&", "\\u0026") }
}

/** Review a native question call before the native form exists. A source
 * finding is tool feedback to the asking agent, never a form submission. */
export async function reviewNativeQuestions(directory: string, args: Record<string, unknown>, context: QuestionReviewContext, reconsidered: Set<string>, reviews = new Map<string, QuestionReview>()) {
  if (!Array.isArray(args.questions) || !args.questions.every(isNativeQuestion)) return
  const questions = args.questions
  const feedback: Parameters<typeof sourceFactsFeedback>[0] = []
  const reviewed: (QuestionReview | undefined)[] = Array(questions.length)
  await parallel(questions.map((question, index) => ({ question, index })), async ({ question, index }) => {
    const candidate: Question = { id: `native-${index}`, stage: "intake", status: "open", question: question.question, context: question.header,
      ...(question.options.length ? { options: question.options.map((option, i) => ({ id: `o${i + 1}`, text: option.label.replace(/\s*\(Recommended\)\s*$/i, "") })) } : {}) }
    const cacheKey = JSON.stringify([question.question, question.header, question.options])
    const result = reviews.get(cacheKey) ?? (reconsidered.size
      ? { advice: { unavailable: { reason: "The evidence review for this turn is complete; this remaining question needs your decision." } } }
      : await reviewProposedQuestion(directory, candidate, { ...context, classifyFact: true }))
    reviews.set(cacheKey, result)
    reviewed[index] = result
    if (result.source && reconsidered.size === 0) {
      feedback.push({ question: question.question, source: result.source })
    }
  })
  if (feedback.length) {
    reconsidered.add("source-evidence-returned")
    throw new Tool.Error({ message: sourceFactsFeedback(feedback), metadata: { question_review: "source_evidence", human_answer: false } })
  }
  args.questions = questions.map((question, index) => {
    const review = reviewed[index]
    const advice = review?.advice
    const options = question.options.map(option => ({ ...option, label: option.label.replace(/\s*\(Recommended\)\s*$/i, "") }))
    if (!advice || !("recommendation" in advice)) return { ...question, options,
      question: `${question.question}\n\n${review?.source ? "Source evidence was returned for reconsideration, but this question still needs your decision." : `Independent review: ${advice && "unavailable" in advice ? advice.unavailable.reason : "No checked recommendation is available."}`}` }
    const recommendation = advice.recommendation
    const selected = (advice.options ?? options.map((option, i) => ({ id: `o${i + 1}`, text: option.label }))).find(option => option.id === recommendation.option)
    if (!selected) return { ...question, options }
    const offered = options.length ? options : (advice.options ?? []).map(option => ({ label: option.text, description: "" }))
    return { ...question, ...(review?.group ? { header: review.group.slice(0, 30) } : {}),
      options: offered.map(option => ({ ...option, label: option.label === selected.text ? `${option.label} (Recommended)` : option.label })),
      question: `${question.question}\n\nRecommendation: ${selected.text}. ${recommendation.rationale}\n${recommendation.basis.map(row => `${row.source_id}: “${row.quote}”${row.relevance ? ` — ${row.relevance}` : ""}`).join("\n")}` }
  })
}

function isNativeQuestion(value: unknown): value is NativeQuestion {
  return typeof value === "object" && value !== null && "question" in value && typeof value.question === "string" &&
    "header" in value && typeof value.header === "string" && "options" in value && Array.isArray(value.options) &&
    value.options.every(option => typeof option === "object" && option !== null && typeof option.label === "string" && typeof option.description === "string")
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
