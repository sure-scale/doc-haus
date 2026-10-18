import { normalize } from "../../lib/quote-verify"
import { handedBack, readQuestions, shortHash, type Question } from "./state"

type Checked<T> = { issues: string[]; value?: T }

/** The intake exchange asks material questions before research. They
 * use the ordinary durable question store and dialog; no draft is written. */
export function intakeQuestions(directory: string, reply: string, stage: Question["stage"] = "intake"): Checked<Question[]> | undefined {
  const text = reply.trim().replace(/^```(?:json)?\s*\n?/, "").replace(/\n?```$/, "").trim()
  if (!text.startsWith("{")) return
  const invalid = { issues: ['Reply with {"questions":[{"question":"...","context":"why this blocks drafting"}]} or the readiness JSON when ready.'] }
  let value: unknown
  try { value = JSON.parse(text) } catch { return invalid }
  if (typeof value !== "object" || value === null) return invalid
  if (!("questions" in value)) return
  if (!Array.isArray(value.questions) || !value.questions.length) return invalid
  const settled = readQuestions(directory).filter(row => row.status === "answered")
  const answered = new Set(settled.filter(row => !handedBack(row)).map(row => normalize(row.question)))
  const left = new Set(settled.filter(handedBack).map(row => normalize(row.question)))
  const questions: Question[] = []
  for (const row of value.questions) {
    if (!row || typeof row.question !== "string" || !row.question.trim() || typeof row.context !== "string" || !row.context.trim()) return invalid
    if (answered.has(normalize(row.question)))
      return { issues: ["That question already has a lawyer's answer. Use the supplied answers and return readiness, or ask only a genuinely unresolved decision."] }
    if (left.has(normalize(row.question)))
      return { issues: ["The lawyer left that question to the drafting to decide. Make the choice in drafting and return readiness, or ask only a genuinely unresolved decision."] }
    const id = `q-${shortHash(normalize(row.question))}`
    if (!questions.some(question => question.id === id))
      questions.push({ id, stage, question: row.question.trim(), context: row.context.trim(), status: "open" })
  }
  return { issues: [], value: questions }
}
