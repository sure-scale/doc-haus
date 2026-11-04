import { matterClient } from "../../lib/matter-client"
import { normalize } from "../../lib/quote-verify"
import { NOT_STATED, nameKey, type ActingFor, type ActingForBasis } from "./plan"
import { quoteFound, quoteMissIssue } from "./quotes"
import { questionKindIssues, readLeanJson, readQuestions, shortHash, type Question } from "./state"

type Checked<T> = { issues: string[]; value?: T }

/** The intake exchange asks material questions before research. They
 * use the ordinary durable question store and dialog; no draft is written. */
export function intakeQuestions(directory: string, reply: string, stage: Question["stage"] = "intake"): Checked<Question[]> | undefined {
  const text = reply.trim().replace(/^```(?:json)?\s*\n?/, "").replace(/\n?```$/, "").trim()
  if (!text.startsWith("{")) return
  const invalid = { issues: ['Reply with {"questions":[{"question":"...","kind":"fact|instruction|judgment","context":"why this blocks drafting"}]} or the readiness JSON when ready.'] }
  let value: unknown
  try { value = JSON.parse(text) } catch { return invalid }
  if (typeof value !== "object" || value === null) return invalid
  if (!("questions" in value)) return
  if (!Array.isArray(value.questions) || !value.questions.length) return invalid
  const answered = new Set(readQuestions(directory).filter(row => row.status === "answered").map(row => normalize(row.question)))
  const questions: Question[] = []
  for (const [index, row] of value.questions.entries()) {
    if (!row || typeof row.question !== "string" || !row.question.trim() || typeof row.context !== "string" || !row.context.trim()) return invalid
    const kind = questionKindIssues(row, `questions[${index}]`)
    if (kind.length) return { issues: kind }
    if (answered.has(normalize(row.question)))
      return { issues: ["That question already has an answer, from the lawyer or settled by the lawyer's instruction. Use the supplied answers and instructions and return readiness, or ask only a genuinely unresolved decision."] }
    const id = `q-${shortHash(normalize(row.question))}`
    if (!questions.some(question => question.id === id))
      questions.push({ id, stage, question: row.question.trim(), context: row.context.trim(), status: "open", kind: row.kind })
  }
  return { issues: [], value: questions }
}

/** Whom intake determined the lawyer acts for, as intake.json records it. */
export type IntakeActingFor = { party: string; basis: ActingForBasis; source_id?: string; quote?: string }
const BASES: ActingForBasis[] = ["stated", "matter", "inferred", "none"]
export const INTAKE_ACTING_FOR_SHAPE = '"acting_for":{"party":"the party, or empty for none","basis":"stated|matter|inferred|none","source_id":"source name, or request","quote":"exact words that show it"}'
const NO_SIDE: IntakeActingFor = { party: "", basis: "none" }

/** The readiness reply's acting_for record, checked: a stated or inferred party
 * carries one quotation found in its source, a matter party is the client the
 * matter records. `value` is the record to save, or no side when it fails. */
export function checkIntakeActingFor(directory: string, value: unknown): Checked<IntakeActingFor> & { value: IntakeActingFor } {
  const fail = (issue: string) => ({ issues: [issue], value: NO_SIDE })
  if (!record(value) || typeof value.party !== "string" || !BASES.includes(value.basis as ActingForBasis))
    return fail(`Add whom the lawyer acts for to the readiness JSON: ${INTAKE_ACTING_FOR_SHAPE}.`)
  const party = value.party.trim()
  if (value.basis === "none") return { issues: [], value: NO_SIDE }
  if (!party) return fail('acting_for names no party. Name the party, or use "basis":"none" when the document serves no one side.')
  if (value.basis === "matter") {
    const client = matterClient(directory)
    return client && nameKey(client) === nameKey(party) ? { issues: [], value: { party: client, basis: "matter" } }
      : fail(client ? `acting_for gives basis "matter", but the client recorded on the matter is ${client}.` : 'acting_for gives basis "matter", but the matter records no client.')
  }
  const quote = { source_id: typeof value.source_id === "string" ? value.source_id.trim() : "", quote: typeof value.quote === "string" ? value.quote.trim() : "" }
  if (!quote.source_id || !quote.quote) return fail(`acting_for names ${party} without the source_id and exact quote that show it.`)
  if (!quoteFound(directory, quote)) return fail(quoteMissIssue(directory, "acting_for", quote, "json"))
  return { issues: [], value: { party, basis: value.basis as ActingForBasis, ...quote } }
}

/** A saved intake.json that records whom the lawyer acts for. */
export function intakeRecordsActingFor(value: unknown) {
  return record(value) && record(value.acting_for) && BASES.includes(value.acting_for.basis as ActingForBasis)
}

/** The party intake took from the sources, with its evidence; otherwise no side.
 * A party the lawyer's words name, and the matter's client, reach the plan
 * through the outline and the matter record. */
export function intakeActingFor(directory: string): ActingFor {
  const value = (readLeanJson(directory, "intake.json") as { acting_for?: IntakeActingFor } | undefined)?.acting_for
  if (!record(value) || value.basis !== "inferred" || typeof value.party !== "string" || !value.party.trim()) return NOT_STATED
  return { party: value.party.trim(), basis: "inferred", ...(value.quote ? { reason: `"${value.quote}" [source: ${value.source_id}]` } : {}) }
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === "object" && value !== null && !Array.isArray(value) }
