import { matterClient } from "../../lib/matter-client"
import { normalize } from "../../lib/quote-verify"
import { NOT_STATED, nameKey, readLeanPlan, type ActingFor, type ActingForBasis, type DocumentDateBasis, type MatterDate, type MatterDates } from "./plan"
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

/** The party intake quoted, stated in the lawyer's words or an instructions
 * document or inferred from the sources, with its evidence and basis; otherwise
 * no side. The matter's client reaches the plan through the matter record. */
export function intakeActingFor(directory: string): ActingFor {
  const value = (readLeanJson(directory, "intake.json") as { acting_for?: IntakeActingFor } | undefined)?.acting_for
  if (!record(value) || !["stated", "inferred"].includes(value.basis) || typeof value.party !== "string" || !value.party.trim()) return NOT_STATED
  return { party: value.party.trim(), basis: value.basis, ...(value.quote ? { reason: `"${value.quote}" [source: ${value.source_id}]` } : {}) }
}

export const INTAKE_DATES_SHAPE = '"dates":{"document_date":{"date":"YYYY-MM-DD, or empty for the drafting day","basis":"request|sources|today","source_id":"source name, or request","quote":"exact words that show it"},"operative_dates":[{"label":"what happens on the date","date":"YYYY-MM-DD","source_id":"source name, or request","quote":"exact words that show it"}]}'
const DATE_BASES: DocumentDateBasis[] = ["request", "sources", "today"]
const DRAFTING_DAY: MatterDates = { document_date: { date: "", basis: "today" }, operative_dates: [] }

/** The readiness reply's dates, checked: each date is a calendar date with one
 * quotation found in its source. `value` is the record to save: an entry that
 * fails is left out, and a document date that fails becomes the drafting day. */
export function checkIntakeDates(directory: string, value: unknown): Checked<MatterDates> & { value: MatterDates } {
  if (!record(value) || !record(value.document_date) || !Array.isArray(value.operative_dates))
    return { issues: [`Add the dates the work turns on to the readiness JSON: ${INTAKE_DATES_SHAPE}.`], value: DRAFTING_DAY }
  const issues: string[] = []
  const quoted = (row: Record<string, unknown>, where: string) => {
    const quote = { source_id: typeof row.source_id === "string" ? row.source_id.trim() : "", quote: typeof row.quote === "string" ? row.quote.trim() : "" }
    const date = typeof row.date === "string" ? row.date.trim() : ""
    const issue = !calendarDate(date) ? `${where} needs its date as YYYY-MM-DD.`
      : !quote.source_id || !quote.quote ? `${where} gives ${date} without the source_id and exact quote that show it.`
      : !quoteFound(directory, quote) ? quoteMissIssue(directory, where, quote, "json") : ""
    if (issue) issues.push(issue)
    return issue ? undefined : { date, ...quote }
  }
  const given = value.document_date
  let document_date = DRAFTING_DAY.document_date
  if (!DATE_BASES.includes(given.basis as DocumentDateBasis)) issues.push(`dates.document_date needs a basis: ${DATE_BASES.join(", ")}.`)
  else if (given.basis !== "today") {
    const checked = quoted(given, "dates.document_date")
    if (checked) document_date = { ...checked, basis: given.basis as DocumentDateBasis }
  }
  const operative_dates: MatterDate[] = []
  for (const [index, row] of value.operative_dates.entries()) {
    const where = `dates.operative_dates[${index}]`
    if (!record(row) || typeof row.label !== "string" || !row.label.trim()) { issues.push(`${where} needs a label saying what happens on the date.`); continue }
    const checked = quoted(row, where)
    if (checked) operative_dates.push({ label: row.label.trim(), ...checked })
  }
  return { issues, value: { document_date, operative_dates } }
}

/** A saved intake.json that records the matter's dates. */
export function intakeRecordsDates(value: unknown) {
  return record(value) && record(value.dates) && record(value.dates.document_date) && Array.isArray(value.dates.operative_dates)
}

/** The matter's dates: as the plan froze them, else as intake recorded them;
 * undefined for a run that recorded none. */
export function matterDates(directory: string): MatterDates | undefined {
  const plan = readLeanPlan(directory)
  if (plan) return plan.dates
  const intake = readLeanJson(directory, "intake.json") as { dates?: MatterDates } | undefined
  return intakeRecordsDates(intake) ? intake!.dates : undefined
}

/** The dates the outline freezes into the plan: intake's, with the drafting day
 * fixed as the day the plan is made, so a run resumed later keeps its date. */
export function frozenDates(directory: string, now = new Date()): MatterDates | undefined {
  const dates = matterDates(directory)
  if (!dates || dates.document_date.date) return dates
  const day = [now.getFullYear(), now.getMonth() + 1, now.getDate()].map(part => String(part).padStart(2, "0")).join("-")
  return { ...dates, document_date: { date: day, basis: "today" } }
}

/** The matter's dates as data every stage after intake reads; empty for a run
 * that recorded none. */
export function matterDatesBlock(directory: string) {
  const dates = matterDates(directory)
  if (!dates) return ""
  const { date, basis, source_id, quote } = dates.document_date
  const dated = basis === "request" ? `${date}, the date the lawyer gives`
    : basis === "sources" ? `${date}, the date the sources fix for it ("${quote}" [source: ${source_id}]), as the lawyer gives none; the drafting notes disclose this once`
    : `${date ? `${date}, ` : ""}the drafting day, as neither the lawyer nor the sources give a date`
  return `<matter_dates>
${[
    `Date the document ${dated}.`,
    ...dates.operative_dates.length ? ["The dates the work turns on:", ...dates.operative_dates.map(row => `- ${row.label}: ${row.date} ("${row.quote}" [source: ${row.source_id}])`)] : [],
    "Compute every elapsed period, deadline and the law in force from these dates, not from today's date in the system prompt, which is the drafting day.",
  ].join("\n")}
</matter_dates>`
}

function calendarDate(value: string) {
  const time = Date.parse(`${value}T00:00:00Z`)
  return /^\d{4}-\d{2}-\d{2}$/.test(value) && !Number.isNaN(time) && new Date(time).toISOString().slice(0, 10) === value
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === "object" && value !== null && !Array.isArray(value) }
