import { boundPlaybookContext } from "../../lib/playbook-source"
import { nameWithProvenance, sourceProvenance } from "../../lib/source-provenance"
import { parseJsonObject } from "./change-checks"
import { excerpt, isAnswerSource, lawyerQuoteSource, matterSourceIDs, quoteFound, quoteMissIssue, REQUEST_SOURCE } from "./quotes"
import type { StageAcceptance } from "./stages"
import { answeredQuestionsText, readLeanJson, readLeanText, readQuestions, shortHash, writeLeanText, type Question } from "./state"

// The lawyer's instructions for this work as a concise list, made before any
// question is asked (the instructions step). One focused job lists them, each
// with the exact words it rests on, which code verifies. A matter document's
// directions are the lawyer's instructions only when the request, a later
// instruction or an answer says to follow that document, or the lawyer
// confirms it at intake (confirmationQuestions); until then they are listed as
// awaiting confirmation and remain evidence. Question stages and the advisor
// read the list; a question one of these instructions decides is settled by it
// rather than asked (settleByInstructions).

export const INSTRUCTIONS_FILE = "instructions.md"
const SAVED = "instructions.json"
const MAX_ITEMS = 80
const MAX_TEXT = 400

export type Instruction = { id: string; text: string; source_id: string; quote: string }
/** A matter document giving the drafter directions for this work, by whom
 * as it states, and the request's words saying to follow it, when they do. */
export type InstructionDocument = { source_id: string; from: string; quote: string; designated_by?: string }
type InstructionList = { items: Instruction[]; documents: InstructionDocument[] }
/** Whether a document's directions are followed: designated or confirmed,
 * awaiting the lawyer's confirmation, or read as evidence only. */
export type DocumentStanding = "follow" | "confirm" | "evidence"

const FOLLOW = { id: "o1", text: "Follow it as my instructions" }
const EVIDENCE = { id: "o2", text: "Treat it as evidence only" }

export function instructionsPrompt(directory: string) {
  const provenance = sourceProvenance(directory)
  const contract = '{"items":[{"text":"one concise instruction in plain words","source_id":"request, or the exact name of a document listed in documents","quote":"the exact words it rests on"}],"documents":[{"source_id":"exact matter source name","from":"who gives these directions and in what role, as the document states","quote":"exact words showing it gives directions for this work","designated_by":"exact words of the request, a later instruction or an answer saying to follow this document, or omit"}]}'
  return [
    boundPlaybookContext(directory),
    "List the lawyer's instructions for this work as a concise list. Each item is one requirement, decision, limit or preference the lawyer has given: what to prepare, for whom, with which terms, and how. Use plain words, keep every condition, figure and exception the instruction states, and never add, infer or resolve anything it does not say. Rest each item on the exact words it comes from.",
    "Sources of instructions: the original request, the lawyer's later instructions and the lawyer's answers below (source_id \"request\"), and any matter document that gives the drafter directions for this work, such as an instruction letter or email from the client or a supervising lawyer. Read the matter sources with search-document and read-document to find such documents. A document only describing a deal, a draft, a form, a term sheet or a counterparty's position gives no directions for this work: it is evidence, and none of its terms is an instruction. List each document that does give directions in documents, with who gives them and in what role, and list its directions as items with that document's name as source_id. When the request, a later instruction or an answer says to follow that document, quote those words in designated_by. Only words that single out that document designate it; words about the attached documents or materials as a whole do not, and such a document is listed without designated_by.",
    "An explicit later instruction supersedes an earlier one on that point: list only the current instruction. Do not list questions, suggestions or open points as instructions.",
    "Original request:", readLeanText(directory, "request.md") ?? "",
    "Later instructions in arrival order:", readLeanText(directory, "updated-notes.md") || "None.",
    "Lawyer's answers:", answeredQuestionsText(directory, { settled: false }) || "None.",
    "Matter source names (use these exact names with search-document and read-document):", ...matterSourceIDs(directory).map(name => `- ${nameWithProvenance(name, provenance)}`),
    `Reply with one JSON object only: ${contract}. documents is empty when no matter document gives directions for this work.`,
  ].filter(Boolean).join("\n\n")
}

/** The list a reply gives, with code-assigned ids, or the issues for a retry. */
export function checkInstructions(directory: string, value: unknown): { issues: string[]; list?: InstructionList } {
  if (!record(value) || !Array.isArray(value.items) || !Array.isArray(value.documents)) return { issues: ['Reply with {"items":[...],"documents":[...]}.'] }
  const names = new Set(matterSourceIDs(directory))
  const issues: string[] = []
  const documents: InstructionDocument[] = []
  for (const [index, row] of value.documents.entries()) {
    const at = `documents[${index}]`
    if (!record(row) || typeof row.source_id !== "string" || typeof row.from !== "string" || !row.from.trim() || typeof row.quote !== "string" || !row.quote.trim()) {
      issues.push(`${at} needs source_id, from and quote.`)
      continue
    }
    if (!names.has(row.source_id)) { issues.push(`${at}.source_id must be an exact matter source name: ${JSON.stringify([...names])}.`); continue }
    if (documents.some(document => document.source_id === row.source_id)) { issues.push(`${at} lists ${row.source_id} again; list each document once.`); continue }
    if (!quoteFound(directory, { source_id: row.source_id, quote: row.quote })) issues.push(quoteMissIssue(directory, `${at}.quote`, { source_id: row.source_id, quote: row.quote }, "json"))
    const designated = typeof row.designated_by === "string" && row.designated_by.trim() ? row.designated_by : undefined
    if (row.designated_by !== undefined && !designated) issues.push(`${at}.designated_by must be the exact words saying to follow the document, or omitted.`)
    if (designated && !quoteFound(directory, { source_id: REQUEST_SOURCE, quote: designated }))
      issues.push(quoteMissIssue(directory, `${at}.designated_by`, { source_id: REQUEST_SOURCE, quote: designated }, "json"))
    documents.push({ source_id: row.source_id, from: row.from.trim(), quote: row.quote, ...(designated ? { designated_by: designated } : {}) })
  }
  if (value.items.length > MAX_ITEMS) issues.push(`List at most ${MAX_ITEMS} items; combine closely related ones.`)
  const items: Instruction[] = []
  for (const [index, row] of value.items.entries()) {
    const at = `items[${index}]`
    if (!record(row) || typeof row.text !== "string" || !row.text.trim() || typeof row.source_id !== "string" || typeof row.quote !== "string" || !row.quote.trim()) {
      issues.push(`${at} needs text, source_id and quote.`)
      continue
    }
    if (row.text.trim().length > MAX_TEXT) issues.push(`${at}.text must be one concise instruction of at most ${MAX_TEXT} characters.`)
    const document = documents.some(document => document.source_id === row.source_id)
    if (row.source_id !== REQUEST_SOURCE && !isAnswerSource(row.source_id) && !document) {
      issues.push(`${at}.source_id must be "request" or a document listed in documents.`)
      continue
    }
    // The lawyer's words are cited where they are: the request (or an
    // update), or the one answer that holds them.
    const source_id = document ? row.source_id : lawyerQuoteSource(directory, row.quote)
    if (!source_id || !quoteFound(directory, { source_id, quote: row.quote }))
      issues.push(quoteMissIssue(directory, `${at}.quote`, { source_id: row.source_id, quote: row.quote }, "json"))
    items.push({ id: `i${items.length + 1}`, text: row.text.trim(), source_id: source_id ?? row.source_id, quote: row.quote })
  }
  return issues.length ? { issues } : { issues: [], list: { items, documents } }
}

export function acceptInstructionsReply(directory: string, reply: string): StageAcceptance {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object: ${parsed.error}.`], outputs: [], openItems: [] }
  const checked = checkInstructions(directory, parsed.value)
  if (!checked.list) return { issues: checked.issues, outputs: [], openItems: [] }
  writeLeanText(directory, SAVED, `${JSON.stringify(checked.list, null, 2)}\n`)
  writeInstructionsFile(directory)
  return { issues: [], outputs: [SAVED, INSTRUCTIONS_FILE], openItems: [] }
}

export function instructionsOutputValid(directory: string) {
  return readInstructionList(directory) !== undefined && readLeanText(directory, INSTRUCTIONS_FILE) === instructionsText(directory)
}

function readInstructionList(directory: string): InstructionList | undefined {
  const value = readLeanJson(directory, SAVED)
  return record(value) && Array.isArray(value.items) && Array.isArray(value.documents) ? value as InstructionList : undefined
}

/** The question confirming whether the drafting follows `source_id`'s directions. */
export function confirmationQuestionID(source_id: string) {
  return `q-${shortHash(`instructions\0${source_id}`)}`
}

/** Each listed document's standing: followed when the lawyer's words
 * designate it or the lawyer chose (or adopted) to follow it; evidence only
 * once the confirmation question has any other answer; otherwise awaiting
 * confirmation. */
export function documentStanding(directory: string, document: InstructionDocument): DocumentStanding {
  if (document.designated_by) return "follow"
  const asked = readQuestions(directory).find(question => question.id === confirmationQuestionID(document.source_id))
  if (asked?.status !== "answered") return "confirm"
  return asked.option_id === FOLLOW.id && asked.answer_kind !== "instruction" ? "follow" : "evidence"
}

/** The current list: items from the lawyer's own words and from followed
 * documents, then those awaiting confirmation; a document read as evidence
 * contributes none. */
export function currentInstructions(directory: string): (Instruction & { standing: DocumentStanding | "lawyer" })[] {
  const list = readInstructionList(directory)
  if (!list) return []
  const standing = new Map(list.documents.map(document => [document.source_id, documentStanding(directory, document)]))
  return list.items.flatMap(item => {
    const lawyer = item.source_id === REQUEST_SOURCE || (isAnswerSource(item.source_id) && !standing.has(item.source_id))
    const of = lawyer ? "lawyer" as const : standing.get(item.source_id) ?? "evidence"
    return of === "evidence" ? [] : [{ ...item, standing: of }]
  })
}

/** instructions.md: the concise list every question stage, the advisor and
 * the compaction handoff read. */
export function instructionsText(directory: string) {
  const list = readInstructionList(directory)
  if (!list) return ""
  const items = currentInstructions(directory)
  const line = (item: Instruction) => `- [${item.id}] ${item.text} (${item.source_id}: "${excerpt(item.quote, 160)}")`
  const followed = items.filter(item => item.standing !== "confirm")
  const awaiting = list.documents.filter(document => documentStanding(directory, document) === "confirm")
  return [
    "# The lawyer's instructions",
    "",
    "Each item is from the lawyer's words (request: the request or a later instruction; answer:<question id>: the lawyer's answer to that question), or a document the lawyer said to follow, quoted after it. An explicit later instruction supersedes an earlier one on that point.",
    "",
    ...(followed.length ? followed.map(line) : ["- None beyond the request."]),
    ...awaiting.flatMap(document => [
      "",
      `## Awaiting the lawyer's confirmation: ${document.source_id}`,
      "",
      `${document.source_id} gives directions for this work (${document.from}). Whether the drafting follows them is asked with intake's questions; until the lawyer answers, they are read as evidence.`,
      "",
      ...items.filter(item => item.source_id === document.source_id).map(line),
    ]),
  ].join("\n") + "\n"
}

export function writeInstructionsFile(directory: string) {
  const text = instructionsText(directory)
  if (text && readLeanText(directory, INSTRUCTIONS_FILE) !== text) writeLeanText(directory, INSTRUCTIONS_FILE, text)
}

/** One code-written question per document awaiting confirmation, asked with
 * intake's questions, never before them. */
export function confirmationQuestions(directory: string): Question[] {
  const list = readInstructionList(directory)
  if (!list) return []
  const items = currentInstructions(directory)
  return list.documents.filter(document => documentStanding(directory, document) === "confirm").map(document => ({
    id: confirmationQuestionID(document.source_id),
    stage: "intake",
    question: `Should the drafting follow ${document.source_id} as your instructions for this work?`,
    context: [
      `${document.source_id} gives directions for this work (${document.from}). Follow it when your client or your own firm wrote it for this work; treat it as evidence when someone else wrote it, it was written for other work, or your own instructions say otherwise. If you follow it, these directions decide the points they cover, and any other source that differs is flagged in the notes. If you treat it as evidence only, the points it would decide are asked.`,
      ...items.filter(item => item.source_id === document.source_id).map(item => `- ${item.text}`),
    ].join("\n"),
    status: "open",
    kind: "instruction",
    options: [FOLLOW, EVIDENCE],
  }))
}

/** `questions` as recorded once advised. A question its advice says one
 * current instruction decides (decided_by) is not asked: one the lawyer's
 * words or a followed document decide is recorded as settled by that
 * instruction, which the lawyer can change, and one a document awaiting
 * confirmation decides is left to that confirmation, asked with it; if the
 * document is read as evidence, intake asks the point again. A fact, a
 * confirmation and a code-written question are always asked. */
export function settleByInstructions(directory: string, questions: Question[]): Question[] {
  const items = new Map(currentInstructions(directory).map(item => [item.id, item]))
  const confirming = new Set(questions.filter(question => question.status === "open").map(question => question.id))
  return questions.flatMap(question => {
    const advice = question.advice
    const recommendation = advice && "recommendation" in advice ? advice.recommendation : undefined
    const item = recommendation?.decided_by ? items.get(recommendation.decided_by) : undefined
    if (!item || question.status !== "open" || question.options || (question.kind !== "instruction" && question.kind !== "judgment") || question.mixed) return [question]
    if (item.standing === "confirm") return confirming.has(confirmationQuestionID(item.source_id)) ? [] : [question]
    const option = advice?.options?.find(row => row.id === recommendation!.option)
    if (!option) return [question]
    return [{ ...question, status: "answered", answer: option.text, answer_kind: "instruction", option_id: option.id, decided_by: { text: item.text, source_id: item.source_id, quote: item.quote }, answered_at: new Date().toISOString() }]
  })
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === "object" && value !== null && !Array.isArray(value) }
