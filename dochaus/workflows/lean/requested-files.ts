import { TOOL_SYNTAX_ISSUE, printedToolCall, record } from "./checks"
import { combinedDelivery, deliveryArtifactEvidence, deliveryStem, memoNotesFile, plannedRegistration, requestedMemo } from "./deliverables"
import { deliveryNameQuoted } from "./delivery-selection"
import { normalizedHeading, plainFileName, readLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { tagged } from "./prompts"
import { REQUEST_SOURCE, excerpt, instructionFound, matterSourceIDs, openItem, quoteMissIssue } from "./quotes"
import { answeredQuestionsText, readLeanJson, readLeanText, readQuestions, shortHash, writeLeanText, type OpenItem, type Question } from "./state"

// Requested files: which files the lawyer asks to receive, and which planned
// documents and the memo go in each. One focused job reads the lawyer's words
// (request, updates and answers) beside the planned documents and memo;
// parseRequestedFiles grounds each file in the lawyer's exact words, its name in
// those words and its contents in the plan. The record is the lawyer's
// packaging, never the plan's. Code does not rewrite the plan's delivery shape:
// requestedFilesIssues gives the outline revision feedback in the outline's own
// terms, and requestedFilesQuestion asks the lawyer about a mismatch the
// revisions leave. A file name is no planning judgement, so withRequestedNames
// gives grounded names to the plan. A shape the run cannot build (several
// documents in one file, the memo after its document or with a redline) is no
// revision issue; like any file the final plan does not deliver as asked, it
// reaches the lawyer as an open item derived from the record
// (requestedFilesItems), never stored.

export const LEAN_REQUESTED_FILES_SCHEMA = "dochaus.lean.requested-files/1"

/** One file the lawyer asks to receive: the planned ids, and "memo" for the
 * memo, it holds in the order the lawyer places them; its name as the lawyer
 * gives it, or null; and the lawyer's exact words that ask for it. */
export type RequestedFile = { holds: string[]; name: string | null; quote: string }
/** `stated` is false, with no files, when the lawyer's words do not say how the
 * planned documents and memo are delivered as files. */
export type RequestedFiles = { stated: boolean; files: RequestedFile[] }
/** The record of the plan whose basis is `basis`. */
export type LeanRequestedFiles = { schema: typeof LEAN_REQUESTED_FILES_SCHEMA; basis: string; stated: boolean; files: RequestedFile[] }

type Checked<T> = { issues: string[]; value?: T }

const MEMO = "memo"
const RECORD_FILE = "requested-files.json"

const filesShape = (memo: boolean) =>
  `Reply with the JSON object only: {"stated":true,"files":[{"holds":["${memo ? "planned id or memo" : "planned id"}"],"name":"file name or null","quote":"..."}]}.`
const STATED_WITH_FILES = '"stated" is false, but "files" lists files. Give "stated": true with the files the lawyer\'s words ask for, or "stated": false with "files": [].'
const STATED_WITHOUT_FILES = '"stated" is true, but "files" is empty. List each file the lawyer\'s words ask for, or reply {"stated":false,"files":[]} when they do not say.'

/** The job's one question, over the lawyer's words and the planned documents and memo. */
export function requestedFilesPrompt(directory: string, documents: PlannedDocument[], memo: string | undefined) {
  const updates = readLeanText(directory, "updated-notes.md")?.trim()
  const answers = answeredQuestionsText(directory)
  const contents = memo ? "the planned documents and the memo" : "the planned documents"
  return [
    `This is a lean drafting run. Read the lawyer's words for one question: which files the lawyer asks to receive, and which of ${contents} go in each.`,
    tagged("request", readLeanText(directory, "request.md") ?? ""),
    updates ? tagged("user_updates", `Later instructions from the lawyer, in arrival order. An explicit later change supersedes the earlier instruction on that point.\n\n${updates}`) : "",
    answers ? tagged("answers", `The lawyer's answers; they remain controlling:\n${answers}`) : "",
    tagged("planned_documents", "The documents the run drafts, by id. Their titles, files and deliverables are the plan's own choices, not the lawyer's words.\n" +
      JSON.stringify(documents.map(filesView), null, 2)),
    memo ? tagged("planned_memo", `The run also writes the memo the lawyer asks for, planned as "${memo}". In "holds", write it as "memo".`) : "",
    tagged("output_contract", [
      "- Answer from the request, the lawyer's updates and answers only.",
      `- When those words say which files the lawyer wants to receive, what goes in each file or what a file is named, reply "stated": true and list each file they ask for. When they do not, reply {"stated":false,"files":[]}.`,
      `- For each file give "holds": the ids of the planned documents${memo ? ' and "memo" for the memo' : ""} that the lawyer's words put in that file` +
      `${memo ? ', in the order the words place them in the file, with "memo" first when the words set no order' : ""}; ` +
      "\"name\": the file name exactly as the lawyer writes it, ending in .docx, adding only .docx to a name the lawyer gives without it, or null when the lawyer gives no name; " +
      "\"quote\": the lawyer's exact words that ask for this file, copied verbatim, including the words that give its name.",
      `- Put each planned document${memo ? " and the memo" : ""} in one file at most. List only the files the lawyer's words ask for, and in each only what those words put in it.`,
      `- Named files are all the files the lawyer asks to receive: when the lawyer's words name the files to deliver, put each planned document${memo ? " and the memo" : ""} ` +
      "in the named file the lawyer's words attach it to; leave one out of every file only when the words leave it out of every named file.",
      `- ${filesShape(Boolean(memo))}`,
      "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks the reply.",
    ].join("\n")),
  ].filter(Boolean).join("\n\n")
}

const filesView = (document: PlannedDocument) =>
  ({ id: document.id, title: document.title, ...(document.file ? { file: document.file } : {}), deliverable: document.deliverable ?? "body" })

/** Reads the job's reply and grounds each file: its quote in the lawyer's own
 * words (instructionFound), a name as a plain Word file name its quote gives
 * (deliveryNameQuoted), its holds as planned ids or "memo" when a memo is
 * planned, each in one file at most. Every defect is retry feedback; the files
 * are returned only when none is left. On the last attempt (`final`) a refused
 * name becomes null and a file whose quote is refused is dropped, so nothing
 * ungrounded is kept; with every file dropped the lawyer's words are taken as
 * not stating the files. Any other defect still fails the last attempt. */
export function parseRequestedFiles(reply: string, directory: string, documents: PlannedDocument[], memo: string | undefined, { final = false } = {}): Checked<RequestedFiles> {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const shape = filesShape(memo !== undefined)
  let value: unknown
  try { value = JSON.parse(reply.trim().replace(/^```(?:json)?\s*/, "").replace(/\s*```$/, "")) } catch {}
  if (!record(value) || typeof value.stated !== "boolean" || !Array.isArray(value.files)) return { issues: [shape] }
  if (!value.stated) return value.files.length ? { issues: [STATED_WITH_FILES] } : { issues: [], value: { stated: false, files: [] } }
  if (!value.files.length) return { issues: [STATED_WITHOUT_FILES] }
  const issues: string[] = []
  const rows: { index: number; file: RequestedFile }[] = []
  value.files.forEach((row: unknown, index) => {
    const where = `files[${index}]`
    if (!record(row) || !Array.isArray(row.holds) || !row.holds.length || !row.holds.every(id => typeof id === "string" && id.trim()) ||
      !(row.name == null || typeof row.name === "string") || typeof row.quote !== "string" || !row.quote.trim())
      return void issues.push(`${where} needs "holds" as a nonempty list of planned ids${memo !== undefined ? ' or "memo"' : ""}, "name" as a file name or null, and a nonempty "quote" string. ${shape}`)
    const file: RequestedFile = { holds: [...new Set((row.holds as string[]).map(id => id.trim()))], name: typeof row.name === "string" && row.name.trim() ? row.name.trim() : null, quote: row.quote.trim() }
    const grounded = instructionFound(directory, file.quote)
    if (!grounded) {
      if (final) return
      issues.push(quoteMissIssue(directory, where, { quote: file.quote, source_id: REQUEST_SOURCE }, "json"))
    }
    // A name is held to its quote once the quote is the lawyer's words.
    const refused = !grounded || file.name === null ? undefined
      : !/\.docx$/i.test(file.name) || !plainFileName(file.name)
        ? `${where} gives the name "${file.name}", which is not a plain Word file name. Give the file name as the lawyer writes it, ending in .docx, without a folder or Markdown, or null when the lawyer gives none.`
      : !deliveryNameQuoted(file.quote, file.name)
        ? `${where} gives the name "${file.name}", but its quote does not give that name. Quote the lawyer's words that name this file, and give the name exactly as they write it, adding only .docx to a name given without it, or null when the lawyer gives none.`
      : undefined
    if (refused && final) file.name = null
    else if (refused) issues.push(refused)
    rows.push({ index, file })
  })
  const ids = documents.map(document => document.id as string)
  const places = new Map<string, number[]>()
  for (const { index, file } of rows) for (const id of file.holds) {
    if (id === MEMO && memo === undefined) issues.push(`files[${index}] holds "memo", but no memo is planned. Use the planned ids ${ids.join(", ")}.`)
    else if (id !== MEMO && !ids.includes(id)) issues.push(`files[${index}] holds "${id}", which is not planned. Use the planned ids ${ids.join(", ")}${memo !== undefined ? ' or "memo"' : ""}.`)
    places.set(id, [...places.get(id) ?? [], index])
  }
  for (const [id, at] of places) if (at.length > 1)
    issues.push(`"${id}" is in ${serialList(at.map(index => `files[${index}]`))}. Put each planned document${memo !== undefined ? " and the memo" : ""} in one file at most, as the lawyer's words place it.`)
  if (issues.length) return { issues }
  return { issues, value: rows.length ? { stated: true, files: rows.map(row => row.file) } : { stated: false, files: [] } }
}

/** What a recorded file asks of the plan: the memo with one document, the memo
 * first; one document alone; or the memo alone. Undefined for a shape no plan
 * builds: several documents in one file, or the memo after its document. */
function fileShape(file: RequestedFile): { document?: DocID; memo: boolean } | undefined {
  const documents = file.holds.filter(id => id !== MEMO) as DocID[]
  const memo = file.holds.includes(MEMO)
  if (documents.length > 1 || (memo && documents.length && file.holds[0] !== MEMO)) return
  return { document: documents[0], memo }
}

/** A recorded file against an outline's documents and memo (`memo`: whether one
 * is planned): unsupported when the plan cannot build it (fileShape, a
 * document not planned, the memo not planned or joined to a redline);
 * otherwise, when the plan's deliverables do not deliver it as asked, the
 * outline revision `issue` and how the plan delivers it instead (`planned`). */
function fileCheck(file: RequestedFile, documents: PlannedDocument[], memo: boolean): { unsupported: true } | { unsupported: false; issue?: string; planned?: string } {
  const shape = fileShape(file)
  const target = shape?.document && documents.find(document => document.id === shape.document)
  if (!shape || (shape.document && !target) || (shape.memo && !memo) || (shape.memo && target && target.deliverable === "redline")) return { unsupported: true }
  const words = `The lawyer's words "${excerpt(file.quote)}"`
  const combined = documents.filter(document => document.deliverable === "combined" && document !== target)
  const body = (document: PlannedDocument) => `"Deliverable: body" under "${document.title}"`
  if (target && shape.memo) {
    const changes = [...(target.deliverable !== "combined" ? [`"Deliverable: combined" under "${target.title}"`] : []), ...combined.map(body)]
    return !changes.length ? { unsupported: false } : {
      unsupported: false,
      issue: `${words} ask for the memo and "${target.title}" in one file. Write ${serialList(changes)}.`,
      planned: target.deliverable === "combined" ? `the memo in one file with ${titles(combined)} as well`
        : combined.length ? `the memo in one file with ${titles(combined)}, and "${target.title}" as a file of its own` : `the memo and "${target.title}" as separate files`,
    }
  }
  if (target) return target.deliverable !== "combined" ? { unsupported: false } : {
    unsupported: false,
    issue: `${words} ask for "${target.title}" in a file of its own, without the memo. Write "Deliverable: body" under "${target.title}" in place of "Deliverable: combined".`,
    planned: `"${target.title}" in one file with the memo`,
  }
  return !combined.length ? { unsupported: false } : {
    unsupported: false,
    issue: `${words} ask for the memo in a file of its own. Write ${serialList(combined.map(body))} in place of "Deliverable: combined".`,
    planned: `the memo in one file with ${titles(combined)}`,
  }
}

/** Outline revision feedback holding the outline's deliverables to the
 * recorded files, in the outline's own terms, quoting the lawyer's words and
 * naming documents by title: the memo with one document needs that document
 * combined and no other; one document alone, or the memo alone, needs no
 * document combined with the memo. A file the plan cannot build gives no issue,
 * since no revision builds it; its index is in `unsupported`, for disclosure.
 * Nothing when the lawyer's words do not state the files. */
export function requestedFilesIssues(documents: PlannedDocument[], memo: string | undefined, requested: RequestedFiles) {
  const issues: string[] = []
  const unsupported: number[] = []
  if (requested.stated) requested.files.forEach((file, index) => {
    const check = fileCheck(file, documents, memo !== undefined)
    if (check.unsupported) unsupported.push(index)
    else if (check.issue) issues.push(check.issue)
  })
  return { issues, unsupported }
}

/** The plan's file names with each grounded name the record gives: a file of
 * one document, alone or after the memo, names that document's file; a file of
 * the memo alone names the notes file. A name is skipped, and noted, when
 * another planned file would have the same name (compared as the plan parser
 * compares file names); every name asked for twice is skipped. A file the plan
 * cannot build (fileShape) names nothing. Without a name of its own, a notes
 * file still named after its document (memoNotesFile) is named again after the
 * named documents. Documents are identified by position (withDocumentIDs). */
export function withRequestedNames<T extends { title: string; file?: string; deliverable?: string }>(
  documents: T[], memo: { name: string; file: string } | undefined, requested: RequestedFiles,
): { documents: T[]; notes_file: string | undefined; applied: string[] } {
  type Wanted = { target: number | typeof MEMO; name: string }
  const wanted = requested.stated ? requested.files.flatMap((file): Wanted[] => {
    const shape = fileShape(file)
    if (!shape || file.name === null) return []
    if (!shape.document) return memo ? [{ target: MEMO, name: file.name }] : []
    const index = documents.findIndex((_, at) => `d${at + 1}` === shape.document)
    return index < 0 ? [] : [{ target: index, name: file.name }]
  }) : []
  const names = [...documents.flatMap((document, index) => wanted.find(row => row.target === index)?.name ?? document.file ?? []), ...wanted.filter(row => row.target === MEMO).map(row => row.name)]
  const repeated = (name: string) => names.filter(other => normalizedHeading(other) === normalizedHeading(name)).length > 1
  const applied: string[] = []
  const label = (target: number | typeof MEMO) => target === MEMO ? "the memo" : `"${documents[target]!.title}"`
  const given = wanted.filter(row => {
    if (repeated(row.name)) applied.push(`"${row.name}" not given to ${label(row.target)}: another planned file has that name`)
    else if (row.name !== (row.target === MEMO ? memo?.file : documents[row.target]!.file)) {
      applied.push(`${label(row.target)}: file "${row.name}"`)
      return true
    }
    return false
  })
  const named = documents.map((document, index) => {
    const name = given.find(row => row.target === index)?.name
    return name ? { ...document, file: name } : document
  })
  const notes = given.find(row => row.target === MEMO)?.name
  const notesFile = !memo ? undefined : notes ?? (memo.file === memoNotesFile(documents, memo.name) ? memoNotesFile(named, memo.name) : memo.file)
  return { documents: named, notes_file: notesFile, applied }
}

/** The lawyer's question on files the outline's revisions still do not
 * deliver as the record asks (`issues`, from requestedFilesIssues on the same
 * documents): the lawyer's words, what they ask for, by title, "the memo" and
 * file name, and how the plan delivers it instead. Its id keys on the record's
 * quotes and the issues, so it is asked once: undefined when there is no issue
 * or a question with that id is already open or answered. */
export function requestedFilesQuestion(directory: string, requested: RequestedFiles, documents: PlannedDocument[], issues: string[]): Question | undefined {
  if (!issues.length) return
  const id = `requested-files-${shortHash(JSON.stringify([requested.files.map(file => file.quote), issues]))}`
  if (readQuestions(directory).some(question => question.id === id)) return
  // Issues arise only for files whose memo is planned.
  const lines = requested.files.flatMap(file => {
    const check = fileCheck(file, documents, true)
    return check.unsupported || !check.planned ? []
      : [`Your words "${excerpt(file.quote)}" ask for ${asked(file, id => documents.find(document => document.id === id))}, but the document plan delivers ${check.planned}.`]
  })
  if (!lines.length) return
  return {
    id,
    stage: "draft",
    question: [...lines, `${lines.length > 1 ? "For each, do" : "Do"} you want it delivered as your words ask, or as planned? If neither, say how it should be delivered.`].join(" "),
    context: "Drafting waits for your answer. Your answer becomes part of the document plan's basis, so the plan is drawn up again with it before any part is drafted, replacing any edits to the saved plan.",
    status: "open",
  }
}

/** What a recorded file asks for, in the lawyer's terms: its contents by title
 * and "the memo", in the lawyer's order, and how they are to be delivered. */
function asked(file: RequestedFile, planned: (id: string) => { title: string } | undefined) {
  const parts = file.holds.map(id => id === MEMO ? "the memo" : planned(id) ? `"${planned(id)!.title}"` : "a document no longer planned")
  const how = file.name ? (parts.length > 1 ? `in one file named "${file.name}"` : `as a file named "${file.name}"`) : parts.length > 1 ? "in one file" : "as a file of its own"
  return `${serialList(parts)} ${how}`
}

export function writeRequestedFiles(directory: string, saved: LeanRequestedFiles) {
  writeLeanText(directory, RECORD_FILE, `${JSON.stringify(saved, null, 2)}\n`)
}

/** The saved record, or undefined when it is missing or malformed. */
export function readRequestedFiles(directory: string): LeanRequestedFiles | undefined {
  const value = readLeanJson(directory, RECORD_FILE)
  if (!record(value) || value.schema !== LEAN_REQUESTED_FILES_SCHEMA || typeof value.basis !== "string" || typeof value.stated !== "boolean" ||
    !Array.isArray(value.files) || !value.files.every(savedFile) || (!value.stated && value.files.length)) return
  return value as LeanRequestedFiles
}

const savedFile = (row: unknown) => record(row) && Array.isArray(row.holds) && row.holds.length > 0 &&
  row.holds.every(id => id === MEMO || (typeof id === "string" && /^d[1-9]\d*$/.test(id))) && (row.name === null || typeof row.name === "string") && typeof row.quote === "string"

/** The open items for recorded files the saved plan does not deliver as the
 * lawyer asked, derived from the record and never stored. Only a record of the
 * saved plan counts (same basis); a plan written by hand, or with no such
 * record, gives none. A file the run cannot build (fileShape, or the memo
 * joined to a redline) always gives one. Otherwise each file is compared with
 * the names registration gives the plan's files (plannedRegistration): the
 * memo with one document must be that document's combined file
 * (combinedDelivery), one document alone must not be combined, the memo alone
 * must be delivered with no document combined, and a requested name must be
 * that file's name. Each item quotes the lawyer's words, says what they ask for
 * and which files the run delivers instead, and belongs to the drafting notes,
 * so a repair goes to the notes writer. Each id keys on the item's content. */
export function requestedFilesItems(directory: string): OpenItem[] {
  const plan = readLeanPlan(directory)
  const saved = readRequestedFiles(directory)
  if (!plan || plan.basis === undefined || !saved || saved.basis !== plan.basis || !saved.stated) return []
  const registered = plannedRegistration(directory, plan)
  // Registration names are lowercased; the run's evidence keeps their spelling.
  const spelled = new Map(deliveryArtifactEvidence(directory).map(row => {
    const name = "download_filename" in row ? row.download_filename : row.planned_filename
    return [name.toLocaleLowerCase(), name] as const
  }))
  const combined = combinedDelivery(directory, plan)
  const memo = requestedMemo(directory, plan)
  const registeredAs = (kind: string, document: DocID | null) => registered.find(row => row.kind === kind && row.document === document)?.name ?? null
  const deliveredIn = (id: string) => {
    if (id === MEMO) return !memo ? null : combined ? registeredAs("combined", combined.document) : registeredAs("drafting_notes", null)
    const document = plan.documents.find(row => row.id === id)
    return !document ? null : combined?.document === document.id ? registeredAs("combined", document.id) : registeredAs(document.deliverable === "redline" ? "redline" : "body", document.id)
  }
  const nameKey = (name: string) => `${deliveryStem({ title: "", file: name })}.docx`.toLocaleLowerCase()
  const lawyerQuote = (words: string) => `"${words.replace(/\s+/g, " ").trim()}"` + (matterSourceIDs(directory).includes(REQUEST_SOURCE) ? "" : ` [source: ${REQUEST_SOURCE}]`)
  const items = saved.files.flatMap(file => {
    const shape = fileShape(file)
    const target = shape?.document ? plan.documents.find(row => row.id === shape.document) : undefined
    const reason = !shape ? file.holds.filter(id => id !== MEMO).length > 1 ? " The run cannot put more than one document in one file." : " The run can join the memo and a document in one file only with the memo first."
      : shape.memo && target?.deliverable === "redline" ? " The run can join the memo and a document in one file only with the document's clean text, not its redline." : ""
    const delivered = !reason && shape && (!shape.document || target) && (!shape.memo || memo) &&
      (shape.memo && target ? combined?.document === target.id : target ? combined?.document !== target.id : !combined) &&
      (file.name === null || deliveredIn(target?.id ?? MEMO) === nameKey(file.name))
    if (delivered) return []
    const files = new Map<string, string[]>()
    for (const id of file.holds) {
      const name = deliveredIn(id)
      if (name) files.set(name, [...files.get(name) ?? [], id === MEMO ? "the memo" : `"${plan.documents.find(row => row.id === id)!.title}"`])
    }
    const phrases = [...[...files].map(([name, parts]) => `${serialList(parts)} in "${spelled.get(name) ?? name}"`), ...(file.holds.includes(MEMO) && !memo ? ["no memo"] : [])]
    return [openItem("finalize", "omission", `Requested files: the lawyer's words ${lawyerQuote(file.quote)} ask for ${asked(file, id => plan.documents.find(row => row.id === id))}, ` +
      `but the run delivers ${phrases.length ? serialList(phrases) : "none of them"}.${reason} The lawyer should confirm how these files are to be delivered.`, { repair_target: "memo" })]
  })
  return [...new Map(items.map(row => [row.id, row])).values()]
}

const titles = (documents: { title: string }[]) => serialList(documents.map(document => `"${document.title}"`))

/** "A", "A and B", or "A, B and C". */
function serialList(items: string[]) {
  return items.length > 1 ? `${items.slice(0, -1).join(", ")} and ${items.at(-1)}` : items.join("")
}
