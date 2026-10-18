import path from "node:path"
import { sha256File } from "../../lib/lineage"
import { normalize, normalizeWithMap } from "../../lib/quote-verify"
import { TOOL_SYNTAX_ISSUE, printedToolCall, record } from "./checks"
import { combinedDelivery, deliveryArtifactEvidence, deliveryStem, memoNotesFile, plannedRegistration, readLeanDocuments, requestedMemo } from "./deliverables"
import { NAMING_WORDS } from "./delivery-selection"
import { plainFileName, readLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { tagged } from "./prompts"
import { REQUEST_SOURCE, excerpt, instructionFound, lawyerWords, matterSourceIDs, openItem, quoteMissIssue } from "./quotes"
import { answeredQuestionsText, readLeanJson, readLeanText, readQuestions, shortHash, writeLeanText, type OpenItem, type Question } from "./state"

// Requested files: which files the lawyer asks to receive, and which planned
// documents and the memo go in each. One focused job reads the lawyer's words
// (request, updates and answers) beside the planned documents' ids and titles
// and the memo; parseRequestedFiles grounds each file in the lawyer's exact
// words, its name in the lawyer's own text where those words stand and its
// contents in the plan. The record is the lawyer's packaging, never the plan's,
// and holds the planned titles and memo it was read for. Code does not rewrite
// the plan's delivery shape: requestedFilesIssues gives the outline revision
// feedback in the outline's own terms, and requestedFilesQuestion asks the
// lawyer about a mismatch the revisions leave, once for the lawyer's words as
// they stand. A file name is no planning judgement, so withRequestedNames gives
// grounded names to the plan. A shape the run cannot build (several documents
// in one file, or the memo after its document) is no revision issue; like any
// file the final plan does not deliver as asked, it reaches the lawyer as an
// open item derived from the record (requestedFilesItems), never stored. A job
// whose reply never passed its check is recorded as unchecked, and disclosed
// the same way.

export const LEAN_REQUESTED_FILES_SCHEMA = "dochaus.lean.requested-files/1"

/** One file the lawyer asks to receive: the planned ids, and "memo" for the
 * memo, it holds, in the lawyer's order or, where the lawyer's words set none,
 * with the memo first; its name in the lawyer's own spelling, or null; and the
 * lawyer's exact words that ask for it. A name a reply gave for these words but
 * the lawyer's words did not confirm is kept as `refused_name`, beside a null
 * name, for disclosure only. */
export type RequestedFile = { holds: string[]; name: string | null; quote: string; refused_name?: string }
/** `stated` is false, with no files, when the lawyer's words do not say how the
 * planned documents and memo are delivered as files. */
export type RequestedFiles = { stated: boolean; files: RequestedFile[] }
/** The record of the plan whose basis is `basis`, read for the planned titles
 * and memo name in `planned` (requestedFilesPlanned). With `status`
 * "unchecked", the files job's reply never passed its check, so it is not known
 * which files the lawyer's words ask for: `stated` is false, with no files. */
export type LeanRequestedFiles = {
  schema: typeof LEAN_REQUESTED_FILES_SCHEMA; basis: string; planned: { titles: string[]; memo: string | null }; status?: "unchecked"; stated: boolean; files: RequestedFile[]
}

type Checked<T> = { issues: string[]; value?: T }

const MEMO = "memo"
const RECORD_FILE = "requested-files.json"
const QUESTION_ID = "requested-files-"

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
    tagged("planned_documents", "The documents the run drafts, by id. Their titles are the plan's own choices, not the lawyer's words.\n" +
      JSON.stringify(documents.map(document => ({ id: document.id, title: document.title })), null, 2)),
    memo ? tagged("planned_memo", `The run also writes the memo the lawyer asks for, planned as "${memo}". In "holds", write it as "memo".`) : "",
    tagged("output_contract", [
      "- Answer from the request, the lawyer's updates and answers only.",
      `- When those words say which files the lawyer wants to receive, what goes in each file or what a file is named, reply "stated": true and list each file they ask for. When they do not, reply {"stated":false,"files":[]}.`,
      `- For each file give "holds": the ids of the planned documents${memo ? ' and "memo" for the memo' : ""} that the lawyer's words put in that file` +
      `${memo ? ', in the order the words place them in the file, with "memo" first when the words set no order' : ""}; ` +
      "\"name\": the file name exactly as the lawyer writes it, or null when the lawyer gives no name; " +
      "\"quote\": the lawyer's exact words that ask for this file, copied verbatim, including the words that give its name.",
      `- Put each planned document${memo ? " and the memo" : ""} in one file at most. List only the files the lawyer's words ask for, and in each only what those words put in it.`,
      "- A file the lawyer's words name as the output of the whole request holds everything the request asks for; a file they name for one document holds only that document. " +
      `Put a planned document${memo ? " or the memo" : ""} in a file only where the lawyer's words place it.`,
      `- ${filesShape(Boolean(memo))}`,
      "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks the reply.",
    ].join("\n")),
  ].filter(Boolean).join("\n\n")
}

/** Reads the job's reply and grounds each file: its quote in the lawyer's own
 * words (instructionFound), a name as a plain Word file name the lawyer's text
 * gives where the quote stands (nameGrounded), kept in the lawyer's spelling,
 * its holds as planned ids or "memo" when a memo is planned, each in one file
 * at most. A name given without an extension is a Word file's: code adds
 * ".docx". Every defect is retry feedback; the files are returned only when
 * none is left. On the last attempt (`final`) a refused name becomes null, kept
 * as `refused_name` for disclosure; any other defect, a quote that is not the
 * lawyer's words included, still fails it, as the coverage check's does, and
 * the job is then unchecked. `refused` carries refused names across one job's
 * attempts, by quote: a later reply that gives these words no name keeps the
 * refused one for disclosure, until a reply gives them a name that passes. */
export function parseRequestedFiles(reply: string, directory: string, documents: PlannedDocument[], memo: string | undefined,
  { final = false, refused = new Map<string, string>() }: { final?: boolean; refused?: Map<string, string> } = {}): Checked<RequestedFiles> {
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
    const given = typeof row.name === "string" ? row.name.trim() : ""
    const name = !given ? null : /\.\p{L}[\p{L}\p{N}]{0,4}$/u.test(given) ? given : `${given}.docx`
    const file: RequestedFile = { holds: [...new Set((row.holds as string[]).map(id => id.trim()))], name, quote: row.quote.trim() }
    rows.push({ index, file })
    if (!instructionFound(directory, file.quote)) return void issues.push(quoteMissIssue(directory, where, { quote: file.quote, source_id: REQUEST_SOURCE }, "json"))
    // A name is held to its quote once the quote is the lawyer's words.
    const key = normalize(file.quote).toLowerCase()
    if (name === null) {
      if (refused.has(key)) file.refused_name = refused.get(key)
      return
    }
    const plain = /\.docx$/i.test(name) && plainFileName(name)
    const spelled = plain ? nameGrounded(directory, file.quote, name) : undefined
    if (spelled) {
      refused.delete(key)
      file.name = spelled
      return
    }
    refused.set(key, name)
    if (final) return void Object.assign(file, { name: null, refused_name: name })
    issues.push(!plain
      ? `${where} gives the name "${given}", which is not a plain Word file name. Give the file name as the lawyer writes it, without a folder or Markdown, or null when the lawyer gives none.`
      : `${where} gives the name "${given}", but its quote does not give that name. Quote the lawyer's words that name this file, and give the name exactly as they write it, or null when the lawyer gives none.`)
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

// A name's edges in the lawyer's text, read as instructionFound reads it
// (normalize: curly quotation marks straight, whitespace runs one space). A
// name starts at the text's start, after a space, a colon, or an opening
// quotation mark, bracket or Markdown mark; an apostrophe opens a name only
// outside a word. A name given without ".docx" follows a naming word
// (NAMING_WORDS, as deliveryNameQuoted grounds a title) and ends at a line
// break, a clause mark, a colon, a closing quotation mark or bracket, a full
// stop before a space, a closing mark or the end, or a dash between spaces.
const WORD = /[\p{L}\p{N}_]/u
const NAME_OPENS = /["«(\[*`:]/
const NAME_END = /^(?:$|[,;!?:"'»)\]]|\.(?:["'»)\]]|\s|$)|\s[-–—](?:\s|$))/
const LINE_END = /^[^\S\r\n]*[\r\n]/

/** Lower case where it keeps the text's length, so offsets still map. */
const fold = (text: string) => text.replace(/[\s\S]/gu, char => {
  const lower = char.toLowerCase()
  return lower.length === char.length ? lower : char
})

/** The name in the lawyer's own spelling when the lawyer's words
 * (lawyerWords, compared as instructionFound compares them: whitespace and case
 * aside) give `name` where they hold `quote`; undefined otherwise. Its stem
 * (the name without ".docx") starts inside the quoted words at a name's start
 * in the lawyer's text and is followed there by ".docx", or, given without it,
 * follows a naming word inside the quote and ends at a name's end. So a name cut
 * short from a longer name is refused, a quote may run on past the name, and a
 * line break inside the name does not matter. The spelling is the lawyer's text
 * from the stem's start, whitespace runs one space, with ".docx" as the lawyer
 * writes it or added. Not caught, since no rule of the text says where such a
 * name starts or ends: a ".docx" name cut short at its start at a space, which
 * leaves a name the lawyer's words also end with ".docx", and a name given
 * without ".docx" cut short at a full stop and a space or at a comma. */
function nameGrounded(directory: string, quote: string, name: string) {
  const stem = fold(normalize(name.replace(/\.docx$/i, "")))
  const words = fold(normalize(quote))
  if (!stem || !words) return
  // Built here: delivery-selection.ts and this module import each other through the stage table.
  const naming = new RegExp(String.raw`(?:^|[^\p{L}\p{N}_])(?:${NAMING_WORDS})\s+["'«]?$`, "u")
  for (const raw of lawyerWords(directory)) {
    const { text: normal, map } = normalizeWithMap(raw)
    const text = fold(normal)
    const starts = (at: number) => at === 0 || text[at - 1] === " " || NAME_OPENS.test(text[at - 1]!) ||
      (text[at - 1] === "'" && !WORD.test(text[at - 2] ?? ""))
    for (let start = text.indexOf(words); start >= 0; start = text.indexOf(words, start + 1)) {
      const end = start + words.length
      for (let at = text.indexOf(stem, start); at >= 0 && at < end; at = text.indexOf(stem, at + 1)) {
        if (!starts(at)) continue
        const close = at + stem.length
        const spelled = (to: number) => raw.slice(map[at], map[to - 1]! + 1).replace(/\s+/g, " ")
        if (/^\.docx(?![\p{L}\p{N}_])/u.test(text.slice(close))) return spelled(close + 5)
        if (close <= end && naming.test(text.slice(start, at)) && (NAME_END.test(text.slice(close)) || LINE_END.test(raw.slice(map[close - 1]! + 1))))
          return `${spelled(close)}.docx`
      }
    }
  }
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

/** A recorded file's shape (fileShape) when an outline with these document ids
 * and memo (`memo`: whether one is planned) can build it: its document planned
 * and its memo planned. Undefined otherwise. */
function buildable(file: RequestedFile, ids: string[], memo: boolean) {
  const shape = fileShape(file)
  return shape && (!shape.document || ids.includes(shape.document)) && (!shape.memo || memo) ? shape : undefined
}

/** A recorded file against an outline's documents and memo (`memo`: whether one
 * is planned): unsupported when the plan cannot build it (buildable);
 * otherwise, when the plan's deliverables do not deliver it as asked, the
 * outline revision `issue` and how the plan delivers it instead (`planned`). */
function fileCheck(file: RequestedFile, documents: PlannedDocument[], memo: boolean): { unsupported: true } | { unsupported: false; issue?: string; planned?: string } {
  const shape = buildable(file, documents.map(document => document.id), memo)
  if (!shape) return { unsupported: true }
  const target = shape.document && documents.find(document => document.id === shape.document)
  const words = `The lawyer's words "${excerpt(file.quote)}"`
  const combined = documents.filter(document => document.deliverable === "combined" && document !== target)
  const body = (document: PlannedDocument) => `"Deliverable: body" under "${document.title}"`
  if (target && shape.memo) {
    const redline = target.deliverable === "redline"
    const changes = [...(target.deliverable !== "combined" ? [`"Deliverable: combined" under "${target.title}"${redline ? ' in place of "Deliverable: redline"' : ""}`] : []), ...combined.map(body)]
    // The run joins the memo only with a document's clean text.
    return !changes.length ? { unsupported: false } : {
      unsupported: false,
      issue: `${words} ask for the memo and "${target.title}" in one file` +
        `${redline ? `, but the plan delivers "${target.title}" as a redline, and the run joins the memo only with a document's clean text` : ""}. Write ${serialList(changes)}.`,
      planned: target.deliverable === "combined" ? `the memo in one file with ${titles(combined)} as well`
        : redline ? `"${target.title}" as a redline and the memo ${combined.length ? `in one file with ${titles(combined)}` : "as a file of its own"}`
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
 * the memo alone names the notes file. Names are compared as registration
 * compares them (deliveryStem, lowercased). The lawyer's name displaces another
 * document's outline file name, which then falls back to its title, and says
 * so. It is skipped, and noted, when a name the plan does not choose has it: a
 * document's title-derived name or the notes file's (unless the memo's own file
 * is named); every name asked for twice is skipped. A file the plan cannot
 * build (buildable) names nothing. Without a name of its own, a notes file
 * still named after its document (memoNotesFile) is named again after the
 * named documents. Documents are identified by position (withDocumentIDs). */
export function withRequestedNames<T extends { title: string; file?: string; deliverable?: string }>(
  documents: T[], memo: { name: string; file: string } | undefined, requested: RequestedFiles,
): { documents: T[]; notes_file: string | undefined; applied: string[] } {
  type Target = number | typeof MEMO
  const ids = documents.map((_, index) => `d${index + 1}`)
  const wanted = requested.stated ? requested.files.flatMap((file): { target: Target; name: string }[] => {
    const shape = buildable(file, ids, memo !== undefined)
    return !shape || file.name === null ? [] : [{ target: shape.document ? ids.indexOf(shape.document) : MEMO, name: file.name }]
  }) : []
  const key = (name: string) => deliveryStem({ title: "", file: name }).toLocaleLowerCase()
  const skipped = new Set(wanted.filter(row => wanted.some(other => other !== row && key(other.name) === key(row.name))))
  // Each skip can restore an outline name or change the notes file's, so the
  // names are worked out again until no given name clashes.
  for (;;) {
    const given = wanted.filter(row => !skipped.has(row))
    const holder = (name: string) => given.find(row => key(row.name) === key(name))
    const named = documents.map((document, index) => {
      const name = given.find(row => row.target === index)?.name
      if (name) return { ...document, file: name }
      if (!document.file || !holder(document.file)) return document
      const { file: _file, ...rest } = document
      return rest as T
    })
    const name = given.find(row => row.target === MEMO)?.name
    const notes = name ?? (!memo ? `${deliveryStem(named[0] ?? { title: "Draft" })} - drafting notes.docx`
      : memo.file === memoNotesFile(documents, memo.name) ? memoNotesFile(named, memo.name) : memo.file)
    const fixed = [
      ...named.flatMap((document, index) => document.file ? [] : [{ target: index as Target, key: deliveryStem({ title: document.title }).toLocaleLowerCase() }]),
      ...(name ? [] : [{ target: MEMO as Target, key: key(notes) }]),
    ]
    const clashes = given.filter(row => fixed.some(other => other.target !== row.target && other.key === key(row.name)))
    if (clashes.length) {
      for (const row of clashes) skipped.add(row)
      continue
    }
    const label = (target: Target) => target === MEMO ? "the memo" : `"${documents[target]!.title}"`
    const applied = [
      ...wanted.flatMap(row => skipped.has(row) ? [`"${row.name}" not given to ${label(row.target)}: another planned file has that name`]
        : row.name !== (row.target === MEMO ? memo?.file : documents[row.target]!.file) ? [`${label(row.target)}: file "${row.name}"`] : []),
      ...documents.flatMap((document, index) => {
        const by = document.file && !given.some(row => row.target === index) ? holder(document.file) : undefined
        return by ? [`"${document.title}": outline file "${document.file}" dropped, as the lawyer's words give that name to ${label(by.target)}`] : []
      }),
    ]
    return { documents: named, notes_file: memo ? notes : undefined, applied }
  }
}

/** The id of the requested-files question on the lawyer's words as they stand:
 * the request and the lawyer's updates. An answer or a later job's quoting does
 * not change it; an update does. */
const questionID = (directory: string) =>
  `${QUESTION_ID}${shortHash(JSON.stringify([readLeanText(directory, "request.md") ?? "", readLeanText(directory, "updated-notes.md") ?? ""]))}`

/** The lawyer's question on files the outline's revisions still do not
 * deliver as the record asks (`issues`, from requestedFilesIssues on the same
 * documents): the lawyer's words, what they ask for, by title, "the memo" and
 * file name, and how the plan delivers it instead. It is asked once for the
 * lawyer's words as they stand (questionID): undefined when there is no issue
 * or that question is already open or answered, since a later job may quote
 * the same words differently. */
export function requestedFilesQuestion(directory: string, requested: RequestedFiles, documents: PlannedDocument[], issues: string[]): Question | undefined {
  if (!issues.length) return
  const id = questionID(directory)
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

/** The planned titles, in order, and memo name a record is read for. */
export function requestedFilesPlanned(documents: { title: string }[], memo: string | undefined): LeanRequestedFiles["planned"] {
  return { titles: documents.map(document => document.title), memo: memo ?? null }
}

export function writeRequestedFiles(directory: string, saved: LeanRequestedFiles) {
  writeLeanText(directory, RECORD_FILE, `${JSON.stringify(saved, null, 2)}\n`)
}

/** The saved record, or undefined when it is missing or malformed. */
export function readRequestedFiles(directory: string): LeanRequestedFiles | undefined {
  const value = readLeanJson(directory, RECORD_FILE)
  if (!record(value) || value.schema !== LEAN_REQUESTED_FILES_SCHEMA || typeof value.basis !== "string" || typeof value.stated !== "boolean" ||
    !(value.status === undefined || (value.status === "unchecked" && !value.stated)) ||
    !record(value.planned) || !Array.isArray(value.planned.titles) || !value.planned.titles.every(title => typeof title === "string") ||
    !(value.planned.memo === null || typeof value.planned.memo === "string") ||
    !Array.isArray(value.files) || !value.files.every(savedFile) || (!value.stated && value.files.length)) return
  return value as LeanRequestedFiles
}

const savedFile = (row: unknown) => record(row) && Array.isArray(row.holds) && row.holds.length > 0 &&
  row.holds.every(id => id === MEMO || (typeof id === "string" && /^d[1-9]\d*$/.test(id))) && (row.name === null || typeof row.name === "string") && typeof row.quote === "string" &&
  (row.refused_name === undefined || typeof row.refused_name === "string")

/** The open items for recorded files the saved plan does not deliver as the
 * lawyer asked, derived from the record and never stored. Only a record of the
 * saved plan counts: same basis, and read for the plan's titles and memo
 * (requestedFilesPlanned); a plan written or edited by hand, or with no such
 * record, gives none. An unchecked record gives one item, as the coverage
 * check's does: the run could not confirm which files the lawyer's words ask to
 * receive. A file the run cannot build (fileShape) always gives
 * one. Otherwise each file is compared with the names registration gives the
 * plan's files (plannedRegistration): the memo with one document must be that
 * document's combined file (combinedDelivery), one document alone must not be
 * combined, the memo alone must be delivered with no document combined, and a
 * requested name, or one the run could not confirm, must be that file's name,
 * free of a matter file the run did not write (the lawyer's accepted Word
 * revision of the run's file is the run's). Each item quotes the lawyer's
 * words, says what they ask for and every file the run delivers instead, by all
 * it holds, and belongs to the drafting notes, so a repair goes to the notes
 * writer. Once the lawyer has answered the question on these words
 * (questionID), an item on a file that question quoted cites the answer. Each
 * id keys on the item's content. */
export function requestedFilesItems(directory: string): OpenItem[] {
  const plan = readLeanPlan(directory)
  const saved = readRequestedFiles(directory)
  if (!plan || plan.basis === undefined || !saved || saved.basis !== plan.basis ||
    JSON.stringify(saved.planned) !== JSON.stringify(requestedFilesPlanned(plan.documents, plan.memo))) return []
  if (saved.status === "unchecked") return [openItem("finalize", "omission", "Requested files could not be checked: the check of the lawyer's words for the files they ask to receive gave no usable result, " +
    "so the run could not confirm which files the lawyer's words ask to receive or what each holds. The lawyer should confirm how the files are to be delivered.", { repair_target: "memo" })]
  if (!saved.stated) return []
  const registered = plannedRegistration(directory, plan)
  // Registration names are lowercased; the run's evidence keeps their spelling.
  const spellings = new Map<string, string>(deliveryArtifactEvidence(directory).map(row =>
    [`${row.kind}:${row.document}`, "download_filename" in row ? row.download_filename : row.planned_filename]))
  const rows = readLeanDocuments(directory)?.documents ?? []
  const combined = combinedDelivery(directory, plan)
  const memo = requestedMemo(directory, plan)
  const title = (id: string) => plan.documents.find(row => row.id === id)?.title
  // The file the run delivers `kind` of `document` as: its registered name and
  // its spelling, with registration's suffix; and whether a matter file the run
  // did not write holds that name (`foreign`), as registration then takes
  // another, with the run's own earlier name for it when the summary has one.
  const delivery = (kind: string, document: DocID | null) => {
    const row = registered.find(other => other.kind === kind && other.document === document)
    if (!row?.name) return null
    const planned = spellings.get(`${kind}:${document}`)
    const spelled = planned && row.wanted === planned.toLocaleLowerCase() && row.name.startsWith(row.wanted.slice(0, -5))
      ? `${planned.slice(0, -5)}${row.name.slice(row.wanted.length - 5)}` : row.name
    const titled = document ? title(document) ?? null : null
    const current = kind === "combined" ? null : sha256File(path.join(directory, spelled))?.sha256
    const owned = rows.some(other => other.title === titled && other.deliverable?.toLocaleLowerCase() === row.name && (other.sha256 === current || other.conflict === "accepted_revision"))
    const stem = row.wanted.slice(0, -5).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    const earlier = rows.find(other => other.kind === kind && other.title === titled && other.deliverable &&
      new RegExp(`^${stem}(?: \\(\\d+\\))?\\.docx$`).test(other.deliverable.toLocaleLowerCase()) && other.deliverable.toLocaleLowerCase() !== row.name)?.deliverable
    return { kind, title: titled, name: row.name, spelled, foreign: Boolean(current) && !owned, earlier }
  }
  const deliveredIn = (id: string) => {
    if (id === MEMO) return !memo ? null : combined ? delivery("combined", combined.document) : delivery("drafting_notes", null)
    const document = plan.documents.find(row => row.id === id)
    return !document ? null : combined?.document === document.id ? delivery("combined", document.id) : delivery(document.deliverable === "redline" ? "redline" : "body", document.id)
  }
  const describe = (file: NonNullable<ReturnType<typeof delivery>>) => {
    const what = file.kind === "combined" ? `the memo and "${file.title}"` : file.kind === "drafting_notes" ? "the memo" : `"${file.title}"${file.kind === "redline" ? " as a redline" : ""}`
    if (file.foreign) return `${what} in ${file.earlier ? `"${file.earlier}"` : "a file of another name"}, since a matter file the run did not write already holds "${file.spelled}"`
    return file.kind === "combined" ? `${what} in one file named "${file.spelled}"` : `${what} in "${file.spelled}"`
  }
  const nameKey = (name: string) => `${deliveryStem({ title: "", file: name })}.docx`.toLocaleLowerCase()
  const lawyerQuote = (words: string) => `"${words.replace(/\s+/g, " ").trim()}"` + (matterSourceIDs(directory).includes(REQUEST_SOURCE) ? "" : ` [source: ${REQUEST_SOURCE}]`)
  const id = questionID(directory)
  // The words the answered question quoted (requestedFilesQuestion); a later
  // job may quote them more or less fully.
  const flat = (words: string) => normalize(words).toLowerCase()
  const quoted = readQuestions(directory).filter(question => question.id === id && question.status === "answered")
    .flatMap(question => [...question.question.matchAll(/Your words "(.*?)" ask for /g)].map(match => flat(match[1]!.replace(/…$/, ""))))
  const answered = (quote: string) => quoted.some(words => words.includes(flat(quote)) || flat(quote).includes(words))
  const items = saved.files.flatMap(file => {
    const check = fileCheck(file, plan.documents, memo !== undefined)
    const shape = fileShape(file)
    const target = shape && deliveredIn(shape.document ?? MEMO)
    const given = file.name ?? file.refused_name
    const named = given === undefined || Boolean(target && !target.foreign && target.name === nameKey(given))
    if (!check.unsupported && !check.issue && named) return []
    const files = new Map<string, NonNullable<ReturnType<typeof delivery>>>()
    for (const held of file.holds) {
      const delivered = deliveredIn(held)
      if (delivered) files.set(delivered.name, delivered)
    }
    const phrases = [...[...files.values()].map(describe), ...(file.holds.includes(MEMO) && !memo ? ["no memo"] : [])]
    const delivers = phrases.length ? serialList(phrases) : "none of them"
    const reason = shape ? "" : file.holds.filter(held => held !== MEMO).length > 1 ? " The run cannot put more than one document in one file."
      : " The run can join the memo and a document in one file only with the memo first."
    const refused = file.name === null && file.refused_name !== undefined && !named
    const closing = answered(file.quote) && !check.unsupported && check.planned
      ? ` The lawyer answered the requested-files question ${id}; the plan drawn up with that answer still delivers them so.`
      : " The lawyer should confirm how these files are to be delivered."
    return [openItem("finalize", "omission", `Requested files: the lawyer's words ${lawyerQuote(file.quote)} ask for ${asked(file, held => plan.documents.find(row => row.id === held))}` +
      (refused ? ` and name this file; the run could not confirm the name "${file.refused_name}" in those words and delivers ${delivers}.` : `, but the run delivers ${delivers}.`) +
      `${reason}${closing}`, { repair_target: "memo" })]
  })
  return [...new Map(items.map(row => [row.id, row])).values()]
}

const titles = (documents: { title: string }[]) => serialList(documents.map(document => `"${document.title}"`))

/** "A", "A and B", or "A, B and C". */
function serialList(items: string[]) {
  return items.length > 1 ? `${items.slice(0, -1).join(", ")} and ${items.at(-1)}` : items.join("")
}
