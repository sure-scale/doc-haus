import { afterEach, describe, expect, test } from "bun:test"
import { TOOL_SYNTAX_ISSUE } from "./checks"
import { writeLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { tagged } from "./prompts"
import { REQUEST_SOURCE, openItem, quoteMissIssue } from "./quotes"
import {
  LEAN_REQUESTED_FILES_SCHEMA, parseRequestedFiles, readRequestedFiles, requestedFilesIssues, requestedFilesItems, requestedFilesPrompt, requestedFilesQuestion,
  withRequestedNames, writeRequestedFiles, type LeanRequestedFiles, type RequestedFile, type RequestedFiles,
} from "./requested-files"
import { writeLeanText, writeQuestions } from "./state"
import { cleanupMatters, matter } from "./test-support"

// Requested files: the files the lawyer asks to receive, grounded in the
// lawyer's words, held against the plan as outline feedback, a question and,
// at the end, open items. Synthetic parties and file names only.

afterEach(cleanupMatters)

const REQUEST = "Prepare a share transfer form and a board minute for Kestrel Works. " +
  "Send the board minute with a cover memo, all in one file named Kestrel Board Pack.docx. " +
  "Deliver the share transfer form separately as Kestrel Transfer.docx."
const BOTH = "Prepare a share transfer form and a board minute for Kestrel Works"
const PACK = "Send the board minute with a cover memo, all in one file named Kestrel Board Pack.docx"
const TRANSFER = "Deliver the share transfer form separately as Kestrel Transfer.docx"
const MEMO = "cover memo"
const SHAPE = 'Reply with the JSON object only: {"stated":true,"files":[{"holds":["planned id or memo"],"name":"file name or null","quote":"..."}]}.'

const outline = (title: string, fields: Partial<Omit<PlannedDocument, "id" | "title">> = {}): Omit<PlannedDocument, "id"> =>
  ({ title, purpose: "", parts: [{ heading: "Terms", scope: "" }], ...fields })
const planned = (id: DocID, title: string, fields: Partial<Omit<PlannedDocument, "id" | "title">> = {}): PlannedDocument => ({ id, ...outline(title, fields) })
const DOCUMENTS = [planned("d1", "Share transfer form"), planned("d2", "Board minute")]
const file = (holds: string[], quote: string, name: string | null = null): RequestedFile => ({ holds, name, quote })
const reply = (stated: boolean, ...files: unknown[]) => JSON.stringify({ stated, files })
const stated = (...files: RequestedFile[]): RequestedFiles => ({ stated: true, files })
const ASKED = stated(file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), file(["d1"], TRANSFER, "Kestrel Transfer.docx"))
const nameNotPlain = (index: number, name: string) =>
  `files[${index}] gives the name "${name}", which is not a plain Word file name. Give the file name as the lawyer writes it, ending in .docx, without a folder or Markdown, or null when the lawyer gives none.`
const nameNotQuoted = (index: number, name: string) =>
  `files[${index}] gives the name "${name}", but its quote does not give that name. Quote the lawyer's words that name this file, and give the name exactly as they write it, adding only .docx to a name given without it, or null when the lawyer gives none.`

function prepared(request = REQUEST) {
  const directory = matter()
  writeLeanText(directory, "request.md", request)
  return directory
}

describe("requested files prompt", () => {
  test("asks one question over the lawyer's words and the planned documents, with the memo only when one is planned", () => {
    const directory = prepared()
    const documents = [planned("d1", "Share transfer form", { file: "Kestrel Transfer.docx" }), planned("d2", "Board minute", { deliverable: "combined" })]
    const prompt = requestedFilesPrompt(directory, documents, MEMO)
    expect(prompt).toContain(tagged("request", REQUEST))
    expect(prompt).toContain(JSON.stringify([
      { id: "d1", title: "Share transfer form", file: "Kestrel Transfer.docx", deliverable: "body" },
      { id: "d2", title: "Board minute", deliverable: "combined" },
    ], null, 2))
    expect(prompt).toContain("<planned_memo>")
    expect(prompt).toContain(`"${MEMO}"`)
    expect(prompt).toContain(SHAPE)
    expect(prompt).not.toContain("<user_updates>")
    const plain = requestedFilesPrompt(directory, DOCUMENTS, undefined)
    expect(plain).not.toContain("<planned_memo>")
    // The request's own words aside, the prompt does not mention a memo.
    expect(plain.replace(tagged("request", REQUEST), "")).not.toMatch(/\bmemo\b/)
    expect(plain).toContain('Reply with the JSON object only: {"stated":true,"files":[{"holds":["planned id"],"name":"file name or null","quote":"..."}]}.')
  })
})

describe("parse requested files", () => {
  test("a reply grounded in the lawyer's words is kept as given, and a reply that the words do not say lists no file", () => {
    const directory = prepared()
    const fenced = "```json\n" + reply(true, file(["memo", "d2", "d2"], ` ${PACK} `, "Kestrel Board Pack.docx"), file(["d1"], TRANSFER, "Kestrel Transfer.docx")) + "\n```"
    expect(parseRequestedFiles(fenced, directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: ASKED })
    expect(parseRequestedFiles(reply(false), directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: { stated: false, files: [] } })
    // No name given is null; an empty name is no name.
    expect(parseRequestedFiles(reply(true, file(["d1"], TRANSFER), { holds: ["d2"], name: " ", quote: BOTH }), directory, DOCUMENTS, MEMO))
      .toEqual({ issues: [], value: stated(file(["d1"], TRANSFER), file(["d2"], BOTH)) })
  })

  test("a name the lawyer gives as a title is grounded with only .docx added", () => {
    const words = "Put the cover memo and the board minute in one file titled Kestrel Board Pack"
    const directory = prepared(`Prepare a board minute for Kestrel Works. ${words}.`)
    expect(parseRequestedFiles(reply(true, file(["memo", "d2"], words, "Kestrel Board Pack.docx")), directory, DOCUMENTS, MEMO))
      .toEqual({ issues: [], value: stated(file(["memo", "d2"], words, "Kestrel Board Pack.docx")) })
  })

  test("a quote that is not the lawyer's words is retried; a last attempt drops that file and never keeps an ungrounded one", () => {
    const directory = prepared()
    const wrong = file(["d1"], "Deliver the transfer form as its own file", "Kestrel Transfer.docx")
    const text = reply(true, file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), wrong)
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO)).toEqual({ issues: [quoteMissIssue(directory, "files[1]", { quote: wrong.quote, source_id: REQUEST_SOURCE }, "json")] })
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO, { final: true })).toEqual({ issues: [], value: stated(file(["memo", "d2"], PACK, "Kestrel Board Pack.docx")) })
    // With every file dropped nothing the lawyer said is left.
    expect(parseRequestedFiles(reply(true, wrong), directory, DOCUMENTS, MEMO, { final: true })).toEqual({ issues: [], value: { stated: false, files: [] } })
    // A dropped file's holds do not refuse the reply.
    expect(parseRequestedFiles(reply(true, file(["memo", "d2"], PACK), { ...wrong, holds: ["d2"] }), directory, DOCUMENTS, MEMO, { final: true }))
      .toEqual({ issues: [], value: stated(file(["memo", "d2"], PACK)) })
  })

  test("a name its quote does not give, or that is not a plain Word file name, is retried; a last attempt keeps the file without it", () => {
    const directory = prepared()
    const text = reply(true, file(["memo", "d2"], PACK, "Kestrel Pack.docx"), file(["d1"], TRANSFER, "transfers/Kestrel Transfer.docx"))
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotQuoted(0, "Kestrel Pack.docx"), nameNotPlain(1, "transfers/Kestrel Transfer.docx")] })
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO, { final: true })).toEqual({ issues: [], value: stated(file(["memo", "d2"], PACK), file(["d1"], TRANSFER)) })
    expect(parseRequestedFiles(reply(true, file(["d1"], TRANSFER, "Kestrel Transfer.pdf")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotPlain(0, "Kestrel Transfer.pdf")] })
  })

  test("a file holds planned documents, the memo only when one is planned, each in one file at most, on every attempt", () => {
    const directory = prepared()
    const check = (text: string, memo: string | null = MEMO) => [parseRequestedFiles(text, directory, DOCUMENTS, memo ?? undefined), parseRequestedFiles(text, directory, DOCUMENTS, memo ?? undefined, { final: true })]
    const twice = (issue: string) => [{ issues: [issue] }, { issues: [issue] }]
    expect(check(reply(true, file(["d3"], TRANSFER)))).toEqual(twice('files[0] holds "d3", which is not planned. Use the planned ids d1, d2 or "memo".'))
    expect(check(reply(true, file(["memo", "d2"], PACK)), null)).toEqual(twice('files[0] holds "memo", but no memo is planned. Use the planned ids d1, d2.'))
    expect(check(reply(true, file(["memo", "d2"], PACK), file(["d2"], BOTH)))).toEqual(twice(
      '"d2" is in files[0] and files[1]. Put each planned document and the memo in one file at most, as the lawyer\'s words place it.'))
    const rowIssue = `files[0] needs "holds" as a nonempty list of planned ids or "memo", "name" as a file name or null, and a nonempty "quote" string. ${SHAPE}`
    expect(check(reply(true, file([], TRANSFER)))).toEqual(twice(rowIssue))
    expect(check(reply(true, { holds: ["d1"], name: "Kestrel Transfer.docx" }))).toEqual(twice(rowIssue))
  })

  test("stated and files must agree; a tool call or another reply is retried", () => {
    const directory = prepared()
    const parse = (text: string) => parseRequestedFiles(text, directory, DOCUMENTS, MEMO, { final: true })
    expect(parse(reply(false, file(["d1"], TRANSFER)))).toEqual({ issues: ['"stated" is false, but "files" lists files. Give "stated": true with the files the lawyer\'s words ask for, or "stated": false with "files": [].'] })
    expect(parse(reply(true))).toEqual({ issues: ['"stated" is true, but "files" is empty. List each file the lawyer\'s words ask for, or reply {"stated":false,"files":[]} when they do not say.'] })
    expect(parse('<tool_call>{"name":"read-document","arguments":{"document":"x"}}</tool_call>')).toEqual({ issues: [TOOL_SYNTAX_ISSUE] })
    expect(parse("The lawyer asks for two files.")).toEqual({ issues: [SHAPE] })
    expect(parse(JSON.stringify({ files: [] }))).toEqual({ issues: [SHAPE] })
  })
})

describe("requested files issues", () => {
  test("the memo with one document asks for that document combined and no other", () => {
    const issues = (documents: PlannedDocument[]) => requestedFilesIssues(documents, MEMO, ASKED)
    expect(issues(DOCUMENTS)).toEqual({ issues: [`The lawyer's words "${PACK}" ask for the memo and "Board minute" in one file. Write "Deliverable: combined" under "Board minute".`], unsupported: [] })
    expect(issues([planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")])).toEqual({ issues: [
      `The lawyer's words "${PACK}" ask for the memo and "Board minute" in one file. Write "Deliverable: combined" under "Board minute" and "Deliverable: body" under "Share transfer form".`,
      `The lawyer's words "${TRANSFER}" ask for "Share transfer form" in a file of its own, without the memo. Write "Deliverable: body" under "Share transfer form" in place of "Deliverable: combined".`,
    ], unsupported: [] })
    expect(issues([planned("d1", "Share transfer form", { deliverable: "redline" }), planned("d2", "Board minute", { deliverable: "combined" })])).toEqual({ issues: [], unsupported: [] })
  })

  test("the memo in a file of its own asks for no document combined", () => {
    const own = stated(file(["memo"], PACK))
    expect(requestedFilesIssues([planned("d1", "Share transfer form"), planned("d2", "Board minute", { deliverable: "combined" })], MEMO, own)).toEqual({ issues: [
      `The lawyer's words "${PACK}" ask for the memo in a file of its own. Write "Deliverable: body" under "Board minute" in place of "Deliverable: combined".`,
    ], unsupported: [] })
    expect(requestedFilesIssues(DOCUMENTS, MEMO, own)).toEqual({ issues: [], unsupported: [] })
  })

  test("a shape the run cannot build is unsupported, never an issue", () => {
    const redline = [planned("d1", "Share transfer form"), planned("d2", "Board minute", { deliverable: "redline" })]
    expect(requestedFilesIssues(DOCUMENTS, MEMO, stated(file(["d1", "d2"], BOTH), file(["d2", "memo"], PACK), file(["d1"], TRANSFER)))).toEqual({ issues: [], unsupported: [0, 1] })
    expect(requestedFilesIssues(redline, MEMO, stated(file(["memo", "d2"], PACK)))).toEqual({ issues: [], unsupported: [0] })
    expect(requestedFilesIssues(DOCUMENTS, undefined, stated(file(["memo", "d2"], PACK)))).toEqual({ issues: [], unsupported: [0] })
    expect(requestedFilesIssues(DOCUMENTS, MEMO, stated(file(["d3"], TRANSFER)))).toEqual({ issues: [], unsupported: [0] })
    expect(requestedFilesIssues(DOCUMENTS, MEMO, { stated: false, files: [] })).toEqual({ issues: [], unsupported: [] })
  })

  test("issues name documents by title, never by id", () => {
    const { issues } = requestedFilesIssues([planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")], MEMO, ASKED)
    expect(issues).toHaveLength(2)
    expect(issues.join(" ")).not.toMatch(/\bd[12]\b/)
  })
})

describe("requested names", () => {
  test("a grounded name becomes its document's file, and the memo's default file follows its document", () => {
    const documents = [outline("Share transfer form"), outline("Board minute", { deliverable: "combined" })]
    expect(withRequestedNames(documents, { name: MEMO, file: "Board minute - cover memo.docx" }, ASKED)).toEqual({
      documents: [outline("Share transfer form", { file: "Kestrel Transfer.docx" }), outline("Board minute", { deliverable: "combined", file: "Kestrel Board Pack.docx" })],
      notes_file: "Kestrel Board Pack - cover memo.docx",
      applied: ['"Board minute": file "Kestrel Board Pack.docx"', '"Share transfer form": file "Kestrel Transfer.docx"'],
    })
    // A notes file named otherwise stays.
    expect(withRequestedNames(documents, { name: MEMO, file: "Kestrel Notes.docx" }, ASKED).notes_file).toBe("Kestrel Notes.docx")
    expect(withRequestedNames(documents, undefined, stated(file(["d1"], TRANSFER, "Kestrel Transfer.docx"))).notes_file).toBeUndefined()
  })

  test("the memo's own name becomes the notes file", () => {
    const documents = [outline("Share transfer form"), outline("Board minute")]
    expect(withRequestedNames(documents, { name: MEMO, file: "Share transfer form - cover memo.docx" }, stated(file(["memo"], PACK, "Kestrel Board Pack.docx"))))
      .toEqual({ documents, notes_file: "Kestrel Board Pack.docx", applied: ['the memo: file "Kestrel Board Pack.docx"'] })
  })

  test("a name another planned file has is skipped and noted; an unbuildable file's name and a stated:false record change nothing", () => {
    const documents = [outline("Share transfer form"), outline("Board minute", { file: "Kestrel Board Pack.docx" })]
    const memo = { name: MEMO, file: "Kestrel Board Pack - cover memo.docx" }
    expect(withRequestedNames(documents, memo, stated(file(["d1"], TRANSFER, "kestrel board pack.docx")))).toEqual({
      documents, notes_file: memo.file, applied: ['"kestrel board pack.docx" not given to "Share transfer form": another planned file has that name'],
    })
    // Two files asked for under one name: neither takes it.
    expect(withRequestedNames(documents.map(row => outline(row.title)), undefined, stated(file(["d1"], TRANSFER, "Kestrel Pack.docx"), file(["d2"], PACK, "Kestrel Pack.docx"))).applied).toEqual([
      '"Kestrel Pack.docx" not given to "Share transfer form": another planned file has that name',
      '"Kestrel Pack.docx" not given to "Board minute": another planned file has that name',
    ])
    expect(withRequestedNames(documents, memo, stated(file(["d1", "d2"], BOTH, "Kestrel Pack.docx"), file(["d2", "memo"], PACK, "Kestrel Pack.docx"))))
      .toEqual({ documents, notes_file: memo.file, applied: [] })
    expect(withRequestedNames(documents, memo, { stated: false, files: [] })).toEqual({ documents, notes_file: memo.file, applied: [] })
  })
})

describe("requested files question", () => {
  const ASK = `Your words "${PACK}" ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", but the document plan delivers the memo and "Board minute" as separate files. ` +
    "Do you want it delivered as your words ask, or as planned? If neither, say how it should be delivered."

  test("a mismatch left after the outline's revisions is asked once, in the lawyer's words and the plan's titles", () => {
    const directory = prepared()
    const { issues } = requestedFilesIssues(DOCUMENTS, MEMO, ASKED)
    const question = requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)!
    expect(question).toEqual({ id: expect.stringMatching(/^requested-files-[0-9a-f]+$/), stage: "draft", question: ASK, context: expect.any(String), status: "open" })
    expect(question.question).not.toMatch(/\bd[12]\b/)
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)!.id).toBe(question.id)
    // Not asked again once open or answered.
    writeQuestions(directory, [question])
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)).toBeUndefined()
    writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)).toBeUndefined()
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, [])).toBeUndefined()
  })

  test("several mismatches are asked together", () => {
    const directory = prepared()
    const documents = [planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")]
    const { issues } = requestedFilesIssues(documents, MEMO, ASKED)
    expect(requestedFilesQuestion(directory, ASKED, documents, issues)!.question).toBe(
      `Your words "${PACK}" ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", but the document plan delivers the memo in one file with "Share transfer form", and "Board minute" as a file of its own. ` +
      `Your words "${TRANSFER}" ask for "Share transfer form" as a file named "Kestrel Transfer.docx", but the document plan delivers "Share transfer form" in one file with the memo. ` +
      "For each, do you want it delivered as your words ask, or as planned? If neither, say how it should be delivered.")
  })
})

describe("requested files items", () => {
  const BASIS = "basis-1"
  const RECORD: LeanRequestedFiles = { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: BASIS, ...ASKED }
  const AS_ASKED = [outline("Share transfer form", { file: "Kestrel Transfer.docx" }), outline("Board minute", { file: "Kestrel Board Pack.docx", deliverable: "combined" })]
  const NOTES = { name: MEMO, file: "Kestrel Board Pack - cover memo.docx" }
  const item = (detail: string) => openItem("finalize", "omission", `Requested files: ${detail} The lawyer should confirm how these files are to be delivered.`, { repair_target: "memo" })

  function saved(documents = AS_ASKED, record: LeanRequestedFiles = RECORD, basis: string | null = BASIS) {
    const directory = prepared()
    writeLeanPlan(directory, documents, basis ?? undefined, undefined, NOTES)
    writeRequestedFiles(directory, record)
    return directory
  }

  test("the record reads back as written; a malformed record reads as none", () => {
    const directory = saved()
    expect(readRequestedFiles(directory)).toEqual(RECORD)
    writeLeanText(directory, "requested-files.json", JSON.stringify({ ...RECORD, files: [{ holds: ["d1"], name: 3, quote: TRANSFER }] }))
    expect(readRequestedFiles(directory)).toBeUndefined()
    expect(requestedFilesItems(directory)).toEqual([])
  })

  test("a plan that delivers the files as asked gives no item", () => {
    expect(requestedFilesItems(saved())).toEqual([])
  })

  test("a plan changed after the record to deliver the memo apart gives one item in the lawyer's words", () => {
    const directory = saved([AS_ASKED[0]!, outline("Board minute", { file: "Kestrel Board Pack.docx", deliverable: "body" })])
    expect(requestedFilesItems(directory)).toEqual([item(`the lawyer's words "${PACK}" [source: request] ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", ` +
      'but the run delivers the memo in "Kestrel Board Pack - cover memo.docx" and "Board minute" in "Kestrel Board Pack.docx".')])
  })

  test("a document delivered under another name gives an item naming the file delivered", () => {
    const directory = saved([outline("Share transfer form"), AS_ASKED[1]!])
    expect(requestedFilesItems(directory)).toEqual([item(`the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file named "Kestrel Transfer.docx", ` +
      'but the run delivers "Share transfer form" in "Share transfer form.docx".')])
  })

  test("a shape the run cannot build always gives an item", () => {
    const record = (...files: RequestedFile[]): LeanRequestedFiles => ({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: BASIS, stated: true, files })
    expect(requestedFilesItems(saved(AS_ASKED, record(file(["d1", "d2"], BOTH))))).toEqual([item(`the lawyer's words "${BOTH}" [source: request] ask for "Share transfer form" and "Board minute" in one file, ` +
      'but the run delivers "Share transfer form" in "Kestrel Transfer.docx" and "Board minute" in "Kestrel Board Pack.docx". The run cannot put more than one document in one file.')])
    expect(requestedFilesItems(saved(AS_ASKED, record(file(["d2", "memo"], PACK, "Kestrel Board Pack.docx"))))).toEqual([item(`the lawyer's words "${PACK}" [source: request] ask for "Board minute" and the memo in one file named "Kestrel Board Pack.docx", ` +
      'but the run delivers "Board minute" and the memo in "Kestrel Board Pack.docx". The run can join the memo and a document in one file only with the memo first.')])
  })

  test("a hand-written plan, or a record of another plan, gives no item", () => {
    const changed = [AS_ASKED[0]!, outline("Board minute", { deliverable: "body" })]
    expect(requestedFilesItems(saved(changed, RECORD, null))).toEqual([])
    expect(requestedFilesItems(saved(changed, { ...RECORD, basis: "basis-0" }))).toEqual([])
    const directory = prepared()
    writeLeanPlan(directory, changed, BASIS, undefined, NOTES)
    expect(requestedFilesItems(directory)).toEqual([])
  })
})
