import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { writeFileSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../../lib/docxodus"
import { sha256File } from "../../lib/lineage"
import { renderMarkdownToDocx } from "../../lib/render"
import { workingDocxKey } from "../../lib/working-docx"
import { TOOL_SYNTAX_ISSUE } from "./checks"
import { checkDeliverySelections } from "./delivery-selection"
import { writeLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { tagged } from "./prompts"
import { REQUEST_SOURCE, openItem, quoteMissIssue } from "./quotes"
import {
  FILES_REVISION_SCOPE, LEAN_REQUESTED_FILES_SCHEMA, parseRequestedFiles, readRequestedFiles, requestedFilesIssues, requestedFilesItems, requestedFilesPrompt, requestedFilesQuestion,
  requestedFilesRevision, withRequestedNames, writeRequestedFiles, type LeanRequestedFiles, type RequestedFile, type RequestedFiles,
} from "./requested-files"
import { writeLeanText, writeQuestions, type Question } from "./state"
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
  `files[${index}] gives the name "${name}", which is not a plain Word file name. Give the file name as the lawyer writes it, without a folder or Markdown, or null when the lawyer gives none.`
const nameNotQuoted = (index: number, name: string) =>
  `files[${index}] gives the name "${name}", but its quote does not give that name. Quote the lawyer's words that name this file, and give the name exactly as they write it, or null when the lawyer gives none.`

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
    // The job sees each planned document's id and title only, never the plan's files or deliverables.
    expect(prompt).toContain(JSON.stringify([{ id: "d1", title: "Share transfer form" }, { id: "d2", title: "Board minute" }], null, 2))
    expect(prompt).not.toMatch(/"(?:file|deliverable)":/)
    expect(prompt).toContain("<planned_memo>")
    expect(prompt).toContain(`"${MEMO}"`)
    expect(prompt).toContain(SHAPE)
    expect(prompt).not.toContain("<user_updates>")
    const plain = requestedFilesPrompt(directory, DOCUMENTS, undefined)
    expect(plain).not.toContain("<planned_memo>")
    // The request's own words aside, the prompt does not mention a memo.
    expect(plain.replace(tagged("request", REQUEST), "")).not.toMatch(/\bmemo\b/)
    expect(plain).toContain('Reply with the JSON object only: {"stated":true,"files":[{"holds":["planned id"],"name":"file name or null","quote":"..."}]}.')
    // The name is copied as the lawyer writes it; code adds Word's extension.
    expect(prompt).toContain('"name": the file name exactly as the lawyer writes it, or null when the lawyer gives no name; ')
    expect(prompt.replace(tagged("request", REQUEST), "")).not.toContain(".docx")
  })

  test("a file named for the whole request holds everything it asks for, a file named for one document only that document", () => {
    const directory = prepared()
    const rule = (contents: string) => "- A file the lawyer's words name as the output of the whole request holds everything the request asks for; " +
      `a file they name for one document holds only that document. Put ${contents} in a file only where the lawyer's words place it.`
    const prompt = requestedFilesPrompt(directory, DOCUMENTS, MEMO)
    expect(prompt).toContain(rule("a planned document or the memo"))
    expect(prompt).not.toContain("Named files are all the files")
    expect(requestedFilesPrompt(directory, DOCUMENTS, undefined)).toContain(rule("a planned document"))
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

  test("a quote that is not the lawyer's words is retried and still fails the last attempt, so no file is dropped unseen", () => {
    const directory = prepared()
    const wrong = file(["d1"], "Deliver the transfer form as its own file", "Kestrel Transfer.docx")
    const text = reply(true, file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), wrong)
    const miss = (index: number) => quoteMissIssue(directory, `files[${index}]`, { quote: wrong.quote, source_id: REQUEST_SOURCE }, "json")
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO)).toEqual({ issues: [miss(1)] })
    // As the coverage check's: the job is then unchecked, and disclosed (requestedFilesItems).
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO, { final: true })).toEqual({ issues: [miss(1)] })
    expect(parseRequestedFiles(reply(true, wrong), directory, DOCUMENTS, MEMO, { final: true })).toEqual({ issues: [miss(0)] })
  })

  test("a name its quote does not give, or that is not a plain Word file name, is retried; a last attempt keeps the file without it", () => {
    const directory = prepared()
    const text = reply(true, file(["memo", "d2"], PACK, "Kestrel Pack.docx"), file(["d1"], TRANSFER, "transfers/Kestrel Transfer.docx"))
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotQuoted(0, "Kestrel Pack.docx"), nameNotPlain(1, "transfers/Kestrel Transfer.docx")] })
    // The refused text is kept beside the file, for disclosure.
    expect(parseRequestedFiles(text, directory, DOCUMENTS, MEMO, { final: true })).toEqual({ issues: [], value: stated(
      { ...file(["memo", "d2"], PACK), refused_name: "Kestrel Pack.docx" }, { ...file(["d1"], TRANSFER), refused_name: "transfers/Kestrel Transfer.docx" }) })
  })

  test("a name is held to the lawyer's own text around the quote, whatever its case or line breaks", () => {
    const directory = prepared(REQUEST.replace("Kestrel Board Pack.docx", "Kestrel Board\nPack.docx"))
    const broken = PACK.replace("Kestrel Board Pack.docx", "Kestrel Board\nPack.docx")
    const parse = (...files: RequestedFile[]) => parseRequestedFiles(reply(true, ...files), directory, DOCUMENTS, MEMO)
    // A line break inside the name, a name in another case, and a quote that runs on past the name.
    expect(parse(file(["memo", "d2"], broken, "Kestrel Board Pack.docx"))).toEqual({ issues: [], value: stated(file(["memo", "d2"], broken, "Kestrel Board Pack.docx")) })
    expect(parse(file(["memo", "d2"], broken, "Kestrel board pack.docx"))).toEqual({ issues: [], value: stated(file(["memo", "d2"], broken, "Kestrel Board Pack.docx")) })
    const sentence = `${TRANSFER}.`
    expect(parse(file(["d1"], sentence, "Kestrel Transfer.docx"))).toEqual({ issues: [], value: stated(file(["d1"], sentence, "Kestrel Transfer.docx")) })
    const titled = "Put the cover memo and the board minute in one file titled Kestrel Board Pack, and send it to the board"
    const plain = prepared(`Prepare a board minute for Kestrel Works. ${titled}.`)
    expect(parseRequestedFiles(reply(true, file(["memo", "d2"], titled, "Kestrel Board Pack.docx")), plain, DOCUMENTS, MEMO))
      .toEqual({ issues: [], value: stated(file(["memo", "d2"], titled, "Kestrel Board Pack.docx")) })
  })

  test("a name cut short from a longer name in the lawyer's words is refused, though the quote ends where the name does", () => {
    const directory = prepared()
    const cut = (name: string) => file(["memo", "d2"], `Send the board minute with a cover memo, all in one file named ${name}`, `${name}.docx`)
    for (const name of ["Kestrel", "Kestrel Board"])
      expect(parseRequestedFiles(reply(true, cut(name)), directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotQuoted(0, `${name}.docx`)] })
    expect(parseRequestedFiles(reply(true, cut("Kestrel")), directory, DOCUMENTS, MEMO, { final: true }))
      .toEqual({ issues: [], value: stated({ ...cut("Kestrel"), name: null, refused_name: "Kestrel.docx" }) })
  })

  test("a name starts at a word's start in the lawyer's text, never inside a word", () => {
    for (const given of ["Draft-Kestrel Pack.docx", "Orchard's-Kestrel Pack.docx", "O'Kestrel Pack.docx", "v1.Kestrel Pack.docx"]) {
      const quote = `Send the board minute with a cover memo, all in one file named ${given}`
      const directory = prepared(`${quote}.`)
      expect(parseRequestedFiles(reply(true, file(["memo", "d2"], quote, "Kestrel Pack.docx")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotQuoted(0, "Kestrel Pack.docx")] })
    }
    // After a space, a colon, or an opening quotation mark or bracket.
    for (const given of ['"Kestrel Pack.docx"', "“Kestrel Pack.docx”", "'Kestrel Pack.docx'", "«Kestrel Pack.docx»", "(Kestrel Pack.docx)", "[Kestrel Pack.docx]", "name:Kestrel Pack.docx"]) {
      const quote = `Send the board minute with a cover memo, all in one file ${given}`
      const directory = prepared(`${quote}.`)
      expect(parseRequestedFiles(reply(true, file(["memo", "d2"], quote, "Kestrel Pack.docx")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: stated(file(["memo", "d2"], quote, "Kestrel Pack.docx")) })
    }
  })

  test("a name given without .docx ends at a line break, a closing mark, a colon, a sentence's end or a spaced dash", () => {
    const words = "Put the cover memo and the board minute in one file"
    const grounded = (request: string, quote: string) => {
      const directory = prepared(request)
      return parseRequestedFiles(reply(true, file(["memo", "d2"], quote, "Kestrel Board Pack.docx")), directory, DOCUMENTS, MEMO)
    }
    for (const [request, quote] of [
      [`Prepare a board minute.\n${words} titled Kestrel Board Pack\nSend it to the board.`, `${words} titled Kestrel Board Pack`],
      [`${words} titled Kestrel Board Pack  \r\nThen send it.`, `${words} titled Kestrel Board Pack`],
      [`${words} titled "Kestrel Board Pack." Then send it.`, `${words} titled "Kestrel Board Pack`],
      [`${words} titled “Kestrel Board Pack.” Then send it.`, `${words} titled “Kestrel Board Pack`],
      [`${words} (titled Kestrel Board Pack) and send it.`, `${words} (titled Kestrel Board Pack`],
      [`${words} titled Kestrel Board Pack: the memo first.`, `${words} titled Kestrel Board Pack`],
      [`${words} titled Kestrel Board Pack — the memo first.`, `${words} titled Kestrel Board Pack`],
      [`${words} titled Kestrel Board Pack - the memo first.`, `${words} titled Kestrel Board Pack`],
      [`${words} titled Kestrel Board Pack – the memo first.`, `${words} titled Kestrel Board Pack`],
    ]) expect(grounded(request!, quote!)).toEqual({ issues: [], value: stated(file(["memo", "d2"], quote!, "Kestrel Board Pack.docx")) })
    // A dash inside a word, or a space before another word, ends no name.
    for (const [request, quote] of [
      [`${words} titled Kestrel Board Pack-Final and send it.`, `${words} titled Kestrel Board Pack`],
      [`${words} titled Kestrel Board Pack Final\nThen send it.`, `${words} titled Kestrel Board Pack`],
    ]) expect(grounded(request!, quote!)).toEqual({ issues: [nameNotQuoted(0, "Kestrel Board Pack.docx")] })
  })

  test("the ordinary ways of naming a file ground a name given without .docx", () => {
    for (const naming of ["name it", "Name the file", "name this file", "name the document", "call it", "call the file", "call the document", "save it as", "save as",
      "save the file as", "title this draft", "File name:", "filename:"]) {
      const quote = `Put the cover memo and the board minute in one file and ${naming} Kestrel Board Pack`
      const directory = prepared(`${quote}.`)
      expect(parseRequestedFiles(reply(true, file(["memo", "d2"], quote, "Kestrel Board Pack.docx")), directory, DOCUMENTS, MEMO))
        .toEqual({ issues: [], value: stated(file(["memo", "d2"], quote, "Kestrel Board Pack.docx")) })
    }
  })

  test("a grounded name is kept in the lawyer's spelling, not the reply's", () => {
    const directory = prepared("Prepare a board minute. Put it with the cover memo in one file named Kestrel Pack.docx.")
    const quote = "Put it with the cover memo in one file named Kestrel Pack.docx"
    expect(parseRequestedFiles(reply(true, file(["memo", "d2"], quote, "kestrel pack.docx")), directory, DOCUMENTS, MEMO))
      .toEqual({ issues: [], value: stated(file(["memo", "d2"], quote, "Kestrel Pack.docx")) })
    const titled = prepared("Prepare a board minute. Put it with the cover memo in one file titled Kestrel Pack.")
    expect(parseRequestedFiles(reply(true, file(["memo", "d2"], "in one file titled Kestrel Pack", "KESTREL PACK.docx")), titled, DOCUMENTS, MEMO))
      .toEqual({ issues: [], value: stated(file(["memo", "d2"], "in one file titled Kestrel Pack", "Kestrel Pack.docx")) })
  })

  test("a name given without .docx may follow its naming word after a colon and an opening mark", () => {
    const words = "Put the board minute in a file of its own"
    const parse = (named: string) => {
      const quote = `${words}, ${named}`
      const directory = prepared(`${quote}.`)
      return { quote, result: parseRequestedFiles(reply(true, file(["d2"], quote, "Kestrel Minute")), directory, DOCUMENTS, MEMO) }
    }
    for (const named of ["save it as **Kestrel Minute**", "named `Kestrel Minute`", "named [Kestrel Minute]", "named (Kestrel Minute)", "titled: Kestrel Minute", "named: Kestrel Minute"]) {
      const { quote, result } = parse(named)
      expect(result).toEqual({ issues: [], value: stated(file(["d2"], quote, "Kestrel Minute.docx")) })
    }
    // A name cut short, or starting inside a word, is still refused.
    for (const named of ["save it as **Kestrel Minute Final**", "named `Kestrel Minutes`", "named **Draft-Kestrel Minute**"])
      expect(parse(named).result).toEqual({ issues: [nameNotQuoted(0, "Kestrel Minute")] })
  })

  test("a name the reply gives without a format's extension is a Word file's; another format the lawyer does not write is retried", () => {
    const directory = prepared()
    expect(parseRequestedFiles(reply(true, file(["memo", "d2"], PACK, "Kestrel Board Pack"), file(["d1"], TRANSFER, " kestrel transfer ")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: ASKED })
    expect(parseRequestedFiles(reply(true, file(["d1"], TRANSFER, "Kestrel Transfer.pdf")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotQuoted(0, "Kestrel Transfer.pdf")] })
    expect(parseRequestedFiles(reply(true, file(["d1"], TRANSFER, "Kestrel Transfer.doc")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [nameNotQuoted(0, "Kestrel Transfer.doc")] })
    // An ending that is no format's is part of the name.
    for (const name of ["Kestrel Holdings S.A", "Kestrel Minute rev.A"]) {
      const quote = `Put the cover memo and the board minute in one file named ${name}, with the memo first`
      expect(parseRequestedFiles(reply(true, file(["memo", "d2"], quote, name)), prepared(`${quote}.`), DOCUMENTS, MEMO))
        .toEqual({ issues: [], value: stated(file(["memo", "d2"], quote, `${name}.docx`)) })
    }
  })

  test("a name the lawyer writes in another format is refused at once in the lawyer's spelling, as no retry makes it a Word file's", () => {
    const quote = "Deliver the share transfer form separately as Kestrel Transfer.pdf"
    const directory = prepared(`${quote}.`)
    for (const name of ["Kestrel Transfer.pdf", "kestrel transfer.PDF"])
      expect(parseRequestedFiles(reply(true, file(["d1"], quote, name)), directory, DOCUMENTS, MEMO))
        .toEqual({ issues: [], value: stated({ ...file(["d1"], quote), refused_name: "Kestrel Transfer.pdf" }) })
  })

  test("a format is no name: a format given as a file's name is refused at once, and no refused name is made up", () => {
    for (const [words, name] of [["save it as PDF", "PDF"], ["save it as a PDF file", "a PDF file"], ["name the file rtf", "rtf"], ["save it as PDF (.pdf)", "PDF (.pdf)"]]) {
      const quote = `Deliver the share transfer form separately and ${words}`
      const directory = prepared(`${quote}.`)
      expect(parseRequestedFiles(reply(true, file(["d1"], quote, name)), directory, DOCUMENTS, MEMO))
        .toEqual({ issues: [], value: stated({ ...file(["d1"], quote), refused_name: name }) })
      expect(parseRequestedFiles(reply(true, file(["d1"], quote)), directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: stated(file(["d1"], quote)) })
    }
  })

  test("Word's own format given as a file's name asks for what the run delivers: no name, nothing refused and no item", () => {
    for (const [words, name] of [["save as Word", "Word"], ["save it as a Word document", "a Word document"], ["save it as a Word document", "Word document.docx"], ["name the file docx", "docx"],
      ["save it as a Word doc", "a Word doc"], ["save it as Word docx", "Word docx"], ["save it as .docx", ".docx"],
      ["save it as Word (.docx)", "Word (.docx)"], ["save it as Word (.docx)", "Word (.docx).docx"]]) {
      const quote = `Deliver the share transfer form separately and ${words}`
      const directory = prepared(`${quote}.`)
      for (const final of [false, true])
        expect(parseRequestedFiles(reply(true, file(["d1"], quote, name)), directory, DOCUMENTS, MEMO, { final })).toEqual({ issues: [], value: stated(file(["d1"], quote)) })
      writeLeanPlan(directory, [outline("Share transfer form"), outline("Board minute")], "basis-1")
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: { titles: ["Share transfer form", "Board minute"], memo: null }, ...stated(file(["d1"], quote)) })
      expect(requestedFilesItems(directory)).toEqual([])
    }
  })

  test("a name the lawyer spells with .docx is a name, though its words also name a format", () => {
    const quote = "Deliver the share transfer form separately and name the file Word Document.docx"
    const directory = prepared(`${quote}.`)
    for (const name of ["Word Document.docx", "Word Document"])
      expect(parseRequestedFiles(reply(true, file(["d1"], quote, name)), directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: stated(file(["d1"], quote, "Word Document.docx")) })
  })

  test("the files job and package review ground a name by one grammar", () => {
    const words = "Deliver the share transfer form separately"
    // The lawyer's words, the quote, the Word file name, and whether the files job grounds that name.
    const cases: [string, string, string, boolean][] = [
      [`${words} as Kestrel Transfer.docx.`, `${words} as Kestrel Transfer.docx`, "Kestrel Transfer.docx", true],
      [`${words}, titled Kestrel Transfer.`, `${words}, titled Kestrel Transfer`, "Kestrel Transfer.docx", true],
      ["Prepare the board minute. Name the file **Kestrel Board Pack**.", "Name the file **Kestrel Board Pack**", "Kestrel Board Pack.docx", true],
      [`${words}, named \`Kestrel Transfer\`.`, `${words}, named \`Kestrel Transfer`, "Kestrel Transfer.docx", true],
      [`${words}, titled: Kestrel Transfer`, `${words}, titled: Kestrel Transfer`, "Kestrel Transfer.docx", true],
      [`${words} (titled Kestrel Transfer) by Friday.`, `${words} (titled Kestrel Transfer`, "Kestrel Transfer.docx", true],
      [`${words}, titled “Kestrel Transfer.” Then send it.`, `${words}, titled “Kestrel Transfer`, "Kestrel Transfer.docx", true],
      [`${words}, titled Kestrel Transfer\nThen send it.`, `${words}, titled Kestrel Transfer`, "Kestrel Transfer.docx", true],
      [`${words}, titled Kestrel Transfer — signed copy to follow.`, `${words}, titled Kestrel Transfer`, "Kestrel Transfer.docx", true],
      [`${words} and name the file Word Document.docx.`, `${words} and name the file Word Document.docx`, "Word Document.docx", true],
      [`${words}, titled Kestrel Transfer Final.`, `${words}, titled Kestrel Transfer`, "Kestrel Transfer.docx", false],
      [`${words} as Draft-Kestrel Transfer.docx.`, `${words} as Draft-Kestrel Transfer.docx`, "Kestrel Transfer.docx", false],
      [`${words}, titled Kestrel Transfer.pdf.`, `${words}, titled Kestrel Transfer`, "Kestrel Transfer.docx", false],
      [`${words} and save it as PDF.`, `${words} and save it as PDF`, "PDF.docx", false],
      [`${words} and save it as a Word doc.`, `${words} and save it as a Word doc`, "a Word doc.docx", false],
      [`${words} and name the file docx.`, `${words} and name the file docx`, "docx.docx", false],
      [`${words} and save it as Kestrel Transfer.doc.`, `${words} and save it as Kestrel Transfer.doc`, "Kestrel Transfer.doc.docx", false],
      [`${words} and save it as Kestrel Transfer.pdf.`, `${words} and save it as Kestrel Transfer.pdf`, "Kestrel Transfer.pdf.docx", false],
    ]
    // Every name the files job grounds, package review accepts; it also accepts what its earlier literal or titled check did.
    for (const [request, quote, name, grounded] of cases) {
      const directory = prepared(request)
      writeLeanPlan(directory, [outline("Share transfer form")])
      const files = parseRequestedFiles(reply(true, file(["d1"], quote, name)), directory, DOCUMENTS, MEMO)
      const review = checkDeliverySelections(directory, [{ document: "d1", role: "body", file: name, instruction_quote: quote }])
      expect({ request, files: files.value?.files[0]?.name === name }).toEqual({ request, files: grounded })
      if (grounded) expect({ request, review }).toEqual({ request, review: [] })
    }
    // Intended difference: the files job keeps the lawyer's spelling of a name given in another case, while
    // package review, which selects the file delivered, takes the lawyer's spelling only.
    const directory = prepared(`${words}, titled Kestrel Transfer.`)
    writeLeanPlan(directory, [outline("Share transfer form")])
    const quote = `${words}, titled Kestrel Transfer`
    expect(parseRequestedFiles(reply(true, file(["d1"], quote, "kestrel transfer")), directory, DOCUMENTS, MEMO)).toEqual({ issues: [], value: stated(file(["d1"], quote, "Kestrel Transfer.docx")) })
    expect(checkDeliverySelections(directory, [{ document: "d1", role: "body", file: "kestrel transfer.docx", instruction_quote: quote }])).not.toEqual([])
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
  const revise = (issue: string, documents: DocID[]) => ({ issue, documents, ask: false })

  test("the memo with one document asks for that document combined and no other", () => {
    const issues = (documents: PlannedDocument[]) => requestedFilesIssues(documents, MEMO, ASKED)
    expect(issues(DOCUMENTS)).toEqual({ issues: [revise(`The lawyer's words "${PACK}" ask for the memo and "Board minute" in one file. Write "Deliverable: combined" under "Board minute".`, ["d2"])], unsupported: [] })
    expect(issues([planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")])).toEqual({ issues: [
      revise(`The lawyer's words "${PACK}" ask for the memo and "Board minute" in one file. Write "Deliverable: combined" under "Board minute" and "Deliverable: body" under "Share transfer form".`, ["d2", "d1"]),
      revise(`The lawyer's words "${TRANSFER}" ask for "Share transfer form" in a file of its own, without the memo. Write "Deliverable: body" under "Share transfer form" in place of "Deliverable: combined".`, ["d1"]),
    ], unsupported: [] })
    expect(issues([planned("d1", "Share transfer form", { deliverable: "redline" }), planned("d2", "Board minute", { deliverable: "combined" })])).toEqual({ issues: [], unsupported: [] })
  })

  test("the memo in a file of its own asks for no document combined", () => {
    const own = stated(file(["memo"], PACK))
    expect(requestedFilesIssues([planned("d1", "Share transfer form"), planned("d2", "Board minute", { deliverable: "combined" })], MEMO, own)).toEqual({ issues: [
      revise(`The lawyer's words "${PACK}" ask for the memo in a file of its own. Write "Deliverable: body" under "Board minute" in place of "Deliverable: combined".`, ["d2"]),
    ], unsupported: [] })
    expect(requestedFilesIssues(DOCUMENTS, MEMO, own)).toEqual({ issues: [], unsupported: [] })
  })

  test("a shape the run cannot build is unsupported, never an issue", () => {
    expect(requestedFilesIssues(DOCUMENTS, MEMO, stated(file(["d1", "d2"], BOTH), file(["d2", "memo"], PACK), file(["d1"], TRANSFER)))).toEqual({ issues: [], unsupported: [0, 1] })
    expect(requestedFilesIssues(DOCUMENTS, undefined, stated(file(["memo", "d2"], PACK)))).toEqual({ issues: [], unsupported: [0] })
    expect(requestedFilesIssues(DOCUMENTS, MEMO, stated(file(["d3"], TRANSFER)))).toEqual({ issues: [], unsupported: [0] })
    expect(requestedFilesIssues(DOCUMENTS, MEMO, { stated: false, files: [] })).toEqual({ issues: [], unsupported: [] })
  })

  test("the memo with a document planned as a redline is the lawyer's call: asked, never revised", () => {
    const directory = prepared()
    const redline = [planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute", { deliverable: "redline" })]
    const { issues } = requestedFilesIssues(redline, MEMO, stated(file(["memo", "d2"], PACK)))
    expect(issues).toEqual([{ issue: `The lawyer's words "${PACK}" ask for the memo and "Board minute" in one file; the plan delivers "Board minute" as a redline, which cannot be joined with the memo.`,
      documents: [], ask: true }])
    expect(requestedFilesRevision(directory, issues)).toBeUndefined()
    // Another file's issue is still revised, and only that one.
    const mixed = requestedFilesIssues(redline, MEMO, stated(file(["memo", "d2"], PACK), file(["d1"], TRANSFER))).issues
    expect(mixed.map(row => row.ask)).toEqual([true, false])
    expect(requestedFilesRevision(directory, mixed)).toEqual({ feedback: [FILES_REVISION_SCOPE, mixed[1]!.issue], documents: new Set(["d1"]) })
  })

  test("the revision request first says that only the delivery lines may change, and none is made once the lawyer has answered", () => {
    const directory = prepared()
    const { issues } = requestedFilesIssues(DOCUMENTS, MEMO, ASKED)
    // Code gives the files their names, so a revision keeps the "File:" lines as they are.
    expect(FILES_REVISION_SCOPE).toBe('These issues concern only how the outline delivers its files: change only the "Deliverable:" lines, ' +
      'and keep every title, part, "File:" line and the "Memo:" line exactly as they are.')
    expect(requestedFilesRevision(directory, issues)).toEqual({ feedback: [FILES_REVISION_SCOPE, issues[0]!.issue], documents: new Set(["d2"]) })
    // Any answered requested-files question, even one on the lawyer's earlier words, stops the revisions.
    writeQuestions(directory, [{ id: "requested-files-0000", stage: "draft", question: "Earlier question.", context: "", status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    expect(requestedFilesRevision(directory, issues)).toBeUndefined()
  })

  test("issues name documents by title, never by id", () => {
    const { issues } = requestedFilesIssues([planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")], MEMO, ASKED)
    expect(issues).toHaveLength(2)
    expect(issues.map(row => row.issue).join(" ")).not.toMatch(/\bd[12]\b/)
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

  test("the lawyer's name displaces another document's outline file name, which falls back to its title", () => {
    const documents = [outline("Share transfer form"), outline("Board minute", { file: "Kestrel Board Pack.docx" })]
    const memo = { name: MEMO, file: "Kestrel Board Pack - cover memo.docx" }
    expect(withRequestedNames(documents, memo, stated(file(["d1"], TRANSFER, "kestrel board pack.docx")))).toEqual({
      documents: [outline("Share transfer form", { file: "kestrel board pack.docx" }), outline("Board minute")],
      notes_file: "kestrel board pack - cover memo.docx",
      applied: ['"Share transfer form": file "kestrel board pack.docx"', '"Board minute": outline file "Kestrel Board Pack.docx" dropped, as the lawyer\'s words give that name to "Share transfer form"'],
    })
  })

  test("a name another planned file has is skipped and noted; an unbuildable file's name and a stated:false record change nothing", () => {
    const documents = [outline("Share transfer form"), outline("Board minute", { file: "Kestrel Board Pack.docx" })]
    const memo = { name: MEMO, file: "Kestrel Board Pack - cover memo.docx" }
    // A document named by its title, or the notes file, keeps its name: compared as registration names files.
    const untitled = documents.map(row => outline(row.title))
    expect(withRequestedNames(untitled, undefined, stated(file(["d1"], TRANSFER, "board MINUTE.docx")))).toEqual({
      documents: untitled, notes_file: undefined, applied: ['"board MINUTE.docx" not given to "Share transfer form": another planned file has that name'],
    })
    expect(withRequestedNames(untitled, undefined, stated(file(["d2"], PACK, "share transfer form - DRAFTING notes.docx")))).toEqual({
      documents: untitled, notes_file: undefined, applied: ['"share transfer form - DRAFTING notes.docx" not given to "Board minute": another planned file has that name'],
    })
    const notes = { name: MEMO, file: "Kestrel Notes.docx" }
    expect(withRequestedNames(untitled, notes, stated(file(["d2"], PACK, "kestrel notes.docx")))).toEqual({
      documents: untitled, notes_file: notes.file, applied: ['"kestrel notes.docx" not given to "Board minute": another planned file has that name'],
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
    expect(question).toEqual({ id: expect.stringMatching(/^requested-files-[0-9a-f]+$/), stage: "draft", question: ASK, context: expect.any(String), status: "open",
      kind: "file_name", options: [{ id: "o1", text: "As my words ask" }, { id: "o2", text: "As planned" }] })
    expect(question.question).not.toMatch(/\bd[12]\b/)
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)!.id).toBe(question.id)
    // Not asked again once open or answered.
    writeQuestions(directory, [question])
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)).toBeUndefined()
    writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, issues)).toBeUndefined()
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, [])).toBeUndefined()
    // A later job quoting the same words differently is not asked again either.
    const requoted = { ...ASKED, files: ASKED.files.map(file => ({ ...file, quote: `${file.quote} ` })) }
    expect(requestedFilesQuestion(directory, requoted, DOCUMENTS, requestedFilesIssues(DOCUMENTS, MEMO, requoted).issues)).toBeUndefined()
  })

  test("after an answer, only words no answered question listed are asked about, under the id of the lawyer's words as they stand", () => {
    const directory = prepared()
    const question = requestedFilesQuestion(directory, ASKED, DOCUMENTS, requestedFilesIssues(DOCUMENTS, MEMO, ASKED).issues)!
    writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    // An update that leaves these words as they were asks nothing again.
    writeLeanText(directory, "updated-notes.md", "<!-- update-1 -->\nAlso prepare a resignation letter.\n\n")
    expect(requestedFilesQuestion(directory, ASKED, DOCUMENTS, requestedFilesIssues(DOCUMENTS, MEMO, ASKED).issues)).toBeUndefined()
    // The answered question did not list the transfer form's words, so they are asked about.
    const documents = [planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")]
    const asked = requestedFilesQuestion(directory, ASKED, documents, requestedFilesIssues(documents, MEMO, ASKED).issues)!
    expect(asked.id).toMatch(/^requested-files-[0-9a-f]+$/)
    expect(asked.id).not.toBe(question.id)
    expect(asked.question).toBe(`Your words "${TRANSFER}" ask for "Share transfer form" as a file named "Kestrel Transfer.docx", but the document plan delivers "Share transfer form" in one file with the memo. ` +
      "Do you want it delivered as your words ask, or as planned? If neither, say how it should be delivered.")
  })

  test("the memo asked for with a document planned as a redline is asked about, saying the memo joins only a document's clean text", () => {
    const directory = prepared()
    const documents = [planned("d1", "Share transfer form"), planned("d2", "Board minute", { deliverable: "redline" })]
    const { issues } = requestedFilesIssues(documents, MEMO, ASKED)
    expect(requestedFilesQuestion(directory, ASKED, documents, issues)!.question).toBe(
      `Your words "${PACK}" ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", but the document plan delivers "Board minute" as a redline and the memo as a file of its own. ` +
      "The memo can be joined in one file only with a document's clean text, not with a redline. " +
      "Do you want it delivered as your words ask, or as planned? If neither, say how it should be delivered.")
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
  const PLANNED = { titles: ["Share transfer form", "Board minute"], memo: MEMO }
  const RECORD: LeanRequestedFiles = { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: BASIS, planned: PLANNED, ...ASKED }
  const AS_ASKED = [outline("Share transfer form", { file: "Kestrel Transfer.docx" }), outline("Board minute", { file: "Kestrel Board Pack.docx", deliverable: "combined" })]
  // The transfer form delivered under its title's name, not the lawyer's.
  const ELSEWHERE = [outline("Share transfer form"), AS_ASKED[1]!]
  const NOTES = { name: MEMO, file: "Kestrel Board Pack - cover memo.docx" }
  const CONFIRM = " The lawyer should confirm how these files are to be delivered."
  const ANSWERED = " The lawyer answered the question about these files; the plan drawn up with that answer still delivers them so."
  // Each item refers to its file by the file's index in the record (in RECORD, the board pack is file 0 and the transfer form
  // file 1) and lists the names it says the run delivers that file's contents in, null for a name not known before it registers.
  const item = (names: (string | null)[], detail: string, closing = CONFIRM, file = 0) =>
    openItem("finalize", "omission", `Requested files: ${detail}${closing}`, { repair_target: "memo", requested_file: file, delivered_names: names })
  const record = (...files: RequestedFile[]): LeanRequestedFiles => ({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: BASIS, planned: PLANNED, stated: true, files })

  function saved(documents = AS_ASKED, value: LeanRequestedFiles = RECORD, basis: string | null = BASIS, notes: { name: string; file: string } | undefined = NOTES) {
    const directory = prepared()
    writeLeanPlan(directory, documents, basis ?? undefined, undefined, notes)
    writeRequestedFiles(directory, value)
    return directory
  }

  test("the record reads back as written; a malformed record reads as none", () => {
    const directory = saved()
    expect(readRequestedFiles(directory)).toEqual(RECORD)
    writeLeanText(directory, "requested-files.json", JSON.stringify({ ...RECORD, files: [{ holds: ["d1"], name: 3, quote: TRANSFER }] }))
    expect(readRequestedFiles(directory)).toBeUndefined()
    expect(requestedFilesItems(directory)).toEqual([])
    // A record without the planned documents it was read for is no record.
    const { planned: _, ...unkeyed } = RECORD
    writeLeanText(directory, "requested-files.json", JSON.stringify(unkeyed))
    expect(readRequestedFiles(directory)).toBeUndefined()
  })

  test("a plan that delivers the files as asked gives no item", () => {
    expect(requestedFilesItems(saved())).toEqual([])
  })

  test("a plan changed after the record to deliver the memo apart gives one item in the lawyer's words", () => {
    const directory = saved([AS_ASKED[0]!, outline("Board minute", { file: "Kestrel Board Pack.docx", deliverable: "body" })])
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Board Pack - cover memo.docx", "Kestrel Board Pack.docx"], `the lawyer's words "${PACK}" [source: request] ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", ` +
      'but the run delivers the memo in "Kestrel Board Pack - cover memo.docx" and "Board minute" in "Kestrel Board Pack.docx".')])
  })

  test("once the lawyer has answered the question on these words, the item cites the answer", () => {
    const directory = saved([AS_ASKED[0]!, outline("Board minute", { file: "Kestrel Board Pack.docx", deliverable: "body" })])
    const question = requestedFilesQuestion(directory, ASKED, DOCUMENTS, requestedFilesIssues(DOCUMENTS, MEMO, ASKED).issues)!
    writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    const detail = `the lawyer's words "${PACK}" [source: request] ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", ` +
      'but the run delivers the memo in "Kestrel Board Pack - cover memo.docx" and "Board minute" in "Kestrel Board Pack.docx".'
    // The answer is cited in plain words: no question id reaches the lawyer.
    const cited = item(["Kestrel Board Pack - cover memo.docx", "Kestrel Board Pack.docx"], detail, ANSWERED)
    expect(requestedFilesItems(directory)).toEqual([cited])
    expect(cited.detail).not.toContain(question.id)
    // An update that leaves these words as they were leaves the answer standing.
    writeLeanText(directory, "updated-notes.md", "<!-- update-1 -->\nAlso prepare a resignation letter.\n\n")
    expect(requestedFilesItems(directory)).toEqual([cited])
  })

  test("an answer to a question that listed more or fewer of the lawyer's words is cited with the words it listed", () => {
    const fewer = "with a cover memo, all in one file named Kestrel Board Pack.docx"
    const plan = [AS_ASKED[0]!, outline("Board minute", { file: "Kestrel Board Pack.docx", deliverable: "body" })]
    for (const [listed, recorded] of [[fewer, PACK], [PACK, fewer]] as const) {
      const directory = saved(plan, record(file(["memo", "d2"], recorded, "Kestrel Board Pack.docx")))
      const listing = stated(file(["memo", "d2"], listed, "Kestrel Board Pack.docx"))
      const question = requestedFilesQuestion(directory, listing, DOCUMENTS, requestedFilesIssues(DOCUMENTS, MEMO, listing).issues)!
      writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
      expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Board Pack - cover memo.docx", "Kestrel Board Pack.docx"], `the lawyer's words "${recorded}" [source: request] ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", ` +
        'but the run delivers the memo in "Kestrel Board Pack - cover memo.docx" and "Board minute" in "Kestrel Board Pack.docx".',
      ` The lawyer answered the question about the files the words "${listed}" [source: request] ask for; the plan drawn up with that answer still delivers them so.`)])
    }
  })

  test("an answer is cited only for the files its question asked about, while the plan delivers them as the question showed", () => {
    const directory = saved([outline("Share transfer form", { deliverable: "combined" }), outline("Board minute", { file: "Kestrel Board Pack.docx" })])
    const answer = (question: Question) => writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    const combined = 'the memo and "Share transfer form" in one file named "Share transfer form.docx"'
    const pack = `the lawyer's words "${PACK}" [source: request] ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", ` +
      `but the run delivers ${combined} and "Board minute" in "Kestrel Board Pack.docx".`
    const transfer = item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file named "Kestrel Transfer.docx", but the run delivers ${combined}.`, CONFIRM, 1)
    // Asked over the plan as it stands, while the files job had recorded the board pack alone, the question lists only the board pack.
    const current = [planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute", { file: "Kestrel Board Pack.docx" })]
    const alone = stated(ASKED.files[0]!)
    answer(requestedFilesQuestion(directory, alone, current, requestedFilesIssues(current, MEMO, alone).issues)!)
    expect(requestedFilesItems(directory)).toEqual([item(["Share transfer form.docx", "Kestrel Board Pack.docx"], pack, ANSWERED), transfer])
    // Asked over an earlier outline, which delivered the board pack's memo and minute as separate files, the answer is not cited
    // now that the plan joins the memo with the transfer form instead.
    const question = requestedFilesQuestion(prepared(), ASKED, DOCUMENTS, requestedFilesIssues(DOCUMENTS, MEMO, ASKED).issues)!
    expect(question.question).not.toContain(TRANSFER)
    answer(question)
    expect(requestedFilesItems(directory)).toEqual([item(["Share transfer form.docx", "Kestrel Board Pack.docx"], pack), transfer])
  })

  test("of two files the same words ask for, an answer is cited only for the one its question asked about", () => {
    const words = "Send the cover memo on its own and the board minute as Kestrel Minute.docx"
    const request = `Prepare a share transfer form and a board minute for Kestrel Works. ${words}.`
    const shared = record(file(["memo"], words), file(["d2"], words, "Kestrel Minute.docx"))
    const directory = saved([outline("Share transfer form"), outline("Board minute", { file: "Kestrel Minute.docx", deliverable: "combined" })], shared)
    writeLeanText(directory, "request.md", request)
    const answer = (question: Question) => writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    const delivered = 'but the run delivers the memo and "Board minute" in one file named "Kestrel Minute.docx".'
    const memo = `the lawyer's words "${words}" [source: request] ask for the memo as a file of its own, ${delivered}`
    const minute = item(["Kestrel Minute.docx"], `the lawyer's words "${words}" [source: request] ask for "Board minute" as a file named "Kestrel Minute.docx", ${delivered}`, CONFIRM, 1)
    // Asked over the plan as it stands, while the files job had recorded the memo's file alone, the question lists only the memo's file.
    const current = [planned("d1", "Share transfer form"), planned("d2", "Board minute", { file: "Kestrel Minute.docx", deliverable: "combined" })]
    const alone = stated(shared.files[0]!)
    answer(requestedFilesQuestion(directory, alone, current, requestedFilesIssues(current, MEMO, alone).issues)!)
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Minute.docx"], memo, ANSWERED), minute])
    // Asked over an earlier outline that joined the memo with the transfer form, the question lists only the memo's file too;
    // its answer is not cited now that the plan joins the memo with the board minute.
    const earlier = [planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")]
    const question = requestedFilesQuestion(prepared(request), shared, earlier, requestedFilesIssues(earlier, MEMO, shared).issues)!
    expect(question.question).not.toContain('ask for "Board minute"')
    answer(question)
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Minute.docx"], memo), minute])
    // Once an update changes the lawyer's words, the file no answered question asked about is asked about.
    writeLeanText(directory, "updated-notes.md", "<!-- update-1 -->\nAlso prepare a resignation letter.\n\n")
    expect(requestedFilesQuestion(directory, shared, current, requestedFilesIssues(current, MEMO, shared).issues)!.question).toBe(
      `Your words "${words}" ask for "Board minute" as a file named "Kestrel Minute.docx", but the document plan delivers "Board minute" in one file with the memo. ` +
      "Do you want it delivered as your words ask, or as planned? If neither, say how it should be delivered.")
  })

  test("an answered file is not asked about again once a later outline renames its document or a later job drops its name; its answer is cited while the plan delivers it as the question showed", () => {
    const documents = [outline("Share transfer form", { deliverable: "combined" }), outline("Board resolution", { file: "Kestrel Board Pack.docx" })]
    const later = { ...record(file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), file(["d1"], TRANSFER)), planned: { titles: ["Share transfer form", "Board resolution"], memo: MEMO } }
    const directory = saved(documents, later)
    // Asked over an outline that already joined the memo with the transfer form and titled the minute "Board minute".
    const earlier = [planned("d1", "Share transfer form", { deliverable: "combined" }), planned("d2", "Board minute")]
    const question = requestedFilesQuestion(directory, ASKED, earlier, requestedFilesIssues(earlier, MEMO, ASKED).issues)!
    expect(question.question).toContain(`Your words "${PACK}" ask for the memo and "Board minute"`)
    expect(question.question).toContain(`Your words "${TRANSFER}" ask for "Share transfer form" as a file named "Kestrel Transfer.docx"`)
    writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    // After an update, neither file is asked about again: the lawyer's words were answered.
    writeLeanText(directory, "updated-notes.md", "<!-- update-1 -->\nAlso prepare a resignation letter.\n\n")
    const current = documents.map((row, index) => ({ ...row, id: `d${index + 1}` as DocID }))
    expect(requestedFilesQuestion(directory, later, current, requestedFilesIssues(current, MEMO, later).issues)).toBeUndefined()
    // The transfer form is delivered as the question showed, so its answer is cited; the board pack's delivery now reads
    // otherwise, under the new title, so the lawyer confirms it.
    const combined = 'the memo and "Share transfer form" in one file named "Share transfer form.docx"'
    expect(requestedFilesItems(directory)).toEqual([
      item(["Share transfer form.docx", "Kestrel Board Pack.docx"], `the lawyer's words "${PACK}" [source: request] ask for the memo and "Board resolution" in one file named "Kestrel Board Pack.docx", ` +
        `but the run delivers ${combined} and "Board resolution" in "Kestrel Board Pack.docx".`),
      item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, but the run delivers ${combined}.`, ANSWERED, 1),
    ])
  })

  test("a document delivered under another name gives an item naming the file delivered", () => {
    const directory = saved([outline("Share transfer form"), AS_ASKED[1]!])
    expect(requestedFilesItems(directory)).toEqual([item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file named "Kestrel Transfer.docx", ` +
      'but the run delivers "Share transfer form" in "Share transfer form.docx".', CONFIRM, 1)])
  })

  test("a file the record lists twice, alike in every field, gives one item, about the first", () => {
    const twice = file(["d1"], TRANSFER, "Kestrel Transfer.docx")
    const directory = saved([outline("Share transfer form"), AS_ASKED[1]!], record(twice, { ...twice }))
    expect(requestedFilesItems(directory)).toEqual([item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" ` +
      'as a file named "Kestrel Transfer.docx", but the run delivers "Share transfer form" in "Share transfer form.docx".')])
  })

  test("a matter file the run did not write at the requested name is not delivery as asked", () => {
    const directory = saved()
    writeFileSync(path.join(directory, "Kestrel Transfer.docx"), "a synthetic document the lawyer saved")
    const asked = `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file named "Kestrel Transfer.docx", but the run delivers "Share transfer form" in `
    expect(requestedFilesItems(directory)).toEqual([item([null], `${asked}a file of another name, since a file already in the matter holds the name "Kestrel Transfer.docx".`, CONFIRM, 1)])
    // The run's own registered file there is its delivery; an earlier suffixed delivery is named.
    const row = (deliverable: string, sha256: string) => ({ kind: "body", document: "d1", title: "Share transfer form", path: ".dochaus/lean/d1/body-applied.docx", sha256, deliverable })
    const summary = (...documents: unknown[]) => writeLeanText(directory, "documents.json", JSON.stringify({ schema: "dochaus.lean.documents/2", status: "completed", documents }))
    // The item reads the same once the run has registered its file under another name.
    summary(row("Kestrel Transfer (2).docx", "0".repeat(64)))
    expect(requestedFilesItems(directory)).toEqual([item([null], `${asked}a file of another name, since a file already in the matter holds the name "Kestrel Transfer.docx".`, CONFIRM, 1)])
    summary(row("Kestrel Transfer.docx", sha256File(path.join(directory, "Kestrel Transfer.docx"))!.sha256))
    expect(requestedFilesItems(directory)).toEqual([])
    // The lawyer's accepted Word revision of the run's file keeps that file the run's.
    summary({ ...row("Kestrel Transfer.docx", "0".repeat(64)), conflict: "accepted_revision" })
    expect(requestedFilesItems(directory)).toEqual([])
  })

  test("the run's earlier suffixed delivery of a document is predicted first, as registration tries it first", () => {
    const directory = saved()
    const row = (deliverable: string, sha256: string) => ({ kind: "body", document: "d1", title: "Share transfer form", path: ".dochaus/lean/d1/body-applied.docx", sha256, deliverable })
    const summary = (...documents: unknown[]) => writeLeanText(directory, "documents.json", JSON.stringify({ schema: "dochaus.lean.documents/2", status: "completed", documents }))
    // The requested name is free, but the run's earlier delivery still holds its bytes under "(2)".
    writeFileSync(path.join(directory, "Kestrel Transfer (2).docx"), "the run's earlier synthetic delivery")
    summary(row("Kestrel Transfer (2).docx", sha256File(path.join(directory, "Kestrel Transfer (2).docx"))!.sha256))
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Transfer (2).docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file named "Kestrel Transfer.docx", ` +
      'but the run delivers "Share transfer form" in "Kestrel Transfer (2).docx".', CONFIRM, 1)])
    // A file of the lawyer's under that earlier name is skipped, as registration skips it.
    writeFileSync(path.join(directory, "Kestrel Transfer (2).docx"), "a synthetic document the lawyer saved")
    expect(requestedFilesItems(directory)).toEqual([])
  })

  test("a run file the lawyer's accepted Word revision or its notes addendum changed is still the run's", () => {
    // The memo alone, asked for under the notes file's name.
    const documents = [outline("Share transfer form", { file: "Kestrel Transfer.docx" }), outline("Board minute")]
    const directory = saved(documents, record(file(["memo"], PACK, "Kestrel Board Pack.docx")), BASIS, { name: MEMO, file: "Kestrel Board Pack.docx" })
    const notes = path.join(directory, "Kestrel Board Pack.docx")
    writeFileSync(notes, "the run's synthetic notes")
    writeLeanText(directory, "documents.json", JSON.stringify({ schema: "dochaus.lean.documents/2", status: "completed", documents: [
      { kind: "drafting_notes", document: null, title: null, path: ".dochaus/lean/drafting-notes.docx", sha256: sha256File(notes)!.sha256, deliverable: "Kestrel Board Pack.docx" },
    ] }))
    expect(requestedFilesItems(directory)).toEqual([])
    // An addendum after a Word accept rewrites the notes as an accepted revision before the summary is next written.
    writeFileSync(notes, "the run's synthetic notes, with an addendum")
    const db = new Database(path.join(directory, ".dochaus", "legal.db"), { create: true })
    db.run("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
    db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run(workingDocxKey(directory, notes), JSON.stringify({ version: 1, source_sha256: "1".repeat(64), projection_sha256: "2".repeat(64) }))
    db.close()
    expect(requestedFilesItems(directory)).toEqual([])
  })

  test("a redline the lawyer's accepted Word revision of its body rebuilt is still the run's, while it is a redline of that body", async () => {
    const plan = [outline("Share transfer form", { file: "Kestrel Transfer.docx", deliverable: "redline" }), AS_ASKED[1]!]
    const directory = saved(plan, record(file(["d1"], TRANSFER, "Kestrel Transfer.docx")))
    const clean = path.join(directory, "Kestrel Transfer - clean.docx")
    const redline = path.join(directory, "Kestrel Transfer.docx")
    const original = (await renderMarkdownToDocx("# Share transfer form\n\nThe transferor transfers ten shares.")).bytes
    const accepted = (await renderMarkdownToDocx("# Share transfer form\n\nThe transferor transfers twelve shares.")).bytes
    // The accept rewrites the body and rebuilds the redline in place; the summary still holds the bytes first registered.
    const rebuilt = await (await docxodus()).compareDocuments(original, accepted, { authorName: "Reviewer" })
    writeFileSync(clean, accepted)
    writeFileSync(redline, rebuilt)
    const db = new Database(path.join(directory, ".dochaus", "legal.db"), { create: true })
    db.run("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
    db.run("CREATE TABLE IF NOT EXISTS documents (doc_path TEXT, role TEXT)")
    const revision = JSON.stringify({ version: 1, source_sha256: "1".repeat(64), projection_sha256: "2".repeat(64) })
    db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run(workingDocxKey(directory, clean), revision)
    for (const name of [clean, redline]) db.query("INSERT INTO documents (doc_path, role) VALUES (?, 'deliverable')").run(name)
    const row = (kind: string, file: string, deliverable: string) => ({ kind, document: "d1", title: "Share transfer form", path: `.dochaus/lean/d1/${file}`, sha256: "0".repeat(64), deliverable })
    const summary = (redlineRow: object) => writeLeanText(directory, "documents.json", JSON.stringify({ schema: "dochaus.lean.documents/2", status: "completed",
      documents: [row("body", "body-applied.docx", "Kestrel Transfer - clean.docx"), redlineRow] }))
    summary(row("redline", "redline.docx", "Kestrel Transfer.docx"))
    expect(requestedFilesItems(directory)).toEqual([])
    // A file there that is no redline of the body is not the run's.
    writeFileSync(redline, accepted)
    const held = [item([null], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file named "Kestrel Transfer.docx", ` +
      'but the run delivers "Share transfer form" as a redline in a file of another name, since a file already in the matter holds the name "Kestrel Transfer.docx".')]
    expect(requestedFilesItems(directory)).toEqual(held)
    // A redline row that records the kept accept owns the file no more than registration does: only while it is a redline of the body.
    summary({ ...row("redline", "redline.docx", "Kestrel Transfer.docx"), conflict: "accepted_revision" })
    expect(requestedFilesItems(directory)).toEqual(held)
    writeFileSync(redline, rebuilt)
    expect(requestedFilesItems(directory)).toEqual([])
    // With its body no longer kept, a redline is kept at its own accepted Word revision, as registration keeps it.
    writeFileSync(redline, accepted)
    db.query("DELETE FROM meta WHERE key = ?").run(workingDocxKey(directory, clean))
    db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run(workingDocxKey(directory, redline), revision)
    db.close()
    expect(requestedFilesItems(directory)).toEqual([])
  })

  test("each delivered file is described by all it holds", () => {
    // The combined file holds the memo and its document, whatever the record asks of either.
    const combined = saved([outline("Share transfer form", { deliverable: "combined" }), outline("Board minute")], record(file(["d1"], TRANSFER)))
    expect(requestedFilesItems(combined)).toEqual([item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run delivers the memo and "Share transfer form" in one file named "Share transfer form.docx".')])
    // A memo-only file under a name a document holds takes a suffix, spelled as planned.
    const suffixed = saved([outline("Share transfer form", { file: "Kestrel Pack.docx" }), outline("Board minute")], record(file(["memo"], PACK, "Kestrel Pack.docx")), BASIS, { name: MEMO, file: "Kestrel Pack.docx" })
    expect(requestedFilesItems(suffixed)).toEqual([item(["Kestrel Pack (2).docx"], `the lawyer's words "${PACK}" [source: request] ask for the memo as a file named "Kestrel Pack.docx", but the run delivers the memo in "Kestrel Pack (2).docx".`)])
  })

  test("the memo with a document planned as a redline gives an item naming the redline, citing the answer once given", () => {
    const documents = [AS_ASKED[0]!, outline("Board minute", { deliverable: "redline" })]
    const directory = saved(documents)
    const detail = `the lawyer's words "${PACK}" [source: request] ask for the memo and "Board minute" in one file named "Kestrel Board Pack.docx", ` +
      'but the run delivers the memo in "Kestrel Board Pack - cover memo.docx" and "Board minute" as a redline in "Board minute.docx". The run cannot join the memo with a redline.'
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Board Pack - cover memo.docx", "Board minute.docx"], detail)])
    const ids = documents.map((row, index) => ({ ...row, id: `d${index + 1}` as DocID }))
    const question = requestedFilesQuestion(directory, ASKED, ids, requestedFilesIssues(ids, MEMO, ASKED).issues)!
    writeQuestions(directory, [{ ...question, status: "answered", answer: "As planned.", answered_at: "2026-10-18T00:00:00.000Z" }])
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Board Pack - cover memo.docx", "Board minute.docx"], detail, ANSWERED)])
  })

  test("a name the run could not confirm is disclosed with the file delivered, whatever that file's name", () => {
    const refused = { ...file(["d1"], TRANSFER), refused_name: "Kestrel Transfer Form.docx" }
    const unconfirmed = (name: string, delivered: string) => item([delivered], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      `but the run could not use "${name}" as its name; it delivers "Share transfer form" in "${delivered}".`)
    expect(requestedFilesItems(saved(ELSEWHERE, record(refused)))).toEqual([unconfirmed("Kestrel Transfer Form.docx", "Share transfer form.docx")])
    // Delivered under the name the lawyer's words give, or under that very name, it is still disclosed: the run could not confirm it.
    expect(requestedFilesItems(saved(AS_ASKED, record(refused)))).toEqual([unconfirmed("Kestrel Transfer Form.docx", "Kestrel Transfer.docx")])
    expect(requestedFilesItems(saved(AS_ASKED, record({ ...refused, refused_name: "Kestrel Transfer.docx" })))).toEqual([unconfirmed("Kestrel Transfer.docx", "Kestrel Transfer.docx")])
  })

  test("a format refused as a name gives an item, though the lawyer's words give it after other words and the file is delivered under a title they give", () => {
    const words = "Deliver the share transfer form as a PDF titled Kestrel Transfer"
    const directory = saved()
    writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${words}.`)
    const parsed = parseRequestedFiles(reply(true, file(["d1"], words, "PDF")), directory, DOCUMENTS, MEMO, { final: true })
    expect(parsed).toEqual({ issues: [], value: stated({ ...file(["d1"], words), refused_name: "PDF" }) })
    writeRequestedFiles(directory, record(...parsed.value!.files))
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${words}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "PDF" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
  })

  test("a refused name gives an item, though the file delivered is named after the end of a longer name the lawyer's words give", () => {
    const words = "Deliver the share transfer form separately as Kestrel Share Transfer Form.docx"
    const directory = saved(ELSEWHERE)
    writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${words}.`)
    const parsed = parseRequestedFiles(reply(true, file(["d1"], words, "Kestrel Share Transfer.docx")), directory, DOCUMENTS, MEMO, { final: true })
    expect(parsed).toEqual({ issues: [], value: stated({ ...file(["d1"], words), refused_name: "Kestrel Share Transfer.docx" }) })
    writeRequestedFiles(directory, record(...parsed.value!.files))
    expect(requestedFilesItems(directory)).toEqual([item(["Share transfer form.docx"], `the lawyer's words "${words}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Share Transfer.docx" as its name; it delivers "Share transfer form" in "Share transfer form.docx".')])
  })

  test("every name the run could not confirm gives one item, whatever the lawyer's words say and the plan delivers", () => {
    const quotes = [TRANSFER, "Deliver the share transfer form as a PDF titled Kestrel Transfer", "Deliver the PDF-ready share transfer form separately as Kestrel Transfer.docx",
      "Deliver the share transfer form separately and save it as a PDF", "Deliver the share transfer form separately"]
    // Each refused name, and whether it is a format or another format's file name, of which the item says the run delivers .docx files only.
    const names: [string, boolean][] = [["PDF", true], ["Word", true], ["a PDF", true], ["Kestrel Transfer.pdf", true], ["Kestrel Transfer.docx", false], ["Kestrel Transfer Form.docx", false]]
    for (const documents of [AS_ASKED, ELSEWHERE]) for (const quote of quotes) for (const [name, format] of names) {
      const directory = saved(documents, record({ ...file(["d1"], quote), refused_name: name }))
      writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${quote}.`)
      const items = requestedFilesItems(directory)
      const delivered = documents === AS_ASKED ? "Kestrel Transfer.docx" : "Share transfer form.docx"
      expect({ quote, name, details: items.map(row => row.detail) }).toEqual({ quote, name, details: [item([delivered], `the lawyer's words "${quote}" [source: request] ask for "Share transfer form" as a file of its own, ` +
        `but the run could not use "${name}" as its name${format ? " and delivers .docx files only" : ""}; it delivers "Share transfer form" in "${delivered}".`).detail] })
    }
  })

  test("a refused name that shares the delivered file's stem, or is that file's very name, is still disclosed", () => {
    // The lawyer asks for Kestrel Transfer v2.docx; the reply's name, refused, shares the delivered file's stem.
    const v2 = "Deliver the share transfer form separately as Kestrel Transfer v2.docx"
    let directory = saved(AS_ASKED, record({ ...file(["d1"], v2), refused_name: "Kestrel Transfer.pdf" }))
    writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${v2}.`)
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${v2}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Transfer.pdf" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
    // The lawyer asks for Kestrel Transfer.doc; the last attempt gives the bare stem, which code makes a Word file's name.
    const doc = "Deliver the share transfer form separately and save it as Kestrel Transfer.doc"
    directory = saved()
    writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${doc}.`)
    const last = parseRequestedFiles(reply(true, file(["d1"], doc, "Kestrel Transfer")), directory, DOCUMENTS, MEMO, { final: true })
    expect(last).toEqual({ issues: [], value: stated({ ...file(["d1"], doc), refused_name: "Kestrel Transfer.docx" }) })
    writeRequestedFiles(directory, record(...last.value!.files))
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${doc}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Transfer.docx" as its name; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
  })

  test("a format given as a name is refused at once and disclosed, wherever the lawyer's words hold it", () => {
    // "PDF" in a description, given by the reply as a name, is refused, and disclosed as a name the run could not confirm.
    const ready = "Deliver the PDF-ready share transfer form separately as Kestrel Transfer.docx"
    const directory = saved()
    writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${ready}.`)
    const parsed = parseRequestedFiles(reply(true, file(["d1"], ready, "PDF")), directory, DOCUMENTS, MEMO)
    expect(parsed).toEqual({ issues: [], value: stated({ ...file(["d1"], ready), refused_name: "PDF" }) })
    writeRequestedFiles(directory, record(...parsed.value!.files))
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${ready}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "PDF" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
    // Asked for after a naming word, it is disclosed alike.
    const save = "Deliver the share transfer form separately and save it as a PDF"
    for (const name of ["PDF", "a PDF"]) {
      const saving = saved(AS_ASKED, record({ ...file(["d1"], save), refused_name: name }))
      writeLeanText(saving, "request.md", `${BOTH}. ${PACK}. ${save}.`)
      expect(requestedFilesItems(saving)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${save}" [source: request] ask for "Share transfer form" as a file of its own, ` +
        `but the run could not use "${name}" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".`)])
    }
  })

  test("a file asked for in another format is disclosed with the .docx files the run delivers, though it delivers a Word file of the same stem", () => {
    const words = (quote: string, name: string) => {
      const directory = saved(AS_ASKED, record({ ...file(["d1"], quote), refused_name: name }))
      writeLeanText(directory, "request.md", `${BOTH}. ${PACK}. ${quote}.`)
      return requestedFilesItems(directory)
    }
    for (const extension of [".pdf", ".doc"]) {
      const quote = TRANSFER.replace(".docx", extension)
      const name = `Kestrel Transfer${extension}`
      expect(words(quote, name)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${quote}" [source: request] ask for "Share transfer form" as a file of its own, ` +
        `but the run could not use "${name}" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".`)])
    }
    // A format the reply gives alone, where the lawyer's words give it only as a name's extension, is disclosed alike.
    const quote = TRANSFER.replace(".docx", ".pdf")
    expect(words(quote, "PDF")).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${quote}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "PDF" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
  })

  test("a refused name in another format the lawyer's words do not give is disclosed as unconfirmed, never as a format the lawyer asked for", () => {
    // The lawyer asks for Kestrel Transfer.docx; the last attempt answers in another format.
    const directory = saved()
    const last = parseRequestedFiles(reply(true, file(["d1"], TRANSFER, "Kestrel Transfer.pdf")), directory, DOCUMENTS, MEMO, { final: true })
    expect(last).toEqual({ issues: [], value: stated({ ...file(["d1"], TRANSFER), refused_name: "Kestrel Transfer.pdf" }) })
    writeRequestedFiles(directory, record(...last.value!.files))
    // Delivered under the name the lawyer gave, it is still disclosed, whatever name the reply gave.
    expect(requestedFilesItems(directory)).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Transfer.pdf" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
    expect(requestedFilesItems(saved(AS_ASKED, record({ ...file(["d1"], TRANSFER), refused_name: "Kestrel Form.pdf" })))).toEqual([item(["Kestrel Transfer.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Form.pdf" as its name and delivers .docx files only; it delivers "Share transfer form" in "Kestrel Transfer.docx".')])
    expect(requestedFilesItems(saved(ELSEWHERE, record({ ...file(["d1"], TRANSFER), refused_name: "Kestrel Form.pdf" })))).toEqual([item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Form.pdf" as its name and delivers .docx files only; it delivers "Share transfer form" in "Share transfer form.docx".')])
  })

  test("a name refused on an earlier attempt stays disclosed when the last attempt gives none, until a name passes", () => {
    const directory = saved(ELSEWHERE)
    const refused = new Map<string, string>()
    const parse = (name: string | null, final = false) =>
      parseRequestedFiles(reply(true, file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), file(["d1"], ` ${TRANSFER} `, name)), directory, DOCUMENTS, MEMO, { final, refused })
    expect(parse("Kestrel Transfer Form.docx")).toEqual({ issues: [nameNotQuoted(1, "Kestrel Transfer Form.docx")] })
    const last = parse(null, true)
    expect(last).toEqual({ issues: [], value: stated(file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), { ...file(["d1"], TRANSFER), refused_name: "Kestrel Transfer Form.docx" }) })
    writeRequestedFiles(directory, record(...last.value!.files))
    expect(requestedFilesItems(directory)).toEqual([item(["Share transfer form.docx"], `the lawyer's words "${TRANSFER}" [source: request] ask for "Share transfer form" as a file of its own, ` +
      'but the run could not use "Kestrel Transfer Form.docx" as its name; it delivers "Share transfer form" in "Share transfer form.docx".', CONFIRM, 1)])
    // A later name the lawyer's words give clears it.
    expect(parse("Kestrel Transfer.docx")).toEqual({ issues: [], value: ASKED })
    expect(parse(null, true)).toEqual({ issues: [], value: stated(file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"), file(["d1"], TRANSFER)) })
  })

  test("a name refused for one file is kept for that file alone, though another file shares its quote", () => {
    const directory = prepared()
    const refused = new Map<string, string>()
    const parse = (final: boolean, ...files: RequestedFile[]) => parseRequestedFiles(reply(true, ...files), directory, DOCUMENTS, MEMO, { final, refused })
    expect(parse(false, file(["d1"], BOTH, "Kestrel Form.docx"), file(["d2"], BOTH))).toEqual({ issues: [nameNotQuoted(0, "Kestrel Form.docx")] })
    expect(parse(true, file(["d1"], BOTH), file(["d2"], BOTH))).toEqual({ issues: [], value: stated({ ...file(["d1"], BOTH), refused_name: "Kestrel Form.docx" }, file(["d2"], BOTH)) })
    // A name that passes for the other file does not clear it.
    refused.clear()
    expect(parse(false, file(["d1"], PACK, "Kestrel Form.docx"), file(["memo", "d2"], PACK, "Kestrel Board Pack.docx"))).toEqual({ issues: [nameNotQuoted(0, "Kestrel Form.docx")] })
    expect(parse(true, file(["d1"], PACK), file(["memo", "d2"], PACK, "Kestrel Board Pack.docx")))
      .toEqual({ issues: [], value: stated({ ...file(["d1"], PACK), refused_name: "Kestrel Form.docx" }, file(["memo", "d2"], PACK, "Kestrel Board Pack.docx")) })
  })

  test("a files job that never passed its check is recorded as unchecked and disclosed as one item", () => {
    const unchecked: LeanRequestedFiles = { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: BASIS, planned: PLANNED, status: "unchecked", stated: false, files: [] }
    const directory = saved(AS_ASKED, unchecked)
    expect(readRequestedFiles(directory)).toEqual(unchecked)
    expect(requestedFilesItems(directory)).toEqual([openItem("finalize", "omission", "Requested files could not be checked: the check of the lawyer's words for the files they ask to receive gave no usable result, " +
      "so the run could not confirm which files the lawyer's words ask to receive or what each holds. The lawyer should confirm how the files are to be delivered.", { repair_target: "memo" })])
    // Only a record of the saved plan counts, as for any record; an unchecked record lists no file.
    expect(requestedFilesItems(saved(AS_ASKED, { ...unchecked, basis: "basis-0" }))).toEqual([])
    writeLeanText(directory, "requested-files.json", JSON.stringify({ ...unchecked, stated: true, files: RECORD.files }))
    expect(readRequestedFiles(directory)).toBeUndefined()
    writeLeanText(directory, "requested-files.json", JSON.stringify({ ...unchecked, status: "checked" }))
    expect(readRequestedFiles(directory)).toBeUndefined()
  })

  test("a shape the run cannot build always gives an item", () => {
    expect(requestedFilesItems(saved(AS_ASKED, record(file(["d1", "d2"], BOTH))))).toEqual([item(["Kestrel Transfer.docx", "Kestrel Board Pack.docx"], `the lawyer's words "${BOTH}" [source: request] ask for "Share transfer form" and "Board minute" in one file, ` +
      'but the run delivers "Share transfer form" in "Kestrel Transfer.docx" and the memo and "Board minute" in one file named "Kestrel Board Pack.docx". The run cannot put more than one document in one file.')])
    expect(requestedFilesItems(saved(AS_ASKED, record(file(["d2", "memo"], PACK, "Kestrel Board Pack.docx"))))).toEqual([item(["Kestrel Board Pack.docx"], `the lawyer's words "${PACK}" [source: request] ask for "Board minute" and the memo in one file named "Kestrel Board Pack.docx", ` +
      'but the run delivers the memo and "Board minute" in one file named "Kestrel Board Pack.docx". The run can join the memo and a document in one file only with the memo first.')])
  })

  test("a hand-written plan, or a record of another plan, gives no item", () => {
    const changed = [AS_ASKED[0]!, outline("Board minute", { deliverable: "body" })]
    expect(requestedFilesItems(saved(changed, RECORD, null))).toEqual([])
    expect(requestedFilesItems(saved(changed, { ...RECORD, basis: "basis-0" }))).toEqual([])
    const directory = prepared()
    writeLeanPlan(directory, changed, BASIS, undefined, NOTES)
    expect(requestedFilesItems(directory)).toEqual([])
    // A plan edited by hand under the same basis: other documents or another memo than the record was read for.
    expect(requestedFilesItems(saved(changed))).toHaveLength(1)
    expect(requestedFilesItems(saved([AS_ASKED[0]!, outline("Board resolution", { deliverable: "body" })]))).toEqual([])
    expect(requestedFilesItems(saved(changed, RECORD, BASIS, { name: "advice memo", file: NOTES.file }))).toEqual([])
  })
})
