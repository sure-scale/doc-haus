import { expect, test } from "bun:test"
import type { RevisionListEntry } from "docxodus"
import { strToU8, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { applyMarkupDecisions, MarkupDecisionError, type MarkupDecision } from "./markup-decisions"

// Docxodus joins adjacent changes by one author of one type into one change once
// the text between them is resolved away. Deciding a change after it has joined
// a held or differently decided one would decide that one too.

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const DATE = "2026-10-20T09:00:00Z"
const by = (author: string, date: string | undefined = DATE) => `w:author="${author}"${date ? ` w:date="${date}"` : ""}`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const ins = (id: number, author: string, text: string, date?: string) => `<w:ins w:id="${id}" ${by(author, date)}>${run(text)}</w:ins>`
const del = (id: number, author: string, text: string, date?: string) =>
  `<w:del w:id="${id}" ${by(author, date)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`
const TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
const docx = (paragraph: string) => zipSync({
  "[Content_Types].xml": strToU8(TYPES),
  "_rels/.rels": strToU8(RELS),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body><w:p>${paragraph}</w:p><w:sectPr/></w:body></w:document>`),
})
const ALPHA_BETA = docx(`${run("Start ")}${ins(1, "Avery Lindqvist", "alpha ")}${del(2, "Morgan Okafor", "middle ")}${ins(3, "Avery Lindqvist", "beta ")}${run("end.")}`)

async function listed(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    return session.listRevisions()
  } finally {
    session.close()
  }
}
/** Decides every listed change of `bytes`, change i taking decisions[i]. */
async function decide(bytes: Uint8Array, ...decisions: MarkupDecision[]) {
  const ids = (await listed(bytes)).map(entry => entry.id)
  expect(ids).toHaveLength(decisions.length)
  return applyMarkupDecisions(bytes, new Map(ids.map((id, at) => [id, decisions[at]!])))
}
const accepted = async (bytes: Uint8Array) => {
  const dx = await docxodus()
  const session = dx.openDocxSession(await dx.docxDiffAcceptRevisions(bytes), { persistAnchorIds: false })
  try {
    return session.grep("^[\\s\\S]*$", { scope: 1 }).map(match => match.text)
  } finally {
    session.close()
  }
}
type Pending = Pick<RevisionListEntry, "type" | "author" | "date" | "text">
const pending = async (bytes: Uint8Array) => (await listed(bytes)).map(({ type, author, date, text }): Pending => ({ type, author, date, text }))

test("a decided change that would join a held one is decided before the others", async () => {
  const cases: Array<[string, Uint8Array, MarkupDecision[], string, Pending]> = [
    ["deletion between insertions", ALPHA_BETA, ["hold", "accept", "reject"], "Start alpha end.",
      { type: "insert", author: "Avery Lindqvist", date: DATE, text: "alpha " }],
    ["one author's insertion, deletion, insertion", docx(`${run("The Supplier ")}${ins(1, "Avery", "promptly ")}${del(2, "Avery", "may ")}${ins(3, "Avery", "shall ")}${run("notify.")}`),
      ["hold", "accept", "accept"], "The Supplier promptly shall notify.", { type: "insert", author: "Avery", date: DATE, text: "promptly " }],
    ["the same without dates", docx(`${run("The Supplier ")}${ins(1, "Avery", "promptly ", "")}${del(2, "Avery", "may ", "")}${ins(3, "Avery", "shall ", "")}${run("notify.")}`),
      ["hold", "accept", "accept"], "The Supplier promptly shall notify.", { type: "insert", author: "Avery", text: "promptly " }],
    ["insertion between deletions", docx(`${run("X ")}${del(1, "Avery", "one ")}${ins(2, "Morgan", "two ")}${del(3, "Avery", "three ")}${run("Y.")}`),
      ["hold", "reject", "accept"], "X Y.", { type: "delete", author: "Avery", date: DATE, text: "one " }],
    ["insertion between insertions", docx(`${run("X ")}${ins(1, "Avery", "one ")}${ins(2, "Morgan", "two ")}${ins(3, "Avery", "three ")}${run("Y.")}`),
      ["hold", "reject", "accept"], "X one three Y.", { type: "insert", author: "Avery", date: DATE, text: "one " }],
  ]
  for (const [label, bytes, decisions, text, held] of cases) {
    const result = await decide(bytes, ...decisions)
    expect(result.outcomes.map(row => row.outcome), label).toEqual(["held", "applied", "applied"])
    expect(await pending(result.bytes), label).toEqual([held])
    expect(await accepted(result.bytes), label).toEqual([text])
  }
})

test("a change that still joins another when decided first is refused, naming both", async () => {
  // Rejecting Morgan's "e " joins Avery's deletions; accepting "d " first instead
  // joins Morgan's insertions, one of them held.
  const bytes = docx(`${run("A ")}${ins(1, "Morgan", "q ")}${del(2, "Avery", "d ")}${ins(3, "Morgan", "e ")}${del(4, "Avery", "f ")}${run("Z.")}`)
  const ids = (await listed(bytes)).map(entry => entry.id)
  const refused = await decide(bytes, "hold", "accept", "reject", "accept").catch(error => error)
  expect(refused).toBeInstanceOf(MarkupDecisionError)
  expect(refused).toMatchObject({ code: "decision_failed", revision: ids[2], mergedWith: [ids[0]], message: expect.stringContaining("Two changes became one") })
})

test("insertions that merge are held together, and so are decided the same way", async () => {
  const held = await decide(ALPHA_BETA, "hold", "accept", "hold")
  expect(held.outcomes.map(row => row.outcome)).toEqual(["held", "applied", "held"])
  expect(await accepted(held.bytes)).toEqual(["Start alpha beta end."])
  const rejected = await decide(ALPHA_BETA, "reject", "accept", "reject")
  expect(rejected.outcomes.map(row => row.outcome)).toEqual(["applied", "applied", "applied"])
  expect(await accepted(rejected.bytes)).toEqual(["Start end."])
})
