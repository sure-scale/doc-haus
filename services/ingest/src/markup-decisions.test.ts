import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { applyMarkupDecisions, MarkupDecisionError, type MarkupDecision } from "./markup-decisions"

// Docxodus joins two adjacent insertions by one author into one change once the
// deletion between them is accepted. Deciding the second insertion after that
// would also decide the first, which the lawyer may have held.

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const by = (author: string) => `w:author="${author}" w:date="2026-10-20T09:00:00Z"`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const BODY = `<w:p>${run("Start ")}<w:ins w:id="1" ${by("Avery Lindqvist")}>${run("alpha ")}</w:ins>` +
  `<w:del w:id="2" ${by("Morgan Okafor")}><w:r><w:delText xml:space="preserve">middle </w:delText></w:r></w:del>` +
  `<w:ins w:id="3" ${by("Avery Lindqvist")}>${run("beta ")}</w:ins>${run("end.")}</w:p>`
const TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
const DOCX = zipSync({
  "[Content_Types].xml": strToU8(TYPES),
  "_rels/.rels": strToU8(RELS),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${BODY}<w:sectPr/></w:body></w:document>`),
})

async function decide(alpha: MarkupDecision, middle: MarkupDecision, beta: MarkupDecision) {
  const session = (await docxodus()).openDocxSession(DOCX, { persistAnchorIds: false })
  const ids = (() => { try { return session.listRevisions().map(entry => entry.id) } finally { session.close() } })()
  expect(ids).toHaveLength(3)
  return applyMarkupDecisions(DOCX, new Map([[ids[0]!, alpha], [ids[1]!, middle], [ids[2]!, beta]]))
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

test("deciding an insertion that has merged with a held one is refused", async () => {
  const refused = await decide("hold", "accept", "reject").catch(error => error)
  expect(refused).toBeInstanceOf(MarkupDecisionError)
  expect(refused).toMatchObject({ code: "decision_failed", message: expect.stringContaining("Two changes became one") })
})

test("insertions that merge are held together, and so are decided the same way", async () => {
  const held = await decide("hold", "accept", "hold")
  expect(held.outcomes.map(row => row.outcome)).toEqual(["held", "applied", "held"])
  expect(await accepted(held.bytes)).toEqual(["Start alpha beta end."])
  const rejected = await decide("reject", "accept", "reject")
  expect(rejected.outcomes.map(row => row.outcome)).toEqual(["applied", "applied", "applied"])
  expect(await accepted(rejected.bytes)).toEqual(["Start end."])
})
