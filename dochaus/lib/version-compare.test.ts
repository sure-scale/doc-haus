import { expect, test } from "bun:test"
import { mkdtemp, readFile, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { compareVersions, roundTrips, type CompareOptions } from "./version-compare"

// Synthetic Word-style fixtures built in code: neutral agreement text between
// fictional parties, returned with tracked changes by fictional reviewers.
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const DATE = "2026-10-19T12:00:00Z"
const PAT = { author: "Pat Reviewer", date: "2026-10-15T09:00:00Z" }
const SAM = { author: "Sam Counsel", date: "2026-10-16T10:00:00Z" }
const US = { author: "Northwind Legal", date: "2026-10-12T08:00:00Z" }
const ALL: CompareOptions = { ignoreFormatting: true, detectMoves: true, headersFooters: true }
const LEFT = "Sent draft", RIGHT = "Their draft"

const escape = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
const run = (text: string, rPr = "") => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${escape(text)}</w:t></w:r>`
let revisionId = 900
const ins = (text: string, by = PAT) => `<w:ins w:id="${revisionId++}" w:author="${by.author}" w:date="${by.date}">${run(text)}</w:ins>`
const del = (text: string, by = PAT) => `<w:del w:id="${revisionId++}" w:author="${by.author}" w:date="${by.date}"><w:r><w:delText xml:space="preserve">${escape(text)}</w:delText></w:r></w:del>`
const para = (inner: string, pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${inner}</w:p>`
const p = (text: string, style?: string) => para(run(text), style ? `<w:pStyle w:val="${style}"/>` : "")

type Spec = { body: string[]; header?: string; footer?: string; numbering?: boolean }
const styles = `<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>`
const numbering = `<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`

function docx(spec: Spec) {
  const type = (part: string, kind: string) => `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`
  const relation = (id: string, kind: string, target: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"/>`
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${type("document.xml", "document.main")}${type("styles.xml", "styles")}${spec.header ? type("header1.xml", "header") : ""}${spec.footer ? type("footer1.xml", "footer") : ""}${spec.numbering ? type("numbering.xml", "numbering") : ""}</Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relation("rId1", "officeDocument", "word/document.xml")}</Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relation("rIdS", "styles", "styles.xml")}${spec.header ? relation("rIdH", "header", "header1.xml") : ""}${spec.footer ? relation("rIdF", "footer", "footer1.xml") : ""}${spec.numbering ? relation("rIdN", "numbering", "numbering.xml") : ""}</Relationships>`,
    "word/styles.xml": styles,
    "word/document.xml": `<w:document ${NS}><w:body>${spec.body.join("")}<w:sectPr>${spec.header ? `<w:headerReference w:type="default" r:id="rIdH"/>` : ""}${spec.footer ? `<w:footerReference w:type="default" r:id="rIdF"/>` : ""}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  }
  if (spec.header) parts["word/header1.xml"] = `<w:hdr ${NS}>${p(spec.header)}</w:hdr>`
  if (spec.footer) parts["word/footer1.xml"] = `<w:ftr ${NS}>${p(spec.footer)}</w:ftr>`
  if (spec.numbering) parts["word/numbering.xml"] = numbering
  return zipSync(Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))
}

const C = {
  title: "Services Agreement",
  intro: "This Services Agreement is made between Northwind Ltd (the Supplier) and Contoso LLC (the Customer).",
  h1: "1. Definitions",
  d11: "1.1 In this Agreement, Services means the consulting services described in Schedule 1.",
  d12: "1.2 Business Day means a day other than a Saturday, Sunday or public holiday in London.",
  d13a: "1.3 Fees means the amounts payable under clause 2.", d13b: " Headings do not affect interpretation.",
  h2: "2. Payment",
  p21: "2.1 The Customer shall pay the Fees within 30 days of the date of each invoice.",
  p22: "2.2 Late payments bear interest at 4 per cent per annum above the base rate.",
  p23: "2.3 All amounts are exclusive of value added tax, which the Customer shall pay in addition.",
  p24a: "2.4 The Supplier may ", p24w: "suspend", p24b: " the Services if any undisputed invoice remains unpaid for 60 days.",
  p25: "2.5 Each party shall bear its own costs in connection with this Agreement.",
}
const INSERTED = " The Services exclude hardware procurement."
const HEADER = "Northwind Ltd / Contoso LLC - Services Agreement (Draft 2)", FOOTER = "Confidential - Subject to contract"
const body = () => [p(C.title, "Title"), p(C.intro), p(C.h1, "Heading1"), p(C.d11), p(C.d12), p(C.d13a + C.d13b),
  p(C.h2, "Heading1"), p(C.p21), p(C.p22), p(C.p23), p(C.p24a + C.p24w + C.p24b), p(C.p25)]
const sent = (rows = body()) => docx({ body: rows, header: HEADER, footer: FOOTER })
// Their return: three tracked changes by Pat Reviewer, as Word writes them.
const tracked = (by = PAT) => {
  const rows = body()
  rows[3] = para(run(C.d11) + ins(INSERTED, by))
  rows[5] = para(run(C.d13a) + del(C.d13b, by))
  rows[10] = para(run(C.p24a) + del(C.p24w, by) + ins("withhold", by) + run(C.p24b))
  return rows
}
const plant = (rows: string[]) => { rows[7] = p(C.p21.replace("within 30 days", "within 45 days")); return rows }
// Our own redline: a proposal of ours still pending when the draft was sent.
const OURS = " A day ends at 5 pm London time."
const proposed = (rows: string[]) => { rows[4] = para(run(C.d12) + ins(OURS, US)); return rows }

// Revision markup of a package: element, author, date and its text.
function revisions(bytes: Uint8Array) {
  return Object.entries(unzipSync(bytes)).filter(([name]) => /^word\/(?:document|header\d*|footer\d*)\.xml$/.test(name)).flatMap(([name, data]) => {
    const doc = new DOMParser().parseFromString(strFromU8(data), "application/xml")
    return Array.from(doc.getElementsByTagNameNS(W, "*")).filter(node => ["ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "moveFromRangeStart", "moveToRangeStart"].includes(node.localName))
      .map(node => ({ part: name, tag: node.localName, author: node.getAttributeNS(W, "author"), date: node.getAttributeNS(W, "date"),
        text: Array.from(node.getElementsByTagNameNS(W, "*")).filter(child => child.localName === "t" || child.localName === "delText").map(child => child.textContent).join("") }))
  })
}
const texted = (bytes: Uint8Array) => revisions(bytes).filter(row => row.text && ["ins", "del", "moveFrom", "moveTo"].includes(row.tag))

// Non-empty paragraph text of every story, accepted view.
async function storyText(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return session.grep("^[\\s\\S]*$", { scope: 7 }).map(match => `${match.enclosingAnchor.scope}: ${match.text}`) }
  finally { session.close() }
}
const accepted = async (bytes: Uint8Array) => storyText(await (await docxodus()).docxDiffAcceptRevisions(bytes))
const rejected = async (bytes: Uint8Array) => storyText(await (await docxodus()).docxDiffRejectRevisions(bytes))
const compare = (left: Uint8Array, right: Uint8Array, options: Partial<CompareOptions> = {}) =>
  compareVersions({ bytes: left, label: LEFT }, { bytes: right, label: RIGHT }, { ...ALL, ...options }, DATE)

test("every inserted, deleted and moved paragraph is a change and a native revision; the blackline rejects to left and accepts to right", async () => {
  const rows = body()
  rows.splice(5, 0, p("1.2A Effective Date means the date of this Agreement."))
  rows.splice(rows.indexOf(p(C.p25)), 1)
  const [moved] = rows.splice(rows.indexOf(p(C.p23)), 1)
  rows.push(moved!)
  rows[rows.indexOf(p(C.p21))] = p(C.p21.replace("30 days", "45 days"))
  const left = sent(), right = sent(rows)
  const result = await compare(left, right)

  expect(result.changes.map(({ kind, before, after }) => ({ kind, before, after }))).toEqual([
    { kind: "inserted", before: undefined, after: "1.2A Effective Date means the date of this Agreement." },
    { kind: "modified", before: C.p21, after: C.p21.replace("30 days", "45 days") },
    { kind: "moved", before: C.p23, after: C.p23 },
    { kind: "deleted", before: C.p25, after: undefined },
  ])
  expect(result.changes.map(hunk => hunk.id)).toEqual(["C1", "C2", "C3", "C4"])
  expect(result.changes.find(hunk => hunk.kind === "inserted")!.location).toEqual({ scope: "body", paragraph: 6, side: "right", heading: C.h1 })
  expect(result.changes.find(hunk => hunk.kind === "deleted")!.location).toEqual({ scope: "body", paragraph: 12, side: "left", heading: C.h2 })
  const markup = texted(result.blackline)
  expect(markup.find(row => row.tag === "ins" && row.text === "1.2A Effective Date means the date of this Agreement.")).toBeDefined()
  expect(markup.find(row => row.tag === "del" && row.text === C.p25)).toBeDefined()
  expect(markup.filter(row => row.tag === "moveFrom" || row.tag === "moveTo").map(row => [row.tag, row.text])).toEqual([["moveFrom", C.p23], ["moveTo", C.p23]])
  expect(await rejected(result.blackline)).toEqual(await storyText(left))
  expect(await accepted(result.blackline)).toEqual(await storyText(right))
})

test("a right side without tracked changes reports every difference as untracked and labels each revision so", async () => {
  const rows = body()
  rows[rows.indexOf(p(C.p21))] = p(C.p21.replace("30 days", "45 days"))
  rows.splice(rows.indexOf(p(C.p25)), 1)
  const result = await compare(sent(), sent(rows))
  expect(result.right).toEqual({ count: 0, authors: [] })
  expect(result.attribution).toBe("authors")
  expect(result.untracked.map(({ id, ...hunk }) => hunk)).toEqual(result.changes.map(({ id, ...hunk }) => hunk))
  expect(result.untracked.map(hunk => hunk.id)).toEqual(["U1", "U2"])
  expect(result.changes.every(hunk => hunk.untracked)).toBe(true)
  expect(new Set(revisions(result.blackline).map(row => row.author))).toEqual(new Set([`Not tracked (${RIGHT})`]))
})

test("a planted untracked edit in a returned draft is the only untracked edit, located and authored apart from the reviewer's tracked changes", async () => {
  const left = sent(), right = sent(plant(tracked()))
  const result = await compare(left, right)

  expect(result.left).toEqual({ count: 0, authors: [] })
  expect(result.right.authors).toEqual(["Pat Reviewer"])
  expect(result.right.count).toBeGreaterThanOrEqual(4)
  expect(result.untracked).toEqual([{ id: "U1", kind: "modified", untracked: true, before: C.p21, after: C.p21.replace("30 days", "45 days"),
    location: { scope: "body", paragraph: 8, side: "right", heading: C.h2 } }])
  expect(result.changes.map(hunk => [hunk.after, hunk.untracked])).toEqual([
    [C.d11 + INSERTED, false], [C.d13a, false], [C.p21.replace("30 days", "45 days"), true], [C.p24a + "withhold" + C.p24b, false]])

  expect(result.attribution).toBe("authors")
  const markup = texted(result.blackline)
  expect(markup.filter(row => row.author === `Not tracked (${RIGHT})`).map(row => [row.tag, row.text, row.date]).toSorted())
    .toEqual([["del", "30", DATE], ["ins", "45", DATE]])
  expect(markup.filter(row => row.author === PAT.author).map(row => [row.tag, row.text, row.date]).toSorted())
    .toEqual([["del", " Headings do not affect interpretation.", PAT.date], ["del", "suspend", PAT.date], ["ins", INSERTED, PAT.date], ["ins", "withhold", PAT.date]])
  // Every revision carries the supplied date unless it took its reviewer's.
  expect(revisions(result.blackline).every(row => row.date === (row.author === PAT.author ? PAT.date : DATE))).toBe(true)
  expect(await rejected(result.blackline)).toEqual(await storyText(left))
  expect(await accepted(result.blackline)).toEqual(await accepted(right))
})

test("our own pending proposals, still pending in their return, are not untracked edits", async () => {
  const result = await compare(sent(proposed(body())), sent(proposed(tracked())))
  expect(result.left).toEqual({ count: 1, authors: [US.author] })
  expect(result.right.authors.toSorted()).toEqual([US.author, PAT.author])
  expect(result.untracked).toEqual([])
  expect(result.changes.map(hunk => [hunk.after, hunk.untracked])).toEqual([
    [C.d11 + INSERTED, false], [C.d13a, false], [C.p24a + "withhold" + C.p24b, false]])
  expect(result.attribution).toBe("authors")
  expect(new Set(texted(result.blackline).map(row => row.author))).toEqual(new Set([PAT.author]))
})

test("beside our own pending proposals, a planted untracked edit is the only untracked edit", async () => {
  const left = sent(proposed(body())), right = sent(plant(proposed(tracked())))
  const result = await compare(left, right)
  expect(result.untracked).toEqual([{ id: "U1", kind: "modified", untracked: true, before: C.p21, after: C.p21.replace("30 days", "45 days"),
    location: { scope: "body", paragraph: 8, side: "right", heading: C.h2 } }])
  expect(result.changes.map(hunk => [hunk.after, hunk.untracked])).toEqual([
    [C.d11 + INSERTED, false], [C.d13a, false], [C.p21.replace("30 days", "45 days"), true], [C.p24a + "withhold" + C.p24b, false]])
  // Both passes read left with our proposal accepted, so the untracked edit past it keys exactly.
  expect(result.attribution).toBe("authors")
  expect(texted(result.blackline).filter(row => row.author === `Not tracked (${RIGHT})`).map(row => row.text).toSorted()).toEqual(["30", "45"])
  expect(await rejected(result.blackline)).toEqual(await accepted(left))
  expect(await accepted(result.blackline)).toEqual(await accepted(right))
})

test("a proposal of ours they accepted is not an untracked edit: the blackline shows no change there", async () => {
  const rows = tracked()
  rows[4] = p(C.d12 + OURS)
  const result = await compare(sent(proposed(body())), sent(rows))
  expect(result.untracked).toEqual([])
  expect(result.changes.map(hunk => [hunk.after, hunk.untracked])).toEqual([
    [C.d11 + INSERTED, false], [C.d13a, false], [C.p24a + "withhold" + C.p24b, false]])
})

test("a proposal of ours they removed leaves no mark, so the deletion the blackline shows is untracked", async () => {
  const rows = tracked()
  rows[4] = p(C.d12)
  const result = await compare(sent(proposed(body())), sent(rows))
  expect(result.untracked.map(({ before, untracked }) => ({ before, untracked }))).toEqual([{ before: C.d12 + OURS, untracked: true }])
  expect(["modified", "deleted"]).toContain(result.untracked[0]!.kind)
  expect(result.changes.map(hunk => [hunk.before, hunk.untracked])).toEqual([
    [C.d11, false], [C.d12 + OURS, true], [C.d13a + C.d13b, false], [C.p24a + C.p24w + C.p24b, false]])
  expect(result.attribution).toBe("authors")
  expect(texted(result.blackline).filter(row => row.text === OURS).map(row => [row.tag, row.author])).toEqual([["del", `Not tracked (${RIGHT})`]])
})

test("an untracked edit beside our pending proposal in the same paragraph flags that paragraph's change", async () => {
  const rows = tracked()
  const working = C.d12.replace("Business Day", "Working Day")
  rows[4] = para(run(working) + ins(OURS, US))
  const result = await compare(sent(proposed(body())), sent(rows))
  expect(result.changes.map(hunk => [hunk.after, hunk.untracked])).toEqual([
    [C.d11 + INSERTED, false], [working + OURS, true], [C.d13a, false], [C.p24a + "withhold" + C.p24b, false]])
  expect(result.untracked.map(({ kind, before, after }) => ({ kind, before, after }))).toEqual([{ kind: "modified", before: C.d12 + OURS, after: working + OURS }])
})

test("revisions are re-attributed only when unambiguous: identical text by two reviewers keeps the side label", async () => {
  const rows = body()
  rows[3] = para(run(C.d11) + ins(" promptly", PAT))
  rows[4] = para(run(C.d12) + ins(" promptly", SAM))
  rows[8] = para(run(C.p22.slice(0, -1)) + ins(" of Northwind Bank", SAM) + run("."))
  const result = await compare(sent(), sent(plant(rows)))
  expect(result.right.authors.toSorted()).toEqual([PAT.author, SAM.author])
  expect(result.attribution).toBe("authors")
  const markup = texted(result.blackline)
  expect(markup.filter(row => row.text === " promptly").map(row => [row.author, row.date])).toEqual([[RIGHT, DATE], [RIGHT, DATE]])
  expect(markup.filter(row => row.text === " of Northwind Bank").map(row => [row.author, row.date])).toEqual([[SAM.author, SAM.date]])
  expect(markup.filter(row => row.author === `Not tracked (${RIGHT})`).map(row => row.text).toSorted()).toEqual(["30", "45"])
})

test("formatting-only differences are ignored on request and reported otherwise", async () => {
  const rows = body()
  rows[8] = para(run("2.2 Late payments bear interest at ", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`)
    + run("4 per cent", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/>`) + run(" per annum above the base rate.", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`))
  rows[9] = para(run(C.p23, "<w:i/>"))
  rows[11] = para(run(C.p25), `<w:jc w:val="both"/>`)
  const left = sent(), right = sent(rows)

  const ignored = await compare(left, right)
  expect(ignored.changes).toEqual([])
  expect(ignored.untracked).toEqual([])
  expect(revisions(ignored.blackline).filter(row => row.tag === "rPrChange" || row.tag === "pPrChange")).toEqual([])
  expect(await rejected(ignored.blackline)).toEqual(await storyText(left))

  const shown = await compare(left, right, { ignoreFormatting: false })
  expect(shown.changes.length).toBeGreaterThan(0)
  expect(shown.changes.every(hunk => hunk.kind === "modified" && hunk.before === hunk.after)).toBe(true)
  expect(shown.untracked).toEqual([])
  expect(revisions(shown.blackline).some(row => row.tag === "rPrChange")).toBe(true)
})

test("header and footer edits are reported with their scope unless headers and footers are excluded", async () => {
  const left = sent(), right = docx({ body: body(), header: HEADER.replace("Draft 2", "Draft 3"), footer: FOOTER.replace("Subject to contract", "Execution version") })
  const result = await compare(left, right)
  expect(result.changes.map(hunk => [hunk.location.scope, hunk.location.paragraph, hunk.before, hunk.after])).toEqual([
    ["header", 1, HEADER, HEADER.replace("Draft 2", "Draft 3")], ["footer", 1, FOOTER, FOOTER.replace("Subject to contract", "Execution version")]])
  expect(result.changes.every(hunk => hunk.location.heading === undefined)).toBe(true)
  expect(new Set(revisions(result.blackline).map(row => row.part))).toEqual(new Set(["word/header1.xml", "word/footer1.xml"]))

  const excluded = await compare(left, right, { headersFooters: false })
  expect(excluded.changes).toEqual([])
  expect(excluded.untracked).toEqual([])
  expect(revisions(excluded.blackline)).toEqual([])
  expect(excluded.attribution).toBe("authors")
})

test("without move detection a moved paragraph is a deletion and an insertion", async () => {
  const rows = body()
  const [moved] = rows.splice(9, 1)
  rows.push(moved!)
  const result = await compare(sent(), sent(rows), { detectMoves: false })
  expect(result.changes.map(hunk => [hunk.kind, hunk.before ?? hunk.after])).toEqual([["deleted", C.p23], ["inserted", C.p23]])
  expect(revisions(result.blackline).some(row => row.tag === "moveFrom" || row.tag === "moveTo")).toBe(false)
})

test("locations carry Word numbering of the clause and of its heading", async () => {
  const numbered = (text: string, level: number, style?: string) => para(run(text), `${style ? `<w:pStyle w:val="${style}"/>` : ""}<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr>`)
  const rows = (fee: string) => [p(C.title, "Title"), numbered("Definitions", 0, "Heading1"), numbered("Services means the consulting services.", 1),
    numbered("Payment", 0, "Heading1"), numbered("The Customer shall pay the Fees.", 1), numbered(`Late payments bear interest at ${fee}.`, 1)]
  const result = await compare(docx({ body: rows("4 per cent"), numbering: true }), docx({ body: rows("5 per cent"), numbering: true }))
  expect(result.changes.map(hunk => hunk.location)).toEqual([{ scope: "body", paragraph: 6, side: "right", heading: "2. Payment", number: "2.2" }])
})

test("the round-trip gate refuses a blackline whose text no longer rejects to left", async () => {
  const left = sent(), right = sent(plant(tracked()))
  const result = await compare(left, right)
  const rightView = await (await docxodus()).docxDiffAcceptRevisions(right)
  expect(await roundTrips(result.blackline, left, rightView, true)).toBe(true)
  const parts = unzipSync(result.blackline)
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("1.2 Business Day", "1.2 Working Day"))
  expect(await roundTrips(zipSync(parts), left, rightView, true)).toBe(false)
})

test("an untracked word typed beside a tracked insertion is re-labelled on its own", async () => {
  // Word keeps right's run boundaries, so each insertion stays its own revision.
  const rows = body()
  rows[3] = para(run(C.d11) + ins(INSERTED) + run(" Travel is excluded."))
  const result = await compare(sent(), sent(rows))
  expect(result.attribution).toBe("authors")
  expect(texted(result.blackline).map(row => [row.tag, row.text, row.author]).toSorted())
    .toEqual([["ins", INSERTED, PAT.author], ["ins", " Travel is excluded.", `Not tracked (${RIGHT})`]])
  expect(result.untracked.map(hunk => [hunk.before, hunk.after])).toEqual([[C.d11, C.d11 + " Travel is excluded."]])
})

test("an untracked edit that cannot be keyed exactly falls back to the plain blackline under the side label", async () => {
  // An untracked full stop deleted against Pat's tracked deletion merges with it
  // in the blackline (both come from one left run), so the untracked edit has no
  // exact counterpart to re-label.
  const cut = C.d13a.slice(0, -1)
  const rows = body()
  rows[5] = para(run(cut) + del(C.d13b))
  const left = sent(), right = sent(rows)
  const result = await compare(left, right)
  expect(result.attribution).toBe("label")
  expect(new Set(revisions(result.blackline).map(row => row.author))).toEqual(new Set([RIGHT]))
  expect(result.untracked.map(hunk => [hunk.before, hunk.after])).toEqual([[C.d13a + C.d13b, cut + C.d13b]])
  expect(result.changes.map(hunk => hunk.untracked)).toEqual([true])
  expect(await rejected(result.blackline)).toEqual(await storyText(left))
  expect(await accepted(result.blackline)).toEqual(await accepted(right))
})

test("a long agreement locates every planted untracked edit under its heading", async () => {
  const fixture = longAgreement(600)
  const result = await compare(fixture.left, fixture.right)
  expect(result.attribution).toBe("authors")
  expect(result.untracked.map(hunk => [hunk.location.paragraph, hunk.location.heading])).toEqual(fixture.untracked.map(row => [row.paragraph, row.heading]))
  expect(result.changes.filter(hunk => hunk.untracked).length).toBe(fixture.untracked.length)
  expect(result.changes.length).toBe(fixture.untracked.length + fixture.tracked)
  expect(new Set(texted(result.blackline).map(row => row.author))).toEqual(new Set([PAT.author, `Not tracked (${RIGHT})`]))
}, 30000)

const soffice = Bun.which("soffice") ?? Bun.which("libreoffice")
test.skipIf(!soffice)("LibreOffice converts the blackline headless", async () => {
  const result = await compare(sent(), sent(plant(tracked())))
  const dir = await mkdtemp(path.join(tmpdir(), "dochaus-blackline-"))
  try {
    const file = path.join(dir, "blackline.docx")
    await Bun.write(file, result.blackline)
    const proc = Bun.spawn([soffice!, "--headless", `-env:UserInstallation=file://${path.join(dir, "profile")}`, "--convert-to", "txt:Text", "--outdir", dir, file], { stdout: "ignore", stderr: "ignore" })
    expect(await proc.exited).toBe(0)
    expect(await readFile(path.join(dir, "blackline.txt"), "utf8")).toContain("2.4 The Supplier may")
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}, 60000)

// About 600 paragraphs under a heading every twelve; twenty tracked edits by Pat
// and ten untracked edits spread through the body.
function longAgreement(count: number) {
  let seed = 7
  const random = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648
  const pick = <T,>(rows: T[]) => rows[Math.floor(random() * rows.length)]!
  const subject = ["The Supplier", "The Customer", "Each party", "Northwind Ltd", "Contoso LLC"]
  const verb = ["shall deliver", "shall maintain", "may request", "shall provide", "shall review"]
  const object = ["the Deliverables", "written notice", "the Services", "an updated schedule", "reasonable assistance"]
  const tail = ["within ten Business Days", "at its own cost", "without undue delay", "on reasonable request", "in each Contract Year"]
  const left = [p("Master Services Agreement", "Title"), p("This Master Services Agreement is made between Northwind Ltd and Contoso LLC.")]
  const headings = [C.title, C.title]
  let section = 0, clause = 0
  for (let index = 0; index < count; index++) {
    if (index % 12 === 0) { section++; clause = 0; left.push(p(`${section}. Section ${section}`, "Heading1")); headings.push(`${section}. Section ${section}`); continue }
    clause++
    const sentences = Array.from({ length: 2 + Math.floor(random() * 3) }, () => `${pick(subject)} ${pick(verb)} ${pick(object)} ${pick(tail)}.`)
    left.push(p(`${section}.${clause} ${sentences.join(" ")} This obligation is reference ${section}-${clause}.`))
    headings.push(`${section}. Section ${section}`)
  }
  const right = [...left]
  const clauses = left.flatMap((row, index) => index > 1 && !row.includes("Heading1") ? [index] : [])
  const chosen = Array.from({ length: 30 }, (_, k) => clauses[Math.floor((k + 0.5) * clauses.length / 30)]!)
  const untracked: { paragraph: number; heading: string }[] = []
  chosen.forEach((index, k) => {
    const text = /<w:t xml:space="preserve">([^<]*)<\/w:t>/.exec(left[index]!)![1]!
    if (k % 3 === 0) { right[index] = p(text.replace("This obligation is reference", "This duty is item")); untracked.push({ paragraph: index + 1, heading: headings[index]! }) }
    else if (k % 3 === 1) right[index] = para(run(text) + ins(` The parties agree amendment ${k}.`))
    else right[index] = para(run(text.slice(0, text.indexOf("This obligation"))) + del("This obligation") + ins("This covenant") + run(text.slice(text.indexOf("This obligation") + "This obligation".length)))
  })
  return { left: docx({ body: left }), right: docx({ body: right }), untracked, tracked: chosen.length - untracked.length }
}
