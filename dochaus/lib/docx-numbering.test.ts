// Run from dochaus/: bun test lib/docx-numbering.test.ts
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { numberDocxClauses } from "./docx-numbering"
import { docxodus } from "./docxodus"
import { ensureParagraphRows } from "./drafting-amendment"
import { renderMarkdownToDocx } from "./render"
import { paragraphRows } from "../workflows/lean/stages"

const BODY = `# Supply Agreement

This Agreement is made between the parties.

## 1. Definitions

1.1 In this Agreement the following terms apply.

1.2 Headings do not affect interpretation, as clause 2.1 confirms.

(a) the first limb;

(b) the second limb; and

(i) a sub-limb.

(c) the third limb.

## 2. Supply

### 2.1 Delivery

The Supplier shall deliver the Goods.

### 2.2 Price

The price is fixed.

## Schedule 1

Goods list.`

let workspace: string
const saved = process.env.WORKSPACE_ROOT
beforeAll(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "docx-numbering-"))
  process.env.WORKSPACE_ROOT = workspace
})
afterAll(() => {
  if (saved === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = saved
  rmSync(workspace, { recursive: true, force: true })
})

async function blocks(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {
    return Object.entries(session.project().anchorIndex).filter(([, anchor]) => ["p", "h", "li"].includes(anchor.kind)).map(([id]) => {
      const info = session.getAnchorInfo(id)!
      const metadata = session.getBlockMetadata(id)!
      return { shown: info.visibleText.replace(/\s+/g, " ").trim(), text: info.textPreview, label: metadata.list?.generatedLabel, style: metadata.styleId, numId: metadata.list?.numId }
    })
  } finally {
    session.close()
  }
}

/** The same Markdown rendered with typed numbers and with Word numbering. */
async function both(markdown: string, styleTemplate?: Uint8Array) {
  const literal = await blocks((await renderMarkdownToDocx(markdown, undefined, undefined, { numbering: false, ...(styleTemplate ? { styleTemplate } : {}) })).bytes)
  const bytes = (await renderMarkdownToDocx(markdown, undefined, undefined, styleTemplate ? { styleTemplate } : {})).bytes
  return { literal, numbered: await blocks(bytes), bytes }
}

const labelled = (rows: Awaited<ReturnType<typeof blocks>>) => rows.filter(row => row.label).map(row => row.shown)

// ---- synthetic Word templates, built here

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
const level = (ilvl: number, format: string, text: string, style?: string) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/>${style ? `<w:pStyle w:val="${style}"/>` : ""}<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${567 * (ilvl + 1)}" w:hanging="567"/></w:pPr></w:lvl>`
const NORMAL = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
const heading = (n: number) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>`
const clauseStyle = (n: number, numId?: string) =>
  `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Level${n}"><w:name w:val="Level ${n}"/><w:basedOn w:val="Normal"/>${numId ? `<w:pPr><w:numPr><w:ilvl w:val="${n - 1}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` : ""}</w:style>`

function template(styles: string, numbering: string) {
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${TYPE}.numbering+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document ${NS}><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`),
    "word/styles.xml": xml(`<w:styles ${NS}>${NORMAL}${heading(1)}${heading(2)}${heading(3)}${styles}</w:styles>`),
    "word/numbering.xml": xml(`<w:numbering ${NS}>${numbering}</w:numbering>`),
  }, { mtime: new Date(1980, 0, 1) })
}

const CLAUSE_LEVELS = level(0, "decimal", "%1.", "Level1") + level(1, "decimal", "%1.%2", "Level2") + level(2, "lowerLetter", "(%3)", "Level3")
const clauseTemplate = (extraStyles = "", extraNumbering = "") => template(
  [1, 2, 3].map(n => clauseStyle(n, "7")).join("") + extraStyles,
  `<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="multilevel"/>${CLAUSE_LEVELS}</w:abstractNum>${extraNumbering}<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>`,
)

const AGREEMENT = `# Agreement

## 1. Definitions

1.1 Terms have their defined meanings.

(a) the first limb.

1.2 Headings are for convenience.

## 2. Supply

2.1 The Supplier shall deliver the Goods.`

// Sequences the list cannot show whole, which stay typed: a level Word would
// number differently ("1.1.1" after "(a)" under 1.1 reads 1.1.2), a skipped
// number, an undotted or bold run-in clause number, a number the list writes
// differently (a stray trailing dot, a zero-padded decimal).
const TYPED = [
  "# Agreement\n\n## 1. One\n\n1.1 The Supplier shall:\n\n(a) first; and\n\n(b) second.\n\n1.1.1 a deeper point.\n\n1.2 next.",
  "# Agreement\n\n## 1. One\n\n1.1 a.\n\n## 2. Two\n\n2.1 b.\n\n## 4. Four\n\n4.1 c.",
  "# Agreement\n\n## 1 One\n\n### 1.1 Sub heading\n\n1.2 text.\n\n## 2 Two\n\n2.1 text.",
  "# Agreement\n\n**1. One**\n\n1.1 text one.\n\n1.2 text.\n\n**2. Two**\n\n2.1 text two.",
  "# Agreement\n\n## 1. One\n\n1.01 A plain decimal level shows 1.1.\n\n## 2. Two\n\n2.1 text.",
  "# Agreement\n\n## 1. One\n\n1.1 text.\n\n1.2. text.\n\n## 2. Two\n\n2.1 text.",
]

// A typed main body (it skips clause 3) before a schedule the list numbers.
const TYPED_BODY_SCHEDULE = "# Agreement\n\n## 1. One\n\n1.1 First.\n\n## 2. Two\n\n2.1 Second.\n\n## 4. Four\n\n4.1 Fourth.\n\n## Schedule 1\n\n### 1. Scope\n\n1.1 Help desk."

describe("numberDocxClauses", () => {
  test("turns literal clause numbers into a Word list bound to the heading style, keeping what a reader sees", async () => {
    const { literal, numbered: after, bytes: numbered } = await both(BODY)
    expect(after.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    const row = (shown: string) => after.find(entry => entry.shown === shown)!
    expect(row("1. Definitions")).toMatchObject({ label: "1.", text: "Definitions", style: "Heading2" })
    expect(row("1.2 Headings do not affect interpretation, as clause 2.1 confirms.")).toMatchObject({ label: "1.2" })
    expect(row("(i) a sub-limb.")).toMatchObject({ label: "(i)", text: "a sub-limb." })
    expect(row("(c) the third limb.")).toMatchObject({ label: "(c)" })
    expect(row("2.2 Price")).toMatchObject({ label: "2.2", text: "Price" })
    expect(row("Schedule 1")).toMatchObject({ label: undefined })

    // The heading style carries the list, so a heading added in Word is numbered.
    const parts = unzipSync(numbered)
    const styles = strFromU8(parts["word/styles.xml"]!)
    const heading2 = styles.match(/<w:style\b[^>]*w:styleId="Heading2"[\s\S]*?<\/w:style>/)![0]
    const numId = heading2.match(/<w:numId w:val="(\d+)"\/>/)![1]
    const numbering = strFromU8(parts["word/numbering.xml"]!)
    expect(numbering).toMatch(new RegExp(`<w:num w:numId="${numId}"><w:abstractNumId w:val="\\d+"/></w:num>`))
    expect(numbering).toContain('<w:pStyle w:val="Heading2"/>')
    const document = strFromU8(parts["word/document.xml"]!)
    const added = document.replace(/(<w:p\b(?:(?!<w:p\b)[\s\S])*?Supply<\/w:t>)/, '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Added</w:t></w:r></w:p>$1')
    expect(added).not.toBe(document)
    const inserted = await blocks(zipSync({ ...parts, "word/document.xml": strToU8(added) }))
    expect(inserted.filter(entry => entry.style === "Heading2" && entry.label).map(entry => entry.shown)).toEqual(["1. Definitions", "2. Added", "3. Supply"])

    // Rows keep the paragraph's own text, which amendments are checked against, and its label.
    const file = path.join(workspace, "numbered.docx")
    writeFileSync(file, numbered)
    const rows = await paragraphRows(file)
    expect(rows.find(entry => entry.text === "In this Agreement the following terms apply.")?.label).toBe("1.1")
    expect(rows.find(entry => entry.text === "Goods list.")?.label).toBeUndefined()
  })

  test("is idempotent and leaves a document without clause numbers untouched", async () => {
    // The apply step numbers every applied body again (`house: false`): a
    // drafted body, typed sequences and template lists included, stays as it is.
    for (const [markdown, styleTemplate] of [BODY, AGREEMENT, TYPED_BODY_SCHEDULE, ...TYPED].map(text => [text, undefined] as const).concat([[AGREEMENT, clauseTemplate()], [TYPED_BODY_SCHEDULE, clauseTemplate()]])) {
      const numbered = (await renderMarkdownToDocx(markdown, undefined, undefined, styleTemplate ? { styleTemplate } : {})).bytes
      expect(numberDocxClauses(numbered)).toBe(numbered)
      expect(numberDocxClauses(numbered, { house: false })).toBe(numbered)
    }
    // The typed body never takes the schedule's list: the schedule keeps its own numbers.
    const shown = (await blocks((await renderMarkdownToDocx(TYPED_BODY_SCHEDULE)).bytes)).filter(row => row.label).map(row => row.shown)
    expect(shown).toEqual(["1. Scope", "1.1 Help desk."])
    const plain = (await renderMarkdownToDocx("# Letter\n\nDear Sir,\n\n(a) a point; and\n\n## Next steps\n\nWe will write again.", undefined, undefined, { numbering: false })).bytes
    expect(numberDocxClauses(plain)).toBe(plain)
  })

  test("a numbered paragraph that repeats its own number as text loses the repeat", async () => {
    const numbered = (await renderMarkdownToDocx(BODY)).bytes
    const parts = unzipSync(numbered)
    const document = strFromU8(parts["word/document.xml"]!)
    const repeated = document.replace(">In this Agreement", ">1.1 In this Agreement").replace(">Price<", ">2.2 Price<")
    expect(repeated).not.toBe(document)
    const amended = zipSync({ ...parts, "word/document.xml": strToU8(repeated) })
    const normalized = await blocks(numberDocxClauses(amended, { house: false }))
    expect(normalized.map(row => row.shown)).toEqual((await blocks(numbered)).map(row => row.shown))
    // Without a clause list of its own, `house: false` numbers nothing.
    const literal = (await renderMarkdownToDocx(BODY, undefined, undefined, { numbering: false })).bytes
    expect(numberDocxClauses(literal, { house: false })).toBe(literal)
  })

  test("parties and recitals lettered before clause 1 stay typed, and the clauses keep their numbers", async () => {
    const { literal, numbered, bytes } = await both(`# Agreement

## Parties

(a) Alpha Limited; and

(b) Beta Limited.

## Recitals

(A) The Supplier provides services.

(B) The Customer wishes to buy them.

${AGREEMENT.split("\n").slice(2).join("\n")}`)
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    expect(labelled(numbered)).toEqual(["1. Definitions", "1.1 Terms have their defined meanings.", "(a) the first limb.", "1.2 Headings are for convenience.", "2. Supply", "2.1 The Supplier shall deliver the Goods."])
    // The Heading 2 of Parties and Recitals before the clauses does not keep the
    // style from carrying the list: a heading added in Word is numbered.
    expect(numbered.find(row => row.shown === "Parties")).toMatchObject({ style: "Heading2", label: undefined })
    expect(strFromU8(unzipSync(bytes)["word/styles.xml"]!).match(/<w:style\b[^>]*w:styleId="Heading2"[\s\S]*?<\/w:style>/)![0]).toContain("<w:numPr>")
  })

  test("a numbered sequence converts whole or not at all", async () => {
    // Every number of a sequence the list cannot show whole stays as typed.
    for (const markdown of TYPED) {
      const { literal, numbered } = await both(markdown)
      expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
      expect(labelled(numbered)).toEqual([])
    }
    // A schedule restarting at 1 is its own sequence, numbered by its own
    // restart of the list; one whose numbers do not fit stays typed alone.
    const { literal, numbered } = await both("# Agreement\n\n## 1. One\n\n1.1 a.\n\n## 2. Two\n\n2.1 b.\n\n# Schedule 1\n\n## 1. Services\n\n1.1 c.\n\n# Schedule 2\n\n## 1. Prices\n\n1.3 d.")
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    expect(labelled(numbered)).toEqual(["1. One", "1.1 a.", "2. Two", "2.1 b.", "1. Services", "1.1 c."])
    const one = numbered.find(row => row.shown === "1. One")!.numId
    expect(numbered.find(row => row.shown === "1. Services")!.numId).not.toBe(one)
  })

  test("an unnumbered heading at the clause headings' level ends the sequence, and the numbers after it stay typed", async () => {
    // A schedule in parts numbered 1.1, 1.2, 2.1 without a "1." heading.
    const { literal, numbered } = await both("# Agreement\n\n## 1. Supply\n\n1.1 The Supplier shall supply the Goods.\n\n## 2. Price\n\n2.1 The price is fixed.\n\n## Schedule 1 – Services\n\n### Part A – Help desk\n\n1.1 The help desk is staffed.\n\n1.2 Calls are logged.\n\n### Part B – On site\n\n2.1 Engineers attend.\n\n## Execution\n\nSigned for the Supplier.")
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    expect(labelled(numbered)).toEqual(["1. Supply", "1.1 The Supplier shall supply the Goods.", "2. Price", "2.1 The price is fixed."])
    // An undotted clause heading is a number the list writes differently, not an end: typed whole.
    const undotted = await both("# Agreement\n\n## 1. Supply\n\n1.1 Goods.\n\n## 2 Price\n\n2.1 Fixed.")
    expect(undotted.numbered.map(row => row.shown)).toEqual(undotted.literal.map(row => row.shown))
    expect(labelled(undotted.numbered)).toEqual([])
  })

  test("a number the list could never show at that point is text, and the sequence around it converts", async () => {
    // One below the current clause, and one in a form no house level has.
    for (const [markdown, text] of [
      ["# Agreement\n\n## 1. One\n\n1.1 a.\n\n## 2. Two\n\n2.1 b.\n\n1.5 times the Charges is payable.\n\n2.2 c.", "1.5 times the Charges is payable."],
      ["# Agreement\n\n## 1. One\n\n1.1 Section 1 provides:\n\n(1) A person commits an offence.\n\n1.2 c.", "(1) A person commits an offence."],
    ]) {
      const { literal, numbered } = await both(markdown!)
      expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
      expect(labelled(numbered)).toEqual(literal.slice(1).map(row => row.shown).filter(shown => shown !== text))
    }
  })

  test("\"(i)\" after \"(h)\" is the next letter, or the first roman sub-limb when \"(ii)\" follows", async () => {
    const limbs = (last: string) => `# Agreement\n\n## 1. One\n\n1.1 The Supplier shall:\n\n${["a", "b", "c", "d", "e", "f", "g"].map(letter => `(${letter}) item;`).join("\n\n")}\n\n${last}\n\n1.2 next.`
    for (const markdown of [limbs("(h) item, including:\n\n(i) a sub-item; and\n\n(ii) another."), limbs("(h) item;\n\n(i) item; and\n\n(j) item.")]) {
      const { literal, numbered } = await both(markdown)
      expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
      expect(labelled(numbered)).toEqual(literal.slice(1).map(row => row.shown))
    }
  })

  test("an ordered list inside a clause stays its own list, and the clause list numbers around it", async () => {
    const { literal, numbered, bytes } = await both("# Agreement\n\n## 1. Supply\n\n1.1 The Supplier shall:\n\n1. deliver the Goods;\n2. invoice the Buyer.\n\n1.2 Payment is due on delivery.\n\n## 2. Term\n\n2.1 One year.")
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    const file = path.join(workspace, "ordered.docx")
    writeFileSync(file, bytes)
    const rows = await ensureParagraphRows(file)
    const clause = (text: string) => rows.find(row => row.text === text)?.clause
    expect(clause("Payment is due on delivery.")).toBe(true)
    expect(clause("deliver the Goods;")).toBeUndefined()
    expect(rows.find(row => row.text === "deliver the Goods;")?.label).toBe("1.")
  })

  test("the house list indents each level below the one before and gives Word nine levels", async () => {
    const numbering = strFromU8(unzipSync((await renderMarkdownToDocx(BODY)).bytes)["word/numbering.xml"]!)
    const house = numbering.match(/<w:abstractNum\b(?:(?!<\/w:abstractNum>)[\s\S])*<w:pStyle w:val="Heading2"\/>[\s\S]*?<\/w:abstractNum>/)![0]
    expect(house).toMatch(/<w:nsid w:val="[0-9A-F]{8}"\/>/)
    expect(house).toMatch(/<w:tmpl w:val="[0-9A-F]{8}"\/>/)
    expect(house.match(/<w:lvl /g)).toHaveLength(9)
    const left = [...house.matchAll(/<w:ind w:left="(\d+)"/g)].map(match => Number(match[1]))
    expect(left[3]).toBeGreaterThan(left[2]!)
    expect(left.slice(2).every((value, at, all) => at === 0 || value > all[at - 1]!)).toBe(true)
  })

  test("a template's clause list is chosen by the numbers it can show, not by a heading-bound schedule or article list", async () => {
    // A "Schedule %1" list bound to a heading-like style beside the clause list.
    const schedule = clauseTemplate(
      `<w:style w:type="paragraph" w:customStyle="1" w:styleId="ScheduleHeading"><w:name w:val="Schedule Heading"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="9"/></w:numPr><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr></w:style>`,
      `<w:abstractNum w:abstractNumId="5"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="ScheduleHeading"/><w:suff w:val="nothing"/><w:lvlText w:val="Schedule %1"/><w:lvlJc w:val="center"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="5"/></w:num>`,
    )
    const { literal, numbered } = await both(AGREEMENT, schedule)
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    expect(labelled(numbered)).toEqual(literal.slice(1).map(row => row.shown))
    expect(numbered.some(row => row.style === "ScheduleHeading")).toBe(false)

    // A list that labels its clauses "Article 1" cannot show "1.": typed.
    const article = template("", `<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/>${level(0, "decimal", "Article %1", "Heading2")}${level(1, "decimal", "%1.%2")}</w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>`)
    const articles = await both("# Agreement\n\n## 1. Definitions\n\n1.1 Terms.\n\n## 2. Supply\n\n2.1 Goods.", article)
    expect(articles.numbered.map(row => row.shown)).toEqual(articles.literal.map(row => row.shown))
    expect(labelled(articles.numbered)).toEqual([])
  })

  test("a template level whose style names the level but carries no numbering keeps the number", async () => {
    const pStyleOnly = template([1, 2, 3].map(n => clauseStyle(n)).join(""), `<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="multilevel"/>${CLAUSE_LEVELS}</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>`)
    const { literal, numbered } = await both(AGREEMENT, pStyleOnly)
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    expect(labelled(numbered)).toEqual(literal.slice(1).map(row => row.shown))
  })

  test("a list defined through a numbering style numbers every clause through the num its clause styles use", async () => {
    // Level 2 (1.1) has no clause style of its own, so those clauses carry the list directly.
    const linked = template(
      clauseStyle(1, "7") + clauseStyle(3, "7") + `<w:style w:type="numbering" w:customStyle="1" w:styleId="ClauseList"><w:name w:val="Clause List"/><w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr></w:style>`,
      `<w:abstractNum w:abstractNumId="3"><w:nsid w:val="11223344"/><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="ClauseList"/>${level(0, "decimal", "%1.", "Level1")}${level(1, "decimal", "%1.%2")}${level(2, "lowerLetter", "(%3)", "Level3")}</w:abstractNum><w:abstractNum w:abstractNumId="4"><w:nsid w:val="55667788"/><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="ClauseList"/></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num><w:num w:numId="8"><w:abstractNumId w:val="3"/></w:num>`,
    )
    const { literal, numbered } = await both(AGREEMENT, linked)
    expect(numbered.map(row => row.shown)).toEqual(literal.map(row => row.shown))
    expect(numbered.filter(row => row.label).map(row => row.numId)).toEqual([7, 7, 7, 7, 7, 7])
  })

  test("a clause heading keeps its heading style on a template whose clause styles are not headings", async () => {
    const { numbered } = await both(AGREEMENT, clauseTemplate())
    expect(numbered.find(row => row.shown === "1. Definitions")).toMatchObject({ style: "Heading2", label: "1." })
    expect(numbered.find(row => row.shown === "1.1 Terms have their defined meanings.")).toMatchObject({ style: "Level2", label: "1.1" })
  })

  test("a later part split from a numbered paragraph leaves the list unless it opens with a deeper number", async () => {
    const bytes = (await renderMarkdownToDocx("# Agreement\n\n## 1. Supply\n\n1.1 Payment is due:\n\n1.2 Late payment under clause 1.1 carries interest.")).bytes
    const parts = unzipSync(bytes)
    const document = strFromU8(parts["word/document.xml"]!)
    // A split copies the paragraph's properties into each part, as the applier does.
    const paragraph = document.match(/<w:p\b(?:(?!<w:p\b)[\s\S])*?Payment is due:[\s\S]*?<\/w:p>/)![0]
    const properties = paragraph.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)![0]
    const part = (text: string, id: string) => `<w:p>${properties}<w:bookmarkStart w:id="${id}" w:name="split_${id}"/><w:bookmarkEnd w:id="${id}"/><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
    const split = document.replace(paragraph, paragraph + part("(a) on delivery; or", "901") + part("(b) on invoice.", "902") + part("Interest runs daily.", "903"))
    const amended = zipSync({ ...parts, "word/document.xml": strToU8(split) })
    const later = (node: Element) => Array.from(node.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "bookmarkStart"))
      .some(mark => mark.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "name")?.startsWith("split_"))
    const shown = (await blocks(numberDocxClauses(amended, { house: false, split: later }))).map(row => row.shown)
    expect(shown).toEqual(["Agreement", "1. Supply", "1.1 Payment is due:", "(a) on delivery; or", "(b) on invoice.", "Interest runs daily.", "1.2 Late payment under clause 1.1 carries interest."])
  })
})
