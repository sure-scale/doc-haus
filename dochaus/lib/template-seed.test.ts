// Run from dochaus/: bun test lib/template-seed.test.ts
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"

// A synthetic firm template: its own body style, fonts and A4 page, and a
// four-level clause list (1. / 1.1 / (a) / (i)) bound to custom clause styles,
// with a header, protection, tracking and document variables of its own.
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const level = (ilvl: number, format: string, text: string, left: number) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:pStyle w:val="Level${ilvl + 1}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="567"/></w:pPr></w:lvl>`
const CLAUSES = `<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="multilevel"/>${level(0, "decimal", "%1.", 567)}${level(1, "decimal", "%1.%2", 567)}${level(2, "lowerLetter", "(%3)", 1134)}${level(3, "lowerRoman", "(%4)", 1701)}</w:abstractNum>`
const clauseStyle = (n: number) =>
  `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Level${n}"><w:name w:val="Level ${n}"/><w:basedOn w:val="BodyText"/><w:pPr><w:numPr><w:ilvl w:val="${n - 1}"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style>`
const heading = (n: number) =>
  `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="BodyText"/><w:pPr><w:keepNext/><w:spacing w:before="480" w:after="240"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:b/></w:rPr></w:style>`

function template() {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${TYPE}.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="${TYPE}.settings+xml"/><Override PartName="/word/header1.xml" ContentType="${TYPE}.header+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rId4" Type="${REL}/header" Target="header1.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document ${NS}><w:body><w:p><w:pPr><w:pStyle w:val="Level1"/></w:pPr><w:r><w:t>Precedent clause</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="567" w:footer="567" w:gutter="0"/><w:titlePg/></w:sectPr></w:body></w:document>`),
    "word/header1.xml": xml(`<w:hdr ${NS}><w:p><w:r><w:t>Precedent party header</w:t></w:r></w:p></w:hdr>`),
    "word/settings.xml": xml(`<w:settings ${NS}><w:documentProtection w:edit="readOnly" w:enforcement="1"/><w:trackRevisions/><w:defaultTabStop w:val="567"/><w:docVars><w:docVar w:name="Reference" w:val="EX-1"/></w:docVars></w:settings>`),
    "word/styles.xml": xml(`<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="BodyText"><w:name w:val="Body Text"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>${heading(1)}${heading(2)}${[1, 2, 3, 4].map(clauseStyle).join("")}</w:styles>`),
    "word/numbering.xml": xml(`<w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>${CLAUSES}<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>`),
  }, { mtime: new Date(1980, 0, 1) })
}

const BODY = `# Supply Agreement

## 1. Definitions

1.1 In this Agreement the following terms apply.

(a) the first limb; and

(i) a sub-limb.

## 2. Supply

2.1 The Supplier shall deliver the Goods.

- a bullet point

## Schedule 1

Goods list.`

async function blocks(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {
    return Object.entries(session.project().anchorIndex).filter(([, anchor]) => ["p", "h", "li"].includes(anchor.kind)).map(([id]) => ({
      shown: session.getAnchorInfo(id)!.visibleText.replace(/\s+/g, " ").trim(),
      style: session.getBlockMetadata(id)!.styleId,
    }))
  } finally {
    session.close()
  }
}

const tight = (xml: string) => xml.replace(/\s+\/>/g, "/>")
const pageSize = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<w:pgSz [^>]*\/>/)?.[0]

let workspace: string
const saved = process.env.WORKSPACE_ROOT
beforeAll(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "template-seed-"))
  process.env.WORKSPACE_ROOT = workspace
})
afterAll(() => {
  if (saved === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = saved
  rmSync(workspace, { recursive: true, force: true })
})

describe("rendering onto a template", () => {
  test("takes the template's numbering definitions, styles, fonts and page set-up, and none of its content", async () => {
    const bytes = (await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: template() })).bytes
    const parts = unzipSync(bytes)
    const numbering = tight(strFromU8(parts["word/numbering.xml"]!))
    expect(numbering).toContain(CLAUSES)
    expect(numbering).toContain('<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>')
    // Clauses are numbered by the template's list through its clause styles;
    // its other headings keep their own spacing.
    expect(await blocks(bytes)).toEqual([
      { shown: "Supply Agreement", style: "Heading1" },
      { shown: "1. Definitions", style: "Level1" },
      { shown: "1.1 In this Agreement the following terms apply.", style: "Level2" },
      { shown: "(a) the first limb; and", style: "Level3" },
      { shown: "(i) a sub-limb.", style: "Level4" },
      { shown: "2. Supply", style: "Level1" },
      { shown: "2.1 The Supplier shall deliver the Goods.", style: "Level2" },
      { shown: "a bullet point", style: undefined },
      { shown: "Schedule 1", style: "Heading2" },
      { shown: "Goods list.", style: undefined },
    ])
    const document = strFromU8(parts["word/document.xml"]!)
    expect(document).not.toMatch(/<w:numPr><w:ilvl w:val="\d"\/><w:numId w:val="7"\/>/)
    expect(document).not.toContain("<w:spacing")
    const styles = tight(strFromU8(parts["word/styles.xml"]!))
    expect(styles).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>')
    expect(styles).toContain(clauseStyle(1))
    expect(pageSize(bytes)).toBe('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(document).toContain('w:left="1418"')

    // The template's text, header, protection, tracking and variables stay behind.
    expect(document).not.toContain("Precedent clause")
    expect(document).not.toMatch(/headerReference|titlePg/)
    expect(Object.keys(parts).filter(name => /header|footer/.test(name))).toEqual([])
    expect(strFromU8(parts["word/settings.xml"]!)).not.toMatch(/documentProtection|trackRevisions|docVars/)

    // A clause added in Word with the template's clause style is numbered, and
    // the clauses after it renumber.
    const added = document.replace(/(<w:p\b(?:(?!<w:p\b)[\s\S])*?>Supply<\/w:t>)/, '<w:p><w:pPr><w:pStyle w:val="Level1"/></w:pPr><w:r><w:t>Added</w:t></w:r></w:p>$1')
    expect(added).not.toBe(document)
    const inserted = await blocks(zipSync({ ...parts, "word/document.xml": strToU8(added) }))
    expect(inserted.filter(row => row.style?.startsWith("Level")).map(row => row.shown)).toEqual([
      "1. Definitions", "1.1 In this Agreement the following terms apply.", "(a) the first limb; and", "(i) a sub-limb.",
      "2. Added", "3. Supply", "3.1 The Supplier shall deliver the Goods.",
    ])
  })

  test("the firm's style template applies to a document without its own; a missing or unreadable one is the house style", async () => {
    const house = pageSize((await renderMarkdownToDocx(BODY)).bytes)
    expect(house).toBe('<w:pgSz w:w="12240" w:h="15840"/>')
    mkdirSync(path.join(workspace, ".templates"), { recursive: true })
    mkdirSync(path.join(workspace, ".preferences"), { recursive: true })
    writeFileSync(path.join(workspace, ".templates", "Firm style.docx"), template())
    const prefer = (styleTemplate: string) => writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify({ styleTemplate }))
    prefer("Firm style.docx")
    expect(pageSize((await renderMarkdownToDocx(BODY)).bytes)).toBe('<w:pgSz w:w="11906" w:h="16838"/>')
    prefer("Removed style.docx")
    expect(pageSize((await renderMarkdownToDocx(BODY)).bytes)).toBe(house)
    expect(pageSize((await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: strToU8("not a Word file") })).bytes)).toBe(house)
  })
})
