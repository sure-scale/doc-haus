// Run from dochaus/: bun test lib/template-seed.test.ts
import { afterAll, beforeAll, describe, expect, spyOn, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { compareDocxWithPointers } from "./drafting-amendment"
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
const heading = (n: number, id = `Heading${n}`) =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="heading ${n}"/><w:basedOn w:val="BodyText"/><w:pPr><w:keepNext/><w:spacing w:before="480" w:after="240"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:b/></w:rPr></w:style>`
const A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="567" w:footer="567" w:gutter="0"/>'

// `options` vary the template: its heading styles, settings, extra document
// relationships and parts, what precedes its body, and its body.
function template(options: { headings?: string; settings?: string; rels?: string; parts?: Record<string, string>; types?: string; before?: string; body?: string } = {}) {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${options.types ?? ""}<Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${TYPE}.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="${TYPE}.settings+xml"/><Override PartName="/word/header1.xml" ContentType="${TYPE}.header+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rId4" Type="${REL}/header" Target="header1.xml"/>${options.rels ?? ""}</Relationships>`),
    "word/document.xml": xml(`<w:document ${NS}>${options.before ?? ""}<w:body>${options.body ?? `<w:p><w:pPr><w:pStyle w:val="Level1"/></w:pPr><w:r><w:t>Precedent clause</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId4"/>${A4}<w:titlePg/></w:sectPr>`}</w:body></w:document>`),
    "word/header1.xml": xml(`<w:hdr ${NS}><w:p><w:r><w:t>Precedent party header</w:t></w:r></w:p></w:hdr>`),
    "word/settings.xml": xml(`<w:settings ${NS}>${options.settings ?? '<w:documentProtection w:edit="readOnly" w:enforcement="1"/><w:trackRevisions/><w:defaultTabStop w:val="567"/><w:docVars><w:docVar w:name="Reference" w:val="EX-1"/></w:docVars>'}</w:settings>`),
    "word/styles.xml": xml(`<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="BodyText"><w:name w:val="Body Text"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>${options.headings ?? heading(1) + heading(2)}${[1, 2, 3, 4].map(clauseStyle).join("")}</w:styles>`),
    "word/numbering.xml": xml(`<w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>${CLAUSES}<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>`),
    ...Object.fromEntries(Object.entries(options.parts ?? {}).map(([name, text]) => [name, xml(text)])),
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
    // Clauses are numbered by the template's list through its clause styles,
    // a clause heading keeping its heading style; headings keep their own spacing.
    expect(await blocks(bytes)).toEqual([
      { shown: "Supply Agreement", style: "Heading1" },
      { shown: "1. Definitions", style: "Heading2" },
      { shown: "1.1 In this Agreement the following terms apply.", style: "Level2" },
      { shown: "(a) the first limb; and", style: "Level3" },
      { shown: "(i) a sub-limb.", style: "Level4" },
      { shown: "2. Supply", style: "Heading2" },
      { shown: "2.1 The Supplier shall deliver the Goods.", style: "Level2" },
      { shown: "a bullet point", style: undefined },
      { shown: "Schedule 1", style: "Heading2" },
      { shown: "Goods list.", style: undefined },
    ])
    const document = strFromU8(parts["word/document.xml"]!)
    // Only the headings carry the list directly; clause styles supply the rest.
    expect(document.match(/<w:numPr><w:ilvl w:val="\d"\/><w:numId w:val="7"\/>/g)).toHaveLength(2)
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
    expect(inserted.filter(row => /^(\d|\()/.test(row.shown)).map(row => row.shown)).toEqual([
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
    // A missing or unreadable style template is reported once, not on every render.
    const warn = spyOn(console, "warn").mockImplementation(() => {})
    try {
      prefer("Removed style.docx")
      expect(pageSize((await renderMarkdownToDocx(BODY)).bytes)).toBe(house)
      expect(pageSize((await renderMarkdownToDocx(BODY)).bytes)).toBe(house)
      expect(pageSize((await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: strToU8("not a Word file") })).bytes)).toBe(house)
      expect(warn.mock.calls.map(call => String(call[0]))).toEqual([expect.stringContaining("Removed style.docx"), expect.stringContaining("not a Word document")])
    } finally {
      warn.mockRestore()
    }
  })

  test("a template whose style ids Word translated keeps its own heading styles", async () => {
    const localized = template({
      headings: heading(1, "berschrift1") + heading(2, "berschrift2")
        + '<w:style w:type="paragraph" w:styleId="Titel"><w:name w:val="Title"/><w:basedOn w:val="BodyText"/></w:style>',
    })
    const parts = unzipSync((await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: localized })).bytes)
    const styles = strFromU8(parts["word/styles.xml"]!)
    const ids = new Set([...styles.matchAll(/w:styleId="([^"]+)"/g)].map(match => match[1]))
    const used = [...strFromU8(parts["word/document.xml"]!).matchAll(/<w:pStyle w:val="([^"]+)"/g)].map(match => match[1]!)
    expect(used.filter(id => !ids.has(id))).toEqual([])
    expect(used).toContain("berschrift1")
    expect(used).toContain("berschrift2")
    expect(ids.has("Heading1")).toBe(false)
    const shown = await blocks(zipSync(parts))
    expect(shown.find(row => row.shown === "1. Definitions")?.style).toBe("berschrift2")
    expect(shown.find(row => row.shown === "1.1 In this Agreement the following terms apply.")?.style).toBe("Level2")

    // The page furniture's header and footer take the template's translated Header and Footer styles.
    const furnished = template({
      headings: heading(2, "berschrift2")
        + '<w:style w:type="paragraph" w:styleId="Kopfzeile"><w:name w:val="header"/><w:basedOn w:val="BodyText"/></w:style>'
        + '<w:style w:type="paragraph" w:styleId="Fuzeile"><w:name w:val="footer"/><w:basedOn w:val="BodyText"/></w:style>',
    })
    const stories = unzipSync((await renderMarkdownToDocx(BODY, undefined, { header: ["Subject to contract"], pageNumber: { format: "page", align: "center" } }, { styleTemplate: furnished })).bytes)
    const storyStyles = Object.keys(stories).filter(name => /^word\/(header|footer)\d*\.xml$/.test(name)).sort()
      .map(name => [name, [...strFromU8(stories[name]!).matchAll(/<w:pStyle w:val="([^"]+)"/g)].map(match => match[1])])
    expect(storyStyles).toEqual([["word/footer1.xml", ["Fuzeile"]], ["word/header1.xml", ["Kopfzeile"]]])
  })

  test("the seed takes the main body's page and leaves out even-page headers, edit history and the printer", async () => {
    const landscape = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
    const seeded = template({
      settings: '<w:evenAndOddHeaders/><w:defaultTabStop w:val="567"/><w:rsids><w:rsidRoot w:val="00AB12CD"/><w:rsid w:val="00AB12CD"/></w:rsids>',
      rels: `<Relationship Id="rId8" Type="${REL}/printerSettings" Target="printerSettings/printerSettings1.bin"/><Relationship Id="rId9" Type="${REL}/image" Target="media/image1.png"/>`,
      parts: { "word/printerSettings/printerSettings1.bin": "FIRM-PRINTER", "word/media/image1.png": "PNG" },
      types: '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings"/><Default Extension="png" ContentType="image/png"/>',
      before: '<w:background w:color="FFFFFF"><v:background xmlns:v="urn:schemas-microsoft-com:vml" id="page" fillcolor="#ffffff"><v:fill r:id="rId9" type="frame"/></v:background></w:background>',
      body: `<w:p><w:pPr><w:sectPr>${A4}<w:printerSettings r:id="rId8"/></w:sectPr></w:pPr><w:r><w:t>Main body</w:t></w:r></w:p><w:p><w:r><w:t>Landscape schedule</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId4"/>${landscape}<w:printerSettings r:id="rId8"/></w:sectPr>`,
    })
    const parts = unzipSync((await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: seeded })).bytes)
    const document = strFromU8(parts["word/document.xml"]!)
    expect(pageSize(zipSync(parts))).toBe('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(document).not.toContain("Main body")
    expect(document).not.toContain("printerSettings")
    const relationships = new Set([...strFromU8(parts["word/_rels/document.xml.rels"]!).matchAll(/Id="([^"]+)"/g)].map(match => match[1]))
    expect([...document.matchAll(/r:(?:id|embed|link)="([^"]+)"/g)].map(match => match[1]).filter(id => !relationships.has(id))).toEqual([])
    expect(strFromU8(parts["word/settings.xml"]!)).not.toMatch(/evenAndOddHeaders|rsid/)
    expect(Object.keys(parts).filter(name => /printerSettings/.test(name))).toEqual([])
  })

  test("a precedent's cover page does not set the draft's page", async () => {
    // A centred cover page numbered in roman, the main body after it, and a closing landscape schedule.
    const cover = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2880" w:right="2880" w:bottom="2880" w:left="2880" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:vAlign w:val="center"/><w:titlePg/>'
    const landscape = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
    const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
    const precedent = template({
      body: `${paragraph("Dated")}<w:p><w:pPr><w:sectPr>${cover}</w:sectPr></w:pPr><w:r><w:t>Supply Agreement</w:t></w:r></w:p>`
        + `${paragraph("1. Definitions")}${paragraph("1.1 Terms.")}${paragraph("2. Supply")}<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>${A4}<w:pgNumType w:start="1"/></w:sectPr></w:pPr><w:r><w:t>2.1 Goods.</w:t></w:r></w:p>`
        + `${paragraph("Schedule 1")}<w:sectPr><w:type w:val="nextPage"/>${landscape}</w:sectPr>`,
    })
    const document = strFromU8(unzipSync((await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: precedent })).bytes)["word/document.xml"]!)
    expect(tight(document.match(/<w:sectPr\b[\s\S]*<\/w:sectPr>/)![0])).toBe(`<w:sectPr>${A4}</w:sectPr>`)
  })

  test("a revised original keeps the header and footer its main body shows, those it continues from an earlier section and their pictures included, the firm's header lines above its own", async () => {
    const story = (root: "hdr" | "ftr", text: string) => `<w:${root} ${NS}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:${root}>`
    const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
    const reference = (kind: "header" | "footer", type: string, id: string) => `<w:${kind}Reference w:type="${type}" r:id="${id}"/>`
    const landscape = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    // A cover section sets the default header and a footer with the firm's logo;
    // the main body, linked to it, sets only its first-page and even-page headers
    // and numbers its pages in roman; a closing schedule has a header of its own.
    const original = template({
      settings: '<w:evenAndOddHeaders/><w:defaultTabStop w:val="567"/>',
      rels: [["rId5", "header", "header2.xml"], ["rId6", "footer", "footer1.xml"], ["rId7", "header", "header3.xml"], ["rId8", "header", "header4.xml"]]
        .map(([id, kind, target]) => `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${target}"/>`).join(""),
      types: '<Default Extension="png" ContentType="image/png"/>' + ["header2", "header3", "header4", "footer1"]
        .map(name => `<Override PartName="/word/${name}.xml" ContentType="${TYPE}.${name.replace(/\d$/, "")}+xml"/>`).join(""),
      parts: {
        "word/header2.xml": story("hdr", "First page header"),
        "word/header3.xml": story("hdr", "Even page header"),
        "word/header4.xml": story("hdr", "Schedule header"),
        "word/footer1.xml": `<w:ftr ${NS}><w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdLogo"/></w:drawing></w:r></w:p></w:ftr>`,
        "word/_rels/footer1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="${REL}/image" Target="media/logo.png"/></Relationships>`,
        "word/media/logo.png": "PNG",
      },
      body: `<w:p><w:pPr><w:sectPr>${reference("header", "default", "rId4")}${reference("footer", "default", "rId6")}${A4}</w:sectPr></w:pPr><w:r><w:t>Cover</w:t></w:r></w:p>`
        + `${paragraph("1. Term")}${paragraph("1.1 One year.")}<w:p><w:pPr><w:sectPr>${reference("header", "first", "rId5")}${reference("header", "even", "rId7")}${A4}<w:pgNumType w:fmt="lowerRoman"/><w:titlePg/></w:sectPr></w:pPr><w:r><w:t>2. Supply</w:t></w:r></w:p>`
        + `${paragraph("Schedule")}<w:sectPr>${reference("header", "default", "rId8")}${landscape}</w:sectPr>`,
    })
    const firm = { header: ["Subject to contract", "Draft: 20 October 2026"], pageNumber: { format: "plain" as const, align: "center" as const } }
    const parts = unzipSync((await renderMarkdownToDocx(BODY, undefined, firm, { styleTemplate: original, revising: true })).bytes)
    const document = strFromU8(parts["word/document.xml"]!)
    const relationships = strFromU8(parts["word/_rels/document.xml.rels"]!)
    const shown = Object.fromEntries([...document.matchAll(/<w:(header|footer)Reference w:type="(\w+)" r:id="(\w+)"\/>/g)]
      .map(([, kind, type, id]) => [`${kind}:${type}`, new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(relationships)?.[1]]))
    expect(shown).toEqual({ "header:default": "header1.xml", "header:first": "header2.xml", "header:even": "header3.xml", "footer:default": "footer1.xml" })
    // Each header shown (default, first page, even pages) takes the firm's lines above its own text; the footer and its logo stay, with no firm page number.
    const texts = (name: string, from = parts) => [...strFromU8(from[name]!).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)</g)].map(match => match[1])
    expect(["word/header1.xml", "word/header2.xml", "word/header3.xml"].map(name => texts(name)))
      .toEqual(["Precedent party header", "First page header", "Even page header"].map(own => [...firm.header, own]))
    const same = (xml: string) => tight(xml).replace(/encoding="[^"]*"/, "")
    expect(same(strFromU8(parts["word/footer1.xml"]!))).toBe(same(strFromU8(unzipSync(original)["word/footer1.xml"]!)))
    expect(Object.keys(parts).filter(name => /header|footer|media/.test(name)).sort())
      .toEqual(["word/_rels/footer1.xml.rels", "word/footer1.xml", "word/header1.xml", "word/header2.xml", "word/header3.xml", "word/media/logo.png"])
    expect(document).toContain('<w:pgNumType w:fmt="lowerRoman"/>')
    expect(document).toContain("<w:titlePg/>")
    expect(document).not.toContain(" PAGE ")
    // An original with only a footer gains a header of the lines alone, its footer kept.
    const footed = template({ body: `${paragraph("1. Term")}<w:sectPr>${reference("footer", "default", "rId6")}${A4}</w:sectPr>`, rels: `<Relationship Id="rId6" Type="${REL}/footer" Target="footer1.xml"/>`,
      types: `<Override PartName="/word/footer1.xml" ContentType="${TYPE}.footer+xml"/>`, parts: { "word/footer1.xml": story("ftr", "Matter PP/7") } })
    const footer = unzipSync((await renderMarkdownToDocx(BODY, undefined, firm, { styleTemplate: footed, revising: true })).bytes)
    expect(Object.keys(footer).filter(name => /header|footer/.test(name)).sort()).toEqual(["word/footer1.xml", "word/header1.xml"])
    expect(texts("word/header1.xml", footer)).toEqual(firm.header)
    expect(texts("word/footer1.xml", footer)).toEqual(["Matter PP/7"])
    expect(tight(strFromU8(parts["word/settings.xml"]!))).toContain("<w:evenAndOddHeaders/>")
    // Drawn on as a template, the same file gives none of its furniture, and the firm's applies.
    const templated = unzipSync((await renderMarkdownToDocx(BODY, undefined, firm, { styleTemplate: original })).bytes)
    expect(Object.keys(templated).filter(name => /header|footer|media/.test(name)).sort()).toEqual(["word/footer1.xml", "word/header1.xml"])
    expect(strFromU8(templated["word/header1.xml"]!)).toContain("Subject to contract")
  })

  test("revising a doc.haus draft repeats no header line and dates its draft line anew; the original's other header text stays", async () => {
    const page = { format: "plain" as const, align: "center" as const }
    const texts = (bytes: Uint8Array, name: string) => [...strFromU8(unzipSync(bytes)[name]!).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)</g)].map(match => match[1])
    const original = (await renderMarkdownToDocx("# Supply terms\n\n## 1. Term\n\n1.1 One year.", undefined, { header: ["Subject to contract", "Draft: 1 October 2026"], pageNumber: page }, { numbering: false })).bytes
    // The redline changes only the draft line's date, or its firm and date.
    for (const [line, deleted, inserted] of [["Draft: 20 October 2026", "1", "20"], ["Doe LLP draft: 20 October 2026", "Draft: 1", "Doe LLP draft: 20"]]) {
      const revised = (await renderMarkdownToDocx("# Supply terms\n\n## 1. Term\n\n1.1 Two years.", undefined, { header: ["subject to  contract", line!], pageNumber: page }, { numbering: false, styleTemplate: original, revising: true })).bytes
      expect(texts(revised, "word/header1.xml")).toEqual(["Subject to contract", line])
      expect(texts(revised, "word/footer1.xml")).toEqual(["1"])
      const redline = unzipSync(await compareDocxWithPointers(original, revised, { authorName: "Test" }))
      const header = strFromU8(redline["word/header1.xml"]!)
      expect([...header.matchAll(/<w:delText[^>]*>([^<]*)</g)].map(match => match[1]!.trim())).toEqual([deleted])
      expect([...header.matchAll(/<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g)].flatMap(match => [...match[0].matchAll(/<w:t[^>]*>([^<]*)</g)].map(text => text[1]!.trim()))).toEqual([inserted])
      expect(header).not.toMatch(/<w:(?:pPrChange|rPrChange)\b/)
      expect(strFromU8(redline["word/footer1.xml"]!)).not.toMatch(/<w:(?:ins|del)\b/)
    }
    // A header line that only starts like a draft line is the original's own text.
    for (const own of ["Draft: for discussion", "Draft: 1 October 2026 (for discussion)"]) {
      const marked = (await renderMarkdownToDocx("# Supply terms", undefined, { header: [own] }, { numbering: false })).bytes
      const kept = (await renderMarkdownToDocx("# Supply terms", undefined, { header: ["Draft: 20 October 2026"] }, { numbering: false, styleTemplate: marked, revising: true })).bytes
      expect(texts(kept, "word/header1.xml")).toEqual(["Draft: 20 October 2026", own])
    }
  })

  test("revising a document doc.haus rendered keeps its headings' spacing", async () => {
    const spacing = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<w:spacing\b/g)?.length ?? 0
    const original = (await renderMarkdownToDocx(BODY)).bytes
    expect(spacing(original)).toBeGreaterThan(0)
    const revised = await renderMarkdownToDocx(BODY, original, undefined, { styleTemplate: original })
    expect(revised.templated).toBe(true)
    expect(spacing(revised.bytes)).toBe(spacing(original))
    // A draft on a firm template, whose heading styles space the headings, gains no spacing when revised.
    const firm = (await renderMarkdownToDocx(BODY, undefined, undefined, { styleTemplate: template() })).bytes
    expect(spacing(firm)).toBe(0)
    expect(spacing((await renderMarkdownToDocx(BODY, firm, undefined, { styleTemplate: firm })).bytes)).toBe(0)
  })
})
