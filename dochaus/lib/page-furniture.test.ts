// Run from dochaus/: bun test lib/page-furniture.test.ts
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { combineMemoAndBody } from "./combined-docx"
import { applyAmendments } from "./drafting-amendment-apply"
import { compareDocxWithPointers } from "./drafting-amendment"
import { DEFAULT_DRAFTING, type DraftingPreferences } from "./drafting-preferences"
import { docxodus } from "./docxodus"
import { applyPageFurniture, NOTES_MARKING, PAGE_NUMBER_AFFIXES, pageFurniture, type PageFurniture } from "./page-furniture"
import { renderMarkdownToDocx } from "./render"
import { presentDraft } from "../workflows/lean/draft-presentation"
import { paragraphRows } from "../workflows/lean/stages"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const R = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
// Midday local time, so the local date is the 20th wherever the test runs.
const DRAFTED = new Date(2026, 9, 20, 12).toISOString()
const BODY = "# Services Agreement\n\n## 1. Term\n\nThe term is one year.\n\n## 2. Fees\n\nThe fees are payable monthly."
const AGREEMENT: PageFurniture = { header: ["Subject to contract", "Draft: October 20, 2026"], pageNumber: { format: "page", align: "center" } }
const LETTER: PageFurniture = { header: [], pageNumber: { format: "plain", align: "center" }, firstPage: "unnumbered" }

// The render lint reads the firm's date style; an empty workspace keeps the default.
let workspace: string
const saved = process.env.WORKSPACE_ROOT
beforeAll(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "page-furniture-"))
  process.env.WORKSPACE_ROOT = workspace
})
afterAll(() => {
  if (saved === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = saved
  rmSync(workspace, { recursive: true, force: true })
})

describe("pageFurniture", () => {
  const prefs = (change: Partial<DraftingPreferences> = {}) => ({ ...DEFAULT_DRAFTING, ...change })

  test("legends are off by default: each form gets the draft line or nothing, and a page number", () => {
    const numbered = { format: "plain", align: "center" }
    expect(pageFurniture("agreement", prefs(), DRAFTED)).toEqual({ header: ["Draft: October 20, 2026"], pageNumber: numbered })
    expect(pageFurniture("advice", prefs(), DRAFTED)).toEqual({ header: ["Draft: October 20, 2026"], pageNumber: numbered })
    expect(pageFurniture("other", prefs(), DRAFTED)).toEqual({ header: [], pageNumber: numbered })
    // A letter, advising or not, is marked as a draft on every page and numbered from page 2.
    for (const kind of ["letter", "advice letter"] as const)
      expect(pageFurniture(kind, prefs(), DRAFTED)).toEqual({ header: ["Draft: October 20, 2026"], pageNumber: numbered, firstPage: "unnumbered" })
    expect(pageFurniture("notes", prefs(), DRAFTED)).toEqual({ header: [NOTES_MARKING], pageNumber: numbered })
  })

  test("each form carries only its own legend, above the draft line", () => {
    const legends = prefs({ agreementLegend: "Subject to contract", adviceLegend: "Privileged & confidential" })
    expect(pageFurniture("agreement", legends, DRAFTED)?.header).toEqual(["Subject to contract", "Draft: October 20, 2026"])
    expect(pageFurniture("advice", legends, DRAFTED)?.header).toEqual(["Privileged & confidential", "Draft: October 20, 2026"])
    // A letter giving advice carries the advice legend; another letter, to anyone, carries none.
    expect(pageFurniture("advice letter", legends, DRAFTED)?.header).toEqual(["Privileged & confidential", "Draft: October 20, 2026"])
    expect(pageFurniture("letter", legends, DRAFTED)?.header).toEqual(["Draft: October 20, 2026"])
    expect(pageFurniture("other", legends, DRAFTED)?.header).toEqual([])
    expect(pageFurniture("notes", legends, DRAFTED)?.header).toEqual([NOTES_MARKING])
  })

  test("no shipped agent or skill hands a model legend text to write: legends are the firm's", () => {
    const root = path.join(import.meta.dir, "..")
    const prompts = ["agent/*.md", "skill/*/SKILL.md"].flatMap((pattern) => [...new Bun.Glob(pattern).scanSync(root)])
    expect(prompts.length).toBeGreaterThan(10)
    // The chat agents whose answers the lawyer copies out are covered by name.
    expect(prompts).toEqual(expect.arrayContaining([path.join("agent", "quick-draft.md"), path.join("agent", "client-summary.md")]))
    const legend = /privileged\s*&\s*confidential|attorney[- ]client communication|attorney work product|subject to contract|for discussion purposes only/i
    expect(prompts.filter((file) => legend.test(readFileSync(path.join(root, file), "utf8")))).toEqual([])
  })

  test("the draft line names the firm when set and follows the firm's date style", () => {
    const line = (change: Partial<DraftingPreferences>) => pageFurniture("agreement", prefs(change), DRAFTED)?.header
    expect(line({ firm: "  Doe & Partners LLP " })).toEqual(["Doe & Partners LLP draft: October 20, 2026"])
    expect(line({ firm: "   " })).toEqual(["Draft: October 20, 2026"])
    expect(line({ dateFormat: "day-month-year" })).toEqual(["Draft: 20 October 2026"])
    expect(line({ dateFormat: "iso" })).toEqual(["Draft: 2026-10-20"])
    expect(line({ firm: "Doe & Partners LLP", draftLineFirm: false })).toEqual(["Draft: October 20, 2026"])
  })

  test("no draft line without a draft time, with an unreadable one, or when turned off", () => {
    expect(pageFurniture("agreement", prefs())?.header).toEqual([])
    expect(pageFurniture("agreement", prefs(), "not a date")?.header).toEqual([])
    expect(pageFurniture("advice", prefs({ draftLine: false }), DRAFTED)?.header).toEqual([])
  })

  test("page numbers follow the format and position, or are left out", () => {
    expect(pageFurniture("other", prefs({ pageNumbers: "page", pageNumberPosition: "right" }))?.pageNumber).toEqual({ format: "page", align: "right" })
    expect(pageFurniture("agreement", prefs({ pageNumbers: "off" }), DRAFTED)).toEqual({ header: ["Draft: October 20, 2026"] })
    expect(pageFurniture("advice", prefs({ pageNumbers: "dashed" }), DRAFTED)?.pageNumber).toEqual({ format: "dashed", align: "center" })
    // Without numbers, a letter has no first page to leave unnumbered.
    expect(pageFurniture("letter", prefs({ pageNumbers: "off" }), DRAFTED)).toEqual({ header: ["Draft: October 20, 2026"] })
    expect(pageFurniture("letter", prefs({ pageNumbers: "off", draftLine: false }), DRAFTED)).toBeUndefined()
    expect(pageFurniture("other", prefs({ pageNumbers: "off" }), DRAFTED)).toBeUndefined()
  })

  test("a firm that does not number page 1 leaves every first page unnumbered; a letter's never is", () => {
    for (const kind of ["agreement", "advice", "other", "notes", "letter"] as const)
      expect(pageFurniture(kind, prefs({ numberFirstPage: false }), DRAFTED)?.firstPage).toBe("unnumbered")
    expect(pageFurniture("letter", prefs({ numberFirstPage: true }), DRAFTED)?.firstPage).toBe("unnumbered")
    expect(pageFurniture("agreement", prefs({ numberFirstPage: false, pageNumbers: "off" }), DRAFTED)).toEqual({ header: ["Draft: October 20, 2026"] })
  })

  test("what the request asks of a document adds to the firm's settings for it, never removing the firm's legend or draft line", () => {
    const firm = prefs({ agreementLegend: "Subject to contract", adviceLegend: "Privileged & confidential" })
    expect(pageFurniture("agreement", firm, DRAFTED, { legend: "Without prejudice" })?.header).toEqual(["Subject to contract", "Without prejudice", "Draft: October 20, 2026"])
    expect(pageFurniture("letter", firm, DRAFTED, { legend: "Without prejudice save as to costs" })?.header).toEqual(["Without prejudice save as to costs", "Draft: October 20, 2026"])
    // The firm's legend asked for again is not repeated.
    expect(pageFurniture("advice letter", firm, DRAFTED, { legend: "PRIVILEGED & CONFIDENTIAL" })?.header).toEqual(["Privileged & confidential", "Draft: October 20, 2026"])
    expect(pageFurniture("advice letter", firm, DRAFTED, { legend: "" })?.header).toEqual(["Privileged & confidential", "Draft: October 20, 2026"])
    expect(pageFurniture("agreement", firm, DRAFTED, { pageNumbers: false })).toEqual({ header: ["Subject to contract", "Draft: October 20, 2026"] })
    expect(pageFurniture("other", prefs(), DRAFTED, { pageNumbers: false })).toBeUndefined()
  })

  test("the notes marking can be turned off", () => {
    expect(pageFurniture("notes", prefs({ notesMarking: false }), DRAFTED)).toEqual({ header: [], pageNumber: { format: "plain", align: "center" } })
    expect(pageFurniture("notes", prefs({ notesMarking: false, pageNumbers: "off" }), DRAFTED)).toBeUndefined()
  })
})

describe("applyPageFurniture", () => {
  test("writes well-formed header and footer parts, related, declared and referenced in schema order", async () => {
    const bytes = applyPageFurniture((await renderMarkdownToDocx(BODY)).bytes, AGREEMENT)
    const parts = unzipSync(bytes)
    for (const [name, data] of Object.entries(parts)) if (/\.(?:xml|rels)$/.test(name)) strict(data)
    const section = sectPr(bytes)
    expect(childNames(section)).toEqual(["headerReference", "footerReference", "pgSz", "pgMar"])
    const relationships = all(strict(parts["word/_rels/document.xml.rels"]!), R, "Relationship")
    expect(new Set(relationships.map(row => row.getAttribute("Id"))).size).toBe(relationships.length)
    const overrides = all(strict(parts["[Content_Types].xml"]!), CT, "Override")
    const target = (reference: Element) => {
      const relationship = relationships.find(row => row.getAttribute("Id") === reference.getAttributeNS(O, "id"))!
      const name = `word/${relationship.getAttribute("Target")}`
      expect(relationship.getAttribute("Type")).toBe(`${O}/${reference.localName.replace("Reference", "")}`)
      expect(overrides.find(row => row.getAttribute("PartName") === `/${name}`)?.getAttribute("ContentType"))
        .toBe(`application/vnd.openxmlformats-officedocument.wordprocessingml.${reference.localName.replace("Reference", "")}+xml`)
      return strict(parts[name]!)
    }
    const [header, footer] = elementChildren(section).slice(0, 2).map(target)
    expect(paragraphs(header!)).toEqual([{ style: "Header", align: "right", text: "Subject to contract" }, { style: "Header", align: "right", text: "Draft: October 20, 2026" }])
    // The page number is its own paragraph: literal "Page ", then a complex
    // PAGE field with a cached result, and no page count.
    expect(paragraphs(footer!)).toEqual([{ style: "Footer", align: "center", text: "Page 1" }])
    expect(all(footer!, W, "instrText").map(node => node.textContent)).toEqual([" PAGE "])
    expect(all(footer!, W, "fldChar").map(node => node.getAttributeNS(W, "fldCharType"))).toEqual(["begin", "separate", "end"])
    expect(strFromU8(parts["word/footer1.xml"]!)).not.toContain("NUMPAGES")

    // The engine reads it and converts it to HTML.
    const dx = await docxodus()
    const session = dx.openDocxSession(bytes, {})
    try {
      expect(session.project().markdown).toContain("Subject to contract")
    } finally {
      session.close()
    }
    expect(await dx.convertDocxToHtml(bytes)).toContain("The term is one year.")
  }, 120000)

  test("an unnumbered first page has its own parts, the same header and an empty footer, and the title-page switch", async () => {
    const bytes = applyPageFurniture((await renderMarkdownToDocx(BODY)).bytes, { ...LETTER, header: ["Private"] })
    const parts = unzipSync(bytes)
    const section = sectPr(bytes)
    expect(childNames(section)).toEqual(["headerReference", "headerReference", "footerReference", "footerReference", "pgSz", "pgMar", "titlePg"])
    const relationships = all(strict(parts["word/_rels/document.xml.rels"]!), R, "Relationship")
    const story = (reference: Element) => strict(parts[`word/${relationships.find(row => row.getAttribute("Id") === reference.getAttributeNS(O, "id"))!.getAttribute("Target")}`]!)
    const byType = Object.fromEntries(elementChildren(section).slice(0, 4).map(node => [`${node.localName}:${node.getAttributeNS(W, "type")}`, paragraphs(story(node))]))
    expect(byType).toEqual({
      "headerReference:default": [{ style: "Header", align: "right", text: "Private" }],
      "headerReference:first": [{ style: "Header", align: "right", text: "Private" }],
      // The default footer first appears on page 2.
      "footerReference:default": [{ style: "Footer", align: "center", text: "2" }],
      "footerReference:first": [{ style: "Footer", align: "center", text: "" }],
    })
  }, 120000)

  test("text goes in as text, never as markup", async () => {
    const text = `Doe & Partners <w:r><w:t>"injected"</w:t></w:r> ' — é`
    const bytes = applyPageFurniture((await renderMarkdownToDocx(BODY)).bytes, { header: [text, "Tab\there\nand a control\u0001character"] })
    const header = strict(unzipSync(bytes)["word/header1.xml"]!)
    expect(all(header, W, "r")).toHaveLength(2)
    expect(paragraphs(header).map(row => row.text)).toEqual([text, "Tab here and a control character"])
  }, 120000)

  test("a dashed page number puts the PAGE field between dashes, each format printing its affixes", async () => {
    const body = (await renderMarkdownToDocx(BODY)).bytes
    const footer = (format: "plain" | "page" | "dashed") => strict(unzipSync(applyPageFurniture(body, { header: [], pageNumber: { format, align: "center" } }))["word/footer1.xml"]!)
    expect(paragraphs(footer("dashed"))).toEqual([{ style: "Footer", align: "center", text: "- 1 -" }])
    expect(all(footer("dashed"), W, "instrText").map(node => node.textContent)).toEqual([" PAGE "])
    // The settings preview prints page numbers from the same affixes.
    for (const format of ["plain", "page", "dashed"] as const) expect(paragraphs(footer(format))[0]?.text).toBe(PAGE_NUMBER_AFFIXES[format].join("1"))
  }, 120000)

  test("without a page number there is no first-page switch, even when asked for an unnumbered first page", async () => {
    const bytes = applyPageFurniture((await renderMarkdownToDocx(BODY)).bytes, { header: ["Draft: October 20, 2026"], firstPage: "unnumbered" })
    expect(childNames(sectPr(bytes))).toEqual(["headerReference", "pgSz", "pgMar"])
  }, 120000)

  test("empty header lines with a page number write only the footer", async () => {
    const bytes = applyPageFurniture((await renderMarkdownToDocx(BODY)).bytes, { header: [], pageNumber: { format: "plain", align: "right" } })
    const parts = unzipSync(bytes)
    expect(Object.keys(parts).filter(name => /header|footer/.test(name))).toEqual(["word/footer1.xml"])
    expect(childNames(sectPr(bytes))).toEqual(["footerReference", "pgSz", "pgMar"])
    expect(strFromU8(parts["[Content_Types].xml"]!)).not.toContain("header+xml")
    expect(paragraphs(strict(parts["word/footer1.xml"]!))).toEqual([{ style: "Footer", align: "right", text: "1" }])
  }, 120000)

  test("new relationship ids never take one in use", async () => {
    const parts = unzipSync((await renderMarkdownToDocx(BODY)).bytes)
    parts["word/_rels/document.xml.rels"] = strToU8(strFromU8(parts["word/_rels/document.xml.rels"]!).replace("</Relationships>",
      `<Relationship Id="rId3" Type="${O}/hyperlink" Target="https://example.com/a" TargetMode="External"/><Relationship Id="rId5" Type="${O}/hyperlink" Target="https://example.com/b" TargetMode="External"/></Relationships>`))
    const bytes = applyPageFurniture(zipSync(parts), AGREEMENT)
    const ids = all(strict(unzipSync(bytes)["word/_rels/document.xml.rels"]!), R, "Relationship").map(row => row.getAttribute("Id"))
    expect(new Set(ids).size).toBe(ids.length)
    expect(elementChildren(sectPr(bytes)).slice(0, 2).map(node => node.getAttributeNS(O, "id"))).toEqual(["rId4", "rId6"])
  }, 120000)

  test("a document that already has a header or footer is returned unchanged", async () => {
    const rendered = (await renderMarkdownToDocx(BODY)).bytes
    for (const own of [AGREEMENT, { header: [], pageNumber: { format: "plain", align: "center" } } satisfies PageFurniture]) {
      const furnished = applyPageFurniture(rendered, own)
      expect(applyPageFurniture(furnished, { header: ["Privileged & confidential"], pageNumber: { format: "page", align: "right" }, firstPage: "unnumbered" })).toBe(furnished)
    }
    // A section break's own header counts too.
    const parts = unzipSync(rendered)
    parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("<w:body>",
      `<w:body><w:p><w:pPr><w:sectPr xmlns:r="${O}"><w:headerReference w:type="default" r:id="rId9"/></w:sectPr></w:pPr></w:p>`))
    const template = zipSync(parts)
    expect(applyPageFurniture(template, AGREEMENT)).toBe(template)
  }, 120000)
})

describe("rendering with furniture", () => {
  test("the same content and furniture render to the same bytes", async () => {
    for (const furniture of [AGREEMENT, LETTER]) {
      const first = (await renderMarkdownToDocx(BODY, undefined, furniture)).bytes
      expect(Buffer.from((await renderMarkdownToDocx(BODY, first, furniture)).bytes).equals(Buffer.from(first))).toBe(true)
    }
  }, 120000)

  test("header and footer paragraphs take no pointer and no paragraph row; the lint never reads them", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "page-furniture-"))
    try {
      // A day-first date breaks the default date style, so the lint flags it in the body.
      const furniture = { ...AGREEMENT, header: ["Draft: 20 October 2026"] }
      expect((await renderMarkdownToDocx(`${BODY}\n\nDraft: 20 October 2026`)).lint.join("\n")).toContain("20 October 2026")
      const rendered = await renderMarkdownToDocx(BODY, undefined, furniture)
      expect(rendered.lint.join("\n")).not.toContain("20 October 2026")
      const parts = unzipSync(rendered.bytes)
      for (const name of ["word/header1.xml", "word/footer1.xml"]) expect(strFromU8(parts[name]!)).not.toContain("_dh_")
      const file = path.join(dir, "body.docx")
      writeFileSync(file, rendered.bytes)
      expect((await paragraphRows(file)).map(row => row.text)).toEqual(["Services Agreement", "1. Term", "The term is one year.", "2. Fees", "The fees are payable monthly."])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  }, 120000)

  test("presentation and amendments leave the header and footer parts byte-identical", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "page-furniture-"))
    try {
      const furniture = { ...AGREEMENT, header: [`Doe & Partners LLP "draft" — <October 20, 2026>`] }
      const rendered = (await renderMarkdownToDocx(BODY, undefined, furniture)).bytes
      const presented = presentDraft(rendered, [{ heading: "1. Term" }, { heading: "2. Fees" }]).bytes
      expectSameStories(rendered, presented)
      const file = path.join(dir, "body.docx")
      writeFileSync(file, presented)
      const row = (await paragraphRows(file)).find(row => row.text === "The term is one year.")!
      const result = await applyAmendments({ docxPath: file, outPath: path.join(dir, "applied.docx"), amendments: [
        { id: "a1", pid: row.pid, existing_sha256: row.sha256, updated: "The term is two years.", reason: "Agreed term.", sources: [] }] })
      expect(result.applied.map(row => row.id)).toEqual(["a1"])
      expectSameStories(readFileSync(file), readFileSync(path.join(dir, "applied.docx")))
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  }, 120000)

  test("the redline of a furnished body keeps its header and footer and shows no change in them", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "page-furniture-"))
    try {
      for (const furniture of [AGREEMENT, LETTER]) {
        const baseline = (await renderMarkdownToDocx(BODY, undefined, furniture)).bytes
        const file = path.join(dir, "body.docx")
        writeFileSync(file, baseline)
        const row = (await paragraphRows(file)).find(row => row.text === "The fees are payable monthly.")!
        await applyAmendments({ docxPath: file, outPath: path.join(dir, "applied.docx"), amendments: [
          { id: "a1", pid: row.pid, existing_sha256: row.sha256, updated: "The fees are payable quarterly.", reason: "Agreed fees.", sources: [] }] })
        const redline = await compareDocxWithPointers(readFileSync(file), readFileSync(path.join(dir, "applied.docx")), { authorName: "Reviewer" })
        const [before, after] = [unzipSync(baseline), unzipSync(redline)]
        expect(strFromU8(after["word/document.xml"]!)).toContain("<w:ins ")
        expect(childNames(sectPr(redline))).toEqual(childNames(sectPr(baseline)))
        const stories = Object.keys(before).filter(name => /^word\/(?:header|footer)\d+\.xml$/.test(name))
        expect(Object.keys(after).filter(name => /^word\/(?:header|footer)\d+\.xml$/.test(name))).toEqual(stories)
        for (const name of stories) {
          expect(strFromU8(after[name]!)).not.toMatch(/<w:(?:ins|del)\b/)
          expect(paragraphs(strict(after[name]!))).toEqual(paragraphs(strict(before[name]!)))
        }
      }
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  }, 120000)

  test("the combined file keeps the body's furniture and gives the memo none", async () => {
    const body = (await renderMarkdownToDocx(BODY, undefined, AGREEMENT)).bytes
    const memo = (await renderMarkdownToDocx("# Drafting notes\n\nThe term follows the instructions.", undefined, { header: [NOTES_MARKING], pageNumber: { format: "plain", align: "center" } })).bytes
    const combined = combineMemoAndBody({ memo, body })
    const [before, after] = [unzipSync(body), unzipSync(combined)]
    expect(Object.keys(after).filter(name => /header|footer/.test(name))).toEqual(["word/header1.xml", "word/footer1.xml"])
    for (const name of ["word/header1.xml", "word/footer1.xml"]) expect(after[name]).toEqual(before[name]!)
    expect(Object.values(after).some(data => strFromU8(data).includes(NOTES_MARKING))).toBe(false)
    // The memo's section break carries only page setup; the body's section keeps
    // its references and restarts numbering at 1.
    const document = strict(after["word/document.xml"]!)
    const sections = all(document, W, "sectPr")
    expect(sections).toHaveLength(2)
    expect(childNames(sections[0]!)).toEqual(["type", "pgSz", "pgMar"])
    expect(childNames(sections[1]!)).toEqual(["headerReference", "footerReference", "pgSz", "pgMar", "pgNumType"])
  }, 120000)

  test.skipIf(!Bun.which("soffice"))("a letter has no number on its first page and numbers the rest", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "page-furniture-"))
    try {
      const text = "The parties record the arrangements set out in this letter for the services described below. ".repeat(4)
      const letter = ["# Engagement letter", ...Array.from({ length: 20 }, (_, index) => `Paragraph ${index + 1}. ${text}`)].join("\n\n")
      writeFileSync(path.join(dir, "letter.docx"), (await renderMarkdownToDocx(letter, undefined, LETTER)).bytes)
      const converted = Bun.spawnSync(["soffice", `-env:UserInstallation=file://${dir}/profile`, "--headless", "--convert-to", "pdf", "--outdir", dir, path.join(dir, "letter.docx")])
      expect(converted.exitCode).toBe(0)
      const { extractText, getDocumentProxy } = await import("unpdf")
      const { text: pages } = await extractText(await getDocumentProxy(new Uint8Array(readFileSync(path.join(dir, "letter.pdf")))), { mergePages: false })
      expect(pages.length).toBeGreaterThan(1)
      expect(pages[0]!.split("\n")[0]).toBe("Engagement letter")
      expect(pages[1]!.split("\n")[0]).toBe("2")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  }, 180000)
})

function expectSameStories(before: Uint8Array, after: Uint8Array) {
  const [a, b] = [unzipSync(before), unzipSync(after)]
  const stories = Object.keys(a).filter(name => /^word\/(?:header|footer)\d+\.xml$/.test(name))
  expect(stories.length).toBeGreaterThan(0)
  expect(Object.keys(b).filter(name => /^word\/(?:header|footer)\d+\.xml$/.test(name))).toEqual(stories)
  for (const name of [...stories, "word/_rels/document.xml.rels", "[Content_Types].xml"]) expect(Buffer.from(b[name]!).equals(Buffer.from(a[name]!))).toBe(true)
  expect(elementChildren(sectPr(after)).filter(node => node.localName.endsWith("Reference")).map(node => node.toString()))
    .toEqual(elementChildren(sectPr(before)).filter(node => node.localName.endsWith("Reference")).map(node => node.toString()))
}

function strict(data: Uint8Array) {
  const fail = (message: string) => { throw new Error(message) }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(strFromU8(data), "application/xml")
}

function all(node: Document | Element, namespace: string, name: string) {
  return Array.from(node.getElementsByTagNameNS(namespace, name))
}

function elementChildren(node: Element) {
  return Array.from(node.childNodes).filter((row): row is Element => row.nodeType === 1)
}

function childNames(node: Element) {
  return elementChildren(node).map(row => row.localName)
}

function sectPr(bytes: Uint8Array) {
  const body = all(strict(unzipSync(bytes)["word/document.xml"]!), W, "body")[0]!
  return elementChildren(body).findLast(node => node.localName === "sectPr")!
}

function paragraphs(doc: Document) {
  return all(doc, W, "p").map(paragraph => ({
    style: all(paragraph, W, "pStyle")[0]?.getAttributeNS(W, "val"),
    align: all(paragraph, W, "jc")[0]?.getAttributeNS(W, "val"),
    text: all(paragraph, W, "t").map(node => node.textContent).join(""),
  }))
}
