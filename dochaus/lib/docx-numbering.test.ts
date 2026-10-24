// Run from dochaus/: bun test lib/docx-numbering.test.ts
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { numberDocxClauses } from "./docx-numbering"
import { docxodus } from "./docxodus"
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

### 2.9 Misnumbered

Text under a heading whose number does not follow.

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
      return { shown: info.visibleText.replace(/\s+/g, " ").trim(), text: info.textPreview, label: metadata.list?.generatedLabel, style: metadata.styleId }
    })
  } finally {
    session.close()
  }
}

describe("numberDocxClauses", () => {
  test("turns literal clause numbers into a Word list bound to the heading style, keeping what a reader sees", async () => {
    const literal = (await renderMarkdownToDocx(BODY, undefined, undefined, { numbering: false })).bytes
    const numbered = (await renderMarkdownToDocx(BODY)).bytes
    const before = await blocks(literal)
    const after = await blocks(numbered)
    expect(after.map(row => row.shown)).toEqual(before.map(row => row.shown))
    const row = (shown: string) => after.find(entry => entry.shown === shown)!
    expect(row("1. Definitions")).toMatchObject({ label: "1.", text: "Definitions", style: "Heading2" })
    expect(row("1.2 Headings do not affect interpretation, as clause 2.1 confirms.")).toMatchObject({ label: "1.2" })
    expect(row("(i) a sub-limb.")).toMatchObject({ label: "(i)", text: "a sub-limb." })
    expect(row("(c) the third limb.")).toMatchObject({ label: "(c)" })
    expect(row("2.2 Price")).toMatchObject({ label: "2.2", text: "Price" })
    // A number that does not follow the sequence, and an unnumbered heading, stay as written.
    expect(row("2.9 Misnumbered")).toMatchObject({ label: undefined, text: "2.9 Misnumbered" })
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
    const numbered = (await renderMarkdownToDocx(BODY)).bytes
    expect(numberDocxClauses(numbered)).toBe(numbered)
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
})
