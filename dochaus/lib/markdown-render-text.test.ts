import { expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "./render"
import { prepareMarkdownRenderText } from "./markdown-render-text"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

test("escaped pipes and emphasis stay literal inside rendered DOCX table cells", async () => {
  const markdown = "| Field | Value |\n|---|---|\n| Pipe | A\\|B |\n| Literal | \\*term\\* and \\[source: note\\] |\n| Entity | A&#124;B |"
  const { bytes } = await renderMarkdownToDocx(markdown, undefined, undefined, { numbering: false, styleTemplate: false })
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!)
  const doc = new DOMParser().parseFromString(xml, "application/xml")
  const tables = Array.from(doc.getElementsByTagNameNS(W, "tbl"))
  expect(tables).toHaveLength(1)
  const rows = Array.from(tables[0]!.getElementsByTagNameNS(W, "tr"))
  expect(rows).toHaveLength(4)
  expect(rows.map(row => Array.from(row.getElementsByTagNameNS(W, "tc")).map(cell =>
    Array.from(cell.getElementsByTagNameNS(W, "t")).map(text => text.textContent).join(""))))
    .toEqual([["Field", "Value"], ["Pipe", "A|B"], ["Literal", "*term* and [source: note]"], ["Entity", "A|B"]])
  expect(xml).not.toContain("DocHausText")
})

test("only text escapes are protected; code and Markdown link destinations retain their syntax", () => {
  const source = "Literal \\*term\\* \\| &#124; [label](https://example.invalid/a\\(b\\))\n\n`A\\|B`\n\n```text\nA\\|B\n```\n\n    A\\|B"
  const prepared = prepareMarkdownRenderText(source).markdown
  expect(prepared).not.toContain("Literal \\*")
  expect(prepared).toContain("https://example.invalid/a\\(b\\)")
  expect(prepared).toContain("`A\\|B`")
  expect(prepared).toContain("```text\nA\\|B\n```")
  expect(prepared).toContain("    A\\|B")
})
