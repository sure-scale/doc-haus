import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { strToU8, zipSync } from "fflate"
import { extractDocumentText, extractSourceText, matterDocumentName } from "./ingest"
import { chunkText } from "../../../dochaus/lib/chunk-text"
import { liveText } from "../../../dochaus/lib/extract"
import { verifyCitation } from "../../../dochaus/plugin/legal/citations"
import { parseZip } from "../../../dochaus/lib/zip"
import { sectionize, chunkSection } from "./structure"

const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
// Independently authored synthetic correspondence, with repeated wording and
// adversarial markup. It contains no client material or model-expected answers.
const source =
  "# Supplied record\r\n\r\nOn 6 March 2025, Elm Works collected the cartons.\r\n\r\nReceipt was signed later.\r\n\r\nReceipt was signed later.\r\n<script>window.external()</script>\r\n[open](javascript:alert(1))\r\n"

test("Markdown source extraction, index coordinates and live citation verification share one canonical text", async () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "markdown-source-"))
  roots.push(root)
  for (const name of ["Chronology.md", "Record.MARKDOWN"]) {
    const bytes = Buffer.from(source)
    const file = path.join(root, name)
    writeFileSync(file, bytes)
    const extracted = await extractSourceText(name, bytes)
    const read = await extractDocumentText(name, bytes)
    const indexed = chunkText(
      sectionize(extracted.text)
        .flatMap((section) => chunkSection(section))
        .map((chunk) => ({ text: chunk.text, char_start: chunk.charStart })),
    ).text
    expect(extracted.quality.status).toBe("ok")
    expect(read).toBe(extracted.text)
    expect(await liveText(file)).toBe(indexed)
    const excerpt = "Receipt was signed later."
    const charStart = indexed.lastIndexOf(excerpt)
    const citation = await verifyCitation({
      docPath: file,
      documentName: name,
      section: "Record",
      excerpt,
      charStart,
      charEnd: charStart + excerpt.length,
    })
    expect(citation?.charStart).toBe(charStart)
    expect(citation?.reanchored).toBeUndefined()
    expect(indexed).toContain("<script>window.external()</script>")
  }
})

test("Markdown upload registration and archive entries accept only safe supported basenames", () => {
  expect(matterDocumentName("chronology.md")).toBe("chronology.md")
  expect(matterDocumentName("Statement.markdown")).toBe("Statement.markdown")
  expect(matterDocumentName("../chronology.md")).toBeUndefined()
  const zip = parseZip(zipSync({ "bundle/chronology.md": strToU8(source) }))
  expect(zip.entries[0]?.supported).toBe(true)
  expect(zip.entries[0]?.name).toBe("chronology.md")
})

test("empty Markdown is unreadable and invalid UTF-8 is rejected without fabricated text", async () => {
  expect((await extractSourceText("empty.md", Buffer.from("  \r\n"))).quality.status).toBe("no_readable_text")
  await expect(extractSourceText("invalid.md", Buffer.from([0xc3, 0x28]))).rejects.toThrow()
})
