import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { draftingAttribution, withGeneratedDocxAttribution } from "./docx-attribution"
import { exportDocxComments } from "./docx-comments"
import { compareDocxWithPointers } from "./drafting-amendment"
import { renderMarkdownToDocx } from "./render"

const previous = process.env.WORKSPACE_ROOT
const roots: string[] = []
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

afterEach(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
})

function preferences(value?: Record<string, unknown>) {
  const root = mkdtempSync(path.join(tmpdir(), "dh-attribution-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  if (!value) return
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify(value))
}

function property(bytes: Uint8Array, part: string, namespace: string, name: string) {
  const document = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[part]), "application/xml")
  return document.getElementsByTagNameNS(namespace, name)[0]?.textContent
}

test("new DOCX metadata reads saved identity on each render and escapes XML", async () => {
  preferences({ attorney: "  Morgan & <Lee>  ", firm: "  Lee & Partners  ", searchApiKey: "synthetic-secret" })
  const first = (await renderMarkdownToDocx("# Review note\n\nA newly prepared note.")).bytes
  expect(property(first, "docProps/core.xml", "http://purl.org/dc/elements/1.1/", "creator")).toBe("Morgan & <Lee>")
  expect(property(first, "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "lastModifiedBy")).toBe("Morgan & <Lee>")
  expect(property(first, "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")).toBe("Lee & Partners")
  expect(Object.entries(unzipSync(first)).filter(([name]) => name.startsWith("docProps/")).map(([, bytes]) => strFromU8(bytes)).join("\n")).not.toContain("synthetic-secret")
  writeFileSync(path.join(process.env.WORKSPACE_ROOT!, ".preferences/preferences.json"), JSON.stringify({ attorney: "Next Lawyer", firm: "Next Firm" }))
  const next = (await renderMarkdownToDocx("A subsequent note.")).bytes
  expect(property(next, "docProps/core.xml", "http://purl.org/dc/elements/1.1/", "creator")).toBe("Next Lawyer")
  expect(property(next, "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")).toBe("Next Firm")
})

test("missing identity falls back safely, and a firm alone supplies the author", async () => {
  preferences()
  expect(draftingAttribution()).toEqual({ author: "doc.haus", company: "doc.haus" })
  const fallback = (await renderMarkdownToDocx("A note without a configured identity.")).bytes
  expect(property(fallback, "docProps/core.xml", "http://purl.org/dc/elements/1.1/", "creator")).toBe("doc.haus")
  expect(property(fallback, "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")).toBe("doc.haus")
  preferences({ attorney: "   ", firm: "Solo Firm" })
  expect(draftingAttribution()).toEqual({ author: "Solo Firm", company: "Solo Firm" })
  preferences({ attorney: 123, firm: null })
  expect(draftingAttribution()).toEqual({ author: "doc.haus", company: "doc.haus" })
})

test("saved and explicit identities produce XML 1.0 metadata without losing Unicode names", async () => {
  const invalid = [...Array.from({ length: 32 }, (_, code) => code).filter(code => ![9, 10, 13].includes(code)), 0xd800, 0xdfff, 0xfffe, 0xffff]
    .map(code => String.fromCharCode(code)).join("\u0001")
  const author = "Zoë 李😀 & <Lee>"
  const company = "李 & Partners 😀"
  preferences({ attorney: `  Zoë ${invalid}李😀 & <Lee>  `, firm: `  李 & ${invalid}Partners 😀  ` })
  const source = (await renderMarkdownToDocx("A newly prepared note.")).bytes
  expect(draftingAttribution()).toEqual({ author, company })
  for (const bytes of [source, withGeneratedDocxAttribution(source, { author: `${invalid}${author}`, company: `${company}${invalid}` })]) {
    const parts = unzipSync(bytes)
    for (const part of ["docProps/core.xml", "docProps/app.xml"]) {
      const errors: string[] = []
      const xml = strFromU8(parts[part])
      new DOMParser({ errorHandler: { warning: issue => errors.push(issue), error: issue => errors.push(issue), fatalError: issue => errors.push(issue) } })
        .parseFromString(xml, "application/xml")
      expect(errors).toEqual([])
      for (const character of invalid) expect(xml).not.toContain(character)
    }
    expect(property(bytes, "docProps/core.xml", "http://purl.org/dc/elements/1.1/", "creator")).toBe(author)
    expect(property(bytes, "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")).toBe(company)
  }
  const whitespace = withGeneratedDocxAttribution(source, { author: "Zoë\t李\n😀 & <Lee>", company })
  expect(property(whitespace, "docProps/core.xml", "http://purl.org/dc/elements/1.1/", "creator")).toBe("Zoë\t李\n😀 & <Lee>")
  preferences({ attorney: invalid, firm: `Firm${invalid}` })
  expect(draftingAttribution()).toEqual({ author: "Firm", company: "Firm" })
  preferences({ attorney: invalid, firm: invalid })
  expect(draftingAttribution()).toEqual({ author: "doc.haus", company: "doc.haus" })
})

test("stamping a generated derivative preserves prior revision and comment authors", async () => {
  preferences({ attorney: "Current Lawyer", firm: "Current Firm" })
  const before = (await renderMarkdownToDocx("The review period is five days.")).bytes
  const after = (await renderMarkdownToDocx("The review period is seven days.")).bytes
  const redline = await compareDocxWithPointers(before, after, { authorName: "Previous Reviewer" })
  const annotated = exportDocxComments(redline, [{ id: "prior", author: "Previous Commenter", text: "Prior review explanation.", anchor: { quote: "seven days" } }]).bytes
  const parts = unzipSync(annotated)
  const generated = withGeneratedDocxAttribution(annotated)
  const output = unzipSync(generated)
  expect(output["word/document.xml"]).toEqual(parts["word/document.xml"])
  expect(output["word/comments.xml"]).toEqual(parts["word/comments.xml"])
  const doc = new DOMParser().parseFromString(strFromU8(output["word/document.xml"]), "application/xml")
  expect(Array.from(doc.getElementsByTagNameNS(W, "ins")).map(node => node.getAttributeNS(W, "author"))).toContain("Previous Reviewer")
  expect(property(generated, "docProps/core.xml", "http://purl.org/dc/elements/1.1/", "creator")).toBe("Current Lawyer")
})
