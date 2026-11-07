import { describe, expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { readerCitations as reader, readerCitationsInDocx as readerDocx } from "./reader-citations"

const SOURCES = new Set(["term-sheet.docx", "notes.eml", "playbook:playbook-nda", "request", "https://example.test/rule/7", "a.docx", "a&b.docx"])
const known = (name: string) => SOURCES.has(name)
const readerCitations = (text: string) => reader(text, known)
const readerCitationsInDocx = (bytes: Uint8Array) => readerDocx(bytes, known)

describe("readerCitations", () => {
  test("a source marker reads as a parenthetical citation; the quoted words stay", () => {
    expect(readerCitations(`The term sheet sets "a fee of 2%" [source: term-sheet.docx].`)).toBe(`The term sheet sets "a fee of 2%" (term-sheet.docx).`)
  })

  test("adjacent markers join; the request's own marker goes; a playbook is named", () => {
    expect(readerCitations(`"Close by May" [source: notes.eml] [source: playbook:playbook-nda] and "use plain words" [source: request], as asked.`))
      .toBe(`"Close by May" (notes.eml; NDA playbook) and "use plain words", as asked.`)
  })

  test("a parenthesised marker and a URL read the same way", () => {
    expect(readerCitations(`"shall file" (source: https://example.test/rule/7)`)).toBe(`"shall file" (https://example.test/rule/7)`)
  })

  test("a parenthetical naming no known source is the drafter's own words and stays", () => {
    const text = `Revenue was flat (Source: Company records) and "fee" [source: term-sheet.docx].`
    expect(readerCitations(text)).toBe(`Revenue was flat (Source: Company records) and "fee" (term-sheet.docx).`)
  })

  test("text without a closed marker is unchanged", () => {
    const text = "See [source: unfinished\nnext line, and [ref: d1-x] stays."
    expect(readerCitations(text)).toBe(text)
  })
})

function docx(body: string) {
  return zipSync({ "word/document.xml": strToU8(`<w:document><w:body>${body}</w:body></w:document>`) })
}
const xml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]!)

describe("readerCitationsInDocx", () => {
  test("a marker split across runs is rewritten once, in the run where it starts", () => {
    const out = xml(readerCitationsInDocx(docx(`<w:p><w:r><w:t>&quot;Fee&quot; [sou</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>rce: a&amp;b.docx</w:t></w:r><w:r><w:t>].</w:t></w:r></w:p>`)))
    expect(out).toBe(`<w:document><w:body><w:p><w:r><w:t xml:space="preserve">"Fee" (a&amp;b.docx)</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"></w:t></w:r><w:r><w:t xml:space="preserve">.</w:t></w:r></w:p></w:body></w:document>`)
  })

  test("a file without markers is left alone", () => {
    const plain = docx(`<w:p><w:r><w:t>No citation.</w:t></w:r></w:p>`)
    expect(readerCitationsInDocx(plain)).toBe(plain)
  })

  test("a redline's struck text takes reader form apart from its visible text", () => {
    const out = xml(readerCitationsInDocx(docx(`<w:p><w:del><w:r><w:delText>"Old" [source: a.docx]</w:delText></w:r></w:del><w:r><w:t>"New" [source: a.docx]</w:t></w:r></w:p>`)))
    expect(out).toContain(`<w:delText xml:space="preserve">"Old" (a.docx)</w:delText>`)
    expect(out).toContain(`"New" (a.docx)`)
    expect(out).not.toContain("[source:")
  })

  test("an empty self-closing paragraph never joins its neighbours into one", () => {
    const out = xml(readerCitationsInDocx(docx(`<w:p/><w:p><w:r><w:t>"Fee" [sou</w:t></w:r></w:p><w:p><w:r><w:t>rce: a.docx]</w:t></w:r></w:p>`)))
    expect(out).toBe(`<w:document><w:body><w:p/><w:p><w:r><w:t>"Fee" [sou</w:t></w:r></w:p><w:p><w:r><w:t>rce: a.docx]</w:t></w:r></w:p></w:body></w:document>`)
  })
})
