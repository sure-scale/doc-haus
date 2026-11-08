import { expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { withStableRevisionDates } from "./drafting-amendment"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
const FIRST = "R1111111111111111"
const SECOND = "R2222222222222222"
const THIRD = "R3333333333333333"
const DATE = "2026-11-08T10:00:00Z"
const packageBytes = (parts: Record<string, string>) => zipSync(Object.fromEntries(
  Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])))
const body = (tag: string) => `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${tag}</w:body></w:document>`
const relationships = (id: string, target = "https://example.invalid/terms") =>
  `<Relationships xmlns="${PACKAGE}"><Relationship Id="${id}" Type="${R}/hyperlink" Target="${target}" TargetMode="External"/></Relationships>`
const link = (id: string) => `<w:p><w:hyperlink r:id="${id}"><w:r><w:t>Terms</w:t></w:r></w:hyperlink></w:p>`

test("revision-date deduplication preserves opaque R-shaped content-control and extension data", () => {
  for (const content of [
    (id: string) => `<w:sdt><w:sdtPr><w:tag w:val="${id}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Same wording.</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
    (id: string) => `<w:p xmlns:x="urn:literal-extension" x:value="${id}"><w:r><w:t>Same wording.</w:t></w:r></w:p>`,
    (id: string) => `<w:p xmlns:x="urn:foreign-relations"><w:r x:id="${id}"><w:t>Same wording.</w:t></w:r></w:p>`,
  ]) {
    const previous = packageBytes({ "word/document.xml": body(content(FIRST)) })
    expect(withStableRevisionDates(previous, DATE, previous)).toBe(previous)
    const current = packageBytes({ "word/document.xml": body(content(SECOND)) })
    const result = withStableRevisionDates(current, DATE, previous)
    expect(result).not.toBe(previous)
    expect(strFromU8(unzipSync(result)["word/document.xml"]!)).toContain(SECOND)
  }
})

test("equivalent declared relationship identities normalize in their own part namespace", () => {
  const parts = (main: string, header: string) => ({
    "word/document.xml": body(link(main)),
    "word/_rels/document.xml.rels": relationships(main),
    "word/header1.xml": `<w:hdr xmlns:w="${W}" xmlns:rel="${R}">${link(header).replaceAll("r:id", "rel:id")}</w:hdr>`,
    "word/_rels/header1.xml.rels": relationships(header),
  })
  const previous = packageBytes(parts(FIRST, FIRST))
  const current = packageBytes(parts(SECOND, THIRD))
  expect(withStableRevisionDates(current, DATE, previous)).toBe(previous)
})

test("R-shaped relationship targets remain meaningful data, even when a real relationship id also changes", () => {
  const parts = (id: string, target: string) => ({
    "word/document.xml": body(link(id)),
    "word/_rels/document.xml.rels": relationships(id, target),
  })
  const previous = packageBytes(parts(FIRST, FIRST))
  const result = withStableRevisionDates(packageBytes(parts(SECOND, SECOND)), DATE, previous)
  expect(result).not.toBe(previous)
  expect(strFromU8(unzipSync(result)["word/_rels/document.xml.rels"]!)).toContain(`Target="${SECOND}"`)
})

test("undeclared relationship references are never guessed equivalent", () => {
  const previous = packageBytes({ "word/document.xml": body(link(FIRST)) })
  const result = withStableRevisionDates(packageBytes({ "word/document.xml": body(link(SECOND)) }), DATE, previous)
  expect(result).not.toBe(previous)
  expect(strFromU8(unzipSync(result)["word/document.xml"]!)).toContain(SECOND)
})

test("normalization never aliases a meaningful named relationship or dangling reference", () => {
  for (const named of ["R0", "DocHausRelationship0"]) {
    for (const declared of [false, true]) {
      const parts = (reference: string) => ({
        "word/document.xml": body(link(reference)),
        "word/_rels/document.xml.rels": relationships(FIRST).replace("</Relationships>",
          `${declared ? `<Relationship Id="${named}" Type="${R}/hyperlink" Target="https://example.invalid/other" TargetMode="External"/>` : ""}</Relationships>`),
      })
      const previous = packageBytes(parts(FIRST))
      const result = withStableRevisionDates(packageBytes(parts(named)), DATE, previous)
      expect(result).not.toBe(previous)
      expect(strFromU8(unzipSync(result)["word/document.xml"]!)).toContain(named)
    }
  }
})

test("opaque hexadecimal Unid attributes are never treated as comparator identities", () => {
  const markup = (id: string) => body(`<w:p xmlns:x="urn:literal-extension" x:Unid="${id}"><w:r><w:t>Same wording.</w:t></w:r></w:p>`)
  const previous = packageBytes({ "word/document.xml": markup("1111") })
  const result = withStableRevisionDates(packageBytes({ "word/document.xml": markup("2222") }), DATE, previous)
  expect(result).not.toBe(previous)
  expect(strFromU8(unzipSync(result)["word/document.xml"]!)).toContain('x:Unid="2222"')
})
