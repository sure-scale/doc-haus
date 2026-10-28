import { expect, test } from "bun:test"
import { sameWordChatOoxml, verifiedNativeWordChatFormatting, wordChatOoxmlDifference } from "./word-chat-host-ooxml"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const styles = '<w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:pPr><w:spacing w:after="160"/></w:pPr></w:style></w:styles>'
function flat(paragraph = '<w:p><w:r><w:t>Reviewed text.</w:t></w:r></w:p>', dependency = styles, padding = "0") {
  return `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body>${paragraph}</w:body></w:document></pkg:xmlData></pkg:part>${dependency ? `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" pkg:padding="${padding}"><pkg:xmlData>${dependency}</pkg:xmlData></pkg:part>` : ""}</pkg:package>`
}

test("dependency diagnostics identify package attributes without ignoring their values", () => {
  const before = flat(), changed = flat(undefined, styles, "hidden-package-value")
  expect(sameWordChatOoxml(before, before, before)).toBe(true)
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  const reason = wordChatOoxmlDifference(before, changed, before)!
  expect(reason).toContain("/word/styles.xml/part/@padding")
  expect(reason).toContain("attribute changed")
  expect(reason).not.toContain("hidden-package-value")
})

test("dependency diagnostics identify meaningful nested style properties and missing captures", () => {
  const before = flat(), changed = flat(undefined, styles.replace('after="160"', 'after="240"'))
  expect(wordChatOoxmlDifference(before, before, before)).toBeUndefined()
  expect(wordChatOoxmlDifference(before, changed, before)).toContain("/styles[1]/style[1]/pPr[1]/spacing/@after")
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  expect(wordChatOoxmlDifference(before, flat(undefined, ""), before)).toBe("/word/styles.xml: dependency missing from current Word capture")
  expect(wordChatOoxmlDifference(flat(undefined, ""), before, before)).toBe("/word/styles.xml: dependency missing from reviewed Word capture")
})

test("main paragraph diagnostics retain the exact existing write-ID guard", () => {
  const before = flat('<w:p w:rsidR="1"><w:r><w:t>Reviewed text.</w:t></w:r></w:p>')
  const observedWrite = before.replace('rsidR="1"', 'rsidR="2"')
  expect(sameWordChatOoxml(before, observedWrite, before)).toBe(true)
  const changed = before.replace('rsidR="1"', 'rsidR="1" w:rsidP="unverified-write-value"')
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  const reason = wordChatOoxmlDifference(before, changed, before)!
  expect(reason).toContain("/word/document.xml[1]/p/@rsidP")
  expect(reason).toContain("attribute added")
  expect(reason).not.toContain("unverified-write-value")
})

test("main paragraph text diagnostics omit source text and values", () => {
  const before = flat('<w:p><w:r><w:t>Private original wording.</w:t></w:r></w:p>')
  const changed = flat('<w:p><w:r><w:t>Private changed wording.</w:t></w:r></w:p>')
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  const reason = wordChatOoxmlDifference(before, changed, before)!
  expect(reason).toContain("/word/document.xml[1]/p[1]/r[1]/t[1]: text changed")
  expect(reason).not.toContain("Private")
  expect(reason).not.toContain("wording")
})

test("native formatting must retain the prior properties and inline revisions and use the approved properties", () => {
  const before = flat('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:ins w:id="4" w:author="Earlier"><w:r><w:t>Reviewed text.</w:t></w:r></w:ins></w:p>')
  const result = flat('<w:p><w:pPr><w:spacing w:after="240"/><w:pPrChange w:id="5" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrChange></w:pPr><w:ins w:id="4" w:author="Earlier"><w:r><w:t>Reviewed text.</w:t></w:r></w:ins></w:p>')
  const native = result.replace('w:author="doc.haus"', 'w:author="Word user"')
  const operations = [{ paragraph: 0, spacingAfter: 12 }]
  expect(verifiedNativeWordChatFormatting(before, result, native, operations)).toEqual(["Word user"])
  for (const corrupted of [native.replace('w:after="120"', 'w:after="140"'), native.replace('w:after="240"', 'w:after="260"'),
    native.replace('w:author="Earlier"', 'w:author="Other"'), native.replace('w:author="Word user"', 'w:author=""')])
    expect(() => verifiedNativeWordChatFormatting(before, result, corrupted, operations)).toThrow()
})
