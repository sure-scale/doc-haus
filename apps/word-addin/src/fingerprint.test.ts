import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { documentFingerprint } from "./fingerprint"

const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const format = '<w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:lang w:val="en-GB"/><w:sz w:val="22"/><w:color w:val="123456"/></w:rPr>'
const run = (text: string, properties = "", space = "") =>
  `<w:r>${properties}<w:t${space ? ` xml:space="${space}"` : ""}>${text}</w:t></w:r>`
const document = (content: string, paragraphAttributes = "") =>
  `<w:document xmlns:w="${word}" xmlns:c="urn:synthetic"><w:body><w:p${paragraphAttributes}>${content}</w:p></w:body></w:document>`
const fingerprint = (content: string, paragraphAttributes = "") =>
  documentFingerprint(zipSync({ "word/document.xml": strToU8(document(content, paragraphAttributes)) }))
const split = (properties = "", first = "Neutral ", middle = "sample", last = " text.") =>
  `${run(first, properties, "preserve")}<w:proofErr w:type="gramStart"/>${run(middle, properties)}<w:proofErr w:type="gramEnd"/>${run(last, properties, "preserve")}`

test("proofing may split equivalent plain or explicitly formatted text without changing the fingerprint", async () => {
  for (const properties of ["", format]) {
    const original = zipSync({ "word/document.xml": strToU8(document(run("Neutral sample text.", properties))) })
    const copy = original.slice()
    const baseline = await documentFingerprint(original)
    expect(await fingerprint(split(properties))).toBe(baseline)
    expect(await fingerprint(split(properties).replaceAll("gramStart", "spellStart").replaceAll("gramEnd", "spellEnd"))).toBe(baseline)
    expect(await fingerprint(run("Neutral sample text.", properties, "preserve"))).toBe(baseline)
    expect(original).toEqual(copy)
  }
})

test("text, complete formatting, revisions and unknown format payload remain significant", async () => {
  const baseline = await fingerprint(run("Neutral sample text.", format))
  expect(await fingerprint(split(format, "Changed "))).not.toBe(baseline)
  expect(await fingerprint(split(format, "Neutral  "))).not.toBe(baseline)
  for (const changed of [
    format.replace("<w:b/>", '<w:b w:val="false"/>'),
    format.replaceAll("Aptos", "Arial"),
    format.replace("en-GB", "en-US"),
    format.replace('w:val="22"', 'w:val="24"'),
    format.replace("123456", "654321"),
  ])
    expect(await fingerprint(`${run("Neutral ", format, "preserve")}${run("sample", changed)}${run(" text.", format, "preserve")}`)).not.toBe(baseline)
  for (const payload of [
    '<w:rPrChange w:id="1" w:author="Synthetic reviewer"><w:rPr><w:b/></w:rPr></w:rPrChange>',
    '<w:unknown w:val="preserved"/>',
    '<w:constructor w:val="preserved"/>',
    '<c:format c:value="preserved"/>',
    '<w:b c:value="preserved"/>',
  ]) {
    const properties = `<w:rPr>${payload}</w:rPr>`
    expect(await fingerprint(split(properties))).not.toBe(await fingerprint(run("Neutral sample text.", properties)))
    expect(await fingerprint(run("Neutral sample text.", properties))).not.toBe(
      await fingerprint(run("Neutral sample text.", properties.replace("preserved", "changed").replace('w:id="1"', 'w:id="2"'))),
    )
  }
  const revision = (id: string, author: string) => `<w:ins w:id="${id}" w:author="${author}">${run("Neutral sample text.")}</w:ins>`
  expect(await fingerprint(revision("1", "Synthetic reviewer"))).not.toBe(await fingerprint(revision("2", "Synthetic reviewer")))
  expect(await fingerprint(revision("1", "Synthetic reviewer"))).not.toBe(await fingerprint(revision("1", "Another reviewer")))
})

test("coalescing does not cross structural boundaries or discard unknown run and text attributes", async () => {
  for (const boundary of [
    '<w:bookmarkStart w:id="1" w:name="Synthetic"/>',
    '<w:bookmarkEnd w:id="1"/>',
    '<w:commentRangeStart w:id="1"/>',
    '<w:commentRangeEnd w:id="1"/>',
    '<w:ins w:id="1" w:author="Synthetic reviewer"/>',
    '<w:del w:id="1" w:author="Synthetic reviewer"/>',
    '<w:hyperlink w:anchor="Synthetic"/>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:r><w:instrText> DATE </w:instrText></w:r>',
    '<w:r><w:tab/></w:r>',
    '<c:boundary/>',
    '<!-- preserved boundary -->',
  ])
    expect(await fingerprint(`${run("one")}${boundary}${run("two")}`)).not.toBe(await fingerprint(`${run("onetwo")}${boundary}`))
  for (const content of [
    '<w:r w:unknown="preserved"><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r>',
    '<w:r c:value="preserved"><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r>',
    '<w:r><w:t c:value="preserved">one</w:t></w:r><w:r><w:t>two</w:t></w:r>',
    '<w:r><w:t>one<c:data/></w:t></w:r><w:r><w:t>two</w:t></w:r>',
  ])
    expect(await fingerprint(content)).not.toBe(await fingerprint(run("onetwo")))
})

test("boundary whitespace, default and inherited preservation cannot conceal spacing differences", async () => {
  const joined = await fingerprint(run("one two"))
  expect(await fingerprint(`${run("one ", "", "preserve")}${run("two")}`)).toBe(joined)
  for (const space of ["", "default"])
    expect(await fingerprint(`${run("one ", "", space)}${run("two")}`)).not.toBe(joined)
  expect(await fingerprint(run(" one", "", "preserve"))).not.toBe(await fingerprint(run(" one")))
  expect(await fingerprint(run("one ", "", "preserve"))).not.toBe(await fingerprint(run("one ", "", "default")))
  expect(await fingerprint(run(" one"), ' xml:space="preserve"')).not.toBe(
    await fingerprint(run(" one"), ' xml:space="default"'),
  )
  expect(await fingerprint(`${run("one\t")}${run("two")}`)).not.toBe(await fingerprint(run("onetwo")))
})

test("only standard empty proofing anchors in Word flow are ignored", async () => {
  for (const marker of [
    '<w:proofErr w:type="future"/>',
    '<w:proofErr w:type="gramStart" w:unknown="preserved"/>',
    '<w:proofErr w:type="gramStart" c:value="preserved"/>',
    '<w:proofErr w:type="gramStart"><c:data/></w:proofErr>',
    '<w:proofErr w:type="gramStart">preserved</w:proofErr>',
    '<c:proofErr w:type="gramStart"/>',
  ])
    expect(await fingerprint(`${run("one")}${marker}`)).not.toBe(await fingerprint(run("one")))
  const plain = document(run("Neutral sample text."))
  const fragmented = document(split())
  expect(await documentFingerprint(zipSync({ "customXml/item1.xml": strToU8(plain) }))).not.toBe(
    await documentFingerprint(zipSync({ "customXml/item1.xml": strToU8(fragmented) })),
  )
  expect(await documentFingerprint(zipSync({ "word/custom.xml": strToU8(plain) }))).not.toBe(
    await documentFingerprint(zipSync({ "word/custom.xml": strToU8(fragmented) })),
  )
  const foreign = (content: string) =>
    documentFingerprint(zipSync({ "word/document.xml": strToU8(document(`<c:data><w:p>${content}</w:p></c:data>`)) }))
  expect(await foreign(run("Neutral sample text."))).not.toBe(await foreign(split()))
})
