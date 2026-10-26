import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { documentFingerprint } from "./fingerprint"

const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const format =
  '<w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:lang w:val="en-GB"/><w:sz w:val="22"/><w:color w:val="123456"/></w:rPr>'
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
    expect(
      await fingerprint(split(properties).replaceAll("gramStart", "spellStart").replaceAll("gramEnd", "spellEnd")),
    ).toBe(baseline)
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
    expect(
      await fingerprint(
        `${run("Neutral ", format, "preserve")}${run("sample", changed)}${run(" text.", format, "preserve")}`,
      ),
    ).not.toBe(baseline)
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
      await fingerprint(
        run("Neutral sample text.", properties.replace("preserved", "changed").replace('w:id="1"', 'w:id="2"')),
      ),
    )
  }
  const revision = (id: string, author: string) =>
    `<w:ins w:id="${id}" w:author="${author}">${run("Neutral sample text.")}</w:ins>`
  expect(await fingerprint(revision("1", "Synthetic reviewer"))).not.toBe(
    await fingerprint(revision("2", "Synthetic reviewer")),
  )
  expect(await fingerprint(revision("1", "Synthetic reviewer"))).not.toBe(
    await fingerprint(revision("1", "Another reviewer")),
  )
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
    "<w:r><w:instrText> DATE </w:instrText></w:r>",
    "<w:r><w:tab/></w:r>",
    "<c:boundary/>",
    "<!-- preserved boundary -->",
  ])
    expect(await fingerprint(`${run("one")}${boundary}${run("two")}`)).not.toBe(
      await fingerprint(`${run("onetwo")}${boundary}`),
    )
  for (const content of [
    '<w:r w:unknown="preserved"><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r>',
    '<w:r c:value="preserved"><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r>',
    '<w:r><w:t c:value="preserved">one</w:t></w:r><w:r><w:t>two</w:t></w:r>',
    "<w:r><w:t>one<c:data/></w:t></w:r><w:r><w:t>two</w:t></w:r>",
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

const coreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
const termsNamespace = "http://purl.org/dc/terms/"
const coreProperties = (modified: string, options: Record<string, string | undefined> = {}) =>
  `<p:coreProperties xmlns:p="${coreNamespace}" xmlns:t="${termsNamespace}" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns:c="urn:synthetic"><dc:creator>${options.creator ?? "Synthetic author"}</dc:creator><p:lastModifiedBy>${options.lastModifiedBy ?? "Synthetic editor"}</p:lastModifiedBy><t:created i:type="t:W3CDTF">${options.created ?? "2025-01-01T00:00:00Z"}</t:created><p:lastPrinted>${options.lastPrinted ?? "2025-02-01T00:00:00Z"}</p:lastPrinted><p:revision>${options.revision ?? "1"}</p:revision><t:modified i:type="t:W3CDTF">${modified}</t:modified></p:coreProperties>`

test("a changed standard save timestamp does not hide real text edits or other core metadata", async () => {
  const fixture = (modified: string, options: Record<string, string | undefined> = {}, text = "Neutral text.") =>
    zipSync({
      "docProps/core.xml": strToU8(coreProperties(modified, options)),
      "word/document.xml": strToU8(document(run(text))),
    })
  const first = fixture("2026-01-02T03:04:05Z")
  const copy = first.slice()
  const baseline = await documentFingerprint(first)
  expect(await documentFingerprint(fixture("2026-02-03T04:05:06Z"))).toBe(baseline)
  expect(await documentFingerprint(fixture("2024-02-29T03:04:05.123+01:00"))).toBe(baseline)
  expect(await documentFingerprint(fixture("2026-02-03T04:05:06Z", {}, "Changed text."))).not.toBe(baseline)
  for (const options of [
    { creator: "Another author" },
    { lastModifiedBy: "Another editor" },
    { created: "2025-03-01T00:00:00Z" },
    { lastPrinted: "2025-04-01T00:00:00Z" },
  ])
    expect(await documentFingerprint(fixture("2026-02-03T04:05:06Z", options))).not.toBe(baseline)
  expect(first).toEqual(copy)
})

test("modified timestamp comparison preserves nonstandard roots, opaque data, extensions and invalid dates", async () => {
  const first = "2026-01-02T03:04:05Z"
  const second = "2026-02-03T04:05:06Z"
  const cases: [string, (time: string) => string][] = [
    ["customXml/item1.xml", (time) => coreProperties(time)],
    ["docProps/custom.xml", (time) => coreProperties(time)],
    ["docProps/core.xml", (time) => coreProperties(time).replace(coreNamespace, "urn:synthetic:core")],
    ["docProps/core.xml", (time) => coreProperties(time).replaceAll("p:coreProperties", "p:other")],
    ["docProps/core.xml", (time) => coreProperties(time).replaceAll("t:modified", "c:modified")],
    [
      "docProps/core.xml",
      (time) =>
        coreProperties(time)
          .replace("<t:modified", "<p:other><t:modified")
          .replace("</t:modified>", "</t:modified></p:other>"),
    ],
    ["docProps/core.xml", (time) => coreProperties(time).replace("<t:modified", '<t:modified c:flag="preserved"')],
    [
      "docProps/core.xml",
      (time) => coreProperties(time).replace(`>${time}</t:modified>`, `><c:data>${time}</c:data></t:modified>`),
    ],
    ["docProps/core.xml", (time) => coreProperties(time).replaceAll('i:type="t:W3CDTF"', 'i:type="c:W3CDTF"')],
    ["docProps/core.xml", (time) => coreProperties(time).replaceAll('i:type="t:W3CDTF"', 'i:type="t:unknown"')],
    [
      "docProps/core.xml",
      (time) =>
        coreProperties(time)
          .replace("<p:coreProperties ", `<p:coreProperties xmlns="${termsNamespace}" `)
          .replaceAll('i:type="t:W3CDTF"', 'i:type=":W3CDTF"'),
    ],
    [
      "docProps/core.xml",
      (time) =>
        coreProperties(time).replace(
          `<t:modified i:type="t:W3CDTF">${time}</t:modified>`,
          `<t:modified i:type="t:W3CDTF">${time}</t:modified><t:modified i:type="t:W3CDTF">${time}</t:modified>`,
        ),
    ],
  ]
  for (const [name, xml] of cases) {
    const hash = (time: string) => documentFingerprint(zipSync({ [name]: strToU8(xml(time)) }))
    expect(await hash(first)).not.toBe(await hash(second))
  }
  for (const [left, right] of [
    ["future-one", "future-two"],
    ["2026-02-30T03:04:05Z", "2026-02-31T03:04:05Z"],
    ["2026-01-02T25:04:05Z", "2026-01-02T26:04:05Z"],
  ]) {
    const hash = (time: string) => documentFingerprint(zipSync({ "docProps/core.xml": strToU8(coreProperties(time)) }))
    expect(await hash(left!)).not.toBe(await hash(right!))
  }
})

test("observed cached counts and numeric package revision can change without hiding content or other metadata", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  const app = `<e:Properties xmlns:e="${extended}"><e:Words>3</e:Words><e:Characters>15</e:Characters><e:Paragraphs>2</e:Paragraphs><e:CharactersWithSpaces>17</e:CharactersWithSpaces><e:Pages>1</e:Pages><e:Lines>2</e:Lines><e:Company>Synthetic company</e:Company><e:Unknown>preserved</e:Unknown></e:Properties>`
  const core = coreProperties("2026-01-02T03:04:05Z")
  const fixture = (application = app, properties = core, content = run("Neutral text.")) =>
    zipSync({
      "docProps/app.xml": strToU8(application),
      "docProps/core.xml": strToU8(properties),
      "word/document.xml": strToU8(document(content)),
    })
  const original = fixture()
  const copy = original.slice()
  const baseline = await documentFingerprint(original)
  for (const property of ["Words", "Characters", "Paragraphs", "CharactersWithSpaces"])
    expect(
      await documentFingerprint(
        fixture(app.replace(new RegExp(`<e:${property}>\\d+</e:${property}>`), `<e:${property}>99</e:${property}>`)),
      ),
    ).toBe(baseline)
  expect(await documentFingerprint(fixture(app, core.replace("<p:revision>1", "<p:revision>2")))).toBe(baseline)
  expect(await documentFingerprint(fixture(app, core, run("Changed text.")))).not.toBe(baseline)
  expect(await documentFingerprint(fixture(app, core, run("Neutral text.", format)))).not.toBe(baseline)
  for (const edit of [
    app.replace("Synthetic company", "Another company"),
    app.replace("preserved", "changed"),
    app.replace("<e:Pages>1", "<e:Pages>2"),
    app.replace("<e:Lines>2", "<e:Lines>3"),
  ])
    expect(await documentFingerprint(fixture(edit))).not.toBe(baseline)
  expect(original).toEqual(copy)
})

test("numeric metadata comparison preserves duplicate leaves, extensions, unknown XML and invalid values", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  for (const [part, namespace, root, property] of [
    ["docProps/app.xml", extended, "Properties", "Words"],
    ["docProps/app.xml", extended, "Properties", "Characters"],
    ["docProps/app.xml", extended, "Properties", "Paragraphs"],
    ["docProps/app.xml", extended, "Properties", "CharactersWithSpaces"],
    ["docProps/core.xml", coreNamespace, "coreProperties", "revision"],
  ] as const) {
    const wrapper = (child: string) => `<p:${root} xmlns:p="${namespace}" xmlns:c="urn:synthetic">${child}</p:${root}>`
    const leaf = (value: string) => `<p:${property}>${value}</p:${property}>`
    const cases: [string, (value: string) => string][] = [
      ["customXml/item1.xml", (value) => wrapper(leaf(value))],
      [part, (value) => wrapper(leaf(value)).replace(namespace, "urn:synthetic:wrong")],
      [part, (value) => wrapper(leaf(value)).replaceAll(`p:${root}`, "p:UnknownRoot")],
      [part, (value) => wrapper(leaf(value)).replaceAll(`p:${property}`, `c:${property}`)],
      [part, (value) => wrapper(`<p:Other>${leaf(value)}</p:Other>`)],
      [part, (value) => wrapper(leaf(value) + leaf(value))],
      [part, (value) => wrapper(`<p:${property} c:flag="preserved">${value}</p:${property}>`)],
      [part, (value) => wrapper(`<p:${property}><c:data>${value}</c:data></p:${property}>`)],
      [part, (value) => wrapper(leaf(`custom-${value}`))],
      [part, (value) => wrapper(leaf(`-${value}`))],
      [part, (value) => wrapper(leaf(`${value}.5`))],
    ]
    for (const [name, xml] of cases) {
      const hash = (value: string) => documentFingerprint(zipSync({ [name]: strToU8(xml(value)) }))
      expect(await hash("1")).not.toBe(await hash("2"))
    }
  }
})
