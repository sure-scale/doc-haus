import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { documentFingerprint } from "./fingerprint"
import { WORD_ADDIN_ID } from "../../../dochaus/lib/word-addin-contract"

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

const word2010 = "http://schemas.microsoft.com/office/word/2010/wordml"
const separatorSpacing = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
const notePart = (kind: "footnote" | "endnote", paragraphId: string, extra = "", textId = "77777777") =>
  `<w:${kind}s xmlns:w="${word}" xmlns:w14="${word2010}"><w:${kind} w:type="separator" w:id="-1"><w:p w14:paraId="${paragraphId}" w14:textId="${textId}"><w:r><w:separator/></w:r>${extra}</w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p w14:paraId="${paragraphId}" w14:textId="${textId}"><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}></w:${kind}s>`

test("Word may regenerate paragraph IDs on boilerplate note separators without changing the document", async () => {
  const fixture = (id: string, content = run("Neutral text.")) => zipSync({
    "word/document.xml": strToU8(document(content)),
    "word/footnotes.xml": strToU8(notePart("footnote", id)),
    "word/endnotes.xml": strToU8(notePart("endnote", id)),
  })
  const original = fixture("00000001")
  const copy = original.slice()
  const baseline = await documentFingerprint(original)
  expect(await documentFingerprint(fixture("00000002"))).toBe(baseline)
  for (const content of [
    run("Changed text."),
    run("Neutral text.", "<w:rPr><w:b/></w:rPr>"),
    '<w:ins w:id="1" w:author="Synthetic reviewer">' + run("Neutral text.") + "</w:ins>",
    '<w:commentRangeStart w:id="1"/>' + run("Neutral text.") + '<w:commentRangeEnd w:id="1"/>',
    run("Neutral text.") + '<w:r><w:footnoteReference w:id="1"/></w:r>',
    run("Neutral text.") + '<w:r><w:endnoteReference w:id="1"/></w:r>',
  ]) expect(await documentFingerprint(fixture("00000002", content))).not.toBe(baseline)
  for (const kind of ["footnote", "endnote"] as const) {
    const reference = (id: number) => run("Neutral text.") + `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`
    expect(await documentFingerprint(fixture("00000002", reference(2)))).not.toBe(
      await documentFingerprint(fixture("00000001", reference(1))),
    )
  }
  expect(original).toEqual(copy)
})

test("formatted note separators retain spacing while Word regenerates their paragraph IDs", async () => {
  const formatted = (kind: "footnote" | "endnote", id: string, properties = separatorSpacing) =>
    notePart(kind, id).replaceAll("<w:r>", `${properties}<w:r>`)
  const fixture = (id: string, properties = separatorSpacing, content = run("Neutral text.")) => zipSync({
    "word/document.xml": strToU8(document(content)),
    "word/footnotes.xml": strToU8(formatted("footnote", id, properties)),
    "word/endnotes.xml": strToU8(formatted("endnote", id, properties)),
  })
  const original = fixture("00000001")
  const copy = original.slice()
  const baseline = await documentFingerprint(original)
  expect(await documentFingerprint(fixture("00000002"))).toBe(baseline)
  for (const properties of [
    separatorSpacing.replace('w:after="0"', 'w:after="120"'),
    separatorSpacing.replace('w:line="240"', 'w:line="360"'),
    separatorSpacing.replace('w:lineRule="auto"', 'w:lineRule="exact"'),
  ]) expect(await documentFingerprint(fixture("00000002", properties))).not.toBe(baseline)
  for (const content of [
    run("Changed text."),
    run("Neutral text.", "<w:rPr><w:b/></w:rPr>"),
    '<w:ins w:id="1" w:author="Synthetic reviewer">' + run("Neutral text.") + "</w:ins>",
    '<w:commentRangeStart w:id="1"/>' + run("Neutral text.") + '<w:commentRangeEnd w:id="1"/>',
    run("Neutral text.") + '<w:r><w:footnoteReference w:id="1"/></w:r>',
    run("Neutral text.") + '<w:r><w:endnoteReference w:id="1"/></w:r>',
  ]) expect(await documentFingerprint(fixture("00000002", separatorSpacing, content))).not.toBe(baseline)
  for (const kind of ["footnote", "endnote"] as const) {
    const hash = (xml: string) => documentFingerprint(zipSync({ [`word/${kind}s.xml`]: strToU8(xml) }))
    const first = formatted(kind, "00000001")
    const second = formatted(kind, "00000002")
    expect(await hash(first)).not.toBe(await hash(second.replace('w:id="-1"', 'w:id="-2"')))
    for (const transform of [
      (xml: string) => xml.replaceAll(' w:type="separator"', '').replaceAll(' w:type="continuationSeparator"', ''),
      (xml: string) => xml.replace("<w:r>", '<w:commentRangeStart w:id="1"/><w:r>'),
      (xml: string) => xml.replace("<w:separator/>", "<w:separator/><w:t>Custom separator</w:t>"),
      (xml: string) => xml.replace("<w:pPr>", '<w:pPr w:unknown="preserved">'),
      (xml: string) => xml.replace("</w:pPr>", '<w:unknown w:val="preserved"/></w:pPr>'),
      (xml: string) => xml.replace("</w:pPr>", '<w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr>'),
      (xml: string) => xml.replace('w:after="0"', 'w:after="0" w:unknown="preserved"'),
    ]) expect(await hash(transform(first))).not.toBe(await hash(transform(second)))
  }
  expect(original).toEqual(copy)
})

test("note separator normalization preserves real paragraph and comment identity and all other data", async () => {
  const hash = (part: string, xml: string) => documentFingerprint(zipSync({ [part]: strToU8(xml) }))
  for (const kind of ["footnote", "endnote"] as const) {
    const part = `word/${kind}s.xml`
    const first = notePart(kind, "00000001")
    const second = notePart(kind, "00000002")
    expect(await hash(part, first)).not.toBe(await hash(part, notePart(kind, "00000002", "", "88888888")))
    // A real note, a comment anchor or custom separator content makes paragraph identity significant.
    for (const transform of [
      (xml: string) => xml.replace(' w:type="separator"', '').replace(' w:type="continuationSeparator"', '').replace('w:id="-1"', 'w:id="1"').replace('w:id="0"', 'w:id="2"'),
      (xml: string) => xml.replace("<w:r>", '<w:commentRangeStart w:id="1"/><w:r>'),
      (xml: string) => xml.replace("<w:separator/>", "<w:separator/><w:t>Custom separator</w:t>"),
      (xml: string) => xml.replace('w14:textId="77777777"', 'w14:textId="77777777" w:unknown="preserved"'),
    ]) expect(await hash(part, transform(first))).not.toBe(await hash(part, transform(second)))
    expect(await hash("customXml/item1.xml", first)).not.toBe(await hash("customXml/item1.xml", second))
    expect(await hash(part, first)).not.toBe(await hash(part, second.replace("<w:separator/>", "<w:separator/><w:t>Changed separator</w:t>")))
  }
  for (const part of ["word/document.xml", "word/comments.xml"]) {
    const root = part === "word/document.xml" ? "document" : "comments"
    const wrapper = root === "document" ? ["<w:body>", "</w:body>"] : ['<w:comment w:id="1" w:author="Synthetic reviewer">', "</w:comment>"]
    const xml = (id: string) => `<w:${root} xmlns:w="${word}" xmlns:w14="${word2010}">${wrapper[0]}<w:p w14:paraId="${id}">${run("Neutral text.")}</w:p>${wrapper[1]}</w:${root}>`
    expect(await hash(part, xml("00000001"))).not.toBe(await hash(part, xml("00000002")))
  }
})

const relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
const officeRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
const coreRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
const taskpaneRelationship = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
function rootRelationships(reopened = false): string {
  const records = reopened
    ? [["rId3", coreRelationship, "docProps/core.xml"], ["rId2", taskpaneRelationship, "NULL"], ["rId1", officeRelationship + "officeDocument", "word/document.xml"], ["rId4", officeRelationship + "extended-properties", "docProps/app.xml"]]
    : [["rId3", officeRelationship + "extended-properties", "docProps/app.xml"], ["rIdDocHaus1", taskpaneRelationship, "NULL"], ["rId2", coreRelationship, "docProps/core.xml"], ["rId1", officeRelationship + "officeDocument", "word/document.xml"]]
  return `<Relationships xmlns="${relationshipNamespace}">${records.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("")}</Relationships>`
}
const rootFingerprint = (relationships: string) => documentFingerprint(zipSync({ "_rels/.rels": strToU8(relationships) }))

test("package root relationship order and local IDs may rotate on native reopen without changing targets", async () => {
  const original = rootRelationships(), reopened = rootRelationships(true)
  const baseline = await rootFingerprint(original)
  expect(await rootFingerprint(original)).toBe(baseline)
  expect(await rootFingerprint(reopened)).toBe(baseline)
  for (const changed of [
    reopened.replace('Target="word/document.xml"', 'Target="word/other.xml"'),
    reopened.replace('Target="NULL"', 'Target="word/webextensions/taskpanes.xml"'),
    reopened.replace('Target="docProps/app.xml"', 'Target="docProps/app.xml" TargetMode="External"'),
  ]) expect(await rootFingerprint(changed)).not.toBe(baseline)
})

test("root relationship comparison preserves unknown payloads, signature relationships and duplicate IDs", async () => {
  for (const transform of [
    (xml: string) => xml.replace("<Relationship ", '<Relationship xmlns:c="urn:synthetic" c:metadata="retained" '),
    (xml: string) => xml.replace(taskpaneRelationship, "urn:synthetic:relationship"),
    (xml: string) => xml.replace(taskpaneRelationship, "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"),
    (xml: string) => xml.replace(/Id="[^"]+"/g, 'Id="duplicate"'),
    (xml: string) => xml.replace(relationshipNamespace, "urn:synthetic"),
  ]) expect(await rootFingerprint(transform(rootRelationships()))).not.toBe(await rootFingerprint(transform(rootRelationships(true))))
})

test("native reopen may recount Lines and remove only the standard magnification leaf", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  const packageBytes = (lines: number, zoom: string, pages = 1) => zipSync({
    "docProps/app.xml": strToU8(`<Properties xmlns="${extended}"><Lines>${lines}</Lines><Pages>${pages}</Pages></Properties>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${word}">${zoom}<w:defaultTabStop w:val="720"/></w:settings>`),
  })
  const baseline = await documentFingerprint(packageBytes(2, '<w:zoom w:percent="125"/>'))
  expect(await documentFingerprint(packageBytes(8, '<w:zoom w:percent="125"/>'))).toBe(baseline)
  expect(await documentFingerprint(packageBytes(2, ""))).toBe(baseline)
  expect(await documentFingerprint(packageBytes(8, '<w:zoom w:val="bestFit" w:percent="100"/>'))).toBe(baseline)
  expect(await documentFingerprint(packageBytes(8, "", 2))).not.toBe(baseline)
})

test("Lines and magnification normalization leave unknown, nested, duplicate and malformed values significant", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  for (const leaf of [
    (value: string) => `<Lines xmlns="urn:synthetic">${value}</Lines>`,
    (value: string) => `<Lines xmlns:c="urn:synthetic" c:metadata="retained">${value}</Lines>`,
    (value: string) => `<Other><Lines>${value}</Lines></Other>`,
    (value: string) => `<Lines><Other>${value}</Other></Lines>`,
    (value: string) => `<Lines>${value}</Lines><Lines>1</Lines>`,
    (value: string) => `<Lines>custom-${value}</Lines>`,
  ]) {
    const bytes = (value: string) => zipSync({ "docProps/app.xml": strToU8(`<Properties xmlns="${extended}">${leaf(value)}</Properties>`) })
    expect(await documentFingerprint(bytes("2"))).not.toBe(await documentFingerprint(bytes("8")))
  }
  for (const leaf of [
    (value: string) => `<c:zoom xmlns:c="urn:synthetic" w:percent="${value}"/>`,
    (value: string) => `<w:zoom xmlns:c="urn:synthetic" w:percent="${value}" c:metadata="retained"/>`,
    (value: string) => `<w:other><w:zoom w:percent="${value}"/></w:other>`,
    (value: string) => `<w:zoom w:percent="${value}"><w:other/></w:zoom>`,
    (value: string) => `<w:zoom w:percent="${value}"/><w:zoom w:percent="100"/>`,
    (value: string) => `<w:zoom w:percent="custom-${value}"/>`,
  ]) {
    const bytes = (value: string) => zipSync({ "word/settings.xml": strToU8(`<w:settings xmlns:w="${word}">${leaf(value)}</w:settings>`) })
    expect(await documentFingerprint(bytes("125"))).not.toBe(await documentFingerprint(bytes("150")))
  }
})

test("native reopen metadata churn cannot hide wording, formatting, numbering or tracked-revision changes", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  const body = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Notice: </w:t></w:r><w:del w:id="1" w:author="doc.haus"><w:r><w:delText>five</w:delText></w:r></w:del><w:ins w:id="2" w:author="doc.haus"><w:r><w:t>ten</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> working days.</w:t></w:r></w:p>`
  const packageBytes = (reopened: boolean, content = body) => zipSync({
    "_rels/.rels": strToU8(rootRelationships(reopened)),
    "docProps/app.xml": strToU8(`<Properties xmlns="${extended}"><Lines>${reopened ? 8 : 2}</Lines><Pages>1</Pages></Properties>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${word}">${reopened ? "" : '<w:zoom w:percent="125"/>'}<w:defaultTabStop w:val="720"/></w:settings>`),
    "word/document.xml": strToU8(document(content)),
  })
  const bytes = packageBytes(false), original = bytes.slice(), baseline = await documentFingerprint(bytes)
  expect(await documentFingerprint(packageBytes(true))).toBe(baseline)
  for (const changed of [body.replace(">ten<", ">twenty<"), body.replace("<w:b/>", "<w:i/>"), body.replace('w:numId w:val="1"', 'w:numId w:val="2"'), body.replace('w:id="2"', 'w:id="3"'), body.replace('w:author="doc.haus"', 'w:author="Other reviewer"')])
    expect(await documentFingerprint(packageBytes(true, changed))).not.toBe(baseline)
  expect(bytes).toEqual(original)
})

const stylePart = (sessionId: string, reopened = false) => `<w:styles xmlns:w="${word}" xmlns:c="urn:synthetic"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rsid w:val="${sessionId}"/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>${reopened ? "" : `<w:rsid w:val="${sessionId}"/>`}</w:style><w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>${reopened ? "" : `<w:rsid w:val="${sessionId}"/>`}</w:style></w:styles>`

test("native reopen may omit style session IDs while real style and document changes remain significant", async () => {
  const content = `<w:pPr><w:pStyle w:val="Normal"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:ins w:id="4" w:author="Synthetic reviewer">${run("Neutral text.")}</w:ins>`
  const bytes = (styles: string, body = content) => zipSync({
    "word/styles.xml": strToU8(styles), "word/document.xml": strToU8(document(body)),
    "word/numbering.xml": strToU8(`<w:numbering xmlns:w="${word}"><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`),
  })
  const saved = bytes(stylePart("00112233")), original = saved.slice(), baseline = await documentFingerprint(saved)
  expect(await documentFingerprint(saved)).toBe(baseline)
  // Word may keep some style IDs while omitting others after Save/close/reopen.
  const reopened = stylePart("00112233", true)
  expect(await documentFingerprint(bytes(reopened))).toBe(baseline)
  expect(await documentFingerprint(bytes(stylePart("00445566", true)))).toBe(baseline)
  for (const styles of [
    reopened.replace('w:val="22"', 'w:val="24"'),
    reopened.replace('w:after="160"', 'w:after="240"'),
    reopened.replace('w:type="numbering"', 'w:type="paragraph"'),
    reopened.replace('w:styleId="Normal"', 'w:styleId="Other"'),
    reopened.replace('<w:rPr>', '<w:rPr><w:b/>'),
    reopened.replace('<w:pPr>', '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr>'),
    reopened.replace('<w:uiPriority w:val="99"/>', '<w:uiPriority w:val="99"/><c:metadata>retained</c:metadata>'),
  ]) expect(await documentFingerprint(bytes(styles))).not.toBe(baseline)
  for (const body of [
    content.replace("Neutral", "Changed"), content.replace('w:numId w:val="1"', 'w:numId w:val="2"'),
    content.replace('w:id="4"', 'w:id="5"'), content.replace("Synthetic reviewer", "Another reviewer"),
  ]) expect(await documentFingerprint(bytes(reopened, body))).not.toBe(baseline)
  expect(saved).toEqual(original)
})

test("style session ID normalization preserves malformed, nested, duplicate, foreign and extended leaves", async () => {
  const hash = (part: string, styles: string) => documentFingerprint(zipSync({ [part]: strToU8(styles) }))
  for (const transform of [
    (xml: string) => xml.replaceAll('<w:rsid ', '<w:rsid c:metadata="retained" '),
    (xml: string) => xml.replaceAll('<w:rsid ', '<c:rsid '),
    (xml: string) => xml.replaceAll(/<w:rsid ([^>]+)\/>/g, '<w:rsid $1><c:data/></w:rsid>'),
    (xml: string) => xml.replaceAll(/<w:rsid ([^>]+)\/>/g, '<w:rsid $1/><w:rsid w:val="77889900"/>'),
    (xml: string) => xml.replaceAll(/<w:rsid ([^>]+)\/>/g, '<w:other><w:rsid $1/></w:other>'),
    (xml: string) => xml.replaceAll('w:val="00', 'w:val="invalid-00'),
    (xml: string) => xml.replaceAll('<w:styles ', '<w:other ' ).replaceAll('</w:styles>', '</w:other>'),
    (xml: string) => xml.replaceAll('<w:style ', '<c:style ').replaceAll('</w:style>', '</c:style>'),
  ]) expect(await hash("word/styles.xml", transform(stylePart("00112233")))).not.toBe(await hash("word/styles.xml", transform(stylePart("00445566"))))
  for (const part of ["customXml/item1.xml", "word/custom.xml"])
    expect(await hash(part, stylePart("00112233"))).not.toBe(await hash(part, stylePart("00445566")))
})

test("DOCX fingerprint ignores ZIP timestamps, compression and entry order but detects changed content", async () => {
  const a = strToU8("<document>one</document>")
  const b = strToU8("<styles/>")
  const first = zipSync({ "word/document.xml": a, "word/styles.xml": b }, { mtime: new Date(2020, 0, 1), level: 0 })
  const second = zipSync({ "word/styles.xml": b, "word/document.xml": a }, { mtime: new Date(2025, 0, 1), level: 6 })
  const changed = zipSync({ "word/document.xml": strToU8("<document>two</document>"), "word/styles.xml": b })
  expect(await documentFingerprint(first)).toBe(await documentFingerprint(second))
  expect(await documentFingerprint(changed)).not.toBe(await documentFingerprint(first))
})

test("fingerprint excludes only an acknowledged link base while preserving other settings and markup", async () => {
  const ns = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
  const fixture = (base: string, token = "capability", extra = "keep", marker = "one") => {
    const link = JSON.stringify({ version: 1, origin: "https://example.invalid", apiOrigin: "https://localhost:4612", matter: "matter", document: "draft", baseVersion: base, token, extra }).replaceAll("&", "&amp;").replaceAll('"', "&quot;")
    return zipSync({ "word/webextensions/webextension1.xml": strToU8(`<we:webextension xmlns:we="${ns}" marker="${marker}"><we:reference id="${WORD_ADDIN_ID}"/><we:properties><we:property name="dochaus.link" value="${link}"/></we:properties></we:webextension>`) })
  }
  const before = await documentFingerprint(fixture("v1"))
  expect(await documentFingerprint(fixture("v2"))).toBe(before)
  expect(await documentFingerprint(fixture("v2", "different"))).not.toBe(before)
  expect(await documentFingerprint(fixture("v2", "capability", "changed"))).not.toBe(before)
  expect(await documentFingerprint(fixture("v2", "capability", "keep", "two"))).not.toBe(before)
})

test("fingerprint refuses excessive parts and inflated sizes before unpacking", async () => {
  const tooMany = Object.fromEntries(Array.from({ length: 1001 }, (_, index) => [`part${index}.xml`, strToU8("<p/>")]))
  await expect(documentFingerprint(zipSync(tooMany))).rejects.toThrow("package size limit")
  const bytes = zipSync({ "word/document.xml": strToU8("<document/>") })
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let at = 0; at < bytes.length - 46; at++) {
    if (view.getUint32(at, true) === 0x02014b50) {
      view.setUint32(at + 24, 257 * 1024 * 1024, true)
      break
    }
  }
  await expect(documentFingerprint(bytes)).rejects.toThrow("package size limit")
})

test("Word session metadata may rotate without hiding text, formatting, comments or tracked edits", async () => {
  const fixture = (
    options: {
      session?: string
      text?: string
      bold?: boolean
      inserted?: string
      deleted?: string
      comment?: string
      changeId?: string
      author?: string
      foreign?: string
      unknown?: string
    } = {},
  ) => {
    const session = options.session ?? "00AAAA01"
    return zipSync({
      "word/document.xml": strToU8(
        `<m:document xmlns:m="${word}" xmlns:other="urn:synthetic:other"><m:body><m:p m:rsidR="${session}" m:rsidRDefault="${session}" m:rsidP="${session}" other:rsidRDefault="${options.foreign ?? "preserved"}" m:rsidUnknown="${options.unknown ?? "preserved"}"><m:pPr><m:jc m:val="left"/></m:pPr><m:r m:rsidRPr="${session}"><m:rPr>${options.bold ? "<m:b/>" : ""}</m:rPr><m:t>${options.text ?? "Neutral text."}</m:t></m:r><m:ins m:id="${options.changeId ?? "1"}" m:author="${options.author ?? "Synthetic reviewer"}" m:date="2026-10-26T10:00:00Z"><m:r><m:t>${options.inserted ?? "Inserted text."}</m:t></m:r></m:ins><m:del m:id="2" m:author="Synthetic reviewer" m:date="2026-10-26T10:00:00Z" m:rsidDel="${session}"><m:r><m:delText>${options.deleted ?? "Deleted text."}</m:delText></m:r></m:del></m:p><m:tbl><m:tr m:rsidTr="${session}"><m:tc><m:p/></m:tc></m:tr></m:tbl><m:sectPr m:rsidSect="${session}"/></m:body></m:document>`,
      ),
      "word/settings.xml": strToU8(
        `<m:settings xmlns:m="${word}"><m:zoom m:percent="100"/><m:rsids><m:rsidRoot m:val="00BBBB02"/><m:rsid m:val="${session}"/>${session === "00AAAA01" ? "" : '<m:rsid m:val="00CCCC03"/>'}</m:rsids></m:settings>`,
      ),
      "word/comments.xml": strToU8(
        `<m:comments xmlns:m="${word}"><m:comment m:id="0" m:author="Synthetic reviewer"><m:p><m:r><m:t>${options.comment ?? "A neutral comment."}</m:t></m:r></m:p></m:comment></m:comments>`,
      ),
    })
  }
  const baseline = fixture()
  const original = baseline.slice()
  const expected = await documentFingerprint(baseline)
  expect(await documentFingerprint(fixture({ session: "00DDDD04" }))).toBe(expected)
  for (const edit of [
    { text: "Changed text." },
    { bold: true },
    { inserted: "Different insertion." },
    { deleted: "Different deletion." },
    { comment: "A changed comment." },
    { changeId: "3" },
    { author: "Another synthetic reviewer" },
    { foreign: "different" },
    { unknown: "different" },
  ])
    expect(await documentFingerprint(fixture({ ...edit, session: "00DDDD04" }))).not.toBe(expected)
  expect(baseline).toEqual(original)
})

test("metadata normalization preserves opaque XML and is restricted to a direct Word settings list", async () => {
  const cases: [string, (value: string) => string][] = [
    ["word/settings.xml", (value) => `<w:settings xmlns:w="${word}" xmlns:c="urn:synthetic"><c:rsids c:value="${value}"/></w:settings>`],
    ["word/settings.xml", (value) => `<w:settings xmlns:w="${word}"><w:compat><w:rsids w:val="${value}"/></w:compat></w:settings>`],
    ["word/document.xml", (value) => `<w:document xmlns:w="${word}"><w:rsids w:val="${value}"/><w:body/></w:document>`],
    ["customXml/item1.xml", (value) => `<w:document xmlns:w="${word}" w:rsidR="${value}"/>`],
    ["customXml/item1.xml", (value) => `<c:data xmlns:c="urn:synthetic" xmlns:w="${word}" w:rsidR="${value}"/>`],
    ["word/custom.xml", (value) => `<c:data xmlns:c="urn:synthetic" xmlns:w="${word}"><w:p w:rsidR="${value}"/></c:data>`],
    ["word/document.xml", (value) => `<w:document xmlns:w="${word}" xmlns:c="urn:synthetic"><c:data w:rsidR="${value}"/></w:document>`],
  ]
  for (const [name, xml] of cases) {
    const hash = (value: string) => documentFingerprint(zipSync({ [name]: strToU8(xml(value)) }))
    expect(await hash("one")).not.toBe(await hash("two"))
  }
})

test("tracked formatting payload stays significant while strict Word session attributes do not", async () => {
  const strictWord = "http://purl.oclc.org/ooxml/wordprocessingml/main"
  const fixture = (session: string, color = "000000") =>
    zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="${strictWord}"><w:body><w:p w:rsidR="${session}"><w:r><w:rPr><w:rPrChange w:id="1" w:author="Synthetic reviewer" w:date="2026-10-26T10:00:00Z"><w:rPr><w:color w:val="${color}"/></w:rPr></w:rPrChange></w:rPr><w:t>Neutral text.</w:t></w:r></w:p></w:body></w:document>`,
      ),
    })
  const baseline = await documentFingerprint(fixture("00AAAA01"))
  expect(await documentFingerprint(fixture("00BBBB02"))).toBe(baseline)
  expect(await documentFingerprint(fixture("00BBBB02", "FFFFFF"))).not.toBe(baseline)
})

test("editing-duration-only changes are ignored while other properties and document text remain significant", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  const fixture = (time: number, manager = "Synthetic manager", text = "Neutral text.", attribute = "preserved") =>
    zipSync({
      "docProps/app.xml": strToU8(
        `<ep:Properties xmlns:ep="${extended}" xmlns:c="urn:synthetic"><ep:TotalTime c:metadata="${attribute}">${time}</ep:TotalTime><ep:Manager>${manager}</ep:Manager><ep:Pages>1</ep:Pages></ep:Properties>`,
      ),
      "word/document.xml": strToU8(
        `<w:document xmlns:w="${word}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
      ),
    })
  const bytes = fixture(3)
  const original = bytes.slice()
  const baseline = await documentFingerprint(bytes)
  expect(await documentFingerprint(fixture(4))).toBe(baseline)
  expect(await documentFingerprint(fixture(4, "Different manager"))).not.toBe(baseline)
  expect(await documentFingerprint(fixture(4, "Synthetic manager", "Changed text."))).not.toBe(baseline)
  expect(await documentFingerprint(fixture(4, "Synthetic manager", "Neutral text.", "changed"))).not.toBe(baseline)
  expect(bytes).toEqual(original)
})

test("TotalTime normalization does not hide custom XML, foreign namespaces, nested or nonnumeric data", async () => {
  const extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  const cases: [string, (value: string) => string][] = [
    ["customXml/item1.xml", (value) => `<Properties xmlns="${extended}"><TotalTime>${value}</TotalTime></Properties>`],
    ["docProps/custom.xml", (value) => `<Properties xmlns="${extended}"><TotalTime>${value}</TotalTime></Properties>`],
    ["docProps/app.xml", (value) => `<Properties xmlns="urn:synthetic"><TotalTime>${value}</TotalTime></Properties>`],
    ["docProps/app.xml", (value) => `<Properties xmlns="${extended}" xmlns:c="urn:synthetic"><c:TotalTime>${value}</c:TotalTime></Properties>`],
    ["docProps/app.xml", (value) => `<Properties xmlns="${extended}"><Other><TotalTime>${value}</TotalTime></Other></Properties>`],
    ["docProps/app.xml", (value) => `<Properties xmlns="${extended}"><TotalTime><Other>${value}</Other></TotalTime></Properties>`],
  ]
  for (const [name, xml] of cases) {
    const hash = (value: string) => documentFingerprint(zipSync({ [name]: strToU8(xml(value)) }))
    expect(await hash("3")).not.toBe(await hash("4"))
  }
  const invalid = (value: string) =>
    documentFingerprint(zipSync({ "docProps/app.xml": strToU8(`<Properties xmlns="${extended}"><TotalTime>${value}</TotalTime></Properties>`) }))
  expect(await invalid("custom-one")).not.toBe(await invalid("custom-two"))
})

test("proofing completion cache may appear or change without hiding edits or proofing preferences", async () => {
  const fixture = (proofState = "", text = "Neutral text.", language = "en-GB", proofError = "spellStart") =>
    zipSync({
      "word/settings.xml": strToU8(`<m:settings xmlns:m="${word}"><m:zoom m:percent="100"/>${proofState}</m:settings>`),
      "word/document.xml": strToU8(
        `<m:document xmlns:m="${word}"><m:body><m:p><m:proofErr m:type="${proofError}"/><m:r><m:rPr><m:lang m:val="${language}"/></m:rPr><m:t>${text}</m:t></m:r></m:p></m:body></m:document>`,
      ),
    })
  const bytes = fixture()
  const original = bytes.slice()
  const baseline = await documentFingerprint(bytes)
  for (const cache of [
    '<m:proofState m:spelling="clean" m:grammar="clean"/>',
    '<m:proofState m:spelling="dirty" m:grammar="dirty"/>',
    '<m:proofState m:grammar="clean"/>',
    "<m:proofState/>",
    '<m:proofState m:spelling="clean"> \n </m:proofState>',
  ])
    expect(await documentFingerprint(fixture(cache))).toBe(baseline)
  const cache = '<m:proofState m:spelling="clean" m:grammar="clean"/>'
  expect(await documentFingerprint(fixture(cache, "Changed text."))).not.toBe(baseline)
  expect(await documentFingerprint(fixture(cache, "Neutral text.", "en-US"))).not.toBe(baseline)
  expect(await documentFingerprint(fixture(cache, "Neutral text.", "en-GB", "futureProofingType"))).not.toBe(baseline)
  expect(bytes).toEqual(original)
})

test("proofState normalization preserves unknown attributes, nested data and opaque XML", async () => {
  const settings = (child: string) => `<w:settings xmlns:w="${word}" xmlns:c="urn:synthetic">${child}</w:settings>`
  const cases: [string, (value: string) => string][] = [
    ["word/settings.xml", (value) => settings(`<w:proofState w:spelling="clean" w:unknown="${value}"/>`)],
    ["word/settings.xml", (value) => settings(`<w:proofState w:spelling="clean" c:value="${value}"/>`)],
    ["word/settings.xml", (value) => settings(`<w:proofState w:spelling="clean"><c:data>${value}</c:data></w:proofState>`)],
    ["word/settings.xml", (value) => settings(`<w:proofState w:spelling="clean">${value}</w:proofState>`)],
    ["word/settings.xml", (value) => settings(`<c:proofState w:spelling="${value}"/>`)],
    ["word/settings.xml", (value) => settings(`<w:compat><w:proofState w:spelling="${value}"/></w:compat>`)],
    ["word/document.xml", (value) => `<w:document xmlns:w="${word}"><w:proofState w:spelling="${value}"/></w:document>`],
    ["customXml/item1.xml", (value) => settings(`<w:proofState w:spelling="${value}"/>`)],
    ["word/settings.xml", (value) => `<c:settings xmlns:c="urn:synthetic" xmlns:w="${word}"><w:proofState w:spelling="${value}"/></c:settings>`],
  ]
  for (const [name, xml] of cases) {
    const hash = (value: string) => documentFingerprint(zipSync({ [name]: strToU8(xml(value)) }))
    expect(await hash("clean")).not.toBe(await hash("dirty"))
  }
  const invalid = (value: string) =>
    documentFingerprint(zipSync({ "word/settings.xml": strToU8(settings(`<w:proofState w:spelling="${value}"/>`)) }))
  expect(await invalid("future-one")).not.toBe(await invalid("future-two"))
})
