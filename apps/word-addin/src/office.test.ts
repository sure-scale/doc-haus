import { expect, test } from "bun:test"
import { zipSync, strToU8 } from "fflate"
import { documentFingerprint } from "./fingerprint"
import { exportFile, type OfficeFile } from "./office"

test("DOCX fingerprint ignores ZIP timestamps, compression and entry order but detects changed content", async () => {
  const a = strToU8("<document>one</document>")
  const b = strToU8("<styles/>")
  const first = zipSync({ "word/document.xml": a, "word/styles.xml": b }, { mtime: new Date(2020, 0, 1), level: 0 })
  const second = zipSync({ "word/styles.xml": b, "word/document.xml": a }, { mtime: new Date(2025, 0, 1), level: 6 })
  const changed = zipSync({ "word/document.xml": strToU8("<document>two</document>"), "word/styles.xml": b })
  expect(await documentFingerprint(first)).toBe(await documentFingerprint(second))
  expect(await documentFingerprint(changed)).not.toBe(await documentFingerprint(first))
})

test("sliced export assembles the file and always closes its handle on success or failure", async () => {
  for (const fails of [false, true]) {
    let closed = 0
    const file: OfficeFile = {
      size: 3,
      sliceCount: 2,
      getSliceAsync: (index, callback) =>
        callback({ status: fails ? "failed" : "succeeded", value: { data: index === 0 ? [1, 2] : [3] } }),
      closeAsync: (callback) => {
        closed++
        callback({ status: "succeeded", value: undefined })
      },
    }
    if (fails) await expect(exportFile(file)).rejects.toThrow("could not read")
    else expect(await exportFile(file)).toEqual(new Uint8Array([1, 2, 3]))
    expect(closed).toBe(1)
  }
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
  const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
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
  const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const cases: [string, (value: string) => string][] = [
    [
      "word/settings.xml",
      (value) => `<w:settings xmlns:w="${word}" xmlns:c="urn:synthetic"><c:rsids c:value="${value}"/></w:settings>`,
    ],
    [
      "word/settings.xml",
      (value) => `<w:settings xmlns:w="${word}"><w:compat><w:rsids w:val="${value}"/></w:compat></w:settings>`,
    ],
    ["word/document.xml", (value) => `<w:document xmlns:w="${word}"><w:rsids w:val="${value}"/><w:body/></w:document>`],
    ["customXml/item1.xml", (value) => `<w:document xmlns:w="${word}" w:rsidR="${value}"/>`],
    ["customXml/item1.xml", (value) => `<c:data xmlns:c="urn:synthetic" xmlns:w="${word}" w:rsidR="${value}"/>`],
    [
      "word/custom.xml",
      (value) => `<c:data xmlns:c="urn:synthetic" xmlns:w="${word}"><w:p w:rsidR="${value}"/></c:data>`,
    ],
    [
      "word/document.xml",
      (value) => `<w:document xmlns:w="${word}" xmlns:c="urn:synthetic"><c:data w:rsidR="${value}"/></w:document>`,
    ],
  ]
  for (const [name, xml] of cases) {
    const fingerprint = (value: string) => documentFingerprint(zipSync({ [name]: strToU8(xml(value)) }))
    expect(await fingerprint("one")).not.toBe(await fingerprint("two"))
  }
})

test("tracked formatting payload stays significant while strict Word session attributes do not", async () => {
  const word = "http://purl.oclc.org/ooxml/wordprocessingml/main"
  const fixture = (session: string, color = "000000") =>
    zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="${word}"><w:body><w:p w:rsidR="${session}"><w:r><w:rPr><w:rPrChange w:id="1" w:author="Synthetic reviewer" w:date="2026-10-26T10:00:00Z"><w:rPr><w:color w:val="${color}"/></w:rPr></w:rPrChange></w:rPr><w:t>Neutral text.</w:t></w:r></w:p></w:body></w:document>`,
      ),
    })
  const baseline = await documentFingerprint(fixture("00AAAA01"))
  expect(await documentFingerprint(fixture("00BBBB02"))).toBe(baseline)
  expect(await documentFingerprint(fixture("00BBBB02", "FFFFFF"))).not.toBe(baseline)
})
