import { describe, expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { prepareNativeRow, queueNativeRow, type NativeRowContext } from "./word-chat-native-row"
import {
  assertNativeWordChatRowDeletionTarget,
  verifyImportedWordChatDeletedRow,
  verifyNativeWordChatDeletedRow,
  verifyReviewedWordChatRowDeletionOracle,
} from "./word-chat-native-row-delete"
import { nativeWordChatRowRefitWidths, verifyNativeWordChatInsertedRow } from "./word-chat-native-row-verification"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"

describe("native row batching", () => {
  function fixture() {
    const batches: string[][] = [], queued: string[] = []
    const table = (prefix: string) => ({
      rowCount: 1, isUniform: true, values: [[`${prefix}1`, `${prefix}2`]], load() {},
      getCell(row: number, column: number) {
        return {
          body: {
            paragraphs: { items: [{ uniqueLocalId: `${prefix}${column}`, text: `${prefix}${column + 1}`, tableNestingLevel: 1 }], load() {} },
            getRange(location: string) {
              return {
                getReviewedText() { return { value: `${prefix}${column + 1}` } },
                font: new Proxy({ bold: true, italic: true, underline: "Single", load() {} }, {
                  set(target, key, value) { queued.push(`cell:${row}:${column}:${location}:${String(key)}:${value}`); (target as Record<PropertyKey, unknown>)[key] = value; return true },
                }),
              }
            },
            insertText(text: string, location: string) {
              queued.push(`text:${row}:${column}:${text}:${location}`)
              const font = new Proxy({ bold: true, italic: true, underline: "Single", load() {} }, {
                set(target, key, value) { queued.push(`font:${String(key)}:${value}`); (target as Record<PropertyKey, unknown>)[key] = value; return true },
              })
              return { font }
            },
          },
        }
      },
      addRows(location: string, count: number, values: string[][]) { queued.push(`row:${location}:${count}:${JSON.stringify(values)}`) },
    })
    const document = {
      mode: "Off", formatting: false, load() {}, body: { tables: { items: [table("a"), table("b")], load() {} } },
      get changeTrackingMode() { return this.mode }, set changeTrackingMode(value: string) { queued.push(`tracking:${value}`); this.mode = value },
      get trackFormatting() { return this.formatting }, set trackFormatting(value: boolean) { queued.push(`formatting:${value}`); this.formatting = value },
    }
    const context: NativeRowContext = { document, async sync() { batches.push(queued.splice(0)) } }
    const binding = { tableIndex: 1, rows: [[{ id: "b0", text: "b1" }, { id: "b1", text: "b2" }]] }
    return { context, document, batches, queued, binding }
  }

  test("binds the reviewed cells among multiple tables and queues one mutation flush", async () => {
    const f = fixture(), prepared = await prepareNativeRow(f.context, f.binding)
    expect(f.batches).toEqual([[], [], []])
    const witness = queueNativeRow(f.context, prepared, ["First", "Second"])
    expect(witness.ranges).toHaveLength(2)
    expect(f.document.changeTrackingMode).toBe("Off")
    expect(f.document.trackFormatting).toBe(false)
    await f.context.sync()
    expect(f.batches[3]).toEqual([
      "tracking:TrackAll", 'row:End:1:[["",""]]',
      "cell:1:0:Whole:bold:false", "cell:1:0:Whole:italic:false", "cell:1:0:Whole:underline:None",
      "text:1:0:First:End", "font:bold:false", "font:italic:false", "font:underline:None",
      "cell:1:1:Whole:bold:false", "cell:1:1:Whole:italic:false", "cell:1:1:Whole:underline:None",
      "text:1:1:Second:End", "font:bold:false", "font:italic:false", "font:underline:None", "tracking:Off", "formatting:false",
    ])
  })

  test("rejects a stale ordinal or changed cell without queuing a mutation", async () => {
    for (const change of ["ordinal", "text", "shape"] as const) {
      const f = fixture()
      if (change === "ordinal") f.binding.tableIndex = 0
      if (change === "text") f.binding.rows[0]![0]!.text = "changed"
      if (change === "shape") f.context.document.body.tables.items[1]!.isUniform = false
      await expect(prepareNativeRow(f.context, f.binding)).rejects.toThrow("reviewed Word table")
      expect(f.queued).toEqual([])
    }
  })

  test("a formatted row inserts each reviewed run with explicit flags in the same tracked flush", async () => {
    const f = fixture(), prepared = await prepareNativeRow(f.context, f.binding)
    const run = (text: string, bold = false, italic = false, underline = false) => ({ text, bold, italic, underline })
    const witness = queueNativeRow(f.context, prepared, ["Fee: 100", "Note"], [[run("Fee:", true), run(" 100")], [run("Note", false, true, true)]])
    expect(witness.ranges).toHaveLength(3)
    await f.context.sync()
    expect(f.batches[3]).toEqual([
      "tracking:TrackAll", 'row:End:1:[["",""]]',
      "cell:1:0:Whole:bold:false", "cell:1:0:Whole:italic:false", "cell:1:0:Whole:underline:None",
      "text:1:0:Fee::End", "font:bold:true", "font:italic:false", "font:underline:None",
      "text:1:0: 100:End", "font:bold:false", "font:italic:false", "font:underline:None",
      "cell:1:1:Whole:bold:false", "cell:1:1:Whole:italic:false", "cell:1:1:Whole:underline:None",
      "text:1:1:Note:End", "font:bold:false", "font:italic:true", "font:underline:Single",
      "tracking:Off", "formatting:false",
    ])
    const g = fixture(), ready = await prepareNativeRow(g.context, g.binding)
    queueNativeRow(g.context, ready, ["Fee:", ""], [[run("Fee:", true)], []])
    await g.context.sync()
    expect(g.batches[3]!.slice(9)).toEqual([
      "cell:1:1:Whole:bold:false", "cell:1:1:Whole:italic:false", "cell:1:1:Whole:underline:None",
      "text:1:1::End", "font:bold:false", "font:italic:false", "font:underline:None", "tracking:Off", "formatting:false",
    ])
  })

  test("formatting that does not match its cells refuses before queuing any write", async () => {
    const run = (text: string, bold = false) => ({ text, bold, italic: false, underline: false })
    for (const runs of [
      [[run("Fee:", true), run(" 10")], []],
      [[run("Fee:", true), run(" 100")]],
      [[{ ...run("Fee:"), bold: "yes" as unknown as boolean }, run(" 100")], []],
      [[run("", true), run("Fee: 100")], []],
    ]) {
      const f = fixture(), prepared = await prepareNativeRow(f.context, f.binding)
      expect(() => queueNativeRow(f.context, prepared, ["Fee: 100", ""], runs)).toThrow("formatting does not match")
      expect(f.queued).toEqual([])
    }
  })
})

describe("native row deletion verification", () => {
  const attrs = `xmlns:w="${W}"`
  const row = (a: string, b: string) => `<w:tr><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${a}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>${b}</w:t></w:r></w:p></w:tc></w:tr>`
  const beforeRow = row("old", "value")
  const mark = (id: number) => `w:id="${id}" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"`
  const deletedRow = `<w:tr><w:trPr><w:del ${mark(1)}/></w:trPr><w:tc><w:p><w:pPr><w:rPr><w:del ${mark(2)}/></w:rPr></w:pPr><w:del ${mark(3)}><w:r><w:rPr><w:b/></w:rPr><w:delText>old</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p><w:pPr><w:rPr><w:del ${mark(4)}/></w:rPr></w:pPr><w:del ${mark(5)}><w:r><w:delText>value</w:delText></w:r></w:del></w:p></w:tc></w:tr>`
  const body = (r: string) => `<w:document ${attrs}><w:body><w:tbl>${row("head", "caption")}${r}</w:tbl><w:sectPr/></w:body></w:document>`
  const flat = (r: string) => `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${body(r)}</pkg:xmlData></pkg:part></pkg:package>`
  const before = [{ id: "a", text: "head" }, { id: "b", text: "caption" }, { id: "c", text: "old" }, { id: "d", text: "value" }].map(x => ({ ...x, tableNestingLevel: 1 }))
  const after = before.map((x, i) => ({ ...x, text: i >= 2 ? "" : x.text }))
  const proposal = { kind: "table_row_delete" as const, id: "p", documentId: "d", explanation: "", openItems: [], createdAt: "", table: { tableIndex: 0, rows: [before.slice(0, 2), before.slice(2)] }, rowIndex: 1, cells: ["old", "value"] }

  test("proves a formatted native row deletion and rejects changes to its preimage, carriers and remainder", () => {
    assertNativeWordChatRowDeletionTarget(flat(beforeRow), before, proposal)
    expect(verifyNativeWordChatDeletedRow(flat(beforeRow), flat(deletedRow), before, after, proposal)).toEqual({ authors: ["doc.haus"] })
    expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow), flat(deletedRow.replace("<w:b/>", "<w:i/>")), before, after, proposal)).toThrow()
    expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow), flat(deletedRow.replace(mark(5), mark(4))), before, after, proposal)).toThrow()
    expect(() => assertNativeWordChatRowDeletionTarget(flat(beforeRow.replace("<w:b/>", "<w:vanish/>")), before, proposal)).toThrow()
    expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow), flat(deletedRow).replace("head", "changed"), before, after, proposal)).toThrow()
    expect(() => assertNativeWordChatRowDeletionTarget(flat(beforeRow).replace("<w:tbl>", '<w:bookmarkStart w:id="8" w:name="crossing"/><w:tbl>').replace("</w:tbl>", '</w:tbl><w:bookmarkEnd w:id="8"/>'), before, proposal)).toThrow()
    expect(() => assertNativeWordChatRowDeletionTarget(flat(beforeRow.replace("<w:tr>", `<w:tr><w:trPr><w:ins ${mark(9)}/></w:trPr>`)), before, proposal)).toThrow()
    expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow), flat(deletedRow.replace('w:author="doc.haus"', 'w:author="Other"')), before, after, proposal)).toThrow()
  })

  test("the SDK oracle is source-bound and preserves every unrelated package part", async () => {
    const source = zipSync({ "word/document.xml": strToU8(body(beforeRow)), "word/settings.xml": strToU8(`<w:settings ${attrs}/>`), other: strToU8("keep") })
    const authored = (other = "keep") => zipSync({ "word/document.xml": strToU8(body(deletedRow)), "word/settings.xml": strToU8(`<w:settings ${attrs}><w:trackRevisions/></w:settings>`), other: strToU8(other) })
    const hash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", source))).map(x => x.toString(16).padStart(2, "0")).join("")
    const oracle = { kind: "table_row_delete" as const, proposalId: proposal.id, sourceSha256: hash, table: proposal.table, rowIndex: 1, cells: proposal.cells, author: "doc.haus" as const, writerRequest: { source_sha256: hash, target_handle: "bound", author: "doc.haus" }, authoredBase64: Buffer.from(authored()).toString("base64"), revisions: [{ id: "r", family: "row_delete" }] }
    await verifyReviewedWordChatRowDeletionOracle(source, oracle, proposal)
    await expect(verifyReviewedWordChatRowDeletionOracle(source, { ...oracle, authoredBase64: Buffer.from(authored("altered")).toString("base64") }, proposal)).rejects.toThrow()
    await expect(verifyReviewedWordChatRowDeletionOracle(source, { ...oracle, sourceSha256: "bad" }, proposal)).rejects.toThrow()
  })

  test("a new Word people relationship must be unique, internal and target the new people part", () => {
    const relns = "http://schemas.openxmlformats.org/package/2006/relationships"
    const people = "http://schemas.microsoft.com/office/2011/relationships/people"
    const pns = "http://schemas.microsoft.com/office/word/2012/wordml"
    const part = (name: string, xml: string) => `<pkg:part pkg:name="${name}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`
    const relationship = (id: string, type: string, target: string, extra = "") => `<Relationship Id="${id}" Type="${type}" Target="${target}"${extra}/>`
    const rels = (children: string) => part("/word/_rels/document.xml.rels", `<Relationships xmlns="${relns}">${children}</Relationships>`)
    const person = part("/word/people.xml", `<w15:people xmlns:w15="${pns}"><w15:person w15:author="doc.haus"><w15:presenceInfo w15:providerId="x" w15:userId="y"/></w15:person></w15:people>`)
    const theme = (id: string) => relationship(id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme.xml")
    const wrap = (xml: string, extra: string) => xml.replace("</pkg:package>", `${extra}</pkg:package>`)
    const original = wrap(flat(beforeRow), rels(theme("r1")))
    const observed = (link: string) => wrap(flat(deletedRow), rels(theme("r2") + link) + person)
    const valid = relationship("r1", people, "people.xml")
    expect(verifyNativeWordChatDeletedRow(original, observed(valid), before, after, proposal).authors).toEqual(["doc.haus"])
    for (const link of [
      relationship("r1", people, "other.xml"),
      relationship("r1", people, "people.xml", ' TargetMode="External"'),
      valid + relationship("r3", people, "people.xml"),
      relationship("r2", people, "people.xml"),
    ]) expect(() => verifyNativeWordChatDeletedRow(original, observed(link), before, after, proposal)).toThrow()
    expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow), flat(deletedRow.replace(mark(5), mark(4).replace('w:id="4"', 'w:id="04"'))), before, after, proposal)).toThrow()
  })

  test("parent import maps runtime IDs only inside its complete table and keeps physical XML strict", () => {
    const outside = (xml: string) => xml.replace("<w:body>", "<w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>").replace("<w:sectPr/>", "<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/>")
    const first = { id: "outside-first", text: "Before", tableNestingLevel: 0 }, last = { id: "outside-last", text: "After", tableNestingLevel: 0 }
    const original = [first, ...before, last], observed = [first, ...after.map(p => ({ ...p, id: `imported-${p.id}` })), last]
    expect(verifyImportedWordChatDeletedRow(outside(flat(beforeRow)), outside(flat(deletedRow)), original, observed, proposal)).toEqual({ authors: ["doc.haus"] })
    expect(() => verifyNativeWordChatDeletedRow(outside(flat(beforeRow)), outside(flat(deletedRow)), original, observed, proposal)).toThrow()
    expect(() => verifyImportedWordChatDeletedRow(outside(flat(beforeRow)), outside(flat(deletedRow)), original, observed.map((p, i) => i === 0 ? { ...p, id: "changed-outside" } : p), proposal)).toThrow()
    expect(() => verifyImportedWordChatDeletedRow(outside(flat(beforeRow)), outside(flat(deletedRow)), original, observed.map((p, i) => i === 1 ? { ...p, id: "outside-first" } : p), proposal)).toThrow()
    const identified = (xml: string, id: string) => xml.replace("<w:tc><w:p>", `<w:tc><w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${id}">`)
    expect(() => verifyImportedWordChatDeletedRow(outside(identified(flat(beforeRow), "A0000001")), outside(identified(flat(deletedRow), "A0000002")), original, observed, proposal)).toThrow()
    expect(() => verifyImportedWordChatDeletedRow(outside(flat(beforeRow)), outside(flat(deletedRow).replaceAll('w:author="doc.haus"', 'w:author="Other"')), original, observed, proposal)).toThrow()
  })

  test("parent import binds persisted row and paragraph IDs bijectively before strict preservation proof", () => {
    const w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    const identified = (xml: string, start: number) => {
      let next = start
      return xml.replace("<w:document ", `<w:document xmlns:w14="${w14}" `)
        .replace(/<w:(tr|p)>/g, (_, name: string) => `<w:${name} w14:paraId="${(++next).toString(16).toUpperCase()}"${name === "p" ? ' w14:textId="33333333"' : ""}>`)
        .replace("<w:tbl>", '<w:tbl><w:tblPr><w:tblW w:w="3600"/></w:tblPr>')
        .replace("<w:r><w:rPr><w:b/></w:rPr><w:t>head</w:t></w:r>", '<w:ins w:id="9" w:author="Reviewer" w:date="2025-01-01T00:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:t>head</w:t></w:r></w:ins>')
    }
    const outside = (xml: string) => xml.replace("<w:body>", '<w:body><w:p w14:paraId="30000001"><w:r><w:t>Before</w:t></w:r></w:p>').replace("<w:sectPr/>", '<w:p w14:paraId="30000002"><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/>')
    const source = outside(identified(flat(beforeRow), 0x10000000)), current = outside(identified(flat(deletedRow), 0x20000000))
    const first = { id: "outside-first", text: "Before", tableNestingLevel: 0 }, last = { id: "outside-last", text: "After", tableNestingLevel: 0 }
    const original = [first, ...before, last], observed = [first, ...after.map(p => ({ ...p, id: `imported-${p.id}` })), last]
    const verify = (a = source, b = current) => verifyImportedWordChatDeletedRow(a, b, original, observed, proposal)
    expect(verify()).toEqual({ authors: ["doc.haus"] })
    expect(current).toContain('w14:paraId="20000001"')
    expect(() => verifyNativeWordChatDeletedRow(source, current, original, [first, ...after, last], proposal)).toThrow()
    for (const changed of [
      current.replace('w14:paraId="20000002"', 'w14:paraId="20000001"'),
      current.replace(' w14:paraId="20000002"', ""),
      current.replace('w14:paraId="20000002"', 'w14:paraId="not-hex"'),
      current.replace('w14:paraId="20000002"', 'w14:paraId="00000000"'),
      current.replace('w14:paraId="20000002"', 'w14:paraId="80000000"'),
      current.replace('w14:paraId="20000002"', 'w14:paraId="30000001"'),
      current.replace('w14:paraId="30000001"', 'w14:paraId="10000001"'),
      current.replace('w14:paraId="30000001"', 'w14:paraId="40000001"'),
      current.replace('w14:textId="33333333"', 'w14:textId="33333334"'),
      current.replace('w:w="3600"', 'w:w="4000"'),
      current.replace("<w:t>head</w:t>", "<w:t>changed</w:t>"),
      current.replace("<w:delText>old</w:delText>", "<w:delText>changed</w:delText>"),
      current.replace('w:author="Reviewer"', 'w:author="Other"'),
      current.replace('w:date="2025-01-01T00:00:00Z"', 'w:date="2025-01-02T00:00:00Z"'),
    ]) expect(() => verify(source, changed)).toThrow()
    expect(() => verify(source.replace(' w14:paraId="10000002"', ""), current)).toThrow()
    for (const id of ["10000002", "20000002"]) {
      for (const site of ["<w:tbl>", '<w:p w14:paraId="30000001"']) {
        const reference = (xml: string) => xml.replace(site, site.replace(/>$/, "") + ` data-reference="prefix(${id.toLowerCase()})"` + (site.endsWith(">") ? ">" : ""))
        expect(() => verify(reference(source), reference(current))).toThrow()
      }
      const part = `<pkg:part pkg:name="/word/metadata.xml"><pkg:xmlData><meta reference="${id}"/></pkg:xmlData></pkg:part>`
      const reference = (xml: string) => xml.replace("</pkg:package>", part + "</pkg:package>")
      expect(() => verify(reference(source), reference(current))).toThrow()
    }
  })
})

describe("native row insertion verification", () => {
  const wrap = (body: string, comment = "Text", rsid = "00112233") =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="1" w:author="Reviewer"><w:p w:rsidR="${rsid}" w:rsidRDefault="${rsid}"><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><w:settings><w:rsids><w:rsid w:val="${rsid}"/></w:rsids></w:settings></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:style w:styleId="Normal"/></w:styles></pkg:xmlData></pkg:part></pkg:package>`
  const stamp = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
  const cell = (text: string, mark: number, content: number, change = false) =>
    `<w:tc><w:tcPr/><w:p><w:pPr><w:rPr><w:ins w:id="${mark}" ${stamp}/></w:rPr></w:pPr><w:ins w:id="${content}" ${stamp}><w:r>${change ? `<w:rPr><w:rPrChange w:id="9" ${stamp}><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>` : ""}<w:t>${text}</w:t></w:r></w:ins></w:p></w:tc>`
  const old = "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>Old two</w:t></w:r></w:p></w:tc></w:tr>"
  const row = `<w:tr><w:trPr><w:ins w:id="1" ${stamp}/></w:trPr>${cell("New", 2, 3, true)}${cell("Value", 4, 5)}</w:tr>`
  const body = (extra: string) => `<w:tbl><w:tblPr/><w:tblGrid/>${old}${extra}</w:tbl>`
  const before = wrap(body("")), after = wrap(body(row), "Text", "AABBCCDD")
  const verify = (xml: string) => verifyNativeWordChatInsertedRow(before, xml, 0, 1, ["New", "Value"])

  test("native inserted row verifies plain text, tracked carriers and narrow volatile dependencies", () => {
    expect(verify(after)).toEqual({ authors: ["Reviewer"] })
    for (const changed of [
      after.replace('w:id="1" ' + stamp, 'w:id="1" w:author="Other" w:date="2026-01-01T00:00:00Z"'),
      after.replace(">New<", ">Wrong<"),
      after.replace('<w:ins w:id="1"', '<w:del w:id="1"').replace("/></w:trPr>", "/></w:trPr>"),
      after.replace("<w:rPr><w:rPrChange", "<w:rPr><w:b/><w:rPrChange"),
      after.replace('w:id="9"', 'w:id="3"'),
      after.replace(">Text<", ">Changed<"),
      after.replace('w:styleId="Normal"', 'w:styleId="Other"'),
      after.replace("AABBCCDD", "arbitrary"),
      after.replace("<w:tcPr/><w:p><w:pPr>", "<w:tcPr/><w:drawing/><w:p><w:pPr>"),
      after.replace("<w:tcPr/><w:p><w:pPr>", '<w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:pPr>'),
      after.replace("<w:pPr><w:rPr>", '<w:pPr w:foo="1"><w:rPr>'),
    ]) expect(() => verify(changed)).toThrow()
  })

  test("new cell paragraph marks are plain even when Word records clearing inherited formatting", () => {
    const marker = `<w:ins w:id="2" ${stamp}/>`
    const changed = (props: string) => after.replace(marker, marker + props)
    const record = `<w:rPrChange w:id="10" ${stamp}><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr></w:rPrChange>`
    expect(verify(after)).toEqual({ authors: ["Reviewer"] })
    expect(verify(changed('<w:b w:val="0"/><w:i w:val="false"/><w:u w:val="none"/>' + record))).toEqual({ authors: ["Reviewer"] })
    for (const props of [
      "<w:b/>", "<w:bCs/>", "<w:i/>", '<w:u w:val="single"/>',
      record.replace('w:id="10"', 'w:id="3"'),
      record.replace("<w:b/>", '<w:sz w:val="28"/>'),
      record.replace("<w:b/>", '<w:b w:val="maybe"/>'),
      record.replace('w:author="Reviewer"', 'w:author="Other"'),
      record + record,
    ]) expect(() => verify(changed(props))).toThrow()
  })

  test("clearing a newly inserted paragraph mark may record its prior insertion carrier", () => {
    const marker = `<w:ins w:id="2" ${stamp}/>`
    const prior = `<w:ins w:id="11" ${stamp}/>`
    const history = (oldXml: string) => `<w:rPrChange w:id="10" ${stamp}><w:rPr>${oldXml}</w:rPr></w:rPrChange>`
    const changed = (oldXml: string, current = "") => after.replace(marker, marker + current + history(oldXml))
    expect(verify(changed(prior + "<w:b/>"))).toEqual({ authors: ["Reviewer"] })
    expect(verify(changed(prior))).toEqual({ authors: ["Reviewer"] })
    for (const oldXml of [
      prior.replace('w:author="Reviewer"', 'w:author="Other"') + "<w:b/>",
      prior.replace("2026-01-01T00:00:00Z", "2026-01-02T00:00:00Z") + "<w:b/>",
      prior.replace('w:id="11"', 'w:id="3"') + "<w:b/>",
      prior + prior,
      prior + prior.replace('w:id="11"', 'w:id="12"'),
      prior.replace("/>", "><w:b/></w:ins>"),
      prior.replace("w:ins", "w:del") + "<w:b/>",
      prior.replace("/>", ' w:extra="value"/>') + "<w:b/>",
      prior + '<w:sz w:val="24"/>',
      prior + '<w:b w:val="maybe"/>',
      prior + "<w:b/><w:b/>",
      prior + '<w:b/><w:bCs w:val="0"/>',
    ]) expect(() => verify(changed(oldXml))).toThrow()
    for (const current of ["<w:b/>", "<w:i/>", '<w:u w:val="single"/>']) expect(() => verify(changed(prior + "<w:b/>", current))).toThrow()
    expect(() => verify(after.replace("<w:rPr><w:b/></w:rPr></w:rPrChange>", `<w:rPr>${prior}<w:b/></w:rPr></w:rPrChange>`))).toThrow()
  })

  const space = (text: string) => (/^\s|\s$/.test(text) ? ' xml:space="preserve"' : "")
  const fcell = (runs: Array<[string, string]>, mark: number, content: number) =>
    `<w:tc><w:tcPr/><w:p><w:pPr><w:rPr><w:ins w:id="${mark}" ${stamp}/></w:rPr></w:pPr><w:ins w:id="${content}" ${stamp}>${runs.map(([props, text]) => `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t${space(text)}>${text}</w:t></w:r>`).join("")}</w:ins></w:p></w:tc>`
  const frow = (a: Array<[string, string]>, b: Array<[string, string]>) =>
    wrap(body(`<w:tr><w:trPr><w:ins w:id="1" ${stamp}/></w:trPr>${fcell(a, 2, 3)}${fcell(b, 4, 5)}</w:tr>`), "Text", "AABBCCDD")
  const span = (text: string, bold = false, italic = false, underline = false) => ({ text, bold, italic, underline })
  const cells = ["Fee: 100", "Note"], planned = [[span("Fee:", true), span(" 100")], [span("Note", false, true, true)]]
  const formatted = (xml: string, runs: typeof planned | undefined = planned) => verifyNativeWordChatInsertedRow(before, xml, 0, 1, cells, runs)
  const BI = "<w:b/><w:bCs/>", IU = '<w:i/><w:iCs/><w:u w:val="single"/>'

  test("formatted row admits on values only on the planned spans, across Word run splits", () => {
    expect(formatted(frow([[BI, "Fee:"], ["", " 100"]], [[IU, "Note"]]))).toEqual({ authors: ["Reviewer"] })
    expect(formatted(frow(
      [["<w:b/>", "Fe"], ['<w:b w:val="1"/><w:bCs w:val="true"/>', "e:"], ['<w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="0"/><w:u w:val="none"/>', " 1"], ["", "00"]],
      [['<w:i/><w:u w:val="single"/>', "No"], [`<w:i w:val="on"/><w:iCs/><w:u w:val="single"/><w:rPrChange w:id="9" ${stamp}><w:rPr/></w:rPrChange>`, "te"]],
    ))).toEqual({ authors: ["Reviewer"] })
    for (const [a, b] of [
      [[[BI, "Fee: "], ["", "100"]], [[IU, "Note"]]],
      [[["", "Fee:"], ["", " 100"]], [[IU, "Note"]]],
      [[[BI, "Fee:"], ["<w:i/>", " 100"]], [[IU, "Note"]]],
      [[[BI + '<w:sz w:val="24"/>', "Fee:"], ["", " 100"]], [[IU, "Note"]]],
      [[['<w:rStyle w:val="Strong"/>', "Fee:"], ["", " 100"]], [[IU, "Note"]]],
      [[['<w:b/><w:bCs w:val="0"/>', "Fee:"], ["", " 100"]], [[IU, "Note"]]],
      [[['<w:b w:val="maybe"/>', "Fee:"], ["", " 100"]], [[IU, "Note"]]],
      [[[BI, "Fee:"], ["", " 100"]], [['<w:i/><w:u w:val="double"/>', "Note"]]],
      [[[BI, "Fee:"], ["", " 100"]], [["<w:i/><w:u/>", "Note"]]],
      [[['<w:b w:color="FF0000"/>', "Fee:"], ["", " 100"]], [[IU, "Note"]]],
      [[[BI, "Fee:"], ["", " 10O"]], [[IU, "Note"]]],
    ] as Array<[Array<[string, string]>, Array<[string, string]>]>) expect(() => formatted(frow(a, b))).toThrow()
    const exact = frow([[BI, "Fee:"], ["", " 100"]], [[IU, "Note"]])
    expect(() => verifyNativeWordChatInsertedRow(before, exact, 0, 1, cells)).toThrow()
    expect(() => formatted(exact, [[span("Fee:", true), span(" 100")], [span("Note", false, true)]])).toThrow()
    expect(() => formatted(exact, [[span("Fee:", true), span(" 100")]])).toThrow()
  })

  test("plain row keeps one off span per cell, with explicit off values and split runs", () => {
    const plain = (xml: string) => verifyNativeWordChatInsertedRow(before, xml, 0, 1, ["New", "Value"])
    expect(plain(frow([['<w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="false"/><w:u w:val="none"/>', "Ne"], ["", "w"]], [["", "Value"]]))).toEqual({ authors: ["Reviewer"] })
    for (const props of ["<w:b/>", '<w:u w:val="single"/>', '<w:i w:val="1"/>', '<w:sz w:val="24"/>']) expect(() => plain(frow([[props, "New"]], [["", "Value"]]))).toThrow()
  })

  test("first revision by a Word user may add its people part and move only the unreferenced theme relationship", () => {
    const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    const W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
    const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${T}/${type}" Target="${target}"/>`
    const rels = (items: string) => `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData><Relationships xmlns="${REL}">${items}</Relationships></pkg:xmlData></pkg:part>`
    const peopleRel = '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>'
    const person = (author: string) => `<pkg:part pkg:name="/word/people.xml"><pkg:xmlData><w15:people xmlns:w15="${W15}"><w15:person w15:author="${author}"><w15:presenceInfo w15:providerId="AD" w15:userId="S-1"/></w15:person></w15:people></pkg:xmlData></pkg:part>`
    const add = (xml: string, parts: string) => xml.replace("</pkg:package>", parts + "</pkg:package>")
    const base = add(before, rels(rel("rId1", "styles", "styles.xml") + rel("rId3", "theme", "theme/theme1.xml")))
    const moved = (author = "Reviewer", relationships = rel("rId1", "styles", "styles.xml") + rel("rId4", "theme", "theme/theme1.xml") + peopleRel) => add(after, rels(relationships) + person(author))
    const check = (xml: string, from = base) => verifyNativeWordChatInsertedRow(from, xml, 0, 1, ["New", "Value"])
    expect(check(moved())).toEqual({ authors: ["Reviewer"] })
    for (const changed of [
      moved("Someone else"),
      moved("Reviewer", rel("rId9", "styles", "styles.xml") + rel("rId3", "theme", "theme/theme1.xml") + peopleRel.replace("rId3", "rId4")),
      moved("Reviewer", rel("rId1", "styles", "styles.xml") + rel("rId4", "theme", "theme/theme1.xml")),
      add(after, rels(rel("rId1", "styles", "styles.xml") + rel("rId3", "theme", "theme/theme1.xml")) + person("Reviewer").replace("/word/people.xml", "/word/other.xml")),
    ]) expect(() => check(changed)).toThrow()
    const cited = base.replace("<w:body>", `<w:body><w:p xmlns:r="${T}" r:id="rId3"/>`), citedAfter = moved().replace("<w:body>", `<w:body><w:p xmlns:r="${T}" r:id="rId3"/>`)
    expect(() => check(citedAfter, cited)).toThrow()
    const kept = moved("Reviewer", rel("rId1", "styles", "styles.xml") + rel("rId3", "theme", "theme/theme1.xml") + peopleRel.replace("rId3", "rId4")).replace("<w:body>", `<w:body><w:p xmlns:r="${T}" r:id="rId3"/>`)
    expect(check(kept, cited)).toEqual({ authors: ["Reviewer"] })
  })

  test("a document that already lists other reviewers may gain only the row's author in its people part", () => {
    const W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
    const person = (author: string) => `<w15:person w15:author="${author}"><w15:presenceInfo w15:providerId="None" w15:userId="${author}"/></w15:person>`
    const people = (persons: string) => `<pkg:part pkg:name="/word/people.xml"><pkg:xmlData><w15:people xmlns:w15="${W15}">${persons}</w15:people></pkg:xmlData></pkg:part>`
    const add = (xml: string, persons: string) => xml.replace("</pkg:package>", people(persons) + "</pkg:package>")
    const prior = add(before, person("Prior"))
    const check = (persons: string) => verifyNativeWordChatInsertedRow(prior, add(after, persons), 0, 1, ["New", "Value"])
    expect(check(person("Prior") + person("Reviewer"))).toEqual({ authors: ["Reviewer"] })
    expect(check(person("Prior"))).toEqual({ authors: ["Reviewer"] })
    for (const persons of [
      person("Prior") + person("Someone else"),
      person("Reviewer"),
      person("Prior").replace('w15:userId="Prior"', 'w15:userId="Changed"') + person("Reviewer"),
      person("Prior") + person("Reviewer") + person("Reviewer"),
    ]) expect(() => check(persons)).toThrow()
  })

  test("an added row may carry the reference row's own paragraph properties, and only those", () => {
    const spacing = '<w:spacing w:after="40"/>'
    const spaced = (xml: string) => xml.replaceAll("<w:p><w:r>", `<w:p><w:pPr>${spacing}</w:pPr><w:r>`).replaceAll("<w:p><w:pPr><w:rPr>", `<w:p><w:pPr>${spacing}<w:rPr>`)
    const from = spaced(before), check = (xml: string, source = from) => verifyNativeWordChatInsertedRow(source, xml, 0, 1, ["New", "Value"])
    expect(check(spaced(after))).toEqual({ authors: ["Reviewer"] })
    for (const [xml, source] of [
      [spaced(after).replaceAll(spacing, '<w:spacing w:after="80"/>').replace(spaced(old).replaceAll(spacing, '<w:spacing w:after="80"/>'), spaced(old)), from],
      [spaced(after).replace(spaced(old), old), before],
      [spaced(after).replace(`${spacing}<w:rPr><w:ins w:id="2"`, `${spacing}<w:jc w:val="center"/><w:rPr><w:ins w:id="2"`), from],
      [spaced(after).replace(`${spacing}<w:rPr><w:ins w:id="2"`, `${spacing}<w:pPrChange w:id="8" ${stamp}><w:pPr/></w:pPrChange><w:rPr><w:ins w:id="2"`), from],
    ] as Array<[string, string]>) expect(() => check(xml, source)).toThrow()
  })

  test("Word's no-op table, grid, row and cell records and its fixed grid-total width settle before the row check", () => {
    const tblPr = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>', grid = '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2600"/></w:tblGrid>', jc = '<w:jc w:val="center"/>'
    const rec = (name: string, id: number, inner: string) => `<w:${name}Change w:id="${id}" ${stamp}>${inner}</w:${name}Change>`
    const ref = old.replace("<w:tr>", `<w:tr><w:trPr>${jc}</w:trPr>`), from = wrap(`<w:tbl>${tblPr}${grid}${ref}</w:tbl>`)
    const width = (w: string, type = "dxa") => `<w:tblPr><w:tblW w:w="${w}" w:type="${type}"/>${rec("tblPr", 20, tblPr)}</w:tblPr>`
    const recorded = (table = width("5000"), newRowPr = jc, cellPr = `<w:tcPr>${rec("tcPr", 22, "<w:tcPr/>")}</w:tcPr>`) => {
      const oldRow = ref.replace(`<w:trPr>${jc}</w:trPr>`, `<w:trPr>${jc}${rec("trPr", 21, `<w:trPr>${jc}</w:trPr>`)}</w:trPr>`).replace("<w:tr>", '<w:tr w:rsidTr="AABBCCDD">').replaceAll("<w:tcPr/>", cellPr)
      const added = row.replace(`<w:trPr><w:ins w:id="1" ${stamp}/></w:trPr>`, `<w:trPr>${newRowPr}<w:ins w:id="1" ${stamp}/>${rec("trPr", 23, `<w:trPr>${newRowPr}</w:trPr>`)}</w:trPr>`)
      return wrap(`<w:tbl>${table}${grid.replace("</w:tblGrid>", rec("tblGrid", 24, grid) + "</w:tblGrid>")}${oldRow}${added}</w:tbl>`, "Text", "AABBCCDD")
    }
    const check = (xml: string) => verifyNativeWordChatInsertedRow(from, xml, 0, 1, ["New", "Value"])
    expect(check(recorded())).toEqual({ authors: ["Reviewer"] })
    for (const newRowPr of ['<w:jc w:val="right"/>', ""]) expect(() => check(recorded(undefined, newRowPr))).toThrow()
  })

  test("a row add that also re-fits the table's column widths as a tracked change is named, not admitted", () => {
    const grid = (a: number, b: number) => `<w:tblGrid><w:gridCol w:w="${a}"/><w:gridCol w:w="${b}"/></w:tblGrid>`
    const table = (g: string, rows: string) => wrap(`<w:tbl><w:tblPr/>${g}${old}${rows}</w:tbl>`)
    const start = table(grid(2400, 2400), "")
    const refit = table('<w:tblGrid><w:gridCol w:w="873"/><w:gridCol w:w="972"/><w:tblGridChange w:id="20"><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid></w:tblGridChange></w:tblGrid>', row)
    expect(nativeWordChatRowRefitWidths(start, refit, 0)).toBe(true)
    expect(nativeWordChatRowRefitWidths(start, table(grid(2400, 2400), row), 0)).toBe(false)
    expect(nativeWordChatRowRefitWidths(start, table(grid(873, 972), row), 0)).toBe(false)
    expect(nativeWordChatRowRefitWidths(start, refit, 1)).toBe(false)
  })
})
