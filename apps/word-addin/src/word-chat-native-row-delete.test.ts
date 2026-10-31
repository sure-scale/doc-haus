import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { assertNativeWordChatRowDeletionTarget, verifyNativeWordChatDeletedRow, verifyReviewedWordChatRowDeletionOracle } from "./word-chat-native-row-delete"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", P = "http://schemas.microsoft.com/office/2006/xmlPackage"
const attrs = `xmlns:w="${W}"`
const row = (a: string, b: string) => `<w:tr><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${a}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>${b}</w:t></w:r></w:p></w:tc></w:tr>`
const beforeRow = row("old", "value")
const mark = (id: number) => `w:id="${id}" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"`
const deletedRow = `<w:tr><w:trPr><w:del ${mark(1)}/></w:trPr><w:tc><w:p><w:pPr><w:rPr><w:del ${mark(2)}/></w:rPr></w:pPr><w:del ${mark(3)}><w:r><w:rPr><w:b/></w:rPr><w:delText>old</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p><w:pPr><w:rPr><w:del ${mark(4)}/></w:rPr></w:pPr><w:del ${mark(5)}><w:r><w:delText>value</w:delText></w:r></w:del></w:p></w:tc></w:tr>`
const body = (r: string) => `<w:document ${attrs}><w:body><w:tbl>${row("head", "caption")}${r}</w:tbl><w:sectPr/></w:body></w:document>`
const flat = (r: string) => `<pkg:package xmlns:pkg="${P}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${body(r)}</pkg:xmlData></pkg:part></pkg:package>`
const before = [{id:"a",text:"head"},{id:"b",text:"caption"},{id:"c",text:"old"},{id:"d",text:"value"}].map(x => ({...x,tableNestingLevel:1}))
const after = before.map((x,i) => ({...x,text:i>=2?"":x.text}))
const proposal = {kind:"table_row_delete" as const,id:"p",documentId:"d",explanation:"",openItems:[],createdAt:"",table:{tableIndex:0,rows:[before.slice(0,2),before.slice(2)]},rowIndex:1,cells:["old","value"]}

test("proves a formatted native row deletion and rejects changes to its preimage, carriers and remainder", () => {
  assertNativeWordChatRowDeletionTarget(flat(beforeRow), before, proposal)
  expect(verifyNativeWordChatDeletedRow(flat(beforeRow),flat(deletedRow),before,after,proposal)).toEqual({authors:["doc.haus"]})
  expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow),flat(deletedRow.replace("<w:b/>","<w:i/>")),before,after,proposal)).toThrow()
  expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow),flat(deletedRow.replace(mark(5),mark(4))),before,after,proposal)).toThrow()
  expect(() => assertNativeWordChatRowDeletionTarget(flat(beforeRow.replace("<w:b/>","<w:vanish/>")),before,proposal)).toThrow()
  expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow),flat(deletedRow).replace("head","changed"),before,after,proposal)).toThrow()
  expect(() => assertNativeWordChatRowDeletionTarget(flat(beforeRow).replace("<w:tbl>", `<w:bookmarkStart w:id="8" w:name="crossing"/><w:tbl>`).replace("</w:tbl>", `</w:tbl><w:bookmarkEnd w:id="8"/>`),before,proposal)).toThrow()
  expect(() => assertNativeWordChatRowDeletionTarget(flat(beforeRow.replace("<w:tr>", `<w:tr><w:trPr><w:ins ${mark(9)}/></w:trPr>`)),before,proposal)).toThrow()
  expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow),flat(deletedRow.replace('w:author="doc.haus"','w:author="Other"')),before,after,proposal)).toThrow()
})

test("the SDK oracle is source-bound and preserves every unrelated package part", async () => {
  const source = zipSync({"word/document.xml":strToU8(body(beforeRow)),"word/settings.xml":strToU8(`<w:settings ${attrs}/>`),"other":strToU8("keep")})
  const authored = (other="keep") => zipSync({"word/document.xml":strToU8(body(deletedRow)),"word/settings.xml":strToU8(`<w:settings ${attrs}><w:trackRevisions/></w:settings>`),"other":strToU8(other)})
  const hash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256",source))).map(x=>x.toString(16).padStart(2,"0")).join("")
  const oracle = {kind:"table_row_delete" as const,proposalId:proposal.id,sourceSha256:hash,table:proposal.table,rowIndex:1,cells:proposal.cells,author:"doc.haus" as const,writerRequest:{source_sha256:hash,target_handle:"bound",author:"doc.haus"},authoredBase64:Buffer.from(authored()).toString("base64"),revisions:[{id:"r",family:"row_delete"}]}
  await verifyReviewedWordChatRowDeletionOracle(source,oracle,proposal)
  await expect(verifyReviewedWordChatRowDeletionOracle(source,{...oracle,authoredBase64:Buffer.from(authored("altered")).toString("base64")},proposal)).rejects.toThrow()
  await expect(verifyReviewedWordChatRowDeletionOracle(source,{...oracle,sourceSha256:"bad"},proposal)).rejects.toThrow()
})

test("a new Word people relationship must be unique, internal and target the new people part", () => {
  const relns = "http://schemas.openxmlformats.org/package/2006/relationships", people = "http://schemas.microsoft.com/office/2011/relationships/people", pns = "http://schemas.microsoft.com/office/word/2012/wordml"
  const part = (name: string, xml: string) => `<pkg:part pkg:name="${name}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`
  const relationship = (id: string, type: string, target: string, extra = "") => `<Relationship Id="${id}" Type="${type}" Target="${target}"${extra}/>`
  const rels = (children: string) => part("/word/_rels/document.xml.rels", `<Relationships xmlns="${relns}">${children}</Relationships>`)
  const person = part("/word/people.xml", `<w15:people xmlns:w15="${pns}"><w15:person w15:author="doc.haus"><w15:presenceInfo w15:providerId="x" w15:userId="y"/></w15:person></w15:people>`)
  const theme = (id: string) => relationship(id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme.xml")
  const wrap = (xml: string, extra: string) => xml.replace("</pkg:package>", `${extra}</pkg:package>`)
  const original = wrap(flat(beforeRow), rels(theme("r1")))
  const observed = (link: string) => wrap(flat(deletedRow), rels(theme("r2") + link) + person)
  const valid = relationship("r1", people, "people.xml")
  expect(verifyNativeWordChatDeletedRow(original,observed(valid),before,after,proposal).authors).toEqual(["doc.haus"])
  for (const link of [relationship("r1", people, "other.xml"), relationship("r1", people, "people.xml", ' TargetMode="External"'), valid + relationship("r3", people, "people.xml"), relationship("r2", people, "people.xml")]) {
    expect(() => verifyNativeWordChatDeletedRow(original,observed(link),before,after,proposal)).toThrow()
  }
  expect(() => verifyNativeWordChatDeletedRow(flat(beforeRow),flat(deletedRow.replace(mark(5),mark(4).replace('w:id="4"','w:id="04"'))),before,after,proposal)).toThrow()
})
