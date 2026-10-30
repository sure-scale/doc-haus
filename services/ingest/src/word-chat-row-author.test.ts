import { expect, test } from "bun:test"
import { zipSync, strToU8, unzipSync, strFromU8 } from "fflate"
import { createHash } from "node:crypto"
import { authorWordChatTableRow } from "./word-chat-amendments"
import type { WordChatSnapshot, WordChatTableRowInsertionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
function fixture(value: string) {
  const p = (s: string) => `<w:p><w:r><w:t>${s}</w:t></w:r></w:p>`
  const cell = (s: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(s)}</w:tc>`
  return zipSync({
    "[Content_Types].xml":strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    "_rels/.rels":strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels":strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/></Relationships>`),
    "word/settings.xml":strToU8(`<w:settings xmlns:w="${W}"/>`),
    "word/document.xml":strToU8(`<w:document xmlns:w="${W}"><w:body><w:p><w:del w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Old</w:delText></w:r></w:del><w:ins w:id="2" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${cell("First")}${cell("Second")}</w:tr><w:tr>${cell(value)}${cell("Fourth")}</w:tr></w:tbl>${p("After")}<w:sectPr/></w:body></w:document>`),
  })
}
function request(value: string) {
  const paragraphs = ["Prior","First","Second",value,"Fourth","After"].map((text, i) => ({id:`native-${i}`,text}))
  const table = {tableIndex:0,rows:[paragraphs.slice(1,3),paragraphs.slice(3,5)]}
  const snapshot: WordChatSnapshot = {text:paragraphs.map(p => p.text).join("\r"),paragraphs,table,saved:false,capturedAt:"now",truncated:false}
  const proposal: WordChatTableRowInsertionProposal = {kind:"table_row_insert",id:"row-1",documentId:"document",explanation:"Add a row",openItems:[],createdAt:"now",position:"end",table,cells:["New","Value"]}
  return {snapshot,proposal}
}

test("row authoring uses the current compressed export and retains prior revisions", async () => {
  const saved = fixture("Saved"), current = fixture("Unsaved"), {snapshot,proposal} = request("Unsaved")
  const oracle = await authorWordChatTableRow(snapshot, proposal, current)
  expect(oracle.sourceSha256).toBe(createHash("sha256").update(current).digest("hex"))
  expect(oracle.sourceSha256).not.toBe(createHash("sha256").update(saved).digest("hex"))
  expect(oracle.writerRequest.position).toBe("after")
  expect(oracle.table).toEqual(snapshot.table!)
  const xml = strFromU8(unzipSync(oracle.authoredBytes)["word/document.xml"]!)
  expect(xml).toContain("Unsaved")
  expect(xml).toContain("Reviewer")
  expect(xml).toContain("delText")
  await expect(authorWordChatTableRow(snapshot, proposal, saved)).rejects.toThrow("word_row_source_changed")
})

test("row authoring refuses altered table binding and oversized or malformed source", async () => {
  const source = fixture("Current"), {snapshot,proposal} = request("Current")
  await expect(authorWordChatTableRow(snapshot, {...proposal,table:{...proposal.table,tableIndex:1}}, source)).rejects.toThrow("word_row_proposal_changed")
  const altered = request("Changed")
  await expect(authorWordChatTableRow(altered.snapshot, altered.proposal, source)).rejects.toThrow("word_row_source_changed")
  await expect(authorWordChatTableRow(snapshot, proposal, new Uint8Array(16 * 1024 * 1024 + 1))).rejects.toThrow("word_row_source_size")
  await expect(authorWordChatTableRow(snapshot, proposal, new Uint8Array([1,2,3]))).rejects.toThrow()
  const oversized = zipSync({"big.xml":new Uint8Array(16 * 1024 * 1024 + 1)})
  await expect(authorWordChatTableRow(snapshot, proposal, oversized)).rejects.toThrow("word_row_source_archive")
})
