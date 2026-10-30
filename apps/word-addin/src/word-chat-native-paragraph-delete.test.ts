import { expect, test } from "bun:test"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import {
  assertNativeWordChatParagraphDeletionTarget,
  verifyNativeWordChatParagraphDeletion,
} from "./word-chat-native-paragraph-delete"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  W14 = "http://schemas.microsoft.com/office/word/2010/wordml",
  PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const flat = (body: string, comment = "Keep") =>
  `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="3" w:author="Earlier"><w:p><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part></pkg:package>`
const targetOld = '<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:t>Original</w:t></w:r></w:p>'
const targetAfter =
  '<w:p w14:paraId="11111111" w14:textId="12345678" w:rsidDel="00112233"><w:pPr><w:rPr><w:del w:id="2" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="3" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r w:rsidDel="00112233"><w:delText>Original</w:delText></w:r></w:del></w:p>'
const rest =
  '<w:p w14:paraId="22222222"><w:r><w:t>Next</w:t></w:r></w:p><w:p w14:paraId="33333333"><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:ins w:id="5" w:author="Earlier" w:date="2025-01-01"><w:r><w:t xml:space="preserve"> Prior </w:t></w:r></w:ins><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>'
const inv = [
  { id: "a", text: "Original", tableNestingLevel: 0 },
  { id: "b", text: "Next", tableNestingLevel: 0 },
  { id: "c", text: " Prior ", tableNestingLevel: 0 },
]
const afterInv = [{ ...inv[0]!, text: "" }, inv[1]!, inv[2]!]
const target = {
  id: "selected",
  text: "Original",
  paragraphs: [{ id: "a", text: "Original" }],
  startOffset: 0,
  endOffset: 8,
  capturedAt: "now",
  saved: false,
}
const before = projectNativeWordChatBody(flat(targetOld + rest), inv)
const after = (body: string, inventory = afterInv, comment = "Keep") =>
  projectNativeWordChatBody(flat(body, comment), inventory)
const check = (body: string, inventory = afterInv, comment = "Keep") =>
  verifyNativeWordChatParagraphDeletion(before, after(body, inventory, comment), target)
test("native deletion retains identity and validates both owned carriers and prior annotations", () => {
  expect(assertNativeWordChatParagraphDeletionTarget(before, target)).toBe(0)
  const good = targetAfter + rest.replace('w:id="5" w:author="Earlier"', 'w:id="7" w:author="Earlier"')
  expect(check(good)).toEqual(["Reviewer"])
  for (const bad of [
    good.replace('w:id="3" w:author="Reviewer"', 'w:id="2" w:author="Reviewer"'),
    good.replace('w:author="Reviewer"', 'w:author="Other"'),
    good.replace(">Original</w:delText>", ">Changed</w:delText>"),
    good.replace("> Prior </w:t>", "> Prior</w:t>"),
    good.replace('w:author="Earlier"', 'w:author="Other"'),
    good.replace('w14:paraId="22222222"', 'w14:paraId="99999999"'),
    good.replace('w:name="Keep"', 'w:name="Changed"'),
    good.replace("<w:pPr>", '<w:pPr><w:spacing w:after="1"/>'),
  ])
    expect(() => check(bad)).toThrow()
  expect(() => check(good, afterInv, "Changed")).toThrow()
  expect(() => check(good, [{ ...afterInv[0]!, id: "changed" }, afterInv[1]!, afterInv[2]!])).toThrow()
  expect(() => assertNativeWordChatParagraphDeletionTarget(before, { ...target, endOffset: 4 })).toThrow()
})

test("a following ordinary paragraph can retain its comment and bookmark graph", () => {
  const following = '<w:p w14:paraId="22222222"><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:r><w:t>Next</w:t></w:r><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>'
  const old = projectNativeWordChatBody(flat(targetOld + following),inv.slice(0,2))
  const checkFollowing = (body: string, comment = "Keep") => verifyNativeWordChatParagraphDeletion(old,projectNativeWordChatBody(flat(body,comment),afterInv.slice(0,2)),target)
  const next = targetAfter + following
  expect(assertNativeWordChatParagraphDeletionTarget(old,target)).toBe(0)
  expect(checkFollowing(next)).toEqual(["Reviewer"])
  expect(() => checkFollowing(next.replace('w:name="Keep"','w:name="Changed"'))).toThrow()
  expect(() => checkFollowing(next.replace('<w:commentRangeEnd w:id="3"/>',''))).toThrow()
  expect(() => checkFollowing(next,"Changed")).toThrow()
})
