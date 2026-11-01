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

test("a doc.haus pointer bookmark on the deleted paragraph survives Word moving its end into the deletion", () => {
  const ptrName = "_dh_p0123456789ab", runs = "<w:r><w:t>Original</w:t></w:r>"
  const oldPointer = targetOld.replace(runs, `<w:bookmarkStart w:id="11" w:name="${ptrName}"/>${runs}<w:bookmarkEnd w:id="11"/>`)
  const old = projectNativeWordChatBody(flat(oldPointer + rest), inv)
  const run = (body: string) => verifyNativeWordChatParagraphDeletion(old, projectNativeWordChatBody(flat(body), afterInv), target)
  // Word renumbers the bookmark, leaves the start before the deletion and moves the end inside it.
  const deleted = targetAfter.replace("<w:del w:id=\"3\"", `<w:bookmarkStart w:id="9" w:name="${ptrName}"/><w:del w:id="3"`).replace("</w:del></w:p>", '<w:bookmarkEnd w:id="9"/></w:del></w:p>')
  const good = deleted + rest.replace('w:id="5" w:author="Earlier"', 'w:id="7" w:author="Earlier"')
  expect(assertNativeWordChatParagraphDeletionTarget(old, target)).toBe(0)
  expect(run(good)).toEqual(["Reviewer"])
  // The end left after the deletion is the same bookmark.
  expect(run(good.replace('<w:bookmarkEnd w:id="9"/></w:del>', "</w:del><w:bookmarkEnd w:id=\"9\"/>"))).toEqual(["Reviewer"])
  for (const bad of [
    good.replace(`w:name="${ptrName}"`, 'w:name="_dh_p00000000ffff"'), // a different bookmark
    good.replace('<w:bookmarkEnd w:id="9"/></w:del>', "</w:del>"), // end lost
    good.replace('<w:bookmarkEnd w:id="9"/>', '<w:bookmarkEnd w:id="8"/>'), // end of another bookmark
    good.replace(`<w:bookmarkStart w:id="9" w:name="${ptrName}"/>`, ""), // start lost
  ])
    expect(() => run(bad)).toThrow()
  // Anything wider than the one pointer pair is still refused up front.
  for (const bad of [targetOld.replace(runs, `<w:bookmarkStart w:id="1" w:name="Keep"/>${runs}<w:bookmarkEnd w:id="1"/>`), targetOld.replace(runs, `<w:bookmarkStart w:id="11" w:name="${ptrName}"/>${runs}`)])
    expect(() => assertNativeWordChatParagraphDeletionTarget(projectNativeWordChatBody(flat(bad + rest), inv), target)).toThrow()
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

test("a numbered heading with formatted runs deletes as one exact tracked paragraph", () => {
  const props = '<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
  const heading = `<w:p w14:paraId="44444444" w14:textId="55555555" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3">${props}<w:r w:rsidRPr="00D4E5F6"><w:rPr><w:b/><w:bCs/></w:rPr><w:t>Termination</w:t></w:r><w:r><w:t xml:space="preserve"> of Agreement</w:t></w:r></w:p>`
  const stamp = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
  const mark = `<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:rPr><w:del w:id="10" ${stamp}/></w:rPr></w:pPr>`
  const bold = '<w:r w:rsidRPr="00D4E5F6" w:rsidDel="00112233"><w:rPr><w:b/><w:bCs/></w:rPr><w:delText>Termination</w:delText></w:r>'
  const plain = '<w:r w:rsidDel="00112233"><w:delText xml:space="preserve"> of Agreement</w:delText></w:r>'
  const deleted = (inner: string, pPr = mark) =>
    `<w:p w14:paraId="44444444" w14:textId="66666666" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3" w:rsidDel="00112233">${pPr}${inner}</w:p>`
  const one = deleted(`<w:del w:id="11" ${stamp}>${bold}${plain}</w:del>`)
  const each = deleted(`<w:del w:id="11" ${stamp}>${bold}</w:del><w:del w:id="12" ${stamp}>${plain}</w:del>`)
  const text = "Termination of Agreement",
    headingInv = [{ ...inv[0]!, id: "h", text }, inv[1]!, inv[2]!],
    deletedInv = [{ ...headingInv[0]!, text: "" }, inv[1]!, inv[2]!]
  const heads = { ...target, text, paragraphs: [{ id: "h", text }], endOffset: text.length }
  const old = projectNativeWordChatBody(flat(heading + rest), headingInv)
  const checkHeading = (body: string) => verifyNativeWordChatParagraphDeletion(old, after(body, deletedInv), heads)
  expect(assertNativeWordChatParagraphDeletionTarget(old, heads)).toBe(0)
  expect(checkHeading(one + rest)).toEqual(["Reviewer"])
  expect(checkHeading(each + rest)).toEqual(["Reviewer"])
  // A mark that already carried run properties keeps them beside Word's w:del.
  const markProps = props.replace("</w:pPr>", "<w:rPr><w:b/></w:rPr></w:pPr>")
  const marked = projectNativeWordChatBody(flat(heading.replace(props, markProps) + rest), headingInv)
  const markedAfter = deleted(`<w:del w:id="11" ${stamp}>${bold}${plain}</w:del>`, mark.replace("</w:rPr>", "<w:b/></w:rPr>"))
  expect(verifyNativeWordChatParagraphDeletion(marked, after(markedAfter + rest, deletedInv), heads)).toEqual(["Reviewer"])
  expect(() => verifyNativeWordChatParagraphDeletion(marked, after(one + rest, deletedInv), heads)).toThrow()
  for (const bad of [
    one + rest.replace('<w:p w14:paraId="22222222">', '<w:p w14:paraId="22222222"><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'),
    one.replace("<w:delText>Termination</w:delText>", "<w:t>Termination</w:t>") + rest,
    one.replace(`<w:rPr><w:del w:id="10" ${stamp}/></w:rPr>`, "") + rest,
    one.replace('<w:pStyle w:val="Heading1"/>', '<w:pStyle w:val="Heading2"/>') + rest,
    one.replace('<w:numId w:val="1"/>', '<w:numId w:val="2"/>') + rest,
    one.replace("<w:b/><w:bCs/>", "<w:bCs/>") + rest,
    one.replace('w:id="11"', 'w:id="10"') + rest,
    one.replace(`w:id="11" ${stamp}`, 'w:id="11" w:author="Other" w:date="2026-01-01T00:00:00Z"') + rest,
    one.replace("Agreement</w:delText>", "Agreement</w:delText><w:tab/>") + rest,
  ])
    expect(() => checkHeading(bad)).toThrow()
  for (const unsupported of [
    heading.replace("<w:r>", '<w:commentRangeStart w:id="3"/><w:r>'),
    heading.replace("<w:r>", '<w:bookmarkStart w:id="8" w:name="_Toc1"/><w:bookmarkEnd w:id="8"/><w:r>'),
    heading.replace(/<w:r>(<w:t xml:space="preserve"> of Agreement<\/w:t>)<\/w:r>/, `<w:ins w:id="9" ${stamp}><w:r>$1</w:r></w:ins>`),
    heading.replace("</w:numPr>", `</w:numPr><w:rPr><w:ins w:id="9" ${stamp}/></w:rPr>`),
    heading.replace("<w:b/>", `<w:b/><w:rPrChange w:id="9" ${stamp}><w:rPr/></w:rPrChange>`),
  ])
    expect(() => assertNativeWordChatParagraphDeletionTarget(projectNativeWordChatBody(flat(unsupported + rest), headingInv), heads)).toThrow()
})
