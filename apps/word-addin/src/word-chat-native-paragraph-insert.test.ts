import { expect, test } from "bun:test"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import {
  assertNativeWordChatParagraphInsertionTarget,
  verifyNativeWordChatParagraphInsertion,
} from "./word-chat-native-paragraph-insert"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  W14 = "http://schemas.microsoft.com/office/word/2010/wordml",
  PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const flat = (body: string, style = "Normal") =>
  `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml" pkg:contentType="style"><pkg:xmlData><w:styles><w:style w:styleId="${style}"/></w:styles></pkg:xmlData></pkg:part></pkg:package>`
const stamp = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
const old =
  '<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:t>Original</w:t></w:r></w:p><w:p w14:paraId="44444444"><w:ins w:id="5" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p>'
const next = `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr><w:r><w:t>Original</w:t></w:r></w:p><w:p w14:paraId="11111111" w14:textId="12345678"><w:ins w:id="2" ${stamp}><w:r><w:t>Added</w:t></w:r></w:ins></w:p><w:p w14:paraId="44444444"><w:ins w:id="7" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p>`
const beforeInv = [
    { id: "orig", text: "Original", tableNestingLevel: 0 },
    { id: "prior", text: "Prior", tableNestingLevel: 0 },
  ],
  afterInv = [
    { id: "new-left", text: "Original", tableNestingLevel: 0 },
    { id: "orig", text: "Added", tableNestingLevel: 0 },
    { id: "prior", text: "Prior", tableNestingLevel: 0 },
  ]
const target = {
  id: "selected",
  text: "Original",
  paragraphs: [{ id: "orig", text: "Original" }],
  startOffset: 0,
  endOffset: 8,
  capturedAt: "now",
  saved: false,
}
const check = (xml: string, inv = afterInv) =>
  verifyNativeWordChatParagraphInsertion(
    projectNativeWordChatBody(flat(old), beforeInv),
    projectNativeWordChatBody(xml, inv),
    target,
    "Added",
  )
test("tracked paragraph insertion validates transferred identity, marks and protected body", () => {
  const source = projectNativeWordChatBody(flat(old), beforeInv)
  expect(assertNativeWordChatParagraphInsertionTarget(source, target)).toBe(0)
  expect(check(flat(next))).toEqual(["Reviewer"])
  for (const bad of [
    flat(next.replace('w:id="2"', 'w:id="1"')),
    flat(next.replace('w:author="Earlier"', 'w:author="Other"')),
    flat(next.replace('w:author="Reviewer"', 'w:author="Other"')),
    flat(next.replace('w14:paraId="11111111"', 'w14:paraId="33333333"')),
    flat(next.replace("<w:t>Added</w:t>", "<w:t>Added</w:t><w:drawing/>")),
    flat(next.replace(">Prior<", ">Changed<")),
    flat(next.replace('w14:paraId="44444444"', 'w14:paraId="55555555"')),
    flat(next, "Changed"),
  ])
    expect(() => check(bad)).toThrow()
  expect(() => assertNativeWordChatParagraphInsertionTarget(source, { ...target, endOffset: 4 })).toThrow()
  expect(() =>
    check(flat(next), [{ ...afterInv[0]!, id: "orig" }, { ...afterInv[1]!, id: "new" }, afterInv[2]!]),
  ).toThrow()
})

test("annotations after the inserted paragraph keep location and payload across ID remapping", () => {
  const wrap = (body: string, commentId: string) =>
    flat(body).replace(
      "</pkg:package>",
      `<pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="${commentId}" w:author="Reviewer"><w:p w:rsidR="00112233" w:rsidRDefault="00112233"><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part></pkg:package>`,
    )
  const anchored = (body: string, commentId: string, bookmarkId: string) =>
    body
      .replace(
        '<w:p w14:paraId="44444444"><w:ins w:id=',
        `<w:p w14:paraId="44444444"><w:bookmarkStart w:id="${bookmarkId}" w:name="Keep"/><w:commentRangeStart w:id="${commentId}"/><w:ins w:id=`,
      )
      .replace(
        ">Prior</w:t></w:r></w:ins></w:p>",
        `>Prior</w:t></w:r></w:ins><w:commentRangeEnd w:id="${commentId}"/><w:bookmarkEnd w:id="${bookmarkId}"/><w:r><w:commentReference w:id="${commentId}"/></w:r></w:p>`,
      )
  const a = wrap(anchored(old, "8", "10"), "8"),
    b = wrap(anchored(next, "9", "11"), "9")
  const source = projectNativeWordChatBody(a, beforeInv),
    destination = projectNativeWordChatBody(b, afterInv)
  expect(verifyNativeWordChatParagraphInsertion(source, destination, target, "Added")).toEqual(["Reviewer"])
  expect(() =>
    verifyNativeWordChatParagraphInsertion(
      source,
      projectNativeWordChatBody(b.replace(">Comment<", ">Tampered<"), afterInv),
      target,
      "Added",
    ),
  ).toThrow()
})

test("only projection-qualified unbound terminal carrier may change physical ID", () => {
  const a = flat(old + '<w:p w14:paraId="AAAAAAAA"/><w:sectPr/>'),
    b = flat(next + '<w:p w14:paraId="BBBBBBBB"/><w:sectPr/>')
  expect(
    verifyNativeWordChatParagraphInsertion(
      projectNativeWordChatBody(a, beforeInv),
      projectNativeWordChatBody(b, afterInv),
      target,
      "Added",
    ),
  ).toEqual(["Reviewer"])
  expect(() =>
    projectNativeWordChatBody(
      flat(next + '<w:p w14:paraId="BBBBBBBB"><w:bookmarkStart w:id="3" w:name="x"/></w:p><w:sectPr/>'),
      afterInv,
    ),
  ).toThrow()
})
