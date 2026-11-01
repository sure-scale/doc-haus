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

test("a doc.haus pointer bookmark on the anchor is accepted, and nothing wider is", () => {
  const ptr = (id: string, name = "_dh_p0123456789ab") => (runs: string) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${runs}<w:bookmarkEnd w:id="${id}"/>`
  const originalRuns = "<w:r><w:t>Original</w:t></w:r>"
  const oldPointer = old.replace(originalRuns, ptr("1")(originalRuns))
  // Word renumbers the bookmark when it adds the tracked insertion, but keeps it on the anchor text.
  const nextPointer = next.replace(originalRuns, ptr("9")(originalRuns))
  const run = (beforeBody: string, afterBody: string, t = target) =>
    verifyNativeWordChatParagraphInsertion(projectNativeWordChatBody(flat(beforeBody), beforeInv), projectNativeWordChatBody(flat(afterBody), afterInv), t, "Added")
  expect(assertNativeWordChatParagraphInsertionTarget(projectNativeWordChatBody(flat(oldPointer), beforeInv), target)).toBe(0)
  expect(run(oldPointer, nextPointer)).toEqual(["Reviewer"])
  for (const bad of [
    old.replace(originalRuns, ptr("1", "Keep")(originalRuns)), // someone else's bookmark
    old.replace(originalRuns, ptr("1", "_dh_pXYZ")(originalRuns)), // not a pointer name
    old.replace(originalRuns, `<w:bookmarkStart w:id="1" w:name="_dh_p0123456789ab"/>${originalRuns}`), // never closed
    old.replace(originalRuns, `${ptr("1")(originalRuns)}<w:bookmarkStart w:id="2" w:name="_dh_p0123456789cd"/><w:bookmarkEnd w:id="2"/>`), // a second one
    old.replace(originalRuns, `<w:bookmarkStart w:id="1" w:name="_dh_p0123456789ab"/><w:bookmarkEnd w:id="1"/>${originalRuns}`), // wraps nothing
  ])
    expect(() => assertNativeWordChatParagraphInsertionTarget(projectNativeWordChatBody(flat(bad), beforeInv), target)).toThrow()
  // The bookmark must stay on the anchor text, not move to the new paragraph.
  expect(() => run(oldPointer, next.replace("<w:t>Added</w:t></w:r>", `<w:t>Added</w:t></w:r><w:bookmarkEnd w:id="9"/>`).replace(originalRuns, ptr("9")(originalRuns)))).toThrow()
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

// Styled anchors: `props` is the anchor pPr body, `runs` its content.
const prior = '<w:p w14:paraId="44444444"><w:r><w:t>Prior</w:t></w:r></w:p>'
const styled = (props: string, runs: string) =>
  `<w:p w14:paraId="11111111" w14:textId="77777777">${props && `<w:pPr>${props}</w:pPr>`}${runs}</w:p>${prior}`
const inserted = (props: string, runs: string, newProps: string, added = "<w:r><w:t>Added</w:t></w:r>", tail = prior) =>
  `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr>${props}<w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr>${runs}</w:p><w:p w14:paraId="11111111" w14:textId="12345678">${newProps && `<w:pPr>${newProps}</w:pPr>`}<w:ins w:id="2" ${stamp}>${added}</w:ins></w:p>${tail}`
const change = (old: string, id = "3", at = "2026-01-01T00:00:01Z", author = "Reviewer") =>
  `<w:pPrChange w:id="${id}" w:author="${author}" w:date="${at}"><w:pPr>${old}</w:pPr></w:pPrChange>`
const verifyStyled = (props: string, runs: string, xml: string, restyled: boolean) =>
  verifyNativeWordChatParagraphInsertion(
    projectNativeWordChatBody(flat(styled(props, runs)), beforeInv),
    projectNativeWordChatBody(flat(xml), afterInv),
    target,
    "Added",
    restyled,
  )
const plain = "<w:r><w:t>Original</w:t></w:r>"

test("a requested style is one tracked property change whose old properties are the anchor's", () => {
  const jc = '<w:jc w:val="both"/>',
    h2 = '<w:pStyle w:val="Heading2"/>'
  expect(verifyStyled(jc, plain, inserted(jc, plain, h2 + jc + change(jc)), true)).toEqual(["Reviewer"])
  for (const [xml, restyled] of [
    [inserted(jc, plain, h2 + jc), true], // restyled without a tracked change
    [inserted(jc, plain, h2 + jc), false],
    [inserted(jc, plain, h2 + jc + change("")), true], // old properties are not the anchor's
    [inserted(jc, plain, h2 + change(jc)), true], // a direct property dropped
    [inserted(jc, plain, jc + change(jc)), true], // no style change recorded
    [inserted(jc, plain, h2 + jc + change(jc, "2")), true], // reused revision id
    [inserted(jc, plain, h2 + jc + change(jc, "3", "2026-01-01T00:05:00Z")), true],
    [inserted(jc, plain, h2 + jc + change(jc, "3", undefined, "Other")), true],
    [inserted(jc, "<w:r><w:rPr><w:b/></w:rPr><w:t>Original</w:t></w:r>", h2 + jc + change(jc)), true], // anchor text altered
    [inserted(jc, plain, h2 + jc + change(jc), undefined, prior.replace("<w:r>", '<w:pPr><w:jc w:val="right"/></w:pPr><w:r>')), true],
    [inserted(jc, plain, h2 + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' + jc + change(jc)), true],
    [inserted(jc, plain, jc), true], // style requested but unchanged
  ] as const)
    expect(() => verifyStyled(jc, plain, xml, restyled)).toThrow("could not be verified")
})

test("Word's editing-session ids on the restyled paragraph are metadata, other attributes are not", () => {
  // Shape captured from Word for Mac: a plain anchor followed by a new Heading 2.
  const h2 = '<w:pStyle w:val="Heading2"/>' + change(""),
    word = inserted("", plain, h2).replace(
      'w14:paraId="11111111" w14:textId="12345678"',
      'w14:paraId="11111111" w14:textId="12345678" w:rsidR="00BB2E32" w:rsidRDefault="009C6BC0" w:rsidP="009C6BC0"',
    )
  expect(verifyStyled("", plain, word, true)).toEqual(["Reviewer"])
  expect(() => verifyStyled("", plain, word.replace('w:rsidP="009C6BC0"', 'w:rsidP="009C6BC0" w:other="1"'), true)).toThrow(
    "could not be verified",
  )
})

test("a heading anchor may be followed by an inheriting or Normal paragraph", () => {
  const h1 = '<w:pStyle w:val="Heading1"/><w:keepNext/>'
  expect(verifyStyled(h1, plain, inserted(h1, plain, h1), false)).toEqual(["Reviewer"])
  expect(verifyStyled(h1, plain, inserted(h1, plain, "<w:keepNext/>" + change(h1)), true)).toEqual(["Reviewer"])
  expect(verifyStyled(h1, plain, inserted(h1, plain, '<w:pStyle w:val="Normal"/><w:keepNext/>' + change(h1)), true)).toEqual([
    "Reviewer",
  ])
  // An untracked style change would survive rejecting the insertion.
  expect(() => verifyStyled(h1, plain, inserted(h1, plain, "<w:keepNext/>"), false)).toThrow()
  expect(() => verifyStyled(h1, plain, inserted(h1, plain, ""), false)).toThrow()
  expect(() => verifyStyled(h1, plain, inserted("<w:keepNext/>", plain, h1), false)).toThrow()
})

test("a multi-run formatted anchor accepts only inherited character properties on the new text", () => {
  const runs = '<w:r><w:t>Orig</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">inal</w:t></w:r>',
    bold = "<w:r><w:rPr><w:b/></w:rPr><w:t>Added</w:t></w:r>"
  expect(verifyStyled("", runs, inserted("", runs, "", bold), false)).toEqual(["Reviewer"])
  expect(verifyStyled("", runs, inserted("", runs, ""), false)).toEqual(["Reviewer"])
  expect(() => verifyStyled("", runs, inserted("", runs, "", bold.replace("<w:b/>", "<w:i/>")), false)).toThrow()
  expect(() => verifyStyled("", runs, inserted("", runs.replace("<w:b/>", "<w:i/>"), "", bold), false)).toThrow()
  const source = (runs: string, props = "") => projectNativeWordChatBody(flat(styled(props, runs)), beforeInv)
  expect(assertNativeWordChatParagraphInsertionTarget(source(runs, "<w:keepNext/>"), target)).toBe(0)
  for (const [refused, props] of [
    ['<w:bookmarkStart w:id="9" w:name="x"/><w:r><w:t>Original</w:t></w:r><w:bookmarkEnd w:id="9"/>', ""],
    ['<w:hyperlink w:anchor="x"><w:r><w:t>Original</w:t></w:r></w:hyperlink>', ""],
    ['<w:fldSimple w:instr="PAGE"><w:r><w:t>Original</w:t></w:r></w:fldSimple>', ""],
    ['<w:commentRangeStart w:id="1"/><w:r><w:t>Original</w:t></w:r><w:commentRangeEnd w:id="1"/>', ""],
    [`<w:ins w:id="4" ${stamp}><w:r><w:t>Original</w:t></w:r></w:ins>`, ""],
    [plain, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'],
  ] as const)
    expect(() => assertNativeWordChatParagraphInsertionTarget(source(refused, props), target)).toThrow("could not be verified")
})
