import { describe, expect, test } from "bun:test"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { normalizeNativeWordChatCommentIds as normalizeCommentIds } from "./word-chat-native-comments"
import { nativeWordChatDecimalList } from "./word-chat-native-list-context"
import {
  assertNativeWordChatListInsertionTarget,
  verifyNativeWordChatListInsertion,
} from "./word-chat-native-list-insert"
import {
  assertNativeWordChatParagraphDeletionTarget,
  verifyNativeWordChatParagraphDeletion,
} from "./word-chat-native-paragraph-delete"
import {
  assertNativeWordChatParagraphInsertionTarget,
  verifyNativeWordChatParagraphInsertion,
} from "./word-chat-native-paragraph-insert"
import {
  assertNativeWordChatParagraphMergeTarget,
  verifyNativeWordChatParagraphMerge,
} from "./word-chat-native-paragraph-merge"
import {
  assertNativeWordChatParagraphSplitTarget,
  verifyNativeWordChatParagraphSplit,
} from "./word-chat-native-paragraph-split"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"

describe("native comments and bookmarks", () => {
  const input = (id: string, content = "Existing comment", author = "Reviewer", rsid = "00112233", extra = "") =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>Anchor</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>${extra}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:p w:rsidR="${rsid}" w:rsidRDefault="${rsid}"><w:r><w:t>${content}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part></pkg:package>`

  test("comment IDs map only with exact matching anchors and comment payload", () => {
    const before = input("1"), after = input("2", "Existing comment", "Reviewer", "AABBCCDD"), out = normalizeCommentIds(before, after)
    expect(out).toMatch(/<w:comment\b[^>]*w:id="1"/)
    expect(out).toContain('w:commentRangeStart w:id="1"')
    expect(out).toContain('w:rsidR="00112233"')
    for (const drift of [
      input("2", "Changed"),
      input("2", "Existing comment", "Other"),
      after.replace("<w:commentRangeStart", "<w:r><w:t>X</w:t></w:r><w:commentRangeStart"),
      after.replace('commentReference w:id="2"', 'commentReference w:id="3"'),
      after.replace('w:rsidR="AABBCCDD"', 'w:rsidR="invalid"'),
      after.replace('<w:comment w:id="2"', '<w:comment w:id="2" w:initials="X"'),
      after.replace("</w:comments>", '<w:comment w:id="2"/></w:comments>'),
    ]) expect(() => normalizeCommentIds(before, drift)).toThrow()
    const outside = after.replace(
      "</pkg:package>",
      `<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData><w:hdr><w:commentReference w:id="2"/></w:hdr></pkg:xmlData></pkg:part></pkg:package>`,
    )
    expect(() => normalizeCommentIds(before, outside)).toThrow()
  })

  test("comments without optional rsid attributes retain their payload and anchors", () => {
    const without = (xml: string) => xml.replaceAll(' w:rsidR="00112233"', "").replaceAll(' w:rsidRDefault="00112233"', "")
    const before = without(input("1")), after = without(input("2"))
    expect(normalizeCommentIds(before, after)).toContain('w:commentRangeStart w:id="1"')
    expect(() => normalizeCommentIds(before, after.replace("Existing comment", "Changed comment"))).toThrow()
  })

  test("bookmark remaps preserve name and actual character range across run splits", () => {
    const base = input("1")
      .replace('<w:commentRangeStart w:id="1"/>', '<w:bookmarkStart w:id="0" w:name="Existing"/><w:commentRangeStart w:id="1"/>')
      .replace('<w:commentRangeEnd w:id="1"/>', '<w:commentRangeEnd w:id="1"/><w:bookmarkEnd w:id="0"/>')
    const revised = base
      .replaceAll('w:id="1"', 'w:id="2"')
      .replaceAll('w:id="0"', 'w:id="3"')
      .replace("<w:r><w:t>Anchor</w:t></w:r>", "<w:r><w:t>Anc</w:t></w:r><w:r><w:t>hor</w:t></w:r>")
    const output = normalizeCommentIds(base, revised)
    expect(output).toContain('w:bookmarkStart w:name="Existing" w:id="0"')
    expect(output).toContain('w:bookmarkEnd w:id="0"')
    for (const bad of [
      revised.replace('w:name="Existing"', 'w:name="Other"'),
      revised.replace('<w:bookmarkEnd w:id="3"/>', '<w:bookmarkEnd w:id="8"/>'),
      revised.replace('<w:bookmarkEnd w:id="3"/>', '<w:bookmarkEnd w:id="3"/><w:bookmarkEnd w:id="3"/>'),
      revised.replace('<w:bookmarkEnd w:id="3"/>', '<w:r><w:t>X</w:t></w:r><w:bookmarkEnd w:id="3"/>'),
    ]) expect(() => normalizeCommentIds(base, bad)).toThrow()
  })
})

describe("native list context", () => {
  const p = (text: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  const body = p("One") + p("Two")
  const numbering = '<w:numbering><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="7"/></w:num></w:numbering>'
  const flat = (b = body, n = numbering) =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${b}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData>${n}</pkg:xmlData></pkg:part></pkg:package>`
  const inv = [{ id: "a", text: "One", tableNestingLevel: 0 }, { id: "b", text: "Two", tableNestingLevel: 0 }]
  const capture = { id: 17, format: "decimal" as const, members: inv.map((item, i) => ({ id: item.id, text: item.text, level: 0, label: `${i + 1}.` })) }

  test("native decimal binding corroborates real definitions and rejects typed or partial membership", () => {
    const check = (b = body, n = numbering, c = capture) => nativeWordChatDecimalList(projectNativeWordChatBody(flat(b, n), inv), c)
    expect(check()).toBe(true)
    expect(check(body.replaceAll('<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>', ""))).toBe(false)
    expect(check(body.replace('<w:numId w:val="5"/>', '<w:numId w:val="6"/>'))).toBe(false)
    expect(check(body, numbering.replace('w:val="decimal"', 'w:val="lowerLetter"'))).toBe(false)
    expect(check(body, numbering.replace('<w:abstractNumId w:val="7"/>', '<w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>'))).toBe(false)
    expect(check(body, numbering.replace('w:val="%1."', 'w:val="Article %1"'))).toBe(false)
    expect(check(body, numbering.replace("</w:abstractNum>", '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum>'))).toBe(true)
    expect(check(body, numbering.replace("</w:abstractNum>", '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'))).toBe(false)
    expect(check(body, numbering, { ...capture, members: [capture.members[0]!] })).toBe(false)
  })

  test("pending inserted mark and text may remain while list definition binding stays explicit", () => {
    const after = body.replace("</w:numPr></w:pPr>", '</w:numPr><w:rPr><w:ins w:id="1" w:author="Reviewer"/></w:rPr></w:pPr>') +
      p("Three").replace("<w:r><w:t>Three</w:t></w:r>", '<w:ins w:id="2" w:author="Reviewer"><w:r><w:t>Three</w:t></w:r></w:ins>')
    const inventory = [...inv, { id: "c", text: "Three", tableNestingLevel: 0 }]
    const c = { ...capture, members: inventory.map((item, i) => ({ id: item.id, text: item.text, level: 0, label: `${i + 1}.` })) }
    expect(nativeWordChatDecimalList(projectNativeWordChatBody(flat(after), inventory), c)).toBe(true)
  })

  test("native capture binds each observed level to the actual list definition without label guesses", () => {
    const nestedBody = p("One") + p("Two").replace('w:ilvl w:val="0"', 'w:ilvl w:val="2"')
    const nestedNumbering = numbering.replace("</w:abstractNum>", '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3)"/></w:lvl></w:abstractNum>')
    const native = { id: 17, format: "native" as const, members: [{ id: "a", text: "One", level: 0, label: "1.", siblingIndex: 0 }, { id: "b", text: "Two", level: 2, label: "i)", siblingIndex: 0 }] }
    const check = (b = nestedBody, n = nestedNumbering, c = native) => nativeWordChatDecimalList(projectNativeWordChatBody(flat(b, n), inv), c)
    expect(check()).toBe(true)
    expect(check(nestedBody, nestedNumbering, { ...native, members: [native.members[0]!, { ...native.members[1]!, label: "arbitrary Word label" }] })).toBe(true)
    expect(check(nestedBody, nestedNumbering.replace('w:ilvl="2"', 'w:ilvl="3"'))).toBe(false)
    expect(check(nestedBody, nestedNumbering, { ...native, members: [native.members[0]!, { ...native.members[1]!, level: 1 }] })).toBe(false)
    expect(check(nestedBody, nestedNumbering, { ...native, members: [native.members[0]!] })).toBe(false)
    expect(check(nestedBody, nestedNumbering.replace('<w:abstractNumId w:val="7"/>', '<w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="2"/>'))).toBe(false)
  })
})

describe("native list insertion", () => {
  const flat = (body: string, style = "Normal") =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml" pkg:contentType="style"><pkg:xmlData><w:styles><w:style w:styleId="${style}"/></w:styles></pkg:xmlData></pkg:part></pkg:package>`
  const stamp = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
  const old =
    '<w:p w14:paraId="11111111" w14:textId="77777777"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Original</w:t></w:r></w:p><w:p w14:paraId="44444444"><w:ins w:id="5" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p>'
  const next = `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr><w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr><w:r><w:t>Original</w:t></w:r></w:p><w:p w14:paraId="11111111" w14:textId="12345678"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:ins w:id="2" ${stamp}><w:r><w:t>Added</w:t></w:r></w:ins></w:p><w:p w14:paraId="44444444"><w:ins w:id="7" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p>`
  const beforeInv = [
    { id: "orig", text: "Original", tableNestingLevel: 0 },
    { id: "prior", text: "Prior", tableNestingLevel: 0 },
  ]
  const afterInv = [
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
    verifyNativeWordChatListInsertion(
      projectNativeWordChatBody(flat(old), beforeInv),
      projectNativeWordChatBody(xml, inv),
      target,
      "Added",
    )

  test("tracked list insertion validates transferred identity, marks and protected body", () => {
    const source = projectNativeWordChatBody(flat(old), beforeInv)
    expect(assertNativeWordChatListInsertionTarget(source, target)).toBe(0)
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
      flat(next.replace('w:numId w:val="5"', 'w:numId w:val="6"')),
      flat(next.replace('w:ilvl w:val="0"', 'w:ilvl w:val="1"')),
      flat(next.replace("<w:numPr>", '<w:pStyle w:val="ListNumber"/><w:numPr>')),
      flat(next.replace(">Prior<", "> Prior<")),
    ]) expect(() => check(bad)).toThrow()
    expect(() => assertNativeWordChatListInsertionTarget(source, { ...target, endOffset: 4 })).toThrow()
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
    const a = wrap(anchored(old, "8", "10"), "8"), b = wrap(anchored(next, "9", "11"), "9")
    const source = projectNativeWordChatBody(a, beforeInv), destination = projectNativeWordChatBody(b, afterInv)
    expect(verifyNativeWordChatListInsertion(source, destination, target, "Added")).toEqual(["Reviewer"])
    expect(() =>
      verifyNativeWordChatListInsertion(
        source,
        projectNativeWordChatBody(b.replace(">Comment<", ">Tampered<"), afterInv),
        target,
        "Added",
      ),
    ).toThrow()
  })

  test("only projection-qualified unbound terminal carrier may change physical ID", () => {
    const a = flat(old + '<w:p w14:paraId="AAAAAAAA"/><w:sectPr/>')
    const b = flat(next + '<w:p w14:paraId="BBBBBBBB"/><w:sectPr/>')
    expect(
      verifyNativeWordChatListInsertion(
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

  test("preflight requires bounded direct list numbering and plain target", () => {
    expect(assertNativeWordChatListInsertionTarget(projectNativeWordChatBody(flat(old.replace('w:ilvl w:val="0"', 'w:ilvl w:val="2"')), beforeInv), target)).toBe(0)
    expect(() => assertNativeWordChatListInsertionTarget(projectNativeWordChatBody(flat(old.replace('w:ilvl w:val="0"', 'w:ilvl w:val="9"')), beforeInv), target)).toThrow()
    for (const source of [
      old.replace('w:numId w:val="5"', 'w:numId w:val="0"'),
      old.replace("<w:numPr>", '<w:pStyle w:val="ListNumber"/><w:numPr>'),
      old.replace("<w:r><w:t>Original", "<w:r><w:rPr><w:b/></w:rPr><w:t>Original"),
    ]) expect(() => assertNativeWordChatListInsertionTarget(projectNativeWordChatBody(flat(source), beforeInv), target)).toThrow()
  })

  test("before-final insertion keeps the anchor identity and verifies both tracked marks", () => {
    const inserted = `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr><w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr><w:ins w:id="2" ${stamp}><w:r><w:t>Added</w:t></w:r></w:ins></w:p>`
    const anchor = old.replace('w14:textId="77777777"', 'w14:textId="ABCDEF12"').replace('w:id="5"', 'w:id="7"')
    const inventory = [{ id: "new", text: "Added", tableNestingLevel: 0 }, ...beforeInv]
    const source = projectNativeWordChatBody(flat(old), beforeInv)
    const checkBefore = (xml: string, inv = inventory) =>
      verifyNativeWordChatListInsertion(source, projectNativeWordChatBody(xml, inv), target, "Added", "before_final")
    expect(checkBefore(flat(inserted + anchor))).toEqual(["Reviewer"])
    expect(() => checkBefore(flat(inserted.replace('w:id="2"', 'w:id="1"') + anchor))).toThrow()
    expect(() => checkBefore(flat(inserted + anchor.replace(">Prior<", ">Changed<")))).toThrow()
    expect(() => checkBefore(flat(inserted + anchor), [{ ...inventory[0]!, id: "orig" }, { ...inventory[1]!, id: "new" }, inventory[2]!])).toThrow()
    expect(() => verifyNativeWordChatListInsertion(source, projectNativeWordChatBody(flat(inserted + anchor), inventory), target, "Added")).toThrow()
  })

  test("whole-story protected ranges must not cross the list target", () => {
    const markerPairs = [
      ['<w:commentRangeStart w:id="8"/>', '<w:commentRangeEnd w:id="8"/>'],
      ['<w:bookmarkStart w:id="8" w:name="Keep"/>', '<w:bookmarkEnd w:id="8"/>'],
      ['<w:permStart w:id="8" w:ed="everyone"/>', '<w:permEnd w:id="8"/>'],
      ['<w:moveFromRangeStart w:id="8" w:name="Move"/>', '<w:moveFromRangeEnd w:id="8"/>'],
      ['<w:moveToRangeStart w:id="8" w:name="Move"/>', '<w:moveToRangeEnd w:id="8"/>'],
      ['<w:customXmlInsRangeStart w:id="8"/>', '<w:customXmlInsRangeEnd w:id="8"/>'],
      ['<w:r><w:fldChar w:fldCharType="begin"/></w:r>', '<w:r><w:fldChar w:fldCharType="end"/></w:r>'],
    ]
    const pre = (inside: string) => `<w:p w14:paraId="AAAAAAAA">${inside}<w:r><w:t>Before</w:t></w:r></w:p>`
    const post = (inside: string) => `<w:p w14:paraId="BBBBBBBB">${inside}<w:r><w:t>After</w:t></w:r></w:p>`
    const inv = [{ id: "pre", text: "Before", tableNestingLevel: 0 }, ...beforeInv, { id: "post", text: "After", tableNestingLevel: 0 }]
    for (const [start, end] of markerPairs) {
      for (const position of ["end", "before_final"] as const) {
        const disjoint = projectNativeWordChatBody(flat(pre(start + end) + old + post("")), inv)
        const crossing = projectNativeWordChatBody(flat(pre(start) + old + post(end)), inv)
        expect(assertNativeWordChatListInsertionTarget(disjoint, target, position)).toBe(1)
        expect(() => assertNativeWordChatListInsertionTarget(crossing, target, position)).toThrow()
      }
    }
  })

  test("End rebinds only a valid regenerated retained-anchor textId", () => {
    const changed = next.replace('w14:paraId="22222222" w14:textId="77777777"', 'w14:paraId="22222222" w14:textId="ABCDEF12"')
    expect(check(flat(changed))).toEqual(["Reviewer"])
    for (const bad of [
      changed.replace('w14:textId="ABCDEF12"', 'w14:textId="invalid"'),
      changed.replace(' w14:textId="ABCDEF12"', ""),
      changed.replace(">Original<", ">Changed<"),
      changed.replace("<w:numPr>", '<w:pStyle w:val="Other"/><w:numPr>'),
      changed.replace('w:author="Earlier"', 'w:author="Other"'),
    ]) expect(() => check(flat(bad))).toThrow()
    const priorWithTextId = old.replace('w14:paraId="44444444"', 'w14:paraId="44444444" w14:textId="11111111"')
    const changedPrior = changed.replace('w14:paraId="44444444"', 'w14:paraId="44444444" w14:textId="22222222"')
    expect(() =>
      verifyNativeWordChatListInsertion(
        projectNativeWordChatBody(flat(priorWithTextId), beforeInv),
        projectNativeWordChatBody(flat(changedPrior), afterInv),
        target,
        "Added",
      ),
    ).toThrow()
    expect(() => check(flat(changed, "Other"))).toThrow()
  })
})

describe("native paragraph deletion", () => {
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
    ]) expect(() => check(bad)).toThrow()
    expect(() => check(good, afterInv, "Changed")).toThrow()
    expect(() => check(good, [{ ...afterInv[0]!, id: "changed" }, afterInv[1]!, afterInv[2]!])).toThrow()
    expect(() => assertNativeWordChatParagraphDeletionTarget(before, { ...target, endOffset: 4 })).toThrow()
  })

  test("a doc.haus pointer bookmark on the deleted paragraph survives Word moving its end into the deletion", () => {
    const ptrName = "_dh_p0123456789ab", runs = "<w:r><w:t>Original</w:t></w:r>"
    const oldPointer = targetOld.replace(runs, `<w:bookmarkStart w:id="11" w:name="${ptrName}"/>${runs}<w:bookmarkEnd w:id="11"/>`)
    const old = projectNativeWordChatBody(flat(oldPointer + rest), inv)
    const run = (body: string) => verifyNativeWordChatParagraphDeletion(old, projectNativeWordChatBody(flat(body), afterInv), target)
    const deleted = targetAfter.replace('<w:del w:id="3"', `<w:bookmarkStart w:id="9" w:name="${ptrName}"/><w:del w:id="3"`).replace("</w:del></w:p>", '<w:bookmarkEnd w:id="9"/></w:del></w:p>')
    const good = deleted + rest.replace('w:id="5" w:author="Earlier"', 'w:id="7" w:author="Earlier"')
    expect(assertNativeWordChatParagraphDeletionTarget(old, target)).toBe(0)
    expect(run(good)).toEqual(["Reviewer"])
    expect(run(good.replace('<w:bookmarkEnd w:id="9"/></w:del>', '</w:del><w:bookmarkEnd w:id="9"/>'))).toEqual(["Reviewer"])
    for (const bad of [
      good.replace(`w:name="${ptrName}"`, 'w:name="_dh_p00000000ffff"'),
      good.replace('<w:bookmarkEnd w:id="9"/></w:del>', "</w:del>"),
      good.replace('<w:bookmarkEnd w:id="9"/>', '<w:bookmarkEnd w:id="8"/>'),
      good.replace(`<w:bookmarkStart w:id="9" w:name="${ptrName}"/>`, ""),
    ]) expect(() => run(bad)).toThrow()
    for (const bad of [
      targetOld.replace(runs, `<w:bookmarkStart w:id="1" w:name="Keep"/>${runs}<w:bookmarkEnd w:id="1"/>`),
      targetOld.replace(runs, `<w:bookmarkStart w:id="11" w:name="${ptrName}"/>${runs}`),
    ]) expect(() => assertNativeWordChatParagraphDeletionTarget(projectNativeWordChatBody(flat(bad + rest), inv), target)).toThrow()
  })

  test("a following ordinary paragraph can retain its comment and bookmark graph", () => {
    const following = '<w:p w14:paraId="22222222"><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:r><w:t>Next</w:t></w:r><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>'
    const old = projectNativeWordChatBody(flat(targetOld + following), inv.slice(0, 2))
    const checkFollowing = (body: string, comment = "Keep") =>
      verifyNativeWordChatParagraphDeletion(old, projectNativeWordChatBody(flat(body, comment), afterInv.slice(0, 2)), target)
    const next = targetAfter + following
    expect(assertNativeWordChatParagraphDeletionTarget(old, target)).toBe(0)
    expect(checkFollowing(next)).toEqual(["Reviewer"])
    expect(() => checkFollowing(next.replace('w:name="Keep"', 'w:name="Changed"'))).toThrow()
    expect(() => checkFollowing(next.replace('<w:commentRangeEnd w:id="3"/>', ""))).toThrow()
    expect(() => checkFollowing(next, "Changed")).toThrow()
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
    const text = "Termination of Agreement"
    const headingInv = [{ ...inv[0]!, id: "h", text }, inv[1]!, inv[2]!]
    const deletedInv = [{ ...headingInv[0]!, text: "" }, inv[1]!, inv[2]!]
    const heads = { ...target, text, paragraphs: [{ id: "h", text }], endOffset: text.length }
    const old = projectNativeWordChatBody(flat(heading + rest), headingInv)
    const checkHeading = (body: string) => verifyNativeWordChatParagraphDeletion(old, after(body, deletedInv), heads)
    expect(assertNativeWordChatParagraphDeletionTarget(old, heads)).toBe(0)
    expect(checkHeading(one + rest)).toEqual(["Reviewer"])
    expect(checkHeading(each + rest)).toEqual(["Reviewer"])
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
    ]) expect(() => checkHeading(bad)).toThrow()
    for (const unsupported of [
      heading.replace("<w:r>", '<w:commentRangeStart w:id="3"/><w:r>'),
      heading.replace("<w:r>", '<w:bookmarkStart w:id="8" w:name="_Toc1"/><w:bookmarkEnd w:id="8"/><w:r>'),
      heading.replace(/<w:r>(<w:t xml:space="preserve"> of Agreement<\/w:t>)<\/w:r>/, `<w:ins w:id="9" ${stamp}><w:r>$1</w:r></w:ins>`),
      heading.replace("</w:numPr>", `</w:numPr><w:rPr><w:ins w:id="9" ${stamp}/></w:rPr>`),
      heading.replace("<w:b/>", `<w:b/><w:rPrChange w:id="9" ${stamp}><w:rPr/></w:rPrChange>`),
    ]) expect(() => assertNativeWordChatParagraphDeletionTarget(projectNativeWordChatBody(flat(unsupported + rest), headingInv), heads)).toThrow()
  })
})

describe("native paragraph insertion", () => {
  const flat = (body: string, style = "Normal") =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml" pkg:contentType="style"><pkg:xmlData><w:styles><w:style w:styleId="${style}"/></w:styles></pkg:xmlData></pkg:part></pkg:package>`
  const stamp = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
  const old =
    '<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:t>Original</w:t></w:r></w:p><w:p w14:paraId="44444444"><w:ins w:id="5" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p>'
  const next = `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr><w:r><w:t>Original</w:t></w:r></w:p><w:p w14:paraId="11111111" w14:textId="12345678"><w:ins w:id="2" ${stamp}><w:r><w:t>Added</w:t></w:r></w:ins></w:p><w:p w14:paraId="44444444"><w:ins w:id="7" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins></w:p>`
  const beforeInv = [
    { id: "orig", text: "Original", tableNestingLevel: 0 },
    { id: "prior", text: "Prior", tableNestingLevel: 0 },
  ]
  const afterInv = [
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
    ]) expect(() => check(bad)).toThrow()
    expect(() => assertNativeWordChatParagraphInsertionTarget(source, { ...target, endOffset: 4 })).toThrow()
    expect(() =>
      check(flat(next), [{ ...afterInv[0]!, id: "orig" }, { ...afterInv[1]!, id: "new" }, afterInv[2]!]),
    ).toThrow()
  })

  test("a doc.haus pointer bookmark on the anchor is accepted, and nothing wider is", () => {
    const ptr = (id: string, name = "_dh_p0123456789ab") => (runs: string) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${runs}<w:bookmarkEnd w:id="${id}"/>`
    const originalRuns = "<w:r><w:t>Original</w:t></w:r>"
    const oldPointer = old.replace(originalRuns, ptr("1")(originalRuns))
    const nextPointer = next.replace(originalRuns, ptr("9")(originalRuns))
    const run = (beforeBody: string, afterBody: string, t = target) =>
      verifyNativeWordChatParagraphInsertion(projectNativeWordChatBody(flat(beforeBody), beforeInv), projectNativeWordChatBody(flat(afterBody), afterInv), t, "Added")
    expect(assertNativeWordChatParagraphInsertionTarget(projectNativeWordChatBody(flat(oldPointer), beforeInv), target)).toBe(0)
    expect(run(oldPointer, nextPointer)).toEqual(["Reviewer"])
    for (const bad of [
      old.replace(originalRuns, ptr("1", "Keep")(originalRuns)),
      old.replace(originalRuns, ptr("1", "_dh_pXYZ")(originalRuns)),
      old.replace(originalRuns, `<w:bookmarkStart w:id="1" w:name="_dh_p0123456789ab"/>${originalRuns}`),
      old.replace(originalRuns, `${ptr("1")(originalRuns)}<w:bookmarkStart w:id="2" w:name="_dh_p0123456789cd"/><w:bookmarkEnd w:id="2"/>`),
      old.replace(originalRuns, `<w:bookmarkStart w:id="1" w:name="_dh_p0123456789ab"/><w:bookmarkEnd w:id="1"/>${originalRuns}`),
    ]) expect(() => assertNativeWordChatParagraphInsertionTarget(projectNativeWordChatBody(flat(bad), beforeInv), target)).toThrow()
    expect(() => run(oldPointer, next.replace("<w:t>Added</w:t></w:r>", '<w:t>Added</w:t></w:r><w:bookmarkEnd w:id="9"/>').replace(originalRuns, ptr("9")(originalRuns)))).toThrow()
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
    const a = wrap(anchored(old, "8", "10"), "8"), b = wrap(anchored(next, "9", "11"), "9")
    const source = projectNativeWordChatBody(a, beforeInv), destination = projectNativeWordChatBody(b, afterInv)
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
    const a = flat(old + '<w:p w14:paraId="AAAAAAAA"/><w:sectPr/>')
    const b = flat(next + '<w:p w14:paraId="BBBBBBBB"/><w:sectPr/>')
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

  const prior = '<w:p w14:paraId="44444444"><w:r><w:t>Prior</w:t></w:r></w:p>'
  const styled = (props: string, runs: string) =>
    `<w:p w14:paraId="11111111" w14:textId="77777777">${props && `<w:pPr>${props}</w:pPr>`}${runs}</w:p>${prior}`
  const inserted = (props: string, runs: string, newProps: string, added = "<w:r><w:t>Added</w:t></w:r>", tail = prior) =>
    `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr>${props}<w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr>${runs}</w:p><w:p w14:paraId="11111111" w14:textId="12345678">${newProps && `<w:pPr>${newProps}</w:pPr>`}<w:ins w:id="2" ${stamp}>${added}</w:ins></w:p>${tail}`
  const change = (oldProps: string, id = "3", at = "2026-01-01T00:00:01Z", author = "Reviewer") =>
    `<w:pPrChange w:id="${id}" w:author="${author}" w:date="${at}"><w:pPr>${oldProps}</w:pPr></w:pPrChange>`
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
    const jc = '<w:jc w:val="both"/>', h2 = '<w:pStyle w:val="Heading2"/>'
    expect(verifyStyled(jc, plain, inserted(jc, plain, h2 + jc + change(jc)), true)).toEqual(["Reviewer"])
    for (const [xml, restyled] of [
      [inserted(jc, plain, h2 + jc), true],
      [inserted(jc, plain, h2 + jc), false],
      [inserted(jc, plain, h2 + jc + change("")), true],
      [inserted(jc, plain, h2 + change(jc)), true],
      [inserted(jc, plain, jc + change(jc)), true],
      [inserted(jc, plain, h2 + jc + change(jc, "2")), true],
      [inserted(jc, plain, h2 + jc + change(jc, "3", "2026-01-01T00:05:00Z")), true],
      [inserted(jc, plain, h2 + jc + change(jc, "3", undefined, "Other")), true],
      [inserted(jc, "<w:r><w:rPr><w:b/></w:rPr><w:t>Original</w:t></w:r>", h2 + jc + change(jc)), true],
      [inserted(jc, plain, h2 + jc + change(jc), undefined, prior.replace("<w:r>", '<w:pPr><w:jc w:val="right"/></w:pPr><w:r>')), true],
      [inserted(jc, plain, h2 + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' + jc + change(jc)), true],
      [inserted(jc, plain, jc), true],
    ] as const) expect(() => verifyStyled(jc, plain, xml, restyled)).toThrow("could not be verified")
  })

  test("Word's editing-session ids on the restyled paragraph are metadata, other attributes are not", () => {
    const h2 = '<w:pStyle w:val="Heading2"/>' + change("")
    const word = inserted("", plain, h2).replace(
      'w14:paraId="11111111" w14:textId="12345678"',
      'w14:paraId="11111111" w14:textId="12345678" w:rsidR="00BB2E32" w:rsidRDefault="009C6BC0" w:rsidP="009C6BC0"',
    )
    expect(verifyStyled("", plain, word, true)).toEqual(["Reviewer"])
    expect(() => verifyStyled("", plain, word.replace('w:rsidP="009C6BC0"', 'w:rsidP="009C6BC0" w:other="1"'), true)).toThrow("could not be verified")
  })

  test("a heading anchor may be followed by an inheriting or Normal paragraph", () => {
    const h1 = '<w:pStyle w:val="Heading1"/><w:keepNext/>'
    expect(verifyStyled(h1, plain, inserted(h1, plain, h1), false)).toEqual(["Reviewer"])
    expect(verifyStyled(h1, plain, inserted(h1, plain, "<w:keepNext/>" + change(h1)), true)).toEqual(["Reviewer"])
    expect(verifyStyled(h1, plain, inserted(h1, plain, '<w:pStyle w:val="Normal"/><w:keepNext/>' + change(h1)), true)).toEqual(["Reviewer"])
    expect(() => verifyStyled(h1, plain, inserted(h1, plain, "<w:keepNext/>"), false)).toThrow()
    expect(() => verifyStyled(h1, plain, inserted(h1, plain, ""), false)).toThrow()
    expect(() => verifyStyled(h1, plain, inserted("<w:keepNext/>", plain, h1), false)).toThrow()
  })

  test("a multi-run formatted anchor accepts only inherited character properties on the new text", () => {
    const runs = '<w:r><w:t>Orig</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">inal</w:t></w:r>'
    const bold = "<w:r><w:rPr><w:b/></w:rPr><w:t>Added</w:t></w:r>"
    expect(verifyStyled("", runs, inserted("", runs, "", bold), false)).toEqual(["Reviewer"])
    expect(verifyStyled("", runs, inserted("", runs, ""), false)).toEqual(["Reviewer"])
    expect(() => verifyStyled("", runs, inserted("", runs, "", bold.replace("<w:b/>", "<w:i/>")), false)).toThrow()
    expect(() => verifyStyled("", runs, inserted("", runs.replace("<w:b/>", "<w:i/>"), "", bold), false)).toThrow()
    const source = (r: string, props = "") => projectNativeWordChatBody(flat(styled(props, r)), beforeInv)
    expect(assertNativeWordChatParagraphInsertionTarget(source(runs, "<w:keepNext/>"), target)).toBe(0)
    for (const [refused, props] of [
      ['<w:bookmarkStart w:id="9" w:name="x"/><w:r><w:t>Original</w:t></w:r><w:bookmarkEnd w:id="9"/>', ""],
      ['<w:hyperlink w:anchor="x"><w:r><w:t>Original</w:t></w:r></w:hyperlink>', ""],
      ['<w:fldSimple w:instr="PAGE"><w:r><w:t>Original</w:t></w:r></w:fldSimple>', ""],
      ['<w:commentRangeStart w:id="1"/><w:r><w:t>Original</w:t></w:r><w:commentRangeEnd w:id="1"/>', ""],
      [`<w:ins w:id="4" ${stamp}><w:r><w:t>Original</w:t></w:r></w:ins>`, ""],
      [plain, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'],
    ] as const) expect(() => assertNativeWordChatParagraphInsertionTarget(source(refused, props), target)).toThrow("could not be verified")
  })
})

describe("native paragraph merge", () => {
  const flat = (body: string, comment = "Keep", style = "same") =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="3" w:author="Earlier"><w:p><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles w:val="${style}"/></pkg:xmlData></pkg:part></pkg:package>`
  const left = '<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:rPr><w:b/></w:rPr><w:t>Left</w:t></w:r></w:p>'
  const right = '<w:p w14:paraId="22222222" w14:textId="77777777"><w:r><w:t>Right</w:t></w:r></w:p>'
  const mark = (id: string, author = "Reviewer") => `w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"`
  const changedLeft = `<w:p w14:paraId="11111111" w14:textId="12345678" w:rsidDel="00112233"><w:pPr><w:rPr><w:del ${mark("2")}/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Left</w:t></w:r></w:p>`
  const changedRight = `<w:p w14:paraId="22222222" w14:textId="12345678"><w:ins ${mark("3")}><w:r><w:t xml:space="preserve"> </w:t></w:r></w:ins><w:r><w:t>Right</w:t></w:r></w:p>`
  const prior = (id: string) =>
    `<w:p><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:ins w:id="${id}" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>`
  const inventory = [{ id: "a", text: "Left", tableNestingLevel: 0 }, { id: "b", text: "Right", tableNestingLevel: 0 }, { id: "c", text: "Prior", tableNestingLevel: 0 }]
  const afterInventory = [inventory[0]!, { ...inventory[1]!, text: " Right" }, inventory[2]!]
  const target = { id: "target", text: "Left\rRight", paragraphs: inventory.slice(0, 2), startOffset: 0, endOffset: 5, capturedAt: "now", saved: false }
  const before = projectNativeWordChatBody(flat(left + right + prior("2")), inventory)
  const after = (body: string, comment = "Keep", style = "same") => projectNativeWordChatBody(flat(body, comment, style), afterInventory)
  const good = changedLeft + changedRight + prior("4")

  test("merge proves the whole inverse and qualified prior review ID mapping", () => {
    expect(assertNativeWordChatParagraphMergeTarget(before, target)).toEqual([0, 1])
    expect(verifyNativeWordChatParagraphMerge(before, after(good), target)).toEqual(["Reviewer"])
    for (const bad of [
      good.replace(mark("3"), mark("02")),
      good.replace(mark("3"), mark("3", "Other")),
      good.replace("<w:b/>", "<w:i/>"),
      good.replace('w:author="Earlier"', 'w:author="Other"'),
      good.replace('w:name="Keep"', 'w:name="Changed"'),
      good.replace('w:id="4" w:author="Earlier"', 'w:id="3" w:author="Earlier"'),
    ]) expect(() => verifyNativeWordChatParagraphMerge(before, after(bad), target)).toThrow()
    expect(() => verifyNativeWordChatParagraphMerge(before, after(good, "Changed"), target)).toThrow()
    expect(() => verifyNativeWordChatParagraphMerge(before, after(good, "Keep", "altered"), target)).toThrow()
    expect(() => assertNativeWordChatParagraphMergeTarget(before, { ...target, endOffset: 4 })).toThrow()
  })

  test("crossing annotations, incompatible paragraph properties and old target revisions refuse", () => {
    const cases = [
      flat('<w:p><w:bookmarkStart w:id="8" w:name="cross"/><w:r><w:t>Prior</w:t></w:r></w:p>' + left + right + '<w:p><w:bookmarkEnd w:id="8"/><w:r><w:t>End</w:t></w:r></w:p>'),
      flat(left + right.replace("<w:r>", '<w:pPr><w:spacing w:after="20"/></w:pPr><w:r>')),
      flat(left.replace("<w:r>", '<w:ins w:id="9" w:author="X" w:date="2025-01-01"><w:r>').replace("</w:r></w:p>", "</w:r></w:ins></w:p>") + right),
    ]
    for (const xml of cases) {
      const inv = xml.includes("cross") ? [{ id: "z", text: "Prior", tableNestingLevel: 0 }, ...inventory.slice(0, 2), { id: "z2", text: "End", tableNestingLevel: 0 }] : inventory.slice(0, 2)
      expect(() => assertNativeWordChatParagraphMergeTarget(projectNativeWordChatBody(xml, inv), target)).toThrow()
    }
  })

  test("a new people part needs one unique internal relationship and preserves other relationships", () => {
    const rn = "http://schemas.openxmlformats.org/package/2006/relationships"
    const pn = "http://schemas.microsoft.com/office/word/2012/wordml"
    const people = "http://schemas.microsoft.com/office/2011/relationships/people"
    const theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    const part = (name: string, xml: string) => `<pkg:part pkg:name="${name}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`
    const rel = (id: string, type: string, to: string, extra = "") => `<Relationship Id="${id}" Type="${type}" Target="${to}"${extra}/>`
    const rels = (x: string) => part("/word/_rels/document.xml.rels", `<Relationships xmlns="${rn}">${x}</Relationships>`)
    const person = part("/word/people.xml", `<w15:people xmlns:w15="${pn}"><w15:person w15:author="Reviewer"><w15:presenceInfo w15:providerId="x" w15:userId="y"/></w15:person></w15:people>`)
    const withParts = (xml: string, x: string) => xml.replace("</pkg:package>", x + "</pkg:package>")
    const b = projectNativeWordChatBody(withParts(flat(left + right + prior("2")), rels(rel("r1", theme, "theme.xml"))), inventory)
    const candidate = (link: string) => projectNativeWordChatBody(withParts(flat(good), rels(rel("r2", theme, "theme.xml") + link) + person), afterInventory)
    const link = rel("r1", people, "people.xml")
    expect(verifyNativeWordChatParagraphMerge(b, candidate(link), target)).toEqual(["Reviewer"])
    for (const bad of [
      rel("r1", people, "other.xml"),
      rel("r1", people, "people.xml", ' TargetMode="External"'),
      link + rel("r3", people, "people.xml"),
      rel("r2", people, "people.xml"),
    ]) expect(() => verifyNativeWordChatParagraphMerge(b, candidate(bad), target)).toThrow()
  })
})

describe("native paragraph split", () => {
  const flat = (body: string, comment = "Keep", style = "same") =>
    `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="3" w:author="Earlier"><w:p><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles w:val="${style}"/></pkg:xmlData></pkg:part></pkg:package>`
  const source = '<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:t>One two</w:t></w:r></w:p>'
  const mark = (id: string, author = "Reviewer") => `w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"`
  const left = `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:rPr><w:ins ${mark("2")}/></w:rPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>`
  const right = `<w:p w14:paraId="11111111" w14:textId="12345678"><w:del ${mark("3")}><w:r w:rsidDel="00112233"><w:delText xml:space="preserve"> </w:delText></w:r></w:del><w:r><w:t>two</w:t></w:r></w:p>`
  const prior = (id: string) =>
    `<w:p><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:ins w:id="${id}" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>`
  const inv = [{ id: "a", text: "One two", tableNestingLevel: 0 }, { id: "b", text: "Prior", tableNestingLevel: 0 }]
  const afterInv = [{ id: "new", text: "One", tableNestingLevel: 0 }, { id: "a", text: "two", tableNestingLevel: 0 }, inv[1]!]
  const target = { id: "target", text: "One two", paragraphs: [inv[0]!], startOffset: 0, endOffset: 7, capturedAt: "now", saved: false }
  const before = projectNativeWordChatBody(flat(source + prior("2")), inv)
  const after = (body: string, comment = "Keep", style = "same", inventory = afterInv) =>
    projectNativeWordChatBody(flat(body, comment, style), inventory)
  const good = left + right + prior("4")

  test("split proves carriers, full inverse and old revision mapping", () => {
    expect(assertNativeWordChatParagraphSplitTarget(before, target)).toBe(0)
    expect(verifyNativeWordChatParagraphSplit(before, after(good), target, "One\rtwo")).toEqual(["Reviewer"])
    for (const bad of [
      good.replace(mark("3"), mark("02")),
      good.replace(mark("3"), mark("3", "Other")),
      good.replace("> </w:delText>", ">x</w:delText>"),
      good.replace("<w:r><w:t>two", "<w:r><w:rPr><w:b/></w:rPr><w:t>two"),
      good.replace('w:name="Keep"', 'w:name="Changed"'),
      good.replace('w:id="4" w:author="Earlier"', 'w:id="3" w:author="Earlier"'),
    ]) expect(() => verifyNativeWordChatParagraphSplit(before, after(bad), target, "One\rtwo")).toThrow()
    expect(() => verifyNativeWordChatParagraphSplit(before, after(good, "Changed"), target, "One\rtwo")).toThrow()
    expect(() => verifyNativeWordChatParagraphSplit(before, after(good, "Keep", "changed"), target, "One\rtwo")).toThrow()
    expect(() => verifyNativeWordChatParagraphSplit(before, after(good, "Keep", "same", [{ ...afterInv[0]!, id: "a" }, { ...afterInv[1]!, id: "new" }, afterInv[2]!]), target, "One\rtwo")).toThrow()
    expect(() => verifyNativeWordChatParagraphSplit(before, after(good), target, "One two")).toThrow()
  })

  test("target accepts plain multiple runs and refuses partial selection and crossing annotation", () => {
    expect(() => assertNativeWordChatParagraphSplitTarget(before, { ...target, endOffset: 4 })).toThrow()
    const multi = source.replace("<w:t>One two</w:t>", '<w:t>One</w:t></w:r><w:r><w:t xml:space="preserve"> two</w:t>')
    expect(assertNativeWordChatParagraphSplitTarget(projectNativeWordChatBody(flat(multi + prior("2")), inv), target)).toBe(0)
    const cross = '<w:p><w:bookmarkStart w:id="8" w:name="x"/><w:r><w:t>Before</w:t></w:r></w:p>' + source + '<w:p><w:bookmarkEnd w:id="8"/><w:r><w:t>After</w:t></w:r></w:p>'
    expect(() => assertNativeWordChatParagraphSplitTarget(projectNativeWordChatBody(flat(cross), [{ id: "x", text: "Before", tableNestingLevel: 0 }, inv[0]!, { id: "y", text: "After", tableNestingLevel: 0 }]), target)).toThrow()
  })

  test("mixed runs preserve the boundary mark and deleted space in the same cell", () => {
    const cell = (content: string, width = "100") => `<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="${width}"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`
    const src = '<w:p w14:paraId="11111111" w14:textId="77777777"><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>One</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> two</w:t></w:r></w:p>'
    const l = `<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:spacing w:after="120"/><w:rPr><w:ins ${mark("2")}/><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>One</w:t></w:r></w:p>`
    const r = `<w:p w14:paraId="11111111" w14:textId="12345678"><w:pPr><w:spacing w:after="120"/></w:pPr><w:del ${mark("3")}><w:r w:rsidDel="00112233"><w:rPr><w:i/></w:rPr><w:delText xml:space="preserve"> </w:delText></w:r></w:del><w:r><w:rPr><w:i/></w:rPr><w:t>two</w:t></w:r></w:p>`
    const original = [{ id: "a", text: "One two", tableNestingLevel: 1 }, { id: "b", text: "Prior", tableNestingLevel: 0 }]
    const changed = [{ id: "new", text: "One", tableNestingLevel: 1 }, { id: "a", text: "two", tableNestingLevel: 1 }, original[1]!]
    const b = projectNativeWordChatBody(flat(cell(src) + prior("4")), original)
    const capture = (xml: string) => projectNativeWordChatBody(flat(xml + prior("4")), changed)
    expect(verifyNativeWordChatParagraphSplit(b, capture(cell(l + r)), target, "One\rtwo")).toEqual(["Reviewer"])
    const bad = [
      cell(l.replace("<w:ins " + mark("2") + "/><w:i/>", "<w:ins " + mark("2") + "/><w:b/>") + r),
      cell(l + r.replace("<w:rPr><w:i/></w:rPr><w:delText", "<w:rPr><w:b/></w:rPr><w:delText")),
      cell(l + r.replace('w:after="120"', 'w:after="121"')),
      cell(l + r, "200"),
    ]
    for (const xml of bad) expect(() => verifyNativeWordChatParagraphSplit(b, capture(xml), target, "One\rtwo")).toThrow()
  })

  test("qualified run session IDs may change without masking formatting or foreign attributes", () => {
    const sourceMeta = source.replace("<w:r>", '<w:r w:rsidR="1234ABCD" w:rsidRPr="00112233">')
    const base = projectNativeWordChatBody(flat(sourceMeta + prior("2")), inv)
    const leftMeta = left.replace("<w:r><w:t>One", '<w:r w:rsidR="ABCDEF01" w:rsidRPr="AABBCCDD"><w:t>One')
    const rightMeta = right.replace('<w:r w:rsidDel="00112233">', '<w:r w:rsidDel="00112233" w:rsidR="ABCDEF01">').replace("<w:r><w:t>two", '<w:r w:rsidRPr="00112233"><w:t>two')
    const candidate = leftMeta + rightMeta + prior("4")
    expect(verifyNativeWordChatParagraphSplit(base, after(candidate), target, "One\rtwo")).toEqual(["Reviewer"])
    for (const xml of [
      candidate.replace('w:rsidR="ABCDEF01"', 'w:rsidR="invalid"'),
      candidate.replace('w:rsidR="ABCDEF01"', 'foreign="ABCDEF01"'),
      candidate.replace("<w:t>two", "<w:rPr><w:b/></w:rPr><w:t>two"),
    ]) expect(() => verifyNativeWordChatParagraphSplit(base, after(xml), target, "One\rtwo")).toThrow()
    expect(() => assertNativeWordChatParagraphSplitTarget(projectNativeWordChatBody(flat(sourceMeta.replace("1234ABCD", "bad") + prior("2")), inv), target)).toThrow()
  })
})
