import { expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { pathToFileURL } from "node:url"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { addCommentReplies, replyInitials } from "./docx-comment-replies"
import { readDocxMarkup, readMarkupComments } from "./docx-markup"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../test/support/counterparty-markup-fixture"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
const CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
const CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const MARKERS = ["commentRangeStart", "commentRangeEnd", "commentReference"]
const BY = { author: "Jordan Reyes", initials: replyInitials("Jordan Reyes"), date: "2026-10-25T14:05:09.250Z" }
const WRITTEN = "2026-10-25T14:05:09Z"
const { comments: COMMENTS, footnoteComment } = COUNTERPARTY_MARKUP

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml")
const all = (node: Document | Element | undefined, namespace: string, name: string) => node ? Array.from(node.getElementsByTagNameNS(namespace, name)) : []
const xml = (node: Node) => new XMLSerializer().serializeToString(node)
const parts = (bytes: Uint8Array) => Object.fromEntries(Object.entries(unzipSync(bytes)).filter(([name]) => /\.(xml|rels)$/.test(name)).map(([name, content]) => [name, parse(strFromU8(content))]))
const text = (bytes: Uint8Array, part: string) => strFromU8(unzipSync(bytes)[part]!)
const idOf = (node: Element) => node.getAttributeNS(W, "id") ?? ""
const lastParaId = (comment: Element) => all(comment, W, "p").at(-1)!.getAttributeNS(W14, "paraId") ?? ""
// Each part's comment references by id, in document order.
const references = (bytes: Uint8Array, part: string) => all(parts(bytes)[part], W, "commentReference").map(idOf)

// A Word 2010 markup: the fixture without its threading records or paraIds, so
// every comment is flat and Morgan's reply stands alone.
function flatMarkup() {
  const files = unzipSync(counterpartyMarkupDocx())
  delete files["word/commentsExtended.xml"]
  delete files["word/people.xml"]
  const edit = (part: string, change: (xml: string) => string) => { files[part] = strToU8(change(strFromU8(files[part]!))) }
  edit("word/comments.xml", xml => xml.replace(/ w14:paraId="\w+"/g, ""))
  edit("word/_rels/document.xml.rels", xml => xml.replace(/<Relationship Id="rId[23]"[^>]*\/>/g, ""))
  edit("[Content_Types].xml", xml => xml.replace(/<Override PartName="\/word\/(commentsExtended|people)\.xml"[^>]*\/>/g, ""))
  return zipSync(files)
}

// The OOXML contract for added replies: each is one new comment whose last
// paragraph has a paraId unique in the package and in range, one commentEx
// linking it to its root's last paragraph, open, and one commentId with a unique
// durable id; its one reference follows its thread's others. Every earlier
// comment, threading entry, person, relationship, content type, marker and
// tracked change is kept; a flat root only gains its last paragraph's paraId.
async function expectReplyContract(before: Uint8Array, after: Uint8Array, added: { thread: number; id: number }[]) {
  const old = parts(before), now = parts(after)
  const comments = all(now["word/comments.xml"], W, "comment"), earlier = all(old["word/comments.xml"], W, "comment")
  const comment = (list: Element[], id: number) => list.filter(node => idOf(node) === String(id))
  const paraIds = Object.values(now).flatMap(doc => all(doc, W, "p").map(node => node.getAttributeNS(W14, "paraId") ?? "").filter(Boolean))
  expect(new Set(paraIds).size).toBe(paraIds.length)
  expect(paraIds.filter(id => !/^[0-9A-F]{8}$/i.test(id) || parseInt(id, 16) === 0 || parseInt(id, 16) >= 0x80000000)).toEqual([])
  const entries = all(now["word/commentsExtended.xml"], W15, "commentEx"), ids = all(now["word/commentsIds.xml"], CID, "commentId")
  const durable = ids.map(node => node.getAttributeNS(CID, "durableId") ?? "")
  expect(new Set(durable).size).toBe(durable.length)
  expect(durable.filter(id => parseInt(id, 16) === 0 || parseInt(id, 16) >= 0x7FFFFFFF)).toEqual([])
  const threads = await readMarkupComments(after)
  for (const reply of added) {
    expect(comment(earlier, reply.id)).toEqual([])
    const own = comment(comments, reply.id), root = comment(comments, reply.thread)
    expect([own.length, root.length]).toEqual([1, 1])
    const paraId = lastParaId(own[0]!)
    expect(entries.filter(node => node.getAttributeNS(W15, "paraId") === paraId).map(node => [node.getAttributeNS(W15, "paraIdParent"), node.getAttributeNS(W15, "done")]))
      .toEqual([[lastParaId(root[0]!), "0"]])
    expect(ids.filter(node => node.getAttributeNS(CID, "paraId") === paraId)).toHaveLength(1)
    const thread = new Set(threads.filter(row => row.id === reply.thread || row.parent === reply.thread).map(row => String(row.id)))
    const holding = Object.values(now).filter(doc => all(doc, W, "commentReference").some(node => idOf(node) === String(reply.id)))
    expect(holding).toHaveLength(1)
    const order = all(holding[0], W, "commentReference").map(idOf).filter(id => thread.has(id))
    expect([order.indexOf(String(reply.id)), order.lastIndexOf(String(reply.id))]).toEqual([order.length - 1, order.length - 1])
  }
  for (const node of earlier) {
    const kept = comment(comments, Number(idOf(node)))[0]!.cloneNode(true) as Element
    if (!lastParaId(node)) all(kept, W, "p").at(-1)!.removeAttributeNS(W14, "paraId")
    expect(xml(kept)).toBe(xml(node))
  }
  for (const [part, namespace, name] of [["word/commentsExtended.xml", W15, "commentEx"], ["word/commentsIds.xml", CID, "commentId"],
    ["word/commentsExtensible.xml", CEX, "commentExtensible"], ["word/people.xml", W15, "person"]] as const) {
    const kept = new Set(all(now[part], namespace, name).map(xml))
    expect(all(old[part], namespace, name).map(xml).filter(entry => !kept.has(entry))).toEqual([])
  }
  const records = (doc: Document | undefined) => [...all(doc, REL, "Relationship"), ...all(doc, CT, "*")]
    .map(node => ["Id", "Type", "Target", "TargetMode", "PartName", "Extension", "ContentType"].map(name => node.getAttribute(name) ?? "").join(" "))
  const earlierIds = new Set(earlier.map(idOf))
  const markers = (part: Document | undefined) => all(part, W, "*").filter(node => MARKERS.includes(node.localName) && earlierIds.has(idOf(node))).map(node => `${node.localName} ${idOf(node)}`)
  for (const [name, doc] of Object.entries(old)) {
    const kept = new Set(records(now[name]))
    expect(records(doc).filter(record => !kept.has(record))).toEqual([])
    expect(markers(now[name])).toEqual(markers(doc))
  }
  expect((await readDocxMarkup(after)).revisions).toEqual((await readDocxMarkup(before)).revisions)
  expect(Object.values(unzipSync(after)).filter(content => strFromU8(content).includes("Unid"))).toEqual([])
}

test("a reply joins its thread after the counterparty's earlier reply, with the lawyer's name, initials and date", async () => {
  const before = counterpartyMarkupDocx()
  const { bytes, added } = await addCommentReplies(before, [{ thread: 0, text: "We can accept forty-five days on that basis." }], BY)
  expect(added).toEqual([{ thread: 0, id: 4 }])
  const reply = (await readMarkupComments(bytes)).find(comment => comment.id === 4)!
  expect(reply).toMatchObject({ author: BY.author, initials: "JR", date: WRITTEN, text: "We can accept forty-five days on that basis.", parent: 0, resolved: false })
  // Docxodus puts the reference after Avery's; Morgan's earlier reply stays ahead of it.
  expect(references(bytes, "word/document.xml")).toEqual(["0", "1", "4", "2", "3"])
  await expectReplyContract(before, bytes, added)
})

test("a reply names the thread's root: a reply's own id, an unknown thread, empty text or no attribution refuses", async () => {
  const bytes = counterpartyMarkupDocx()
  const refused = (replies: { thread: number; text: string }[], by = BY) => addCommentReplies(bytes, replies, by).then(() => "written", error => error.code)
  expect(await refused([{ thread: 1, text: "Agreed." }])).toBe("thread_not_root")
  expect(await refused([{ thread: 9, text: "Agreed." }])).toBe("thread_not_found")
  expect(await refused([{ thread: 0, text: "Agreed." }, { thread: 3, text: " \n\u200B " }])).toBe("empty_reply")
  expect(await refused([{ thread: 0, text: "Agreed." }], { ...BY, author: " " })).toBe("unattributed_reply")
  expect(await refused([{ thread: 0, text: "Agreed." }], { ...BY, date: "" })).toBe("unattributed_reply")
})

test("the reply is plain text: Markdown characters come back literally, and line breaks and paragraphs are kept", async () => {
  const typed = "See *clause 5*, [the schedule](https://example.com) and # 5 <b>&amp;</b> & C:\\Deals_2026 ~x~ |a| 1. +-!\nSecond line\n\nNew paragraph"
  const { bytes, added } = await addCommentReplies(counterpartyMarkupDocx(), [{ thread: 3, text: typed }], BY)
  expect((await readMarkupComments(bytes)).find(comment => comment.id === added[0]!.id)!.text).toBe(typed.replace("\n\n", "\n"))
  const reply = all(parts(bytes)["word/comments.xml"], W, "comment").find(node => idOf(node) === String(added[0]!.id))!
  expect([all(reply, W, "p").length, all(reply, W, "br").length, all(reply, W, "b").length, all(reply, W, "i").length]).toEqual([2, 1, 0, 0])
  expect(text(bytes, "word/comments.xml")).not.toContain("hyperlink")
})

test("Word's initials for a name are the first letter of each word, at most three", () => {
  expect(["Jordan Reyes", " jordan  reyes ", "Ana María de la Cruz", "Dr. (Sam) Lee", "Ō"].map(replyInitials)).toEqual(["JR", "JR", "AMD", "DSL", "Ō"])
})

test("a flat root gains the threading records, with their relationship and content type, and its reply threads under it", async () => {
  const before = flatMarkup()
  expect(Object.keys(unzipSync(before)).filter(name => /comments(Extended|Ids)/.test(name))).toEqual([])
  const { bytes, added } = await addCommentReplies(before, [{ thread: 0, text: "Agreed." }], BY)
  expect((await readMarkupComments(bytes)).find(comment => comment.id === added[0]!.id)).toMatchObject({ parent: 0, resolved: false })
  const relationships = text(bytes, "word/_rels/document.xml.rels"), types = text(bytes, "[Content_Types].xml")
  expect(relationships).toContain('Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended"')
  expect(relationships).toContain('Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"')
  expect(types).toContain('PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"')
  expect(types).toContain('PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"')
  await expectReplyContract(before, bytes, added)
})

test("in a Word 2016+ file every durable id and extensible entry is kept, and each reply gets a new durable id", async () => {
  const before = counterpartyMarkupDocx({ modern: true })
  const durable = (bytes: Uint8Array) => all(parts(bytes)["word/commentsIds.xml"], CID, "commentId").map(node => node.getAttributeNS(CID, "durableId"))
  const { bytes, added } = await addCommentReplies(before, [{ thread: 0, text: "Agreed." }, { thread: 3, text: "Noted." }], BY)
  expect(durable(bytes).slice(0, COMMENTS.length)).toEqual(durable(before))
  expect(durable(bytes)).toHaveLength(COMMENTS.length + added.length)
  await expectReplyContract(before, bytes, added)
})

test("a reply to a comment in a footnote is placed in the footnote, after its root", async () => {
  const before = counterpartyMarkupDocx({ footnote: true })
  const { bytes, added } = await addCommentReplies(before, [{ thread: footnoteComment.id, text: "Schedule 1." }], BY)
  const id = String(added[0]!.id)
  expect(references(bytes, "word/footnotes.xml")).toEqual([String(footnoteComment.id), id])
  expect(references(bytes, "word/document.xml")).not.toContain(id)
  expect((await readMarkupComments(bytes)).find(comment => comment.id === added[0]!.id)).toMatchObject({ parent: footnoteComment.id, scope: "footnote", text: "Schedule 1." })
  await expectReplyContract(before, bytes, added)
})

// LibreOffice 26.2 does not open a .docx with a comment in a footnote, so the
// footnote thread is checked in the package only.
test.skipIf(!Bun.which("soffice"))("LibreOffice shows each reply last in its thread, under the root, as the lawyer", async () => {
  const dir = mkdtempSync(path.join(tmpdir(), "comment-replies-lo-"))
  try {
    const replies = [{ thread: 0, text: "We can accept forty-five days on that basis." }, { thread: 3, text: "Please confirm the two-year term." }]
    writeFileSync(path.join(dir, "replied.docx"), (await addCommentReplies(counterpartyMarkupDocx(), replies, BY)).bytes)
    const converted = Bun.spawnSync(["soffice", "--headless", "--norestore", `-env:UserInstallation=${pathToFileURL(path.join(dir, "lo-profile")).href}`,
      "--convert-to", "odt", "--outdir", dir, path.join(dir, "replied.docx")])
    expect(converted.exitCode).toBe(0)
    const content = parse(strFromU8(unzipSync(readFileSync(path.join(dir, "replied.odt")))["content.xml"]!))
    const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", LOEXT = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
    const annotations = all(content, OFFICE, "annotation").map(node => ({
      name: node.getAttributeNS(OFFICE, "name"), parent: node.getAttributeNS(LOEXT, "parent-name"),
      creator: all(node, "http://purl.org/dc/elements/1.1/", "creator")[0]?.textContent ?? "",
      text: all(node, "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "p").map(paragraph => paragraph.textContent).join("\n"),
    }))
    for (const reply of replies) {
      const root = annotations.find(annotation => annotation.text === COMMENTS[reply.thread]!.text)!
      const thread = annotations.filter(annotation => annotation === root || annotation.parent === root.name)
      expect(thread.at(-1)).toMatchObject({ parent: root.name, creator: BY.author, text: reply.text })
    }
    // Thread 0 keeps Morgan's reply between Avery's root and the lawyer's.
    expect(annotations.filter(annotation => annotation.name === annotations[0]!.name || annotation.parent === annotations[0]!.name).map(annotation => annotation.creator))
      .toEqual([COMMENTS[0]!.author, COMMENTS[1]!.author, BY.author])
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}, 180_000)
