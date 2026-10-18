import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { combineMemoAndBody } from "./combined-docx"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const field = (...codes: string[]) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`
  + codes.map(code => `<w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r>`).join("")
  + `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("1")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
const mark = (id: number, name: string, text: string) =>
  `<w:p><w:bookmarkStart w:id="${id}" w:name="${name}"/>${run(text)}<w:bookmarkEnd w:id="${id}"/></w:p>`

test("a memo bookmark whose name the body uses is renamed, and only the memo's references follow it", () => {
  // The memo refers to its own _Ref1 before it is defined: a split REF field,
  // a PAGEREF and HYPERLINK \l in simple fields, and an internal hyperlink.
  const memo = docx(`<w:p>${run("See ")}${field(" REF ", "_Ref1 \\h ")}${run(" on page ")}`
    + `<w:fldSimple w:instr=" PAGEREF _Ref1 \\h ">${run("1")}</w:fldSimple>`
    + `<w:hyperlink w:anchor="_Ref1">${run("jump")}</w:hyperlink>`
    + `<w:fldSimple w:instr=" HYPERLINK \\l &quot;_Ref1&quot; ">${run("also")}</w:fldSimple></w:p>`
    + mark(0, "_Ref1", "Memo target") + mark(1, "_GoBack", "Memo edit"))
  const body = docx(mark(0, "_Ref1", "Body target") + mark(3, "_GoBack", "Body edit")
    + `<w:p>${field(" REF _Ref1 \\h ")}<w:hyperlink w:anchor="_Ref1">${run("back")}</w:hyperlink></w:p>`)
  const combined = combineMemoAndBody({ memo, body })
  const document = part(combined, "word/document.xml")
  const blocks = bodyBlocks(document)

  const starts = all(document, "bookmarkStart")
  const named = (text: string) => starts.find(start => start.parentNode?.textContent === text)?.getAttributeNS(W, "name")
  const renamed = named("Memo target")!
  expect(named("Body target")).toBe("_Ref1")
  expect(renamed).toBeString()
  expect(renamed).not.toBe("_Ref1")
  expect(renamed.toLowerCase()).not.toBe("_goback")
  expect(renamed.length).toBeLessThanOrEqual(40)
  // Only Word's last-edit mark is dropped; the body keeps its own.
  expect(starts.map(start => start.getAttributeNS(W, "name")).toSorted()).toEqual(["_GoBack", "_Ref1", renamed].toSorted())
  expect(named("Body edit")).toBe("_GoBack")
  expectPaired(document)

  const [references] = blocks
  expect(all(references!, "instrText").map(code => code.textContent).join("")).toBe(` REF ${renamed} \\h `)
  expect(all(references!, "fldSimple").map(node => node.getAttributeNS(W, "instr")))
    .toEqual([` PAGEREF ${renamed} \\h `, ` HYPERLINK \\l "${renamed}" `])
  expect(all(references!, "hyperlink").map(node => node.getAttributeNS(W, "anchor"))).toEqual([renamed])
  expect(blocks.slice(3).map(canonical)).toEqual(bodyBlocks(part(body, "word/document.xml")).map(canonical))
})

test("a renamed bookmark stays within Word's 40 characters and clear of every name in either file", () => {
  const long = `_Toc${"9".repeat(36)}`
  const memo = docx(mark(0, long, "Memo target") + mark(1, `${long.slice(0, 38)}_1`, "Memo other")
    + `<w:p><w:hyperlink w:anchor="${long}">${run("jump")}</w:hyperlink></w:p>`)
  const body = docx(mark(0, long, "Body target"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  const names = all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "name")!)
  expect(new Set(names.map(name => name.toLowerCase())).size).toBe(3)
  expect(names.every(name => name.length <= 40)).toBe(true)
  expect(all(document, "hyperlink")[0]!.getAttributeNS(W, "anchor")).toBe(names[0]!)
  expect(names[0]).not.toBe(long)
  expectPaired(document)
})

test("memo bookmarks that repeat an id each keep one start and one end; a stray end goes", () => {
  const memo = docx(mark(5, "first", "One") + mark(5, "second", "Two")
    + `<w:p><w:bookmarkStart w:id="7" w:name="outer"/><w:bookmarkStart w:id="7" w:name="inner"/>${run("Three")}<w:bookmarkEnd w:id="7"/><w:bookmarkEnd w:id="7"/><w:bookmarkEnd w:id="9"/></w:p>`)
  const body = docx(mark(0, "body", "Body"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  expectPaired(document)
  expect(all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "name"))).toEqual(["first", "second", "outer", "inner", "body"])
})

test("memo bookmark ids stay clear of ids the body uses in its headers, footers and notes", () => {
  // Bookmark ids are numbered package-wide; the body's document part uses only 0.
  const notes = (id: number) => `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:bookmarkStart w:name="_RefNote" w:id="${id}"/>${run("n")}<w:bookmarkEnd w:id="${id}"/></w:p></w:footnote></w:footnotes>`
  const header = `<w:hdr xmlns:w="${W}"><w:p><w:bookmarkStart w:id="2" w:name="_RefHeader"/>${run("h")}<w:bookmarkEnd w:id="2"/></w:p></w:hdr>`
  const memo = docx(mark(0, "one", "One") + mark(1, "two", "Two") + mark(2, "three", "Three"))
  const body = docx(mark(0, "body", "Body"), { "word/footnotes.xml": notes(1), "word/header1.xml": header })
  const combined = combineMemoAndBody({ memo, body })
  const document = part(combined, "word/document.xml")
  expectPaired(document)
  const ids = all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "id"))
  expect(ids).toHaveLength(4)
  expect(ids.filter(id => ["1", "2"].includes(id!))).toEqual([])
  expect(unzipSync(combined)["word/footnotes.xml"]).toEqual(unzipSync(body)["word/footnotes.xml"]!)
})

test("a link into another file keeps its location even when a memo bookmark of that name is renamed", () => {
  const external = "https://example.com/other.docx"
  const memo = docx(`<w:p><w:fldSimple w:instr=" HYPERLINK &quot;${external}&quot; \\l &quot;_Ref1&quot; ">${run("simple")}</w:fldSimple>`
    + field(` HYPERLINK "${external}" `, `\\l "_Ref1" `)
    + `<w:hyperlink r:id="rId1" w:anchor="_Ref1">${run("external")}</w:hyperlink>`
    + `<w:fldSimple w:instr=" HYPERLINK \\l &quot;_Ref1&quot; \\o &quot;tip&quot; ">${run("internal")}</w:fldSimple>`
    + `<w:hyperlink w:anchor="_Ref1">${run("internal")}</w:hyperlink></w:p>`
    + mark(0, "_Ref1", "Memo target"), {
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/hyperlink" Target="${external}" TargetMode="External"/></Relationships>`,
  })
  const body = docx(mark(0, "_Ref1", "Body target"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  const renamed = all(document, "bookmarkStart").find(start => start.parentNode?.textContent === "Memo target")!.getAttributeNS(W, "name")!
  expect(renamed).not.toBe("_Ref1")
  const [references] = bodyBlocks(document)
  expect(all(references!, "fldSimple").map(node => node.getAttributeNS(W, "instr")))
    .toEqual([` HYPERLINK "${external}" \\l "_Ref1" `, ` HYPERLINK \\l "${renamed}" \\o "tip" `])
  expect(all(references!, "instrText").map(code => code.textContent).join("")).toBe(` HYPERLINK "${external}" \\l "_Ref1" `)
  expect(all(references!, "hyperlink").map(node => node.getAttributeNS(W, "anchor"))).toEqual(["_Ref1", renamed])
})

test("a memo bookmark start with no end goes, and a later complete bookmark of its name stays", () => {
  const memo = docx(`<w:p><w:bookmarkStart w:id="4" w:name="shared"/><w:bookmarkStart w:id="5" w:name="lost"/>${run("Open")}</w:p>`
    + mark(6, "shared", "Closed") + `<w:p><w:fldSimple w:instr=" REF shared ">${run("1")}</w:fldSimple></w:p>`)
  const body = docx(mark(0, "body", "Body"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  expectPaired(document)
  const starts = all(document, "bookmarkStart")
  expect(starts.map(start => start.getAttributeNS(W, "name"))).toEqual(["shared", "body"])
  expect(starts[0]!.parentNode?.textContent).toBe("Closed")
})

test("a memo paragraph pointer that the body also uses gets a fresh pointer, and the memo's references follow it", () => {
  const pointer = "_dh_p0123456789ab"
  const memo = docx(mark(0, pointer, "Memo paragraph") + `<w:p><w:fldSimple w:instr=" PAGEREF ${pointer} \\h ">${run("1")}</w:fldSimple></w:p>`)
  const body = docx(mark(0, pointer, "Body paragraph"))
  const combined = combineMemoAndBody({ memo, body })
  const document = part(combined, "word/document.xml")
  expectPaired(document)
  const names = all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "name")!)
  expect(names[1]).toBe(pointer)
  expect(names[0]).toMatch(/^_dh_p[0-9a-f]{12}$/)
  expect(names[0]).not.toBe(pointer)
  expect(all(document, "fldSimple")[0]!.getAttributeNS(W, "instr")).toBe(` PAGEREF ${names[0]} \\h `)
  expect(combineMemoAndBody({ memo, body })).toEqual(combined)
})

test("a field split by a nested field, and a bare bookmark field, follow a renamed bookmark; nested and keyword fields do not", () => {
  const nested = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "_Ref1" </w:instrText></w:r>`
    + `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("_Ref1")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  const split = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK \\o "</w:instrText></w:r>${nested}`
    + `<w:r><w:instrText xml:space="preserve">" \\l "_Ref1" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("jump")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  const memo = docx(`<w:p>${split}${field(" _Ref1 \\h ")}<w:fldSimple w:instr=" _Ref1 ">${run("1")}</w:fldSimple>`
    + `<w:fldSimple w:instr=" PAGE ">${run("1")}</w:fldSimple>${field(" page \\* MERGEFORMAT ")}</w:p>`
    + mark(0, "_Ref1", "Memo target") + mark(1, "Page", "Memo page"))
  const body = docx(mark(0, "_Ref1", "Body target") + mark(1, "page", "Body page"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  const renamed = all(document, "bookmarkStart").find(start => start.parentNode?.textContent === "Memo target")!.getAttributeNS(W, "name")!
  expect(renamed).not.toBe("_Ref1")
  const [references] = bodyBlocks(document)
  expect(all(references!, "instrText").map(code => code.textContent))
    .toEqual([" HYPERLINK \\o \"", ' QUOTE "_Ref1" ', `" \\l "${renamed}" `, ` ${renamed} \\h `, " page \\* MERGEFORMAT "])
  expect(all(references!, "fldSimple").map(node => node.getAttributeNS(W, "instr"))).toEqual([` ${renamed} `, " PAGE "])
  expect(all(references!, "t").map(text => text.textContent)).toContain("_Ref1")
})

test("a memo reference to a bookmark that goes, or that the memo never marks, does not reach a body bookmark of that name", () => {
  // The memo's "shared" start has no end and goes; it never marks "Gone".
  const memo = docx(`<w:p><w:bookmarkStart w:id="4" w:name="shared"/>${run("Memo orphan")}</w:p>`
    + `<w:p><w:fldSimple w:instr=" REF shared ">${run("1")}</w:fldSimple>${field(" PAGEREF Gone \\h ")}`
    + `<w:hyperlink w:anchor="SHARED">${run("jump")}</w:hyperlink></w:p>`)
  const body = docx(mark(0, "shared", "Body shared") + mark(1, "gone", "Body gone")
    + `<w:p><w:fldSimple w:instr=" REF shared ">${run("1")}</w:fldSimple></w:p>`)
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  expectPaired(document)
  const names = all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "name")!.toLowerCase())
  expect(names).toEqual(["shared", "gone"])
  const [, references] = bodyBlocks(document)
  const instructions = [...all(references!, "fldSimple").map(node => node.getAttributeNS(W, "instr")!),
    all(references!, "instrText").map(code => code.textContent).join("")]
  expect(instructions).toEqual([expect.stringMatching(/^ REF \S+ $/), expect.stringMatching(/^ PAGEREF \S+ \\h $/)])
  const targets = [...instructions.map(code => code.trim().split(/\s+/)[1]!), all(references!, "hyperlink")[0]!.getAttributeNS(W, "anchor")!]
  expect(targets.filter(target => names.includes(target.toLowerCase()))).toEqual([])
  expect(targets[2]).toBe(targets[0]!)
  expect(all(bodyBlocks(document).at(-1)!, "fldSimple")[0]!.getAttributeNS(W, "instr")).toBe(" REF shared ")
})

test("a memo bookmark whose name the body uses only in a note or a header is renamed, and the memo's references follow it", () => {
  const notes = `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:bookmarkStart w:id="3" w:name="_Ref555"/>${run("n")}<w:bookmarkEnd w:id="3"/></w:p></w:footnote></w:footnotes>`
  const header = `<w:hdr xmlns:w="${W}"><w:p><w:bookmarkStart w:name='_RefHead' w:id='4'/>${run("h")}<w:bookmarkEnd w:id='4'/></w:p></w:hdr>`
  const memo = docx(mark(0, "_Ref555", "Memo note") + mark(1, "_refhead", "Memo head")
    + `<w:p><w:fldSimple w:instr=" REF _Ref555 ">${run("1")}</w:fldSimple><w:hyperlink w:anchor="_refhead">${run("jump")}</w:hyperlink></w:p>`)
  const body = docx(mark(0, "body", "Body"), { "word/footnotes.xml": notes, "word/header1.xml": header })
  const combined = combineMemoAndBody({ memo, body })
  const document = part(combined, "word/document.xml")
  expectPaired(document)
  const starts = all(document, "bookmarkStart")
  const named = (text: string) => starts.find(start => start.parentNode?.textContent === text)!.getAttributeNS(W, "name")!
  expect(named("Memo note").toLowerCase()).not.toBe("_ref555")
  expect(named("Memo head").toLowerCase()).not.toBe("_refhead")
  const [, , references] = bodyBlocks(document)
  expect(all(references!, "fldSimple")[0]!.getAttributeNS(W, "instr")).toBe(` REF ${named("Memo note")} `)
  expect(all(references!, "hyperlink")[0]!.getAttributeNS(W, "anchor")).toBe(named("Memo head"))
  expect(unzipSync(combined)["word/footnotes.xml"]).toEqual(unzipSync(body)["word/footnotes.xml"]!)
})

test("an escaped quote in a quoted field argument does not hide a link's bookmark", () => {
  const memo = docx(`<w:p><w:fldSimple w:instr=' HYPERLINK \\o "say \\"hi\\"" \\l "_Ref1" '>${run("simple")}</w:fldSimple>`
    + `${field(' HYPERLINK \\o "a \\"', 'b\\"" \\l "_Ref1" ')}</w:p>` + mark(0, "_Ref1", "Memo target"))
  const body = docx(mark(0, "_Ref1", "Body target"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  const renamed = all(document, "bookmarkStart").find(start => start.parentNode?.textContent === "Memo target")!.getAttributeNS(W, "name")!
  expect(renamed).not.toBe("_Ref1")
  const [references] = bodyBlocks(document)
  expect(all(references!, "fldSimple")[0]!.getAttributeNS(W, "instr")).toBe(` HYPERLINK \\o "say \\"hi\\"" \\l "${renamed}" `)
  expect(all(references!, "instrText").map(code => code.textContent)).toEqual([' HYPERLINK \\o "a \\"', `b\\"" \\l "${renamed}" `])
})

test("a field keyword stays when a memo bookmark of that name is renamed, and FTNREF follows its bookmark", () => {
  const names = ["OCX", "FtnRef", "Subscriber", "HTMLControl", "Note"]
  const memo = docx(`<w:p>${field(" OCX ")}<w:fldSimple w:instr=" FTNREF Note \\h ">${run("1")}</w:fldSimple>`
    + `<w:fldSimple w:instr=" SUBSCRIBER ">${run("1")}</w:fldSimple>${field(" htmlcontrol ")}</w:p>`
    + names.map((name, id) => mark(id, name, `Memo ${name}`)).join(""))
  const body = docx(names.map((name, id) => mark(id, name.toUpperCase(), `Body ${name}`)).join(""))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  expectPaired(document)
  const starts = all(document, "bookmarkStart")
  const renamed = names.map(name => starts.find(start => start.parentNode?.textContent === `Memo ${name}`)!.getAttributeNS(W, "name")!)
  expect(renamed.filter((name, index) => name.toLowerCase() === names[index]!.toLowerCase())).toEqual([])
  const [references] = bodyBlocks(document)
  expect(all(references!, "instrText").map(code => code.textContent)).toEqual([" OCX ", " htmlcontrol "])
  expect(all(references!, "fldSimple").map(node => node.getAttributeNS(W, "instr"))).toEqual([` FTNREF ${renamed[4]} \\h `, " SUBSCRIBER "])
})

// Every bookmarkEnd matches exactly one bookmarkStart and each start has one end.
function expectPaired(document: Document) {
  const starts = all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "id"))
  expect(new Set(starts).size).toBe(starts.length)
  expect(all(document, "bookmarkEnd").map(end => end.getAttributeNS(W, "id")).toSorted()).toEqual(starts.toSorted())
}

// A minimal package with no styles or numbering part, and any extra parts given.
function docx(blocks: string, extra: Record<string, string> = {}) {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  return zipSync({
    ...Object.fromEntries(Object.entries(extra).map(([name, text]) => [name, xml(text)])),
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>${blocks}<w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
}

function part(bytes: Uint8Array, name: string) {
  return new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[name]!), "application/xml")
}

function all(node: Document | Element, name: string) {
  return Array.from(node.getElementsByTagNameNS(W, name))
}

function canonical(node: Node) {
  return new XMLSerializer().serializeToString(node)
}

function bodyBlocks(doc: Document) {
  return Array.from(all(doc, "body")[0]!.childNodes)
    .filter((node): node is Element => node.nodeType === 1 && (node as Element).localName !== "sectPr")
}

const names = (document: Document) => all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "name")!)
const simple = (document: Document) => all(document, "fldSimple").map(node => node.getAttributeNS(W, "instr")!)

test("a body bookmark name written as a character reference clashes with the memo's bookmark of that name", () => {
  const header = (name: string) => `<w:hdr xmlns:w="${W}"><w:p><w:bookmarkStart w:id="5" w:name="${name}"/>${run("h")}<w:bookmarkEnd w:id="5"/></w:p></w:hdr>`
  for (const encoded of ["&#220;bersicht", "&#xDC;bersicht", "&#xdc;BERSICHT", "&#x00DC;bersicht"]) {
    const memo = docx(mark(0, "Übersicht", "Memo target") + `<w:p><w:fldSimple w:instr=" REF Übersicht ">${run("1")}</w:fldSimple></w:p>`)
    const body = docx(mark(0, "body", "Body"), { "word/header1.xml": header(encoded) })
    const combined = combineMemoAndBody({ memo, body })
    const document = part(combined, "word/document.xml")
    const memoName = names(document).find(name => name !== "body")!
    expect(memoName.toLowerCase()).not.toBe("übersicht")
    expect(simple(document)).toEqual([` REF ${memoName} `])
  }
  const named = docx(mark(0, "A&amp;B", "Memo target"))
  const clash = docx(mark(0, "body", "Body"), { "word/header1.xml": header("a&amp;b") })
  expect(names(part(combineMemoAndBody({ memo: named, body: clash }), "word/document.xml")).filter(name => name.toLowerCase() === "a&b")).toEqual([])
})

test("fields that read a bookmark by keyword or expression follow a renamed memo bookmark", () => {
  const codes = [" = Fee * 2 ", " =Fee+fee_x ", " IF Fee > 100 \"Fee\" \"low\" ", " GOTOBUTTON Fee Jump ", " TOC \\b Fee \\h ", " SET Fee 5 ", " ASK Fee \"amount?\" ",
    " = SUM(Fee, 3) \\# \"0.00\" \\* Fee ", " = 1e5 + Fee "]
  const escape = (code: string) => code.replace(/"/g, "&quot;")
  const memo = docx(mark(0, "Fee", "10") + `<w:p>${codes.map(code => `<w:fldSimple w:instr="${escape(code)}">${run("x")}</w:fldSimple>`).join("")}</w:p>`
    + `<w:p>${field(" = ", "Fee * 3")}${field(" GOTOBUTTON Fe", "e Jump")}</w:p>`)
  const body = docx(mark(0, "Fee", "999"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  const fresh = all(document, "bookmarkStart").find(start => start.parentNode?.textContent === "10")!.getAttributeNS(W, "name")!
  expect(fresh).not.toBe("Fee")
  expect(simple(document)).toEqual([
    ` = ${fresh} * 2 `, ` =${fresh}+fee_x `, ` IF ${fresh} > 100 "Fee" "low" `, ` GOTOBUTTON ${fresh} Jump `, ` TOC \\b ${fresh} \\h `,
    ` SET ${fresh} 5 `, ` ASK ${fresh} "amount?" `, ` = SUM(${fresh}, 3) \\# "0.00" \\* Fee `, ` = 1e5 + ${fresh} `,
  ])
  expect(all(document, "instrText").map(code => code.textContent).join("")).toBe(` = ${fresh} * 3 GOTOBUTTON ${fresh} Jump`)
})

test("a dangling memo reference is never captured by a minted name", () => {
  const dangling = (name: string) => `<w:p><w:fldSimple w:instr=" REF ${name} ">${run("Error!")}</w:fldSimple></w:p>`
  const memo = docx(mark(0, "Price", "Memo price") + dangling("Price_1"))
  const body = docx(mark(0, "Price", "Body price") + dangling("price_1"))
  const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
  const minted = all(document, "bookmarkStart").map(start => start.getAttributeNS(W, "name")!).filter(name => name !== "Price")
  expect(minted).toHaveLength(1)
  expect(["price_1"]).not.toContain(minted[0]!.toLowerCase())
  expect(simple(document).map(code => code.trim())).toEqual(["REF Price_1", "REF price_1"])
})

test("a memo bookmark is renamed away from a name the body only refers to, and the memo's references follow", () => {
  const memo = docx(mark(0, "Term", "Memo term") + `<w:p><w:fldSimple w:instr=" REF Term ">${run("1")}</w:fldSimple><w:hyperlink w:anchor="Term">${run("j")}</w:hyperlink></w:p>`)
  for (const reference of [`<w:fldSimple w:instr=" REF Term ">${run("x")}</w:fldSimple>`, `<w:hyperlink w:anchor="term">${run("x")}</w:hyperlink>`,
    `<w:fldSimple w:instr=" = Term * 2 ">${run("x")}</w:fldSimple>`, `<w:fldSimple w:instr=" GOTOBUTTON Term go ">${run("x")}</w:fldSimple>`]) {
    const body = docx(mark(0, "body", "Body") + `<w:p>${reference}</w:p>`)
    const document = part(combineMemoAndBody({ memo, body }), "word/document.xml")
    const memoName = names(document).find(name => name !== "body")!
    expect(memoName.toLowerCase()).not.toBe("term")
    expect(simple(document)[0]).toBe(` REF ${memoName} `)
    expect(all(document, "hyperlink")[0]!.getAttributeNS(W, "anchor")).toBe(memoName)
    expect(bodyBlocks(document).at(-1)!.textContent).toBe("x")
    expect(canonical(bodyBlocks(document).at(-1)!)).toBe(canonical(bodyBlocks(part(body, "word/document.xml")).at(-1)!))
  }
})

test("a minted name never splits a surrogate pair", () => {
  const long = `${"a".repeat(37)}😀😀😀`
  const memo = docx(mark(0, long, "Memo target"))
  const body = docx(mark(0, long, "Body target"))
  const minted = names(part(combineMemoAndBody({ memo, body }), "word/document.xml")).find(name => name !== long)!
  expect(minted.length).toBeLessThanOrEqual(40)
  expect(minted).not.toContain("\uFFFD")
  expect(minted).toBe(minted.toWellFormed())
  expect(minted).toEndWith("_1")
})
