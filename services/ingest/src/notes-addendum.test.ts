import { describe, expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { appendAcceptedChangesToNotes, type AcceptedNotesChange } from "./notes-addendum"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"

const HEADING = "Changes accepted after these notes were prepared"
const INTRO = "These changes were accepted in Word after the final check. The final check was not re-run on them. Changes accepted while the notes could not be updated are not listed."

// Synthetic notes rendered by the real drafting-notes renderer.
const notes = (await renderMarkdownToDocx([
  "# Drafting notes",
  "The supply note follows the agreed delivery terms.",
  "## Open points",
  "Confirm the notice address with the client.",
].join("\n\n"))).bytes

const replaced: AcceptedNotesChange = { document: "Supply note.docx", before: "within five days", after: "within ten days", reason: "Client instruction", acceptedAt: "2026-10-17T09:30:00Z" }
const inserted: AcceptedNotesChange = { document: "Supply note.docx", before: "", after: "Notices may be sent by email.", acceptedAt: "2026-10-17T10:00:00Z" }
const deleted: AcceptedNotesChange = { document: "Covering letter.docx", before: "Time is of the essence.", after: "", reason: "Not agreed.", acceptedAt: "2026-10-18T08:00:00+02:00" }

const texts = (bytes: Uint8Array) => readDocxCommentParagraphs(bytes).map(row => row.paragraph_text)
const part = (bytes: Uint8Array, name: string) => strFromU8(unzipSync(bytes)[name]!)

describe("drafting notes addendum for accepted Word changes", () => {
  test("first append adds the heading, the note and each change in order", () => {
    const output = appendAcceptedChangesToNotes(notes, [replaced, inserted, deleted])
    expect(texts(output)).toEqual([
      ...texts(notes),
      HEADING,
      INTRO,
      "2026-10-17 — Supply note.docx: “within five days” → “within ten days”. Reason: Client instruction.",
      "2026-10-17 — Supply note.docx: added “Notices may be sent by email.”.",
      "2026-10-18 — Covering letter.docx: deleted “Time is of the essence.”. Reason: Not agreed.",
    ])
    // The renderer's "##" heading style, so the addendum heading looks native.
    expect(part(output, "word/document.xml")).toContain(`<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">${HEADING}</w:t></w:r></w:p>`)
    expect(part(notes, "word/styles.xml")).toMatch(/w:styleId="Heading2">\s*<w:name w:val="heading 2"\s*\/>/)
  })

  test("a later append adds only its changes after the existing addendum", () => {
    const first = appendAcceptedChangesToNotes(notes, [replaced])
    const second = appendAcceptedChangesToNotes(first, [inserted, deleted])
    expect(texts(second)).toEqual([
      ...texts(first),
      "2026-10-17 — Supply note.docx: added “Notices may be sent by email.”.",
      "2026-10-18 — Covering letter.docx: deleted “Time is of the essence.”. Reason: Not agreed.",
    ])
    expect(texts(second).filter(text => text === HEADING)).toHaveLength(1)
    expect(texts(second).filter(text => text === INTRO)).toHaveLength(1)
  })

  test("quotes long and special text exactly and omits an absent reason", () => {
    const before = `${"The Supplier shall deliver the Goods & Services <as ordered>. ".repeat(40)}End.`
    const output = appendAcceptedChangesToNotes(notes, [
      { document: "Supply note.docx", before, after: "Delivery\tschedule", acceptedAt: "2026-10-17" },
      { document: "Supply note.docx", before: "a", after: "b", reason: "   ", acceptedAt: "2026-10-17T00:00:00.000Z" },
    ])
    expect(texts(output).slice(-2)).toEqual([
      `2026-10-17 — Supply note.docx: “${before}” → “Delivery\tschedule”.`,
      "2026-10-17 — Supply note.docx: “a” → “b”.",
    ])
  })

  test("keeps lawyer-edited content and every other part byte for byte", () => {
    // A lawyer's Word edit: bold wording and an extra paragraph of their own.
    const parts = unzipSync(notes)
    const xml = strFromU8(parts["word/document.xml"]!)
    const edited = xml
      .replace("<w:t xml:space=\"preserve\">Confirm the notice", "<w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">Confirm the notice")
      .replace("<w:sectPr>", "<w:p><w:r><w:t>Lawyer note: call the client on Monday.</w:t></w:r></w:p>\n    <w:sectPr>")
    expect(edited).not.toBe(xml)
    const lawyer = zipSync({ ...parts, "word/document.xml": strToU8(edited) })
    const output = appendAcceptedChangesToNotes(lawyer, [replaced])
    const after = unzipSync(output)
    const original = unzipSync(lawyer)
    expect(Object.keys(after)).toEqual(Object.keys(original))
    for (const name of Object.keys(original).filter(name => name !== "word/document.xml"))
      expect(Buffer.from(after[name]!).equals(Buffer.from(original[name]!))).toBe(true)
    const result = strFromU8(after["word/document.xml"]!)
    const at = edited.lastIndexOf("<w:sectPr>")
    expect(result.slice(0, at)).toBe(edited.slice(0, at))
    expect(result.endsWith(edited.slice(at))).toBe(true)
    expect(result.length).toBeGreaterThan(edited.length)
    expect(texts(output)).toEqual([...texts(lawyer), HEADING, INTRO, "2026-10-17 — Supply note.docx: “within five days” → “within ten days”. Reason: Client instruction."])
  })

  test("output opens in the viewer conversion and the document projection", async () => {
    const output = appendAcceptedChangesToNotes(notes, [replaced, inserted])
    const dx = await docxodus()
    const html = await dx.convertDocxToHtml(output, { renderTrackedChanges: true, showDeletedContent: true, renderHeadersAndFooters: true })
    expect(html).toContain(`<h2 dir="ltr" class="docx-Heading2"><span class="docx-000003">${HEADING}</span></h2>`)
    expect(html).toContain("2026-10-17 — Supply note.docx: added “Notices may be sent by email.”.")
    const session = dx.openDocxSession(output, {})
    try {
      const lines = session.project().markdown.split("\n").map(line => line.replace(/^\{#[^}]+\} /, "").replace(/\\(.)/g, "$1")).filter(Boolean)
      expect(lines.slice(-4)).toEqual([
        `## ${HEADING}`,
        INTRO,
        "2026-10-17 — Supply note.docx: “within five days” → “within ten days”. Reason: Client instruction.",
        "2026-10-17 — Supply note.docx: added “Notices may be sent by email.”.",
      ])
    } finally {
      session.close()
    }
  })

  test("a notes document without a styles part gets an unstyled heading before its section properties", async () => {
    const document = [
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>`,
      `<w:p><w:r><w:t>Drafting notes without styles.</w:t></w:r></w:p>`,
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="1" w:author="Reviewer" w:date="2026-10-01T00:00:00Z"><w:sectPr/></w:sectPrChange></w:sectPr>`,
      `</w:body></w:document>`,
    ].join("")
    const bare = zipSync({
      "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
      "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(document),
    })
    const output = appendAcceptedChangesToNotes(bare, [inserted])
    const xml = part(output, "word/document.xml")
    expect(xml).not.toContain("pStyle")
    expect(xml).toContain(`<w:p><w:r><w:t xml:space="preserve">${HEADING}</w:t></w:r></w:p>`)
    expect(xml.slice(xml.indexOf("<w:sectPr>"))).toBe(document.slice(document.indexOf("<w:sectPr>")))
    expect(texts(output)).toEqual(["Drafting notes without styles.", HEADING, INTRO, "2026-10-17 — Supply note.docx: added “Notices may be sent by email.”."])
    const session = (await docxodus()).openDocxSession(output, {})
    try { expect(session.project().markdown).toContain(HEADING) } finally { session.close() }
  })

  test("returns the notes unchanged with no changes and refuses an undescribable change", () => {
    expect(appendAcceptedChangesToNotes(notes, [])).toBe(notes)
    expect(() => appendAcceptedChangesToNotes(notes, [{ ...replaced, before: "", after: "" }])).toThrow("no text before or after")
    expect(() => appendAcceptedChangesToNotes(notes, [{ ...replaced, acceptedAt: "17 October 2026" }])).toThrow("invalid acceptance time")
  })
})
