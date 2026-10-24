// True-redaction guarantees, exercised through the real redact tool against
// authored synthetic documents rendered by the product helper — no mocks beyond the
// tool ctx. Run from dochaus/: bun test test/redact.test.ts
//
// Lives in test/, not tool/: the engine's ToolRegistry imports every .ts under
// {tool,tools}/ as a custom tool, and importing a bun:test file outside the
// test runner throws ("Cannot use describe outside of the test runner").
import { describe, expect, spyOn, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, realpathSync, symlinkSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "../lib/docxodus"
import { renderMarkdownToDocx } from "../lib/render"
import { scrubDocxPackage, docxPackageResidue, listRedactions } from "../lib/redactions"
import { recordRedline, pendingRedlinesForDoc } from "../lib/redlines"

// The tool reads INGEST_URL at import time; point it at a dead port so a live
// dev ingest server never receives test uploads.
process.env.INGEST_URL = "http://127.0.0.1:1"
const redact = (await import("../tool/redact")).default

const SSN = "123-45-6789"
// Authored test content; the renderer supplies the shipped, source-built DOCX seed.
const fixtureBytes = (await renderMarkdownToDocx([
  "# Synthetic workshop record",
  "## Confidential Information",
  "The sample inventory lists blue paper labels and wooden counters.",
  "## Entire Agreement",
  "These invented instructions group the sample inventory details in one record.",
].join("\n\n"))).bytes
const dx = await docxodus()

function ctxFor(directory: string, ask: (request: { metadata: Record<string, unknown> }) => Promise<void> = async () => {}) {
  return {
    sessionID: "test",
    messageID: "test",
    callID: "test",
    agent: "test",
    directory,
    worktree: directory,
    abort: new AbortController().signal,
    metadata() {},
    ask,
  }
}

// Synthetic content with the SSN planted three times (twice in one heading, once in a
// new paragraph) and a creator metadata entry carrying it too.
function buildFixture() {
  const session = dx.openDocxSession(fixtureBytes, {})
  const heading = session.findByText("Confidential Information")!
  session.replaceTextRange(heading.id, "Confidential Information", `SSN ${SSN} and again ${SSN}`)
  const tail = session.findByText("Entire Agreement")!
  session.insertParagraph(tail.id, "after", `Employee SSN: ${SSN}, do redact.`)
  const bytes = session.save()
  session.close()

  const parts = unzipSync(bytes)
  parts["docProps/core.xml"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Agent ${SSN}</dc:creator><dc:title>Personnel file ${SSN}</dc:title></cp:coreProperties>`,
  )
  return zipSync(parts)
}

function residueParts(bytes: Uint8Array, needle: string) {
  return [...new Set(docxPackageResidue(bytes, needle).hits.map((hit) => hit.part))]
}

async function matterWith(bytes: Uint8Array) {
  const dir = mkdtempSync(path.join(tmpdir(), "redact-matter-"))
  const file = path.join(dir, "personnel.docx")
  await Bun.write(file, bytes)
  return { dir, file }
}

describe("metadata scrub", () => {
  test("removes the needle from docProps text nodes and reports residue correctly", () => {
    const fixture = buildFixture()
    expect(residueParts(fixture, SSN)).toContain("docProps/core.xml")
    const scrub = scrubDocxPackage(fixture, SSN, "[REDACTED]")
    expect(scrub.hits["docProps/core.xml"]).toBe(2)
    expect(residueParts(scrub.bytes, SSN)).not.toContain("docProps/core.xml")
    expect(strFromU8(unzipSync(scrub.bytes)["docProps/core.xml"])).toContain("Agent [REDACTED]")
  })
})

describe("redact tool", () => {
  test("removes every occurrence from content and metadata, logs it, and retires stale redlines", async () => {
    const { dir, file } = await matterWith(buildFixture())
    recordRedline(dir, {
      docPath: file,
      docName: "personnel.docx",
      scope: "phrase",
      findText: `SSN: ${SSN}`,
      oldText: `SSN: ${SSN}`,
      newText: "SSN: [on file]",
      author: "test",
      reviewComment: "Replace personal data with a non-sensitive label.",
      anchorId: "x",
    })

    const result = await redact.execute(
      { document: "personnel.docx", text: SSN, reason: "SSN — PII", author: "A. Lawyer", label: undefined },
      ctxFor(dir),
    )
    expect(result).toBeObject()
    if (!result || typeof result === "string") throw new Error(String(result))
    expect(result.output).toContain("3 text occurrence(s)")
    expect(result.output).toContain("re-indexing failed")
    expect(result.metadata?.reingested).toBe(false)

    // The file itself is clean: raw package residue zero, Docxodus flat-text zero.
    const redacted = await Bun.file(file).bytes()
    expect(docxPackageResidue(redacted, SSN).hits).toEqual([])
    const session = dx.openDocxSession(redacted, {})
    expect(session.grep(SSN.replace(/-/g, "\\-"), { scope: dx.ProjectionScopes.All })).toEqual([])
    // The projection escapes markdown punctuation; normalize before asserting.
    expect(session.project().markdown.replaceAll("\\", "")).toContain("[REDACTED]")
    session.close()

    const log = listRedactions(dir)
    expect(log).toHaveLength(1)
    expect(log[0].reason).toBe("SSN — PII")
    expect(log[0].author).toBe("A. Lawyer")
    expect(log[0].occurrences).toBe(3)
    expect(log[0].metadata_hits).toBe(2)

    expect(pendingRedlinesForDoc(dir, file)).toEqual([])
    const db = new Database(path.join(dir, ".dochaus", "legal.db"), { readonly: true })
    expect((db.query("SELECT status FROM redlines").get() as { status: string }).status).toBe("superseded")
    db.close()
  })

  test("a rejected ask leaves the document and log untouched", async () => {
    const { dir, file } = await matterWith(buildFixture())
    const before = await Bun.file(file).bytes()
    await expect(
      redact.execute(
        { document: "personnel.docx", text: SSN, reason: "SSN — PII", author: undefined, label: undefined },
        ctxFor(dir, async () => {
          throw new Error("rejected")
        }),
      ),
    ).rejects.toThrow("rejected")
    expect(await Bun.file(file).bytes()).toEqual(before)
    expect(listRedactions(dir)).toEqual([])
  })

  test("refuses a document with unresolved tracked changes", async () => {
    const session = dx.openDocxSession(fixtureBytes, {})
    const target = session.findByText("Confidential Information")!
    session.replaceTextRange(target.id, "Confidential Information", `Secret ${SSN}`)
    const modified = session.save()
    session.close()
    const tracked = await dx.compareDocuments(fixtureBytes, modified, { authorName: "test" })

    const { dir } = await matterWith(tracked)
    const result = await redact.execute(
      { document: "personnel.docx", text: SSN, reason: "SSN — PII", author: undefined, label: undefined },
      ctxFor(dir),
    )
    expect(result).toBeString()
    expect(result as string).toContain("unresolved tracked changes")
  })

  test("refuses when the text spans a paragraph boundary", async () => {
    const session = dx.openDocxSession(fixtureBytes, {})
    const tail = session.findByText("Entire Agreement")!
    session.insertParagraph(tail.id, "after", "ends with alpha")
    const first = session.findByText("ends with alpha")!
    session.insertParagraph(first.id, "after", "beta starts")
    const bytes = session.save()
    session.close()

    const { dir } = await matterWith(bytes)
    const result = await redact.execute(
      { document: "personnel.docx", text: "alpha\nbeta", reason: "test", author: undefined, label: undefined },
      ctxFor(dir),
    )
    expect(result).toBeString()
    expect(result as string).toContain("span a paragraph boundary")
  })

  test("reports when the text is nowhere in the document", async () => {
    const { dir } = await matterWith(fixtureBytes)
    const result = await redact.execute(
      { document: "personnel.docx", text: "not in this document", reason: "test", author: undefined, label: undefined },
      ctxFor(dir),
    )
    expect(result).toBeString()
    expect(result as string).toContain("Text not found")
  })

  test("rejects a label that contains the text being redacted", async () => {
    const { dir } = await matterWith(buildFixture())
    const result = await redact.execute(
      { document: "personnel.docx", text: SSN, reason: "test", author: undefined, label: `[was ${SSN}]` },
      ctxFor(dir),
    )
    expect(result).toBeString()
    expect(result as string).toContain("must not contain")
  })

  test("a log write that fails after the file is redacted still re-indexes and says what was not recorded", async () => {
    const { dir, file } = await matterWith(buildFixture())
    recordRedline(dir, {
      docPath: file,
      docName: "personnel.docx",
      scope: "phrase",
      findText: `SSN: ${SSN}`,
      oldText: `SSN: ${SSN}`,
      newText: "SSN: [on file]",
      author: "test",
      reviewComment: "",
      anchorId: "x",
    })
    // Another process holds legal.db's write lock past the writers' short wait,
    // as an ingest upload commit can: every log write after the .docx write
    // fails with SQLITE_BUSY.
    const holder = Bun.spawn(
      [
        process.execPath,
        "-e",
        `const { Database } = require("bun:sqlite")
        const db = new Database(${JSON.stringify(path.join(dir, ".dochaus", "legal.db"))})
        db.run("BEGIN IMMEDIATE")
        db.run("UPDATE redlines SET author = author")
        console.log("locked")
        setTimeout(() => {}, 60_000)`,
      ],
      { stdout: "pipe" },
    )
    const requests: string[] = []
    const fetchSpy = spyOn(globalThis, "fetch").mockImplementation((async (input: string | URL | Request) => {
      requests.push(String(input))
      return new Response("{}")
    }) as typeof fetch)
    let result: Awaited<ReturnType<typeof redact.execute>> | Error
    try {
      const reader = holder.stdout.getReader()
      expect(new TextDecoder().decode((await reader.read()).value)).toContain("locked")
      reader.releaseLock()
      result = await redact
        .execute({ document: "personnel.docx", text: SSN, reason: "SSN — PII", author: "A. Lawyer", label: undefined }, ctxFor(dir))
        .catch((error: Error) => error)
    } finally {
      fetchSpy.mockRestore()
      holder.kill()
      await holder.exited
    }

    // The irreversible removal happened, and the index was asked to drop the text.
    expect(docxPackageResidue(await Bun.file(file).bytes(), SSN).hits).toEqual([])
    expect({
      thrown: result instanceof Error ? result.message : null,
      reindexRequests: requests.filter((url) => url.endsWith("/documents/revisions")).length,
    }).toEqual({ thrown: null, reindexRequests: 1 })
    if (!result || typeof result === "string" || result instanceof Error) throw new Error(String(result))
    expect(result.output).toContain("Permanently removed")
    expect(result.output).not.toContain("Recorded as redaction log entry")
    expect(result.output).not.toContain("Retired")
    expect(result.output).toContain("WARNING: the redaction log entry was not recorded (database is locked)")
    expect(result.output).toContain("WARNING: pending redlines that reference the removed text were not retired (database is locked)")
    expect(result.metadata).toMatchObject({ redaction: null, supersededRedlines: null, reingested: true })
    // Nothing was written to the log or retired: the warnings are the record.
    expect(listRedactions(dir)).toEqual([])
    expect(pendingRedlinesForDoc(dir, file)).toHaveLength(1)
  }, 30_000)

  test("a save that could not delete the document's earlier versions is a warning", async () => {
    const { dir } = await matterWith(buildFixture())
    const failed = "The earlier versions of this document could not be deleted from doc.haus's kept versions; they may still hold the removed text."
    const fetchSpy = spyOn(globalThis, "fetch").mockImplementation((async () => Response.json({ ok: true, purge_failed: failed })) as typeof fetch)
    let result: Awaited<ReturnType<typeof redact.execute>>
    try {
      result = await redact.execute({ document: "personnel.docx", text: SSN, reason: "SSN — PII", author: "A. Lawyer", label: undefined }, ctxFor(dir))
    } finally {
      fetchSpy.mockRestore()
    }
    if (!result || typeof result === "string") throw new Error(String(result))
    expect(result.output).toContain(`WARNING: ${failed}`)
    expect(result.output).not.toContain("re-indexing failed")
  })

  test("a save that fails is sent again once before any warning", async () => {
    const { dir } = await matterWith(buildFixture())
    const origins: (string | null)[] = []
    const fetchSpy = spyOn(globalThis, "fetch").mockImplementation((async (_input: string | URL | Request, init?: RequestInit) => {
      origins.push(new Headers(init?.headers).get("x-dochaus-origin"))
      return origins.length === 1 ? new Response("unavailable", { status: 503 }) : Response.json({ ok: true })
    }) as typeof fetch)
    let result: Awaited<ReturnType<typeof redact.execute>>
    try {
      result = await redact.execute({ document: "personnel.docx", text: SSN, reason: "SSN — PII", author: "A. Lawyer", label: undefined }, ctxFor(dir))
    } finally {
      fetchSpy.mockRestore()
    }
    if (!result || typeof result === "string") throw new Error(String(result))
    expect(origins).toEqual(["redaction", "redaction"])
    expect(result.output).not.toContain("WARNING")
    expect(result.metadata?.reingested).toBe(true)
  })

  test("a save that fails twice says so, and running the redaction again saves it", async () => {
    const { dir, file } = await matterWith(buildFixture())
    const asked: Record<string, unknown>[] = []
    const run = () => redact.execute({ document: "personnel.docx", text: SSN, reason: "SSN — PII", author: "A. Lawyer", label: undefined }, ctxFor(dir, async request => { asked.push(request.metadata) }))
    const down = spyOn(globalThis, "fetch").mockImplementation((async () => new Response("unavailable", { status: 503 })) as typeof fetch)
    let failed: Awaited<ReturnType<typeof redact.execute>>
    try {
      failed = await run()
    } finally {
      down.mockRestore()
    }
    if (!failed || typeof failed === "string") throw new Error(String(failed))
    expect(failed.output).toContain("doc.haus's earlier versions of the document may still hold it. Run this redaction again")
    // The text is gone from the file and logged: running it again sends the
    // redaction save of the file as it is now.
    const redacted = await Bun.file(file).bytes()
    const sent: { origin: string | null; bytes: Uint8Array }[] = []
    const up = spyOn(globalThis, "fetch").mockImplementation((async (_input: string | URL | Request, init?: RequestInit) => {
      const upload = (init?.body as FormData).get("file") as File
      sent.push({ origin: new Headers(init?.headers).get("x-dochaus-origin"), bytes: new Uint8Array(await upload.arrayBuffer()) })
      return Response.json({ ok: true })
    }) as typeof fetch)
    let again: Awaited<ReturnType<typeof redact.execute>>
    try {
      again = await run()
    } finally {
      up.mockRestore()
    }
    expect(sent).toEqual([{ origin: "redaction", bytes: redacted }])
    // The approval card is told this run saves again rather than removes text.
    expect(asked.map(metadata => metadata.resave ?? false)).toEqual([false, true])
    if (!again || typeof again === "string") throw new Error(String(again))
    expect(again.output).toContain("already removed")
    expect(again.output).not.toContain("WARNING")
    expect(listRedactions(dir)).toHaveLength(1)
    // Text never redacted from this document is still not found.
    expect(await redact.execute({ document: "personnel.docx", text: "not in this document", reason: "test", author: undefined, label: undefined }, ctxFor(dir)))
      .toContain("Text not found")
  })

  test("doc.haus's own records are not a document it redacts", async () => {
    const { dir } = await matterWith(buildFixture())
    const kept = path.join(dir, ".dochaus", "revisions", "document", "kept.docx")
    mkdirSync(path.dirname(kept), { recursive: true })
    const bytes = buildFixture()
    await Bun.write(kept, bytes)
    // Links into the records, and the matter folder under another spelling of its path.
    symlinkSync(path.join(dir, ".dochaus"), path.join(dir, "records"))
    symlinkSync(kept, path.join(dir, "alias.docx"))
    const spelled = `${dir}-spelled`
    symlinkSync(dir, spelled)
    const documents = [
      kept, ".dochaus/revisions/document/kept.docx", ".DOCHAUS/revisions/document/kept.docx", ".dochau\u017f/revisions/document/kept.docx",
      path.join(realpathSync(dir), ".dochaus", "revisions", "document", "kept.docx"), "records/revisions/document/kept.docx", "alias.docx",
    ]
    for (const directory of [dir, spelled])
      for (const document of documents)
        expect(await redact.execute({ document, text: SSN, reason: "SSN — PII", author: "A. Lawyer", label: undefined }, ctxFor(directory)))
          .toBe(`Document not found in this matter: ${document}`)
    expect(await Bun.file(kept).bytes()).toEqual(bytes)
  })
})

// A hand-built synthetic package with one fictional surname planted in every
// kind of place a DOCX can carry text: projected story text (body with a value
// split across differently formatted runs, header, footer, comment, footnote),
// docProps, attribute values (image alt text, hyperlink tooltip, mailto target,
// document variable, people-part author under the w15 prefix), a .vml part, and
// two binary parts (an image and a package thumbnail).
const NAME = "Thistlewood"
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const DRAWING_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0))
const JPEG = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0xff, 0xd9])
const para = (inner: string) => `<w:p>${inner}</w:p>`
const run = (text: string, props = "") => `<w:r>${props}<w:t xml:space="preserve">${text}</w:t></w:r>`

function plantedParts(): Record<string, Uint8Array> {
  const body = [
    para(run("Synthetic workshop record.")),
    para(run("Contact ") + run("Thistle") + run("wood", "<w:rPr><w:b/></w:rPr>") + run(" about the record.")),
    para(`<w:hyperlink r:id="rIdLink" w:tooltip="Write to ${NAME}">${run("Email the records desk")}</w:hyperlink>`),
    para(`<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1" descr="Portrait of ${NAME}" title="${NAME}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`),
    para(`<w:commentRangeStart w:id="0"/>${run("Reviewed.")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`),
    para(`${run("See note.")}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>`),
  ].join("")
  const override = (part: string, type: string) => `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${type}"/>`
  return {
    "[Content_Types].xml": strToU8(`${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>${override("word/document.xml", "wordprocessingml.document.main+xml")}${override("word/settings.xml", "wordprocessingml.settings+xml")}${override("word/header1.xml", "wordprocessingml.header+xml")}${override("word/footer1.xml", "wordprocessingml.footer+xml")}${override("word/comments.xml", "wordprocessingml.comments+xml")}${override("word/footnotes.xml", "wordprocessingml.footnotes+xml")}${override("word/people.xml", "wordprocessingml.people+xml")}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${override("docProps/app.xml", "extended-properties+xml")}</Types>`),
    "_rels/.rels": strToU8(`${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/></Relationships>`),
    "word/document.xml": strToU8(`${XML_DECL}<w:document ${W_NS} ${DRAWING_NS}><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSet" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rIdHdr" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdFtr" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rIdCom" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdPeople" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/><Relationship Id="rIdImg" Type="${REL}/image" Target="media/image1.png"/><Relationship Id="rIdLink" Type="${REL}/hyperlink" Target="mailto:records@example.test?subject=${NAME}" TargetMode="External"/></Relationships>`),
    "word/settings.xml": strToU8(`${XML_DECL}<w:settings ${W_NS}><w:docVars><w:docVar w:name="ClientSurname" w:val="${NAME}"/></w:docVars></w:settings>`),
    "word/header1.xml": strToU8(`${XML_DECL}<w:hdr ${W_NS}>${para(run(`${NAME} file`))}</w:hdr>`),
    "word/footer1.xml": strToU8(`${XML_DECL}<w:ftr ${W_NS}>${para(run(`Prepared for ${NAME}`))}</w:ftr>`),
    "word/comments.xml": strToU8(`${XML_DECL}<w:comments ${W_NS}><w:comment w:id="0" w:author="Reviewer" w:initials="R">${para(run(`Ask ${NAME} to confirm.`))}</w:comment></w:comments>`),
    "word/footnotes.xml": strToU8(`${XML_DECL}<w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1">${para("<w:r><w:separator/></w:r>")}</w:footnote><w:footnote w:type="continuationSeparator" w:id="0">${para("<w:r><w:continuationSeparator/></w:r>")}</w:footnote><w:footnote w:id="1">${para(run(`Source: ${NAME} interview.`))}</w:footnote></w:footnotes>`),
    "word/people.xml": strToU8(`${XML_DECL}<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:person w15:author="${NAME}"><w15:presenceInfo w15:providerId="None" w15:userId="${NAME}"/></w15:person></w15:people>`),
    "word/vmlDrawing1.vml": strToU8(`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="s1" alt="${NAME} signature" o:title="${NAME}"><v:textbox><div>Signed by ${NAME}</div></v:textbox></v:shape></xml>`),
    "word/media/image1.png": PNG,
    "docProps/core.xml": strToU8(`${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>${NAME}</dc:creator></cp:coreProperties>`),
    "docProps/app.xml": strToU8(`${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>${NAME} Holdings</Company></Properties>`),
    "docProps/thumbnail.jpeg": JPEG,
  }
}

// A chart title whose value is split across two DrawingML runs: outside the
// Docxodus projection, so only the package walk can see it.
const SPLIT_CHART = strToU8(`${XML_DECL}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Thistle</a:t></a:r><a:r><a:rPr b="1"/><a:t>wood</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)

// Independent oracle: the needle in any part's raw bytes (attributes included)
// or in any part's text once every tag is stripped (so split runs rejoin).
function leaks(bytes: Uint8Array, needle: string) {
  return Object.entries(unzipSync(bytes)).flatMap(([name, data]) => {
    const raw = strFromU8(data, true)
    const decoded = raw.replace(/(?:%[0-9A-Fa-f]{2})+/g, (run) => {
      try {
        return decodeURIComponent(run)
      } catch {
        return run
      }
    })
    return [
      decoded !== raw && decoded.includes(needle) && `${name}: percent-decoded`,
      raw.includes(needle) && `${name}: raw`,
      raw.replace(/<[^>]*>/g, "").includes(needle) && `${name}: joined text`,
    ].filter((leak): leak is string => !!leak)
  })
}

describe("package residue scan", () => {
  test("reads attribute values, .vml parts and text split across runs", () => {
    const parts = plantedParts()
    parts["word/charts/chart1.xml"] = SPLIT_CHART
    const where = docxPackageResidue(zipSync(parts), NAME).hits.map((hit) => `${hit.part} | ${hit.where}`)
    expect(where).toEqual(
      expect.arrayContaining([
        "word/document.xml | attribute descr on <wp:docPr>",
        "word/document.xml | attribute title on <wp:docPr>",
        "word/document.xml | attribute w:tooltip on <w:hyperlink>",
        "word/document.xml | text split across runs in <w:p>",
        "word/_rels/document.xml.rels | attribute Target on <Relationship>",
        "word/settings.xml | attribute w:val on <w:docVar>",
        "word/people.xml | attribute w15:author on <w15:person>",
        "word/vmlDrawing1.vml | attribute alt on <v:shape>",
        "word/vmlDrawing1.vml | text in <div>",
        "word/charts/chart1.xml | text split across runs in <a:p>",
      ]),
    )
  })

  test("a deliberately unscrubbed attribute fails the residue check", () => {
    const parts = unzipSync(scrubDocxPackage(zipSync(plantedParts()), NAME, "[REDACTED]").bytes)
    expect(residueParts(zipSync(parts), NAME)).not.toContain("word/settings.xml")
    parts["word/settings.xml"] = strToU8(`${XML_DECL}<w:settings ${W_NS}><w:docVars><w:docVar w:name="ClientSurname" w:val="${NAME}"/></w:docVars></w:settings>`)
    expect(docxPackageResidue(zipSync(parts), NAME).hits).toContainEqual({ part: "word/settings.xml", where: "attribute w:val on <w:docVar>" })
  })

  test("lists every part it cannot read, including a thumbnail, and flags a literal copy inside one", () => {
    const parts = plantedParts()
    expect(docxPackageResidue(zipSync(parts), NAME).unread.sort()).toEqual(["docProps/thumbnail.jpeg", "word/media/image1.png"])
    parts["docProps/thumbnail.jpeg"] = new Uint8Array([...JPEG.subarray(0, 12), ...strToU8(`Artist ${NAME}`), 0xff, 0xd9])
    expect(docxPackageResidue(zipSync(parts), NAME).hits).toContainEqual({
      part: "docProps/thumbnail.jpeg",
      where: "literal bytes in a part the scan cannot parse",
    })
  })

  test("a number inside a longer typed identifier is not the redacted value", () => {
    const parts = plantedParts()
    parts["word/settings.xml"] = strToU8(`${XML_DECL}<w:settings ${W_NS}><w:rsids><w:rsidRoot w:val="00451200"/></w:rsids></w:settings>`)
    expect(residueParts(zipSync(parts), "4512")).toEqual([])
  })
})

describe("redact tool across the whole package", () => {
  test("removes the value from every readable part and scopes its claim to what the scan read", async () => {
    const { dir, file } = await matterWith(zipSync(plantedParts()))
    expect(leaks(await Bun.file(file).bytes(), NAME).length).toBeGreaterThan(0)

    const result = await redact.execute(
      { document: "personnel.docx", text: NAME, reason: "Client identity", author: undefined, label: undefined },
      ctxFor(dir),
    )
    if (!result || typeof result === "string") throw new Error(String(result))
    const redacted = await Bun.file(file).bytes()
    expect(leaks(redacted, NAME)).toEqual([])
    expect(docxPackageResidue(redacted, NAME).hits).toEqual([])
    // The message claims only what the scan read and names what it could not.
    expect(result.output).toContain("Verified zero residue in the text the scan reads")
    expect(result.output).not.toMatch(/zero residue\.(\s|$)/)
    expect(result.output).toContain("2 part(s) not read")
    expect(result.output).toContain("docProps/thumbnail.jpeg")
    expect(result.output).toContain("word/media/image1.png")
    expect(result.metadata?.unreadParts).toEqual(expect.arrayContaining(["docProps/thumbnail.jpeg", "word/media/image1.png"]))
    // Rewritten parts stay well formed: Docxodus reopens the file and the
    // mailto link and alt text carry the label.
    const session = dx.openDocxSession(redacted, {})
    expect(session.project().markdown).toContain("Contact")
    session.close()
    const parts = unzipSync(redacted)
    expect(strFromU8(parts["word/_rels/document.xml.rels"])).toContain('Target="mailto:records@example.test?subject=[REDACTED]"')
    expect(strFromU8(parts["word/document.xml"])).toContain('descr="Portrait of [REDACTED]"')
  })

  test("aborts naming the part when a split value sits outside the Docxodus projection", async () => {
    const parts = plantedParts()
    parts["word/charts/chart1.xml"] = SPLIT_CHART
    const { dir, file } = await matterWith(zipSync(parts))
    const before = await Bun.file(file).bytes()
    const result = await redact.execute(
      { document: "personnel.docx", text: NAME, reason: "Client identity", author: undefined, label: undefined },
      ctxFor(dir),
    )
    expect(result).toBeString()
    expect(result as string).toContain("Redaction aborted, document unchanged")
    expect(result as string).toContain("word/charts/chart1.xml (text split across runs in <a:p>)")
    expect(await Bun.file(file).bytes()).toEqual(before)
    expect(listRedactions(dir)).toEqual([])
  })
})

// A two-word needle whose only copies are percent-encoded inside URLs: an external
// mailto target, a HYPERLINK field instruction, and an internal relationship
// target (a part name, which cannot be rewritten).
const FULL_NAME = `Orla ${NAME}`
function percentEncodedParts() {
  const parts = plantedParts()
  const rels = strFromU8(parts["word/_rels/document.xml.rels"])
  parts["word/_rels/document.xml.rels"] = strToU8(rels.replace(`subject=${NAME}"`, `subject=Orla%20${NAME}"`))
  const field = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "mailto:records@example.test?subject=Orla%20${NAME}" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("records")}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]).replace("<w:sectPr>", `${field}<w:sectPr>`))
  return parts
}

describe("percent-encoded copies in URLs", () => {
  test("the scan decodes percent escapes, and a part-name copy stays residue after the scrub", () => {
    const parts = percentEncodedParts()
    const where = docxPackageResidue(zipSync(parts), FULL_NAME).hits.map((hit) => `${hit.part} | ${hit.where}`)
    expect(where).toEqual(
      expect.arrayContaining([
        "word/_rels/document.xml.rels | attribute Target on <Relationship> (percent-encoded)",
        "word/document.xml | text in <w:instrText> (percent-encoded)",
      ]),
    )
    const rels = strFromU8(parts["word/_rels/document.xml.rels"])
    parts["word/_rels/document.xml.rels"] = strToU8(rels.replace("media/image1.png", "media/Orla%20Thistlewood.png"))
    const scrubbed = scrubDocxPackage(zipSync(parts), FULL_NAME, "[REDACTED]").bytes
    expect(docxPackageResidue(scrubbed, FULL_NAME).hits).toEqual([
      { part: "word/_rels/document.xml.rels", where: "attribute Target on <Relationship> (percent-encoded)" },
    ])
  })

  test("the redact tool rewrites encoded copies with an encoded label", async () => {
    const { dir, file } = await matterWith(zipSync(percentEncodedParts()))
    expect(leaks(await Bun.file(file).bytes(), FULL_NAME).length).toBeGreaterThan(0)
    const result = await redact.execute(
      { document: "personnel.docx", text: FULL_NAME, reason: "Client identity", author: undefined, label: undefined },
      ctxFor(dir),
    )
    if (!result || typeof result === "string") throw new Error(String(result))
    expect(result.output).toContain("Verified zero residue in the text the scan reads")
    const redacted = await Bun.file(file).bytes()
    expect(leaks(redacted, FULL_NAME)).toEqual([])
    expect(docxPackageResidue(redacted, FULL_NAME).hits).toEqual([])
    const parts = unzipSync(redacted)
    expect(strFromU8(parts["word/_rels/document.xml.rels"])).toContain('Target="mailto:records@example.test?subject=%5BREDACTED%5D"')
    expect(strFromU8(parts["word/document.xml"])).toContain('HYPERLINK "mailto:records@example.test?subject=%5BREDACTED%5D"')
  })
})
