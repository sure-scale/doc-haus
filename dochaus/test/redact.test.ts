// True-redaction guarantees, exercised through the real redact tool against
// authored synthetic documents rendered by the product helper — no mocks beyond the
// tool ctx. Run from dochaus/: bun test test/redact.test.ts
//
// Lives in test/, not tool/: the engine's ToolRegistry imports every .ts under
// {tool,tools}/ as a custom tool, and importing a bun:test file outside the
// test runner throws ("Cannot use describe outside of the test runner").
import { describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdtempSync } from "node:fs"
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

function ctxFor(directory: string, ask: () => Promise<void> = async () => {}) {
  return {
    sessionID: "test",
    messageID: "test",
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

async function matterWith(bytes: Uint8Array) {
  const dir = mkdtempSync(path.join(tmpdir(), "redact-matter-"))
  const file = path.join(dir, "personnel.docx")
  await Bun.write(file, bytes)
  return { dir, file }
}

describe("metadata scrub", () => {
  test("removes the needle from docProps text nodes and reports residue correctly", () => {
    const fixture = buildFixture()
    expect(docxPackageResidue(fixture, SSN)).toContain("docProps/core.xml")
    const scrub = scrubDocxPackage(fixture, SSN, "[REDACTED]")
    expect(scrub.hits["docProps/core.xml"]).toBe(2)
    expect(docxPackageResidue(scrub.bytes, SSN)).not.toContain("docProps/core.xml")
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
    if (typeof result === "string") throw new Error(result)
    expect(result.output).toContain("3 text occurrence(s)")
    expect(result.output).toContain("re-indexing failed")
    expect(result.metadata?.reingested).toBe(false)

    // The file itself is clean: raw package residue zero, Docxodus flat-text zero.
    const redacted = await Bun.file(file).bytes()
    expect(docxPackageResidue(redacted, SSN)).toEqual([])
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
})
