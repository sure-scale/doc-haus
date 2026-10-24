import { describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, realpathSync, symlinkSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { pendingRedlinesForDoc } from "../lib/redlines"
import { renderMarkdownToDocx } from "../lib/render"

// Authored test content; the renderer supplies the shipped, source-built DOCX seed.
const fixtureBytes = (await renderMarkdownToDocx([
  "# Synthetic workshop note",
  "## Confidential Information",
  "The sample workshop inventory records wooden counters and paper labels.",
  "## Entire Agreement",
  "This synthetic note collects the sample inventory instructions in one place.",
].join("\n\n"))).bytes
const trackedChanges = (await import("../tool/tracked-changes")).default
const redline = (await import("../tool/redline")).default

function ctxFor(directory: string) {
  return {
    sessionID: "redline-comment-test",
    messageID: "redline-comment-test",
    agent: "redliner",
    directory,
    worktree: directory,
    abort: new AbortController().signal,
    metadata() {},
    ask: async () => {},
  }
}

async function documentFixture() {
  const directory = mkdtempSync(path.join(tmpdir(), "redline-review-comment-"))
  const document = path.join(directory, "agreement.docx")
  await Bun.write(document, fixtureBytes)
  return { directory, document }
}

describe("redline review comments", () => {
  test("stores a non-operative review comment beside a surgical pending change", async () => {
    const fixture = await documentFixture()
    const result = await trackedChanges.execute(
      {
        document: "agreement.docx",
        find: "Confidential Information",
        replace: "Protected Information",
        comment: "Aligns the defined term with the terminology used throughout the agreement.",
        author: "Reviewing Lawyer",
      },
      ctxFor(fixture.directory),
    )

    expect(result).toBeObject()
    if (typeof result === "string") throw new Error(result)
    expect(result.metadata?.comment).toBe("Aligns the defined term with the terminology used throughout the agreement.")
    expect(pendingRedlinesForDoc(fixture.directory, fixture.document)).toEqual([
      expect.objectContaining({
        author: "Reviewing Lawyer",
        review_comment: "Aligns the defined term with the terminology used throughout the agreement.",
      }),
    ])
  })

  test("stores the review comment separately from clause replacement text", async () => {
    const fixture = await documentFixture()
    const result = await redline.execute(
      {
        document: "agreement.docx",
        clause: "Entire Agreement",
        replacement: "Entire Agreement. This document records the complete agreement between the parties.",
        comment: "Clarifies the integration wording while retaining a single operative clause.",
        author: "Reviewing Lawyer",
      },
      ctxFor(fixture.directory),
    )

    expect(result).toBeObject()
    if (typeof result === "string") throw new Error(result)
    const pending = pendingRedlinesForDoc(fixture.directory, fixture.document)
    expect(pending[0]?.new_text).not.toContain("Clarifies the integration wording")
    expect(pending[0]?.review_comment).toBe(
      "Clarifies the integration wording while retaining a single operative clause.",
    )
  })

  test("a clause copied with tracked-change labels is refused with the reason", async () => {
    const fixture = await documentFixture()
    const result = await redline.execute(
      {
        document: "agreement.docx",
        clause: "The sample workshop inventory records [Tracked insertion by Counterparty: wooden] counters",
        replacement: "The sample workshop inventory records counters.",
        comment: "Restores the earlier inventory wording.",
        author: "Reviewing Lawyer",
      },
      ctxFor(fixture.directory),
    )

    expect(result).toBe("agreement.docx contains tracked changes. Accept or reject them in Word before proposing changes to it.")
    expect(pendingRedlinesForDoc(fixture.directory, fixture.document)).toEqual([])
  })

  test("doc.haus's own records are not a document to propose changes to", async () => {
    const fixture = await documentFixture()
    const kept = path.join(fixture.directory, ".dochaus", "revisions", "document", "kept.docx")
    mkdirSync(path.dirname(kept), { recursive: true })
    await Bun.write(kept, fixtureBytes)
    const comment = "Aligns the defined term with the terminology used throughout the agreement."
    // Links into the records, and the matter folder under another spelling of its path.
    symlinkSync(path.join(fixture.directory, ".dochaus"), path.join(fixture.directory, "records"))
    symlinkSync(kept, path.join(fixture.directory, "alias.docx"))
    const spelled = `${fixture.directory}-spelled`
    symlinkSync(fixture.directory, spelled)
    const documents = [
      kept, ".dochaus/revisions/document/kept.docx", ".DOCHAUS/revisions/document/kept.docx", ".dochau\u017f/revisions/document/kept.docx",
      path.join(realpathSync(fixture.directory), ".dochaus", "revisions", "document", "kept.docx"), "records/revisions/document/kept.docx", "alias.docx",
    ]
    for (const directory of [fixture.directory, spelled])
      for (const document of documents) {
        expect(await trackedChanges.execute({ document, find: "Confidential Information", replace: "Protected Information", comment }, ctxFor(directory)))
          .toBe(`Document not found in this matter: ${document}`)
        expect(await redline.execute({ document, clause: "Entire Agreement", replacement: "Entire Agreement.", comment }, ctxFor(directory)))
          .toBe(`Document not found in this matter: ${document}`)
      }
    expect(pendingRedlinesForDoc(fixture.directory, kept)).toEqual([])
    expect(pendingRedlinesForDoc(fixture.directory, path.join(fixture.directory, "alias.docx"))).toEqual([])
    // A document saved in the matter is found under either spelling.
    expect(await trackedChanges.execute({ document: path.join(realpathSync(fixture.directory), "agreement.docx"), find: "Confidential Information", replace: "Protected Information", comment }, ctxFor(spelled)))
      .toBeObject()
  })
})
