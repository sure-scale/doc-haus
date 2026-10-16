import { describe, expect, test } from "bun:test"
import { mkdtempSync } from "node:fs"
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
})
