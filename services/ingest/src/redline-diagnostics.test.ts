import { expect, setSystemTime, spyOn, test } from "bun:test"
import { Document, Packer, Paragraph } from "docx"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import type { RedlineRow } from "./db"
import { docxodus } from "./docxodus"
import { bake, buildRedlined } from "./redline"

function proposal(scope: RedlineRow["scope"], findText: string, newText: string): RedlineRow {
  return {
    id: 1,
    doc_path: "/synthetic-matter/source.docx",
    doc_name: "source.docx",
    scope,
    find_text: findText,
    old_text: findText,
    new_text: newText,
    author: "Synthetic reviewer",
    review_comment: "",
    anchor_id: null,
    status: "pending",
    created_at: 0,
  }
}

function originalDocument() {
  return Packer.toBuffer(new Document({ sections: [{ children: [
    new Paragraph("Original marker describes the working note."),
  ] }] }))
}

test.each(["phrase", "clause"] as const)("missing %s redlines log fingerprints while preserving skipped bytes and acceptance errors", async scope => {
  const original = await originalDocument()
  const before = Buffer.from(original)
  const canary = `private-anchor-${crypto.randomUUID()}`
  const row = proposal(scope, canary, `private-replacement-${crypto.randomUUID()}`)
  const log = spyOn(console, "warn").mockImplementation(() => {})
  try {
    expect(await buildRedlined(original, [row])).toEqual(before)
    const failure = await bake(original, [row]).then(() => undefined, (error: unknown) => error)
    expect(failure).toBeInstanceOf(Error)
    if (!(failure instanceof Error)) throw new Error("Missing-anchor acceptance did not retain its error")
    expect(failure.message).toContain(canary)
    expect(failure.message).toContain(scope === "clause" ? "Clause no longer found" : "Text no longer found")
    expect(log.mock.calls).toEqual([
      ["[redline] skipping in redlined view", { error_kind: "string", error_ref: draftingTraceReference(failure.message) }],
    ])
    expect(JSON.stringify(log.mock.calls)).not.toContain(canary)
    expect(JSON.stringify(log.mock.calls)).not.toContain(row.new_text)
    expect(JSON.stringify(log.mock.calls)).not.toContain(row.doc_path)
    expect(original).toEqual(before)
  } finally {
    log.mockRestore()
  }
})

test("a skipped redline still allows valid proposals to render without changing canonical input", async () => {
  const original = await originalDocument()
  const before = Buffer.from(original)
  const canary = `private-missing-${crypto.randomUUID()}`
  const replacement = `Updated${crypto.randomUUID().replaceAll("-", "")}`
  const log = spyOn(console, "warn").mockImplementation(() => {})
  try {
    const rendered = await buildRedlined(original, [
      proposal("phrase", canary, "Skipped proposal wording"),
      { ...proposal("phrase", "Original marker", replacement), id: 2 },
    ])
    const revisions = await (await docxodus()).getRevisions(rendered)
    expect(revisions.length).toBeGreaterThan(0)
    expect(JSON.stringify(revisions)).toContain(replacement)
    expect(log.mock.calls).toHaveLength(1)
    expect(log.mock.calls[0]?.[0]).toBe("[redline] skipping in redlined view")
    expect(log.mock.calls[0]?.[1]).toMatchObject({ error_kind: "string", error_ref: expect.stringMatching(/^sha256:[a-f0-9]{64}$/) })
    expect(JSON.stringify(log.mock.calls)).not.toContain(canary)
    expect(original).toEqual(before)
  } finally {
    log.mockRestore()
  }
})

test("a tracked-changes export dates its revisions when exported, not with the comparator's fixed date", async () => {
  setSystemTime(new Date("2026-03-04T05:06:07.890Z"))
  try {
    const rendered = await buildRedlined(await originalDocument(), [proposal("phrase", "Original marker", "Updated marker")])
    const revisions = await (await docxodus()).getRevisions(rendered)
    expect(revisions.length).toBeGreaterThan(0)
    expect(revisions.map(revision => revision.date)).toEqual(revisions.map(() => "2026-03-04T05:06:07Z"))
  } finally {
    setSystemTime()
  }
})
