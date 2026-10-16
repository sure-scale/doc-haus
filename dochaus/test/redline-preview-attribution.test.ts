import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../lib/render"
import { createDocxLocalization } from "../lib/docx-amendments"
import { exportDocxComments } from "../lib/docx-comments"
import { bake, buildRedlined, RedlinePreviewConflictError } from "../../services/ingest/src/redline"

const previous = process.env.WORKSPACE_ROOT
const roots: string[] = []
afterEach(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
})

test("revision-bound preview preserves the stored proposal author and defaults only missing authors", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dh-preview-attribution-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify({ attorney: "Saved Lawyer", firm: "Saved Firm" }))
  const original = (await renderMarkdownToDocx("Current review period is five days.")).bytes
  const snapshot = new Uint8Array(original)
  const packet = await createDocxLocalization(original, [{ id: "period", before_quote: "five days", replacement: "seven days" }])
  const target = packet.candidates[0].target_handle
  const row = {
    id: 1, doc_path: "note.docx", doc_name: "note.docx", scope: "phrase" as const,
    find_text: "five days", old_text: "five days", new_text: "seven days", review_comment: "Instructed period.", anchor_id: null, created_at: 1, status: "pending" as const,
    target_json: JSON.stringify({ version: 1, source_sha256: createHash("sha256").update(original).digest("hex"), target_handle: target, amendment_id: "period" }),
  }
  for (const [author, expected] of [["Explicit Reviewer", "Explicit Reviewer"], ["", "Saved Lawyer"]]) {
    const redline = await buildRedlined(original, [{ ...row, author }])
    const body = new DOMParser().parseFromString(strFromU8(unzipSync(redline)["word/document.xml"]), "application/xml")
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    const revisions = ["ins", "del"].flatMap(name => Array.from(body.getElementsByTagNameNS(W, name)))
    expect(revisions.length).toBeGreaterThan(0)
    expect(new Set(revisions.map(node => node.getAttributeNS(W, "author")))).toEqual(new Set([expected]))
  }
  expect(original).toEqual(snapshot)
})

test("mixed-author previews refuse false shared attribution while exact same-author edits still export", async () => {
  const rendered = (await renderMarkdownToDocx("Current review period is five days and retention is three months.\n\nHistorical review period was five days.")).bytes
  const original = exportDocxComments(rendered, [{ id: "history", text: "Earlier reviewer explanation.", author: "Historical Reviewer", anchor: { quote: "Historical review period" } }]).bytes
  const snapshot = new Uint8Array(original)
  const amendments = [
    { id: "period", before_quote: "five days", replacement: "seven days" },
    { id: "retention", before_quote: "three months", replacement: "six months" },
  ]
  const packet = await createDocxLocalization(original, amendments)
  const rows = amendments.map((amendment, index) => ({
    id: index + 1, doc_path: "note.docx", doc_name: "note.docx", scope: "phrase" as const,
    find_text: amendment.before_quote, old_text: amendment.before_quote, new_text: amendment.replacement,
    author: `Explicit Reviewer ${index === 0 ? "A" : "B"}`, review_comment: "Instructed change.",
    anchor_id: null, created_at: 1, status: "pending" as const,
    target_json: JSON.stringify({ version: 1, source_sha256: packet.revision.source_sha256,
      target_handle: packet.candidates.find(candidate => candidate.amendment_id === amendment.id)!.target_handle, amendment_id: amendment.id }),
  }))
  const stored = structuredClone(rows)
  await expect(buildRedlined(original, rows)).rejects.toBeInstanceOf(RedlinePreviewConflictError)
  await expect(buildRedlined(original, rows)).rejects.toThrow("different authors")
  // Legacy replay has the same comparison-wide author constraint.
  await expect(buildRedlined(original, rows.map(row => ({ ...row, target_json: null })))).rejects.toThrow("different authors")
  expect(rows).toEqual(stored)
  expect(original).toEqual(snapshot)
  expect(strFromU8(unzipSync(original)["word/comments.xml"]!)).toContain('w:author="Historical Reviewer"')

  const sameAuthor = rows.map(row => ({ ...row, author: "Explicit Reviewer A" }))
  const preview = await buildRedlined(original, sameAuthor)
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const body = new DOMParser().parseFromString(strFromU8(unzipSync(preview)["word/document.xml"]!), "application/xml")
  const revisions = ["ins", "del"].flatMap(name => Array.from(body.getElementsByTagNameNS(W, name)))
  expect(revisions.length).toBeGreaterThan(0)
  expect(new Set(revisions.map(node => node.getAttributeNS(W, "author")))).toEqual(new Set(["Explicit Reviewer A"]))
  const accepted = await bake(original, rows)
  const text = new DOMParser().parseFromString(strFromU8(unzipSync(accepted)["word/document.xml"]!), "application/xml")
  expect(Array.from(text.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join(""))
    .toBe("Current review period is seven days and retention is six months.Historical review period was five days.")
  // A stale second author is not part of the preview; exact target validation
  // still takes precedence over the mixed-author export limitation.
  await expect(buildRedlined(original, [rows[0]!, { ...rows[1]!, target_json: JSON.stringify({ ...JSON.parse(rows[1]!.target_json), source_sha256: "0".repeat(64) }) }])).resolves.toBeInstanceOf(Uint8Array)
  await expect(buildRedlined(original, [rows[0]!, { ...rows[1]!, target_json: JSON.stringify({ ...JSON.parse(rows[1]!.target_json), target_handle: "unknown-target" }) }])).rejects.toThrow("unknown_or_conflicting_target_handle")
  await expect(buildRedlined(original, [rows[0]!, { ...rows[0]!, id: 3, author: "Explicit Reviewer B" }])).rejects.toThrow("overlapping_docx_amendments")
  expect(original).toEqual(snapshot)
})
