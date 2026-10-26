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
  // A review comment needs its stored author, so only an unannotated proposal takes the saved default.
  await expect(buildRedlined(original, [{ ...row, author: "" }])).rejects.toThrow("stored author")
  for (const [author, review_comment, expected] of [["Explicit Reviewer", "Instructed period.", "Explicit Reviewer"], ["", "", "Saved Lawyer"]]) {
    const redline = await buildRedlined(original, [{ ...row, author, review_comment }])
    const body = new DOMParser().parseFromString(strFromU8(unzipSync(redline)["word/document.xml"]), "application/xml")
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    const revisions = ["ins", "del"].flatMap(name => Array.from(body.getElementsByTagNameNS(W, name)))
    expect(revisions.length).toBeGreaterThan(0)
    expect(new Set(revisions.map(node => node.getAttributeNS(W, "author")))).toEqual(new Set([expected]))
  }
  expect(original).toEqual(snapshot)
})

test("a mixed-author preview writes each change under its own author; an older replay refuses false shared attribution", async () => {
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
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const revisions = (bytes: Uint8Array) => Array.from(new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
    .getElementsByTagNameNS(W, "*")).filter(node => node.localName === "ins" || node.localName === "del")
    .map(node => [node.localName, node.getAttributeNS(W, "author"), node.textContent])
  // Each change is written under its own proposal's author, as Accept writes it.
  expect(revisions(await buildRedlined(original, rows))).toEqual([["del", "Explicit Reviewer A", "five"], ["ins", "Explicit Reviewer A", "seven"],
    ["del", "Explicit Reviewer B", "three"], ["ins", "Explicit Reviewer B", "six"]])
  // Legacy replay without review notes is a comparison, which takes one author for every revision.
  await expect(buildRedlined(original, rows.map(row => ({ ...row, target_json: null, review_comment: "" })))).rejects.toBeInstanceOf(RedlinePreviewConflictError)
  await expect(buildRedlined(original, rows.map(row => ({ ...row, target_json: null, review_comment: "" })))).rejects.toThrow("different authors")
  expect(rows).toEqual(stored)
  expect(original).toEqual(snapshot)
  expect(strFromU8(unzipSync(original)["word/comments.xml"]!)).toContain('w:author="Historical Reviewer"')

  const sameAuthor = rows.map(row => ({ ...row, author: "Explicit Reviewer A" }))
  const preview = await buildRedlined(original, sameAuthor)
  expect(revisions(preview).length).toBeGreaterThan(0)
  expect(new Set(revisions(preview).map(([, author]) => author))).toEqual(new Set(["Explicit Reviewer A"]))
  const accepted = await bake(original, rows)
  const text = new DOMParser().parseFromString(strFromU8(unzipSync(accepted)["word/document.xml"]!), "application/xml")
  expect(Array.from(text.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join(""))
    .toBe("Current review period is seven days and retention is six months.Historical review period was five days.")
  // A stale, unknown or overlapping second target fails the whole preview closed.
  await expect(buildRedlined(original, [rows[0]!, { ...rows[1]!, target_json: JSON.stringify({ ...JSON.parse(rows[1]!.target_json), source_sha256: "0".repeat(64) }) }])).rejects.toThrow("no longer resolves")
  await expect(buildRedlined(original, [rows[0]!, { ...rows[1]!, target_json: JSON.stringify({ ...JSON.parse(rows[1]!.target_json), target_handle: "unknown-target" }) }])).rejects.toThrow("no longer resolves")
  await expect(buildRedlined(original, [rows[0]!, { ...rows[0]!, id: 3, author: "Explicit Reviewer B" }])).rejects.toThrow("no longer resolves")
  expect(original).toEqual(snapshot)
})
