import { expect, test } from "bun:test"
import { Hono } from "hono"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { Document, Packer, Paragraph } from "docx"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { compareParagraphs, readComparisonParagraphs } from "../../../dochaus/lib/version-compare"
import { readDocxMarkup } from "../../../dochaus/lib/docx-markup"
import { correspondHistoryParagraphs } from "../../../dochaus/lib/clause-history"
import { openDb } from "./db"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { readCurrentVersion, readDocumentVersions, readVersionBytes, recordVersionEvent } from "./document-versions"
import { clauseHistoryRoutes, installWordClauseHistory, readClauseHistory, readClauseHistorySource } from "./clause-history"
import type { WordGrant } from "./word-grants"
import { prepareDocumentReceivedEvent, readDocumentExchangeRequest, relinkDocumentReceivedDocument } from "./document-exchanges"
import { createDocxLocalization } from "../../../dochaus/lib/docx-amendments"
import { paragraphTargets } from "../../../dochaus/lib/drafting-amendment"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
async function docx(paragraphs: { text: string; id?: string }[], comment?: string) {
  const parts = unzipSync(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs.map(row => new Paragraph(row.text)) }] })))
  const document = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  document.documentElement.setAttribute("xmlns:w14", W14)
  const nodes = Array.from(document.getElementsByTagNameNS(W, "p"))
  paragraphs.forEach((row, index) => { if (row.id) nodes[index]!.setAttributeNS(W14, "w14:paraId", row.id) })
  if (comment) {
    const start = document.createElementNS(W, "w:commentRangeStart"); start.setAttributeNS(W, "w:id", "7")
    const end = document.createElementNS(W, "w:commentRangeEnd"); end.setAttributeNS(W, "w:id", "7")
    const reference = document.createElementNS(W, "w:commentReference"); reference.setAttributeNS(W, "w:id", "7")
    const run = document.createElementNS(W, "w:r"); run.appendChild(reference)
    nodes[0]!.insertBefore(start, nodes[0]!.firstChild); nodes[0]!.appendChild(end); nodes[0]!.appendChild(run)
    const escaped = comment.replace(/&/g, "&amp;").replace(/</g, "&lt;")
    parts["word/comments.xml"] = strToU8(`<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="Recorded reviewer" w:date="2026-11-01T12:00:00Z"><w:p><w:r><w:t>${escaped}</w:t></w:r></w:p></w:comment></w:comments>`)
    // Packer already declares the comments content type and relationship.
  }
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(document))
  return Buffer.from(zipSync(parts))
}

test("#201 inventories include unchanged full paragraphs and actual persisted identities", async () => {
  const left = await docx([{ text: "First clause remains.", id: "12345678" }, { text: "Second clause is Monday.", id: "23456789" }])
  const right = await docx([{ text: "Second clause is Tuesday.", id: "23456789" }, { text: "First clause remains.", id: "12345678" }])
  const compared = await compareParagraphs(left, right)
  expect(compared.left.map(row => row.text)).toEqual(["First clause remains.", "Second clause is Monday."])
  expect(compared.right.map(row => row.text)).toEqual(["Second clause is Tuesday.", "First clause remains."])
  expect(compared.left[1]!.nativeId).toBe("23456789")
  expect(correspondHistoryParagraphs(compared, [compared.left[1]!])).toMatchObject({ status: "verified", evidence: "word-paragraph-id", paragraphs: [{ text: "Second clause is Tuesday." }] })
}, 30_000)

test("actual #201 split edges cannot silently reduce a clause to its first fragment", async () => {
  const left = await docx([{ text: "Delivery is Monday. Exceptions require written agreement.", id: "12345678" }])
  const right = await docx([{ text: "Delivery is Monday.", id: "12345678" }, { text: "Exceptions require written agreement.", id: "23456789" }])
  const compared = await compareParagraphs(left, right)
  const relation = correspondHistoryParagraphs(compared, compared.left)
  expect(relation.status === "unresolved" || relation.paragraphs.length === 2).toBe(true)
  if (relation.status === "verified") expect(relation.paragraphs.map(paragraph => paragraph.text)).toEqual(["Delivery is Monday.", "Exceptions require written agreement."])
}, 30_000)

test("full history wording keeps literal tabs and soft breaks while page and column breaks stay layout", async () => {
  const parts = unzipSync(await docx([{ text: "Placeholder", id: "12345678" }]))
  const content = '<w:r><w:t xml:space="preserve">Alpha 💬</w:t><w:tab/><w:t>Beta</w:t><w:br/><w:t>Next</w:t><w:cr/><w:t>After CR</w:t><w:br w:type="page"/><w:t>Page</w:t><w:br w:type="column"/><w:t>Column</w:t><w:br w:type="textWrapping"/><w:t>End</w:t></w:r>'
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace(/<w:r>[\s\S]*?<\/w:r>/, content))
  const paragraphs = await readComparisonParagraphs(zipSync(parts))
  expect(paragraphs[0]!.text).toBe("Alpha 💬\tBeta\nNext\nAfter CRPageColumn\nEnd")
}, 30_000)

test("linked returns retain physical source identity and require exact wording plus a live association for history sources", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-linked-clause-"))
  try {
    await writeDeliverable(directory, "Target.docx", await docx([{ text: "Delivery is Tuesday.", id: "12345678" }]))
    const target = readCurrentVersion(directory, "Target.docx")!
    await writeDeliverable(directory, "Uploaded copy.docx", await docx([{ text: "Delivery is Monday.", id: "12345678" }]))
    const source = readCurrentVersion(directory, "Uploaded copy.docx")!
    await writeDeliverableRevision(directory, "Uploaded copy.docx", await docx([{ text: "Delivery is Tuesday.", id: "12345678" }], "Tuesday is the recorded delivery note."), {
      version: { origin: "revision", event: prepareDocumentReceivedEvent(directory, { baseVersionId: source.current.id, requestId: "linked-receipt" }) },
    })
    const receipt = readDocumentExchangeRequest(directory, "linked-receipt")!
    relinkDocumentReceivedDocument(directory, { exchangeId: receipt.id, targetDocumentId: target.document.id, baseVersionId: target.current.id, expectedDocumentId: source.document.id, expectedBaseVersionId: source.current.id, requestId: "link-clause" })
    const request = { anchorVersion: target.current.id, paragraph: { scope: "body" as const, text: "Delivery is Tuesday." } }
    const result = await readClauseHistory(directory, "Target.docx", request)
    const linked = result.rounds.find(round => round.version.linkedReturn)!
    expect(linked.version).toMatchObject({ id: receipt.versionId, number: 2, linkedReturn: { sourceDocumentId: source.document.id, sourceDocumentName: "Uploaded copy.docx", targetDocumentId: target.document.id, exchangeId: receipt.id } })
    expect(linked.evidence).toBe("exact-wording")
    expect(linked.reasons[0]!.source.exchangeId).toBe(receipt.id)
    const record = await readClauseHistorySource(directory, "Target.docx", linked.reasons[0]!.source)
    expect(record.reason?.text).toBe("Tuesday is the recorded delivery note.")
    expect(record.version.linkedReturn?.sourceDocumentName).toBe("Uploaded copy.docx")
    await expect(readClauseHistorySource(directory, "Target.docx", { ...linked.reasons[0]!.source, exchangeId: undefined })).rejects.toThrow("not in this document")
    expect(readDocumentVersions(directory, "Target.docx").versions).toHaveLength(1)
    await writeDeliverable(directory, "Different copy.docx", await docx([{ text: "Delivery is Monday.", id: "12345678" }]))
    const different = readCurrentVersion(directory, "Different copy.docx")!
    await writeDeliverableRevision(directory, "Different copy.docx", await docx([{ text: "Delivery is Friday.", id: "12345678" }]), {
      version: { origin: "revision", event: prepareDocumentReceivedEvent(directory, { baseVersionId: different.current.id, requestId: "different-linked-receipt" }) },
    })
    const differentReceipt = readDocumentExchangeRequest(directory, "different-linked-receipt")!
    relinkDocumentReceivedDocument(directory, { exchangeId: differentReceipt.id, targetDocumentId: target.document.id, baseVersionId: target.current.id, expectedDocumentId: different.document.id, expectedBaseVersionId: different.current.id, requestId: "link-different-clause" })
    const unresolved = await readClauseHistory(directory, "Target.docx", request)
    expect(unresolved.rounds.find(round => round.version.id === differentReceipt.versionId)).toMatchObject({ status: "unresolved", paragraphs: [], reasons: [] })
    const db = openDb(directory)
    try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [receipt.versionId]) } finally { db.close() }
    const removed = await readClauseHistory(directory, "Target.docx", request)
    expect(removed.rounds.find(round => round.version.linkedReturn)).toMatchObject({ status: "unavailable", paragraphs: [], reasons: [] })
    await expect(readClauseHistorySource(directory, "Target.docx", linked.reasons[0]!.source)).rejects.toThrow("no longer available")
  } finally { rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("history returns exact rounds and source-bound comments without changing bytes, versions or ownership", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-clause-history-"))
  try {
    await writeDeliverable(directory, "Draft.docx", await docx([{ text: "Delivery is Monday.", id: "12345678" }, { text: "The rest remains.", id: "23456789" }], "Keep Monday because the receiving team is available."), { version: { origin: "upload", actor: "Uploader" } })
    const first = readCurrentVersion(directory, "Draft.docx")!.current
    await writeDeliverableRevision(directory, "Draft.docx", await docx([{ text: "The rest remains.", id: "23456789" }, { text: "Delivery is Tuesday.", id: "12345678" }]), { version: { origin: "word", actor: "Editor" } })
    const current = readCurrentVersion(directory, "Draft.docx")!.current
    expect((await readDocxMarkup(readVersionBytes(directory, first)!)).comments).toHaveLength(1)
    expect((await compareParagraphs(readVersionBytes(directory, current)!, readVersionBytes(directory, first)!)).right).toHaveLength(2)
    const before = readFileSync(path.join(directory, "Draft.docx"))
    const count = readDocumentVersions(directory, "Draft.docx").versions.length
    const result = await readClauseHistory(directory, "Draft.docx", { anchorVersion: current.id, paragraph: { scope: "body", text: "Delivery is Tuesday." } })
    expect(result.anchor.status).toBe("verified")
    expect(result.rounds.map(round => [round.status, round.paragraphs.map(paragraph => paragraph.text)])).toEqual([["verified", ["Delivery is Monday."]], ["verified", ["Delivery is Tuesday."]]])
    expect(result.rounds[0]!.version.actor).toBe("Uploader")
    expect(result.rounds[0]!.reasons).toMatchObject([{ kind: "comment", text: "Keep Monday because the receiving team is available.", author: "Recorded reviewer", source: { versionId: first.id, hash: first.sha256, id: "7" } }])
    const source = await readClauseHistorySource(directory, "Draft.docx", result.rounds[0]!.reasons[0]!.source)
    expect(source.reason?.text).toBe("Keep Monday because the receiving team is available.")
    expect(result.rounds[1]!.reasonStatus).toBe("not-recorded")
    expect(readFileSync(path.join(directory, "Draft.docx"))).toEqual(before)
    expect(readDocumentVersions(directory, "Draft.docx").versions).toHaveLength(count)
    const db = openDb(directory)
    try {
      expect(db.query("SELECT COUNT(*) AS n FROM word_edit_leases").get()).toEqual({ n: 0 })
      expect(db.query("SELECT COUNT(*) AS n FROM word_sessions").get()).toEqual({ n: 0 })
      db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [first.id])
    } finally { db.close() }
    const redacted = await readClauseHistory(directory, "Draft.docx", { anchorVersion: current.id, paragraph: { scope: "body", text: "Delivery is Tuesday." } })
    expect(redacted.rounds[0]).toMatchObject({ status: "unavailable", paragraphs: [], reasons: [] })
    await expect(readClauseHistorySource(directory, "Draft.docx", result.rounds[0]!.reasons[0]!.source)).rejects.toThrow("removed by redaction")
    const unsaved = await readClauseHistory(directory, "Draft.docx", { anchorVersion: current.id, paragraph: { scope: "body", text: "Unsaved different wording." } })
    expect(unsaved.anchor.status).toBe("unresolved")
    expect(unsaved.rounds.every(round => round.paragraphs.length === 0)).toBe(true)
  } finally { rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("tracked authors and local decisions retain exact source identity without inventing rationale", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-clause-decision-"))
  try {
    const parts = unzipSync(await docx([{ text: "Delivery is Monday.", id: "12345678" }]))
    const xml = strFromU8(parts["word/document.xml"]!).replace("</w:p>", '<w:ins w:id="91" w:author="Revision reviewer" w:date="2026-11-02T10:00:00Z"><w:r><w:t xml:space="preserve"> Confirmed in writing.</w:t></w:r></w:ins></w:p>')
    parts["word/document.xml"] = strToU8(xml)
    const tracked = Buffer.from(zipSync(parts))
    const markup = await readDocxMarkup(tracked)
    expect(markup.revisions).toHaveLength(1)
    await writeDeliverable(directory, "Draft.docx", tracked)
    const first = readCurrentVersion(directory, "Draft.docx")!.current
    const wording = "Delivery is Monday. Confirmed in writing."
    await writeDeliverableRevision(directory, "Draft.docx", await docx([{ text: wording, id: "12345678" }]), {
      version: { origin: "accept", actor: "Local lawyer", event: { kind: "accept", actor: "Local lawyer", detail: { markup: [{ id: markup.revisions[0]!.id, decision: "accept", outcome: "applied" }] } } },
    })
    const current = readCurrentVersion(directory, "Draft.docx")!.current
    const result = await readClauseHistory(directory, "Draft.docx", { anchorVersion: current.id, paragraph: { scope: "body", text: wording } })
    expect(result.rounds[0]!.changes).toMatchObject([{ author: "Revision reviewer", date: "2026-11-02T10:00:00Z", text: " Confirmed in writing.", source: { versionId: first.id, kind: "revision" } }])
    expect(result.rounds[0]!.reasons).toMatchObject([{ kind: "decision", text: "", author: "Local lawyer", decision: "accept", outcome: "applied" }])
    expect(result.rounds[0]!.reasonStatus).toBe("not-recorded")
    expect((await readClauseHistorySource(directory, "Draft.docx", result.rounds[0]!.reasons[0]!.source)).reason?.decision).toBe("accept")
  } finally { rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("existing accepted/rejected redline IDs show proposed reasons separately from human decisions and pin mutable source records", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-clause-proposal-"))
  try {
    const name = "Draft.docx"
    await writeDeliverable(directory, name, await docx([{ text: "Delivery is Monday.", id: "12345678" }, { text: "The rest remains.", id: "23456789" }]))
    const first = readCurrentVersion(directory, name)!.current
    const firstAnchor = (await readComparisonParagraphs(readVersionBytes(directory, first)!))[0]!.anchorId
    const insert = (find: string, replacement: string, reason: string, anchor: string | null, target: unknown, status: string) => {
      const db = openDb(directory)
      try { return Number(db.run("INSERT INTO redlines(doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,target_json,status,created_at) VALUES(?,?,'phrase',?,?,?,?,?,?,?,?,?)", [path.join(directory, name), name, find, find, replacement, "Proposal author", reason, anchor, target ? JSON.stringify(target) : null, status, 100]).lastInsertRowid) }
      finally { db.close() }
    }
    const accepted = insert("Monday", "Tuesday", "Proposed rationale: allow the receiving team another day.", firstAnchor, null, "accepted")
    await writeDeliverableRevision(directory, name, await docx([{ text: "Delivery is Tuesday.", id: "12345678" }, { text: "The rest remains.", id: "23456789" }]), {
      version: { origin: "accept", actor: "Human reviewer", event: { kind: "accept", actor: "Human reviewer", detail: { redline_ids: [accepted] } } },
    })
    const current = readCurrentVersion(directory, name)!.current, bytes = readVersionBytes(directory, current)!
    const packet = await createDocxLocalization(bytes, [{ id: "stored-proposal", before_quote: "Tuesday", replacement: "Wednesday" }])
    const rejected = insert("Tuesday", "Wednesday", "Proposed rationale: a later day may be convenient.", null,
      { version: 1, source_sha256: current.sha256, amendment_id: "stored-proposal", target_handle: packet.candidates[0]!.target_handle }, "rejected")
    const stale = insert("Tuesday", "Friday", "This stale rationale must not appear.", null,
      { version: 1, source_sha256: first.sha256, amendment_id: "stale", target_handle: packet.candidates[0]!.target_handle }, "rejected")
    const targets = await paragraphTargets(bytes), pointed = targets[0]!
    const paragraphBound = insert("Tuesday", "Thursday", "Proposed rationale bound to this exact paragraph.", null,
      { version: 2, amendment_id: "paragraph-bound", ...(pointed.pid ? { pid: pointed.pid } : { para_id: pointed.para_id }), paragraph_sha256: pointed.sha256, occurrence: 0, base_sha256: current.sha256 }, "rejected")
    const db = openDb(directory)
    try { recordVersionEvent(db, directory, name, { kind: "reject", actor: "Another human reviewer", detail: { redline_ids: [rejected, stale, paragraphBound] } }) } finally { db.close() }
    const request = { anchorVersion: current.id, paragraph: { scope: "body" as const, text: "Delivery is Tuesday." } }
    const result = await readClauseHistory(directory, name, request)
    expect(result.rounds[0]!.reasons).toMatchObject([
      { kind: "proposal", author: "Proposal author", text: "Proposed rationale: allow the receiving team another day.", proposal: { id: accepted, oldText: "Monday", newText: "Tuesday" } },
      { kind: "decision", author: "Human reviewer", text: "", decision: "accept" },
    ])
    const latest = result.rounds.find(round => round.version.id === current.id)!
    expect(latest.reasons.filter(reason => reason.kind === "proposal").map(reason => reason.proposal!.id)).toEqual([rejected, paragraphBound])
    expect(latest.reasons.filter(reason => reason.kind === "decision")).toMatchObject([{ author: "Another human reviewer", decision: "reject" }, { author: "Another human reviewer", decision: "reject" }])
    expect(JSON.stringify(result)).not.toContain("This stale rationale must not appear.")
    const reason = result.rounds[0]!.reasons[0]!
    expect(reason.source.recordHash).toHaveLength(64)
    expect((await readClauseHistorySource(directory, name, reason.source)).reason?.proposal?.oldText).toBe("Monday")
    await expect(readClauseHistorySource(directory, name, { ...reason.source, recordHash: undefined })).rejects.toThrow("record changed")
    const change = openDb(directory)
    try { change.run("UPDATE redlines SET review_comment='Changed proposed rationale' WHERE id=?", [accepted]) } finally { change.close() }
    await expect(readClauseHistorySource(directory, name, reason.source)).rejects.toThrow("record changed")
    const redaction = openDb(directory)
    try { redaction.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [first.id]) } finally { redaction.close() }
    await expect(readClauseHistorySource(directory, name, reason.source)).rejects.toThrow("removed by redaction")
  } finally { rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("web and Word routes are read-only and reject wrong documents, revoked grants and damaged current files", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-clause-history-http-"))
  try {
    await writeDeliverable(directory, "Draft.docx", await docx([{ text: "A full saved paragraph.", id: "12345678" }]))
    const current = readCurrentVersion(directory, "Draft.docx")!
    const request = { anchorVersion: current.current.id, paragraph: { scope: "body", text: "A full saved paragraph." } }
    const app = new Hono(); app.route("/matters", clauseHistoryRoutes({ directory: () => directory }))
    let revoked = false
    const grant: WordGrant = { document_id: current.document.id, issued_version: current.current.id, origin: "https://fixture.invalid", token_hash: "fixture", revoked_at: null, overlap: 0 }
    installWordClauseHistory(app, () => revoked ? undefined : { matter: { id: "one", title: "Synthetic", dir: directory }, grant, target: { name: "Draft.docx", docPath: path.join(directory, "Draft.docx") } })
    const post = (url: string, body: object) => app.request(url, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
    expect((await post("/matters/one/documents/clause-history?name=Draft.docx", request)).status).toBe(200)
    expect((await post("/word/chat/history", request)).status).toBe(200)
    expect((await post("/word/chat/history", { ...request, anchorVersion: "not-this-document" })).status).toBe(404)
    revoked = true
    expect((await post("/word/chat/history", request)).status).toBe(401)
    revoked = false
    const interrupted = new Hono()
    let authentications = 0
    installWordClauseHistory(interrupted, () => ++authentications > 1 ? undefined : { matter: { id: "one", title: "Synthetic", dir: directory }, grant, target: { name: "Draft.docx", docPath: path.join(directory, "Draft.docx") } })
    const lost = await interrupted.request("/word/chat/history", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(request) })
    expect(lost.status).toBe(401)
    expect(await lost.text()).not.toContain("A full saved paragraph.")
    writeFileSync(path.join(directory, "Draft.docx"), "unregistered outside edit")
    expect((await post("/word/chat/history", request)).status).toBe(409)
    expect((await post("/matters/one/documents/clause-history?name=../Draft.docx", request)).status).toBe(400)
    const paragraphs = await readComparisonParagraphs(await docx([{ text: "Same." }, { text: "Same." }]))
    expect(paragraphs).toHaveLength(2)
  } finally { rmSync(directory, { recursive: true, force: true }) }
}, 30_000)
