import { afterAll, beforeAll, expect, test } from "bun:test"
import { Hono } from "hono"
import { mkdtempSync, mkdirSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { installWordChat } from "./word-chat"
import { wordChatMaterials, wordChatSuggestionBatches, validateWordChatContext, type WordChatAuth } from "./word-chat-context"
import { claimWordEditLease, clearWordEditLease } from "../../../dochaus/lib/word-edit-lease"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { openDb } from "./db"
import { readCurrentVersion } from "./document-versions"
import type { WordChatSend, WordChatTurn } from "../../../dochaus/lib/word-chat-contract"

type Session = { id: string; parentID?: string; metadata: Record<string, unknown>; location: { directory: string }; title: string; agent: string; model: { providerID: string; id: string }; time: { created: string; updated: string; idle?: string }; outcome?: string }
const sessions = new Map<string, Session>(), messages = new Map<string, Record<string, unknown>[]>()
const prompts: string[] = []
let directory: string, app: Hono, auth: WordChatAuth, next = 0, generations = 0, loseResponse = false, hold = false, inventCitation = false, amend = false, approveReview = true, deletion = false
let modelProposal: { before: string; after: string } | undefined
const oldRuntime = process.env.DOCHAUS_RUNTIME_DIR, oldWorkspace = process.env.WORKSPACE_ROOT
const transport = Object.assign(async (input: RequestInfo | URL, init?: RequestInit) => {
  const url = new URL(String(input)); const method = init?.method ?? "GET"
  const payload = typeof init?.body === "string" ? JSON.parse(init.body) as Record<string, unknown> : {}
  const match = /^\/api\/session\/([^/]+)(.*)$/.exec(url.pathname)
  const envelope = (data: unknown) => Response.json({ data })
  if ((url.pathname === "/api/session" || url.pathname === "/api/dochaus/session") && method === "POST") {
    const id = String(payload.id ?? `ses_test_${++next}`)
    if (sessions.has(id)) return envelope(sessions.get(id))
    const parent = sessions.get(String(payload.parentID)); const time = new Date(Date.now() + next++).toISOString()
    const session = { ...payload, id, location: parent?.location ?? payload.location, time: { created: time, updated: time } } as Session
    sessions.set(id, session); messages.set(id, []); return envelope(session)
  }
  if (url.pathname === "/api/session" && method === "GET") return Response.json({ data: [...sessions.values()].filter(session =>
    url.searchParams.get("parentID") === "null" ? !session.parentID : url.searchParams.has("parentID") ? session.parentID === url.searchParams.get("parentID") : true), cursor: {} })
  if (!match) throw new Error(`Unexpected engine route ${method} ${url.pathname}`)
  const session = sessions.get(match[1]!)
  if (!session) return Response.json({}, { status: 404 })
  const tail = match[2]
  if (!tail && method === "GET") return envelope(session)
  if (!tail && method === "PATCH") { Object.assign(session, payload); return new Response(null, { status: 204 }) }
  if (tail === "/message") return Response.json({ data: [...messages.get(session.id)!].reverse(), cursor: {} })
  if (tail === "/inbox") return envelope([])
  if (tail === "/interrupt") { session.outcome = "interrupted"; return Response.json({ interrupted: true }) }
  if (tail === "/prompt") {
    const rows = messages.get(session.id)!
    const admission = { id: payload.id, sessionID: session.id, type: "user", payload: { text: payload.text }, delivery: payload.delivery }
    if (!rows.some(message => message.id === payload.id)) {
      generations++; prompts.push(String(payload.text))
      rows.push({ id: payload.id, type: "user", text: payload.text, time: { created: new Date().toISOString() } })
      if (!hold) {
        const sources = (session.metadata.sources ?? []) as { id: string; quote: string }[]
        const source = sources.find(source => source.id !== "word") ?? sources[0]
        const output: Record<string, unknown> = session.metadata.type === "word-chat-review" ? { approved: approveReview, issues: approveReview ? [] : ["The replacement loses a qualification."] } : { answer: `Answer ${source?.quote ?? ""}`, citations: source ? [{ sourceId: inventCitation ? "foreign-source" : source.id, quote: source.quote.slice(0, 100) }] : [], suggestions: [] }
        if (amend && session.metadata.type !== "word-chat-review") output.proposal = { before: deletion ? "Tuesday " : context.text, after: deletion ? "" : context.text.replace("Tuesday", "Wednesday"), ...modelProposal, explanation: "Requested delivery update", openItems: ["Confirm the date with the client"] }
        rows.push({ id: `msg_answer_${next++}`, type: "assistant", content: [{ type: "text", text: JSON.stringify(output) }], time: { created: new Date().toISOString(), completed: new Date().toISOString() } })
        session.time.idle = new Date().toISOString(); session.outcome = "succeeded"
      }
    }
    if (loseResponse) { loseResponse = false; throw new Error("Lost admitted response") }
    return envelope(admission)
  }
  throw new Error(`Unexpected engine route ${method} ${url.pathname}`)
}, { preconnect() {} }) as typeof fetch
const post = (action: string, body: object = {}) => app.request(`/word/chat/${action}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
const context = { text: "The current Word buffer uses Tuesday delivery.", paragraphs: [{ id: "11", text: "The current Word buffer uses Tuesday delivery." }], saved: false, capturedAt: "2026-01-01T00:00:00Z", truncated: false }
const send = (conversationId: string, scope: "document" | "matter", text: string, requestId = crypto.randomUUID(), attachments: WordChatSend["attachments"] = []) => post("send", { conversationId, scope, text, requestId, attachments, context })
const newConversation = async () => (await (await post("conversation")).json() as { conversation: { id: string } }).conversation.id

beforeAll(async () => {
  directory = realpathSync(mkdtempSync(path.join(tmpdir(), "word-chat-scoped-")))
  process.env.DOCHAUS_RUNTIME_DIR = path.join(directory, "runtime"); process.env.WORKSPACE_ROOT = directory
  mkdirSync(path.join(directory, "runtime", "config", "opencode"), { recursive: true })
  writeFileSync(path.join(directory, "runtime", "config", "opencode", "opencode.json"), JSON.stringify({ model: "synthetic/model" }))
  const docx = (text: string) => Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] }))
  await writeDeliverable(directory, "Draft.docx", await docx("The saved version uses Monday delivery."))
  await writeDeliverable(directory, "Other.docx", await docx("DISTINCTIVE_EVIDENCE is a blue agreement."))
  const current = readCurrentVersion(directory, "Draft.docx")!
  const db = openDb(directory)
  try {
    const other = db.query("SELECT id,doc_path FROM documents WHERE name='Other.docx'").get() as { id: number; doc_path: string }
    db.run("INSERT INTO chunks(document_id,doc_path,doc_name,section,chunk_index,text,char_start,char_end,embedding) VALUES (?,?,?,?,?,?,?,?,?)", [other.id, other.doc_path, "Other.docx", "Terms", 0, "DISTINCTIVE_EVIDENCE is a blue agreement.", 0, 40, new Uint8Array(16)])
  } finally { db.close() }
  auth = { matter: { id: "synthetic", title: "Synthetic matter", dir: directory }, target: { name: "Draft.docx", docPath: path.join(directory, "Draft.docx") },
    grant: { token_hash: "synthetic", document_id: current.document.id, origin: "https://localhost", issued_version: current.current.id, revoked_at: null, overlap: 0 } }
  app = new Hono(); installWordChat(app, () => auth, { fetch: transport, engineURL: "http://synthetic.invalid" })
}, 30_000)
afterAll(() => {
  if (oldRuntime === undefined) delete process.env.DOCHAUS_RUNTIME_DIR; else process.env.DOCHAUS_RUNTIME_DIR = oldRuntime
  if (oldWorkspace === undefined) delete process.env.WORKSPACE_ROOT; else process.env.WORKSPACE_ROOT = oldWorkspace
  rmSync(directory, { recursive: true, force: true })
})

test("document-only cannot inherit matter evidence; explicit attachment deliberately admits it", async () => {
  const id = await newConversation()
  expect((await send(id, "matter", "What is DISTINCTIVE_EVIDENCE?")).status).toBe(200)
  expect(prompts.at(-1)).toContain("DISTINCTIVE_EVIDENCE is a blue agreement.")
  expect((await send(id, "document", "Explain the current delivery day.")).status).toBe(200)
  expect(prompts.at(-1)).not.toContain("DISTINCTIVE_EVIDENCE")
  expect(prompts.at(-1)).not.toContain("Monday delivery")
  expect(prompts.at(-1)).toContain("Tuesday delivery")
  const source = wordChatMaterials(auth).find(material => material.name === "Other.docx")!
  expect((await send(id, "document", "Explain the attached evidence", crypto.randomUUID(), [source])).status).toBe(200)
  expect(prompts.at(-1)).toContain("DISTINCTIVE_EVIDENCE is a blue agreement.")
  expect((await send(id, "document", "Explain this", crypto.randomUUID(), [{ ...source, hash: "0".repeat(64) }])).status).toBe(409)
})
test("native root and child ownership refuse another document; invented citations never reach the answer", async () => {
  const id = await newConversation()
  const old = auth
  auth = { ...auth, grant: { ...auth.grant, document_id: "other-document" } }
  expect((await post("conversation", { conversationId: id })).status).toBe(404)
  auth = old
  inventCitation = true
  const response = await send(id, "document", "Explain delivery")
  inventCitation = false
  const turn = (await response.json() as { turn: WordChatTurn }).turn
  expect(turn.status).toBe("error"); expect(turn.answer).toBe(""); expect(turn.sources).toEqual([])
})
test("lost prompt admission is reconciled with the exact native ID and generates once", async () => {
  const id = await newConversation(), requestId = crypto.randomUUID()
  const before = generations
  loseResponse = true
  expect((await send(id, "document", "Explain delivery", requestId)).status).toBe(503)
  expect((await send(id, "document", "Explain delivery", requestId)).status).toBe(200)
  expect(generations - before).toBe(1)
  expect((await send(id, "document", "A changed request", requestId)).status).toBe(409)
})
test("Stop cancels one scoped request and later polling cannot expose its late answer", async () => {
  const id = await newConversation(); hold = true
  const response = await send(id, "document", "Explain delivery")
  hold = false
  const turn = (await response.json() as { turn: WordChatTurn }).turn
  expect(turn.status).toBe("running")
  const stopped = await post("cancel", { conversationId: id, turnId: turn.id })
  expect((await stopped.json() as { turn: WordChatTurn }).turn.status).toBe("cancelled")
  const polled = await post("poll", { conversationId: id, turnId: turn.id })
  expect((await polled.json() as { turn: WordChatTurn }).turn.answer).toBe("")
})
test("suggestion passage batches stay small, sample across a large document, and disclose coverage", () => {
  const paragraphs = Array.from({ length: 30 }, (_, index) => ({ id: String(index), text: `Passage ${index}. ${"x".repeat(2000)}` }))
  const split = wordChatSuggestionBatches({ ...context, paragraphs, text: paragraphs.map(paragraph => paragraph.text).join("\n") }, false)
  expect(split.batches).toHaveLength(3); expect(split.total).toBeGreaterThan(3)
  expect(split.batches[0]?.[0]?.id).toBe("0"); expect(split.batches.at(-1)?.at(-1)?.id).toBe("29")
  expect(split.batches.every(batch => batch.reduce((size, paragraph) => size + paragraph.text.length, 0) <= 8000)).toBe(true)
})

test("amendment preparation is checked by a separate denied-tool native reviewer before Apply", async () => {
  const id = await newConversation(); amend = true
  const before = generations
  const response = await send(id, "document", "Rewrite this for Wednesday delivery")
  const turn = (await response.json() as { turn: WordChatTurn }).turn
  amend = false
  expect(turn.proposal?.before).toBe(context.text)
  expect(turn.proposal?.after).toContain("Wednesday")
  expect(turn.proposal?.openItems).toContain("Confirm the date with the client")
  expect(generations - before).toBe(2)
  const review = [...sessions.values()].find(session => session.parentID === turn.id)!
  expect(review.metadata.type).toBe("word-chat-review")
  expect((review as unknown as { permissions: unknown }).permissions).toEqual([{ action: "*", resource: "*", effect: "deny" }])
  expect((await post("prepare", { conversationId: id, turnId: turn.id, ooxml: "not authorized" })).status).toBe(409)
  approveReview = false; amend = true
  const rejected = (await (await send(id, "document", "Rewrite delivery for next week")).json() as { turn: WordChatTurn }).turn
  amend = false; approveReview = true
  expect(rejected.proposal).toBeUndefined()
  expect(rejected.error).toContain("needs revision")
})
test("source version changes invalidate late results and remove access to earlier quotes", async () => {
  const id = await newConversation()
  const source = wordChatMaterials(auth).find(material => material.name === "Other.docx")!
  const turn = (await (await send(id, "document", "Explain attached source", crypto.randomUUID(), [source])).json() as { turn: WordChatTurn }).turn
  expect(turn.status).toBe("complete")
  writeFileSync(path.join(directory, "Other.docx"), "outside replacement")
  const polled = (await (await post("poll", { conversationId: id, turnId: turn.id })).json() as { turn: WordChatTurn }).turn
  expect(polled.status).toBe("error"); expect(polled.answer).toBe(""); expect(polled.sources).toEqual([])
})
test("chunked multipart and JSON payloads are bounded before parsing", async () => {
  for (const [route, length, contentType] of [["send", 2 * 1024 * 1024 + 1, "application/json"], ["upload", 17 * 1024 * 1024 + 1, "multipart/form-data; boundary=synthetic"]] as const) {
    const stream = new ReadableStream<Uint8Array>({ start(controller) { controller.enqueue(new Uint8Array(length)); controller.close() } })
    const response = await app.request(`/word/chat/${route}`, { method: "POST", headers: { "content-type": contentType }, body: stream, duplex: "half" } as RequestInit)
    expect(response.status).toBe(413)
  }
})

test("deleting selected wording prepares an exact reviewed empty replacement", async () => {
  const id = await newConversation(); amend = true; deletion = true
  const response = await send(id, "document", "Remove Tuesday from the delivery wording")
  amend = false; deletion = false
  const turn = (await response.json() as { turn: WordChatTurn }).turn
  expect(turn.status).toBe("complete")
  expect(turn.proposal?.before).toBe("Tuesday ")
  expect(turn.proposal?.after).toBe("")
})

test("model line feeds normalize to exact Word paragraph boundaries before review and refinement", async () => {
  const id = await newConversation()
  const multiline = { ...context, text: "Delivery is Tuesday.\rNotice is written.", paragraphs: [{ id: "11", text: "Delivery is Tuesday." }, { id: "12", text: "Notice is written." }] }
  amend = true; modelProposal = { before: "Delivery is Tuesday.\nNotice is written.", after: "Delivery is Wednesday.\r\nNotice is written." }
  const response = await post("send", { conversationId: id, scope: "document", text: "Revise delivery to Wednesday", requestId: crypto.randomUUID(), attachments: [], context: multiline })
  amend = false; modelProposal = undefined
  const turn = (await response.json() as { turn: WordChatTurn }).turn
  expect(turn.proposal?.before).toBe(multiline.text)
  expect(turn.proposal?.after).toBe("Delivery is Wednesday.\rNotice is written.")
  expect(turn.proposal?.target.paragraphs.map(paragraph => paragraph.id)).toEqual(["11", "12"])
  amend = true; modelProposal = { before: "Delivery is Tuesday.\nNotice is written.", after: "Delivery is Thursday.\nNotice is written." }
  const refined = await post("send", { conversationId: id, scope: "document", text: "Make it Thursday instead", requestId: crypto.randomUUID(), attachments: [], context: multiline, proposal: turn.proposal })
  amend = false; modelProposal = undefined
  const revised = (await refined.json() as { turn: WordChatTurn }).turn
  expect(revised.proposal?.target).toEqual(turn.proposal!.target)
  expect(revised.proposal?.after).toBe("Delivery is Thursday.\rNotice is written.")
})
test("a long single Word paragraph is accepted within the complete bounded snapshot", () => {
  const text = "A".repeat(100_000)
  expect(validateWordChatContext({ ...context, text, paragraphs: [{ id: "long", text }] })).toBe(true)
  const oversized = "A".repeat(512_001)
  expect(validateWordChatContext({ ...context, text: oversized, paragraphs: [{ id: "large", text: oversized }] })).toBe(false)
})

test("prepare rechecks ownership and canonical base after asynchronous authoring", async () => {
  const id = await newConversation(); amend = true
  const turn = (await (await send(id, "document", "Rewrite delivery to Wednesday")).json() as { turn: WordChatTurn }).turn
  amend = false
  const current = readCurrentVersion(directory, "Draft.docx")!
  const db = openDb(directory)
  db.run("INSERT INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)", [auth.grant.token_hash, auth.grant.document_id, auth.grant.origin, auth.grant.issued_version, Date.now()])
  const identity = { document: auth.grant.document_id, tokenHash: auth.grant.token_hash, session: "synthetic-pane-session" }
  let owner = claimWordEditLease(db, identity)
  let duringAuthor: (() => Promise<void>) | undefined
  const preparing = new Hono()
  installWordChat(preparing, () => auth, { fetch: transport, engineURL: "http://synthetic.invalid", authorAmendment: async (_ooxml, proposal) => {
    await duringAuthor?.()
    return { proposalId: proposal.id, ooxml: "synthetic-pure-artifact", paragraphs: [proposal.after], revisions: [], author: "doc.haus" }
  } })
  const prepare = () => preparing.request("/word/chat/prepare", { method: "POST", headers: { "content-type": "application/json", "x-dochaus-word-session": owner.session, "x-dochaus-word-lease": owner.generation, "x-dochaus-word-base-version": current.current.id },
    body: JSON.stringify({ conversationId: id, turnId: turn.id, ooxml: "synthetic-original-ooxml" }) })
  // Starting authority passes; only the change during the await causes refusal.
  expect((await prepare()).status).toBe(200)
  duringAuthor = async () => { clearWordEditLease(db, auth.grant.document_id) }
  expect((await prepare()).status).toBe(409)
  owner = claimWordEditLease(db, identity)
  duringAuthor = async () => {
    const bytes = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("A concurrent acknowledged Word save.")] }] }))
    await writeDeliverableRevision(directory, "Draft.docx", bytes, { priorSha256: current.current.sha256, keepWorkingRecord: true, wordLeaseOwner: owner, version: { origin: "word" } })
  }
  expect((await prepare()).status).toBe(409)
  clearWordEditLease(db, auth.grant.document_id); db.close()
})
