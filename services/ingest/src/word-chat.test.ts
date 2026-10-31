import { afterAll, beforeAll, expect, test } from "bun:test"
import { Hono } from "hono"
import { createHash, randomBytes } from "node:crypto"
import { wordDocumentRoutes } from "./word-documents"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { mkdtempSync, mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { installWordChat } from "./word-chat"
import { WordFormattingError } from "./word-chat-amendments"
import { wordChatMaterials, wordChatSuggestionBatches, validateWordChatContext, type WordChatAuth } from "./word-chat-context"
import { beginWordNativeOperation, settleWordNativeOperation, claimWordEditLease, clearWordEditLease } from "../../../dochaus/lib/word-edit-lease"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { openDb } from "./db"
import { readCurrentVersion } from "./document-versions"
import { matterDocumentName } from "./ingest"
import { writeWordChatPreferences } from "./word-chat-preferences"
import type { WordChatMaterial, WordChatSend, WordChatTurn } from "../../../dochaus/lib/word-chat-contract"

type Session = { id: string; parentID?: string; metadata: Record<string, unknown>; location: { directory: string }; title: string; agent: string; model: { providerID: string; id: string }; time: { created: string; updated: string; idle?: string }; outcome?: string }
const sessions = new Map<string, Session>(), messages = new Map<string, Record<string, unknown>[]>(), inboxes = new Map<string, Record<string, unknown>[]>()
const prompts: string[] = [], routerPrompts: string[] = [], titleActions: string[] = [], admissions: { session: string; tail: string; payload: Record<string, unknown> }[] = []
let directory: string, app: Hono, auth: WordChatAuth, next = 0, generations = 0, loseResponse = false, hold = false, inventCitation = false, amend = false, approveReview = true, deletion = false
let modelProposal: { before: string; after?: string; formatting?: unknown } | undefined
let modelAnswer: string | undefined
let modelRawOutput: string | undefined, reviewerRawOutput: string | undefined
let modelStructuredRawOutput: string | undefined
let routerAgent: "qa" | "redliner" = "qa", routerOutput: string | undefined, loseContextResponse = false
let reviewerOutcome: "succeeded" | "failed" | "interrupted" = "succeeded"
let mainOutcome: "succeeded" | "failed" | "interrupted" = "succeeded", queueReview = false
const stoppedInputs: string[] = []
// Existing synthetic fixtures describe answers in the previous test shape.
// Emit the admitted structured wire format; legacy sessions are tested explicitly.
function structuredFixture(raw: string): string {
  try {
    const value = JSON.parse(raw) as {answer?: unknown;citations?: unknown[];answerParts?: unknown}
    if (typeof value.answer !== "string" || value.answerParts !== undefined || !Array.isArray(value.citations)) return raw
    const {answer,citations,...rest} = value
    return JSON.stringify({...rest,answerParts:[{type:"text",text:answer},...citations.map(citation => ({type:"reference",...(citation as object)}))]})
  } catch { return raw }
}
let engineHook: ((session: Session, tail: string, method: string, payload: Record<string, unknown>) => Promise<void>) | undefined
const oldRuntime = process.env.DOCHAUS_RUNTIME_DIR, oldWorkspace = process.env.WORKSPACE_ROOT
const transport = Object.assign(async (input: RequestInfo | URL, init?: RequestInit) => {
  const url = new URL(String(input)); const method = init?.method ?? "GET"
  const payload = typeof init?.body === "string" ? JSON.parse(init.body) as Record<string, unknown> : {}
  const match = /^\/api\/session\/([^/]+)(.*)$/.exec(url.pathname)
  const envelope = (data: unknown) => Response.json({ data })
  if (/^\/api\/experimental\/session\/[^/]+\/wait$/.test(url.pathname)) return new Response(null, { status: 204 })
  if ((url.pathname === "/api/session" || url.pathname === "/api/dochaus/session") && method === "POST") {
    const id = String(payload.id ?? `ses_test_${++next}`)
    if (sessions.has(id)) return envelope(sessions.get(id))
    const parent = sessions.get(String(payload.parentID)); const time = new Date(Date.now() + next++).toISOString()
    const session = { ...payload, id, location: parent?.location ?? payload.location, time: { created: time, updated: time } } as Session
    sessions.set(id, session); messages.set(id, []); inboxes.set(id, []); return envelope(session)
  }
  if (url.pathname === "/api/session" && method === "GET") return Response.json({ data: [...sessions.values()].filter(session =>
    url.searchParams.get("parentID") === "null" ? !session.parentID : url.searchParams.has("parentID") ? session.parentID === url.searchParams.get("parentID") : true), cursor: {} })
  if (!match) throw new Error(`Unexpected engine route ${method} ${url.pathname}`)
  const session = sessions.get(match[1]!)
  if (!session) return Response.json({}, { status: 404 })
  const tail = match[2]
  await engineHook?.(session, tail, method, payload)
  if (!tail && method === "GET") return envelope(session)
  if (!tail && method === "PATCH") {
    if (typeof payload.expectedTitle === "string" && session.title !== payload.expectedTitle) return new Response(null, { status: 204 })
    if (payload.title === "") { titleActions.push(session.id); Object.assign(session, payload, { title: "Delivery clause review" }) }
    else Object.assign(session, payload)
    return new Response(null, { status: 204 })
  }
  if (!tail && method === "DELETE") { sessions.delete(session.id); return new Response(null, { status: 204 }) }
  if (tail === "/agent") { session.agent = String(payload.agent); return new Response(null, { status: 204 }) }
  if (tail === "/model") { session.model = payload.model as Session["model"]; return new Response(null, { status: 204 }) }
  if (tail === "/message") return Response.json({ data: [...messages.get(session.id)!].reverse(), cursor: {} })
  if (tail === "/inbox") return envelope(inboxes.get(session.id)!)
  if (tail?.startsWith("/inbox/") && method === "DELETE") {
    const id = decodeURIComponent(tail.slice("/inbox/".length))
    stoppedInputs.push(`${session.id}:delete:${id}`)
    inboxes.set(session.id, inboxes.get(session.id)!.filter(value => value.id !== id)); return new Response(null, { status: 204 })
  }
  if (tail === "/interrupt") { stoppedInputs.push(`${session.id}:interrupt`); session.outcome = "interrupted"; return Response.json({ interrupted: true }) }
  if (tail === "/synthetic") {
    admissions.push({ session: session.id, tail, payload })
    const rows = messages.get(session.id)!
    if (!rows.some(message => message.id === payload.id) && !inboxes.get(session.id)!.some(message => message.id === payload.id)) {
      prompts.push(String(payload.text))
      if (queueReview && session.metadata.type === "word-chat-review") inboxes.get(session.id)!.push({ id: payload.id, type: "synthetic" })
      else rows.push({ id: payload.id, type: "synthetic", text: payload.text, time: { created: new Date().toISOString() } })
    }
    if (loseContextResponse) { loseContextResponse = false; throw new Error("Lost context admission response") }
    return envelope({ id: payload.id, sessionID: session.id, type: "synthetic", payload: { text: payload.text }, delivery: payload.delivery })
  }
  if (tail === "/prompt") {
    admissions.push({ session: session.id, tail, payload })
    const rows = messages.get(session.id)!
    const admission = { id: payload.id, sessionID: session.id, type: "user", payload: { text: payload.text }, delivery: payload.delivery }
    if (!rows.some(message => message.id === payload.id) && !inboxes.get(session.id)!.some(message => message.id === payload.id)) {
      if (session.agent === "router") routerPrompts.push(String(payload.text)); else generations++
      if (!session.metadata.contextMode && session.agent !== "router") prompts.push(String(payload.text))
      if (queueReview && session.metadata.type === "word-chat-review") {
        inboxes.get(session.id)!.push({ id: payload.id, type: "user" })
        return envelope(admission)
      }
      rows.push({ id: payload.id, type: "user", text: payload.text, time: { created: new Date().toISOString() } })
      if (!hold || session.agent === "router") {
        const sources = (session.metadata.sources ?? []) as { id: string; quote: string }[]
        const source = sources.find(source => source.id !== "word") ?? sources[0]
        const output: Record<string, unknown> = session.metadata.type === "word-chat-review" ? { approved: approveReview, issues: approveReview ? [] : ["The replacement loses a qualification."] } : { answer: modelAnswer ?? `Answer ${source?.quote ?? ""}`, citations: source ? [{ sourceId: inventCitation ? "foreign-source" : source.id, quote: source.quote.slice(0, 100) }] : [], suggestions: [] }
        if (amend && session.metadata.type !== "word-chat-review") output.proposal = { before: deletion ? "Tuesday " : context.text, after: deletion ? "" : context.text.replace("Tuesday", "Wednesday"), ...modelProposal, explanation: "Requested delivery update", openItems: ["Confirm the date with the client"] }
        const raw = (session.metadata.type === "word-chat-review" ? reviewerRawOutput : modelRawOutput) ?? JSON.stringify(output)
        rows.push({ id: `msg_answer_${next++}`, type: "assistant", content: [{ type: "text", text: session.agent === "router" ? routerOutput ?? JSON.stringify({ route: amend ? "redliner" : routerAgent }) : session.metadata.type === "word-chat-review" ? raw : modelStructuredRawOutput ?? (session.metadata.answerFormat === "segments-v1" ? structuredFixture(raw) : raw) }], time: { created: new Date().toISOString(), completed: new Date().toISOString() } })
        session.time.idle = new Date().toISOString(); session.outcome = session.metadata.type === "word-chat-review" ? reviewerOutcome : session.agent === "router" ? "succeeded" : mainOutcome
      }
    }
    if (loseResponse && session.agent !== "router") { loseResponse = false; throw new Error("Lost admitted response") }
    return envelope(admission)
  }
  throw new Error(`Unexpected engine route ${method} ${url.pathname}`)
}, { preconnect() {} }) as typeof fetch
const post = (action: string, body: object = {}) => app.request(`/word/chat/${action}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
const context = { text: "The current Word buffer uses Tuesday delivery.", paragraphs: [{ id: "11", text: "The current Word buffer uses Tuesday delivery." }], saved: false, capturedAt: "2026-01-01T00:00:00Z", truncated: false }
const send = (conversationId: string, scope: "document" | "matter", text: string, requestId = crypto.randomUUID(), attachments: WordChatSend["attachments"] = []) => post("send", { conversationId, scope, text, requestId, attachments, context })
const suppliedContext = (prompt: string) => JSON.parse(prompt.split("\n").at(-2)!).context
const sessionContext = (id: string) => String(messages.get(id)!.find(message => message.type === "synthetic")!.text)
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
  expect(routerPrompts.at(-1)).not.toContain("DISTINCTIVE_EVIDENCE")
  const source = wordChatMaterials(auth).find(material => material.name === "Other.docx")!
  expect((await send(id, "document", "Explain the attached evidence", crypto.randomUUID(), [source])).status).toBe(200)
  expect(prompts.at(-1)).toContain("DISTINCTIVE_EVIDENCE is a blue agreement.")
  const routed = routerPrompts.length
  expect((await send(id, "document", "Explain this", crypto.randomUUID(), [{ ...source, hash: "0".repeat(64) }])).status).toBe(409)
  expect(routerPrompts).toHaveLength(routed)
})
test("conversation naming uses the native title action once from clean user input, not a first-question prefix", async () => {
  const id = await newConversation()
  expect(sessions.get(id)?.title).toBe("Word · Draft.docx")
  expect((await send(id, "document", "Explain\n  the delivery clause")).status).toBe(200)
  await new Promise(resolve => setTimeout(resolve, 0))
  expect(sessions.get(id)?.title).toBe("Delivery clause review")
  expect((await send(id, "document", "What needs clarification?")).status).toBe(200)
  expect(sessions.get(id)?.title).toBe("Delivery clause review")
  const list = await (await post("conversations")).json() as { conversations: { id: string; title: string }[] }
  expect(list.conversations.find(item => item.id === id)?.title).toBe("Delivery clause review")
  const turns = [...sessions.values()].filter(session => session.parentID === id && session.metadata.type === "word-chat-turn")
  expect(titleActions.filter(value => turns.some(turn => turn.id === value))).toEqual([turns[0]!.id])
  const admitted = messages.get(turns[0]!.id)!.filter(message => message.type === "user")
  expect(admitted.map(message => message.text)).toEqual(["Explain\n  the delivery clause"])
  expect(sessionContext(turns[0]!.id)).toContain("supplied-conversation-context")
  // Historical placeholders no longer fabricate a title from user text.
  sessions.get(id)!.title = "Word · Draft.docx"
  const opened = await (await post("conversation", { conversationId: id })).json() as { conversation: { title: string } }
  expect(opened.conversation.title).toBe("Word · Draft.docx")
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
test("lost synthetic-context admission resumes with one context and one clean user message", async () => {
  const id = await newConversation(), requestId = crypto.randomUUID(), before = generations
  loseContextResponse = true
  expect((await send(id, "document", "Explain this selected delivery wording", requestId)).status).toBe(503)
  const child = [...sessions.values()].find(value => value.parentID === id && value.metadata.type === "word-chat-turn")!
  expect(messages.get(child.id)!.map(message => message.type)).toEqual(["synthetic"])
  expect((await send(id, "document", "Explain this selected delivery wording", requestId)).status).toBe(200)
  expect(generations - before).toBe(1)
  const rows = messages.get(child.id)!
  expect(rows.filter(message => message.type === "synthetic")).toHaveLength(1)
  expect(rows.filter(message => message.type === "user").map(message => message.text)).toEqual(["Explain this selected delivery wording"])
  const contextAdmissions = admissions.filter(value => value.session === child.id && value.tail === "/synthetic")
  expect(contextAdmissions).toHaveLength(2)
  expect(new Set(contextAdmissions.map(value => value.payload.id)).size).toBe(1)
  expect(contextAdmissions[0]!.payload).toMatchObject({ delivery: "steer", resume: false })
  expect(admissions.find(value => value.session === child.id && value.tail === "/prompt")!.payload).toMatchObject({ text: "Explain this selected delivery wording", delivery: "steer", resume: true })
})

test("Auto obeys the shared native router while explicit modes and pinned refinement bypass it", async () => {
  const id = await newConversation(), before = routerPrompts.length
  const request = (text: string, extra: object = {}) => post("send", { conversationId: id, scope: "document", text, requestId: crypto.randomUUID(), attachments: [], context, ...extra })
  try {
    // A router decision, not a Word regex, owns the operation even when the
    // question resembles a known editing command.
    routerOutput = '{"route":"qa"}'
    const answer = (await (await request("Rewrite this clause")).json() as { turn: WordChatTurn }).turn
    expect(answer.kind).toBe("chat")
    expect(sessions.get(answer.id)?.agent).toBe("qa")
    expect(routerPrompts.at(-1)).toContain("- qa:")
    expect(routerPrompts.at(-1)).toContain("- redliner:")
    expect(routerPrompts.at(-1)).not.toContain("- draft:")
    routerOutput = '{"route":"redliner"}'
    modelRawOutput = JSON.stringify({ answer: "Prepared a layout proposal.", citations: [], proposal: { before: context.text, after: context.text, formatting: [{ paragraph: 0, spacingAfter: 6 }], explanation: "Separate this paragraph.", openItems: [] } })
    const change = (await (await request("Could this presentation be clearer?")).json() as { turn: WordChatTurn }).turn
    expect(sessions.get(change.id)?.agent).toBe("redliner")
    expect(sessions.get(change.id)?.metadata.kind).toBe("amendment")
    expect(change.kind).toBe("formatting")
    expect(change.proposal?.formatting).toEqual([{ paragraph: 0, spacingAfter: 6 }])
    const count = routerPrompts.length
    routerOutput = '{"route":"unsupported"}'
    const explicit = (await (await request("Rewrite this", { mode: "qa", kind: "formatting" })).json() as { turn: WordChatTurn }).turn
    expect(explicit.kind).toBe("chat")
    expect(explicit.proposal).toBeUndefined()
    const refined = (await (await request("Keep that paragraph and use a little more space", { mode: "qa", proposal: change.proposal })).json() as { turn: WordChatTurn }).turn
    expect(sessions.get(refined.id)?.agent).toBe("redliner")
    expect(refined.proposal?.target).toEqual(change.proposal?.target)
    expect(routerPrompts).toHaveLength(count)
    const prior = generations
    const failed = await request("Explain the document", { mode: "auto" })
    expect(failed.status).toBe(503)
    expect((await failed.json() as { error: string }).error).toContain("select Q&A or Redline")
    expect(generations).toBe(prior)
    expect([...sessions.values()].some(value => value.parentID === id && value.metadata.type === "word-chat-router")).toBe(false)
    expect(routerPrompts.length).toBeGreaterThan(before)
    for (const child of [...sessions.values()].filter(value => value.parentID === id)) {
      expect(child.metadata.contextMode).toBe("supplied")
      expect((child as unknown as { permissions: unknown }).permissions).toEqual([{ action: "*", resource: "*", effect: "deny" }])
    }
  } finally { routerOutput = undefined; modelRawOutput = undefined }
})

test("legacy admitted turns remain readable and reconcile their original prompt without a new context admission", async () => {
  const id = await newConversation(), requestId = crypto.randomUUID()
  const response = await send(id, "document", "Explain delivery", requestId)
  const turn = (await response.json() as { turn: WordChatTurn }).turn
  const child = sessions.get(turn.id)!
  delete child.metadata.suppliedContext; delete child.metadata.contextMode; delete child.metadata.agent
  child.agent = "word-chat"
  const original = String(child.metadata.prompt)
  const rows = messages.get(child.id)!.filter(message => message.type !== "synthetic")
  rows.find(message => message.type === "user")!.text = original
  messages.set(child.id, rows)
  const before = generations, contextCount = admissions.filter(value => value.tail === "/synthetic").length
  expect((await send(id, "document", "Explain delivery", requestId)).status).toBe(200)
  expect(generations).toBe(before)
  expect(admissions.filter(value => value.tail === "/synthetic")).toHaveLength(contextCount)
  expect(admissions.at(-1)?.payload.text).toBe(original)
})
test("legacy proposal kinds retain their original wording and formatting boundaries", async () => {
  const id = await newConversation(), before = context.text, after = before.replace("Tuesday", "Wednesday")
  const formatting = [{ paragraph: 0, spacingAfter: 6 }]
  routerAgent = "redliner"; amend = true; modelProposal = { before, after, formatting }
  try {
    const turn = (await (await send(id, "document", "Rewrite and adjust spacing")).json() as { turn: WordChatTurn }).turn
    expect(turn.proposal?.after).toBe(after)
    const child = sessions.get(turn.id)!
    delete child.metadata.suppliedContext; delete child.metadata.contextMode
    const answer = messages.get(turn.id)!.find(message => message.type === "assistant")!
    const setProposal = (replacement: string, properties?: unknown) => {
      answer.content = [{ type: "text", text: structuredFixture(JSON.stringify({ answer: "Prepared.", citations: [], proposal: { before, after: replacement, ...(properties === undefined ? {} : { formatting: properties }), explanation: "Review", openItems: [] } })) }]
    }
    const poll = async () => (await (await post("poll", { conversationId: id, turnId: turn.id })).json() as { turn: WordChatTurn }).turn
    child.metadata.kind = "formatting"
    setProposal(before, formatting)
    expect((await poll()).proposal?.formatting).toEqual(formatting)
    setProposal(after, formatting)
    expect((await poll()).proposal).toBeUndefined()
    child.metadata.kind = "amendment"
    setProposal(after)
    expect((await poll()).proposal?.after).toBe(after)
    setProposal(after, formatting)
    expect((await poll()).proposal).toBeUndefined()
  } finally { routerAgent = "qa"; amend = false; modelProposal = undefined }
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
  const split = wordChatSuggestionBatches({ ...context, paragraphs, text: paragraphs.map(paragraph => paragraph.text).join("\n") })
  expect(split.batches).toHaveLength(3); expect(split.total).toBeGreaterThan(3)
  expect(split.batches[0]?.[0]?.id).toBe("0"); expect(split.batches.at(-1)?.at(-1)?.id).toBe("29")
  expect(split.batches.every(batch => batch.reduce((size, paragraph) => size + paragraph.text.length, 0) <= 8000)).toBe(true)
})

test("one suggestion flow retains whole-document passages and a crossing selection without duplicated ideas or naming", async () => {
  const id = await newConversation()
  const paragraphs = Array.from({ length: 30 }, (_, index) => ({ id: String(index), text: `Passage ${index}. ${"x".repeat(980)}` }))
  const selected = paragraphs.slice(15, 17), text = selected.map(paragraph => paragraph.text).join("\r")
  const selection = { id: "crossing", text, paragraphs: selected, startOffset: 0, endOffset: selected[1]!.text.length, saved: false, capturedAt: context.capturedAt }
  const current = { ...context, text: paragraphs.map(paragraph => paragraph.text).join("\r"), paragraphs, selection }
  const named = titleActions.length
  modelRawOutput = JSON.stringify({ answer: "", citations: [], suggestions: [{ title: "Clarify this passage", explanation: "Confirm the intended wording.", passage: text }] })
  try {
    const response = await post("suggest", { conversationId: id, scope: "document", text: "", requestId: crypto.randomUUID(), attachments: [], context: current })
    expect(response.status).toBe(200)
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.suggestionCoverage).toEqual({ reviewed: 3, total: 4 })
    expect(turn.suggestions).toHaveLength(1)
    expect(turn.suggestions?.[0]?.passage).toBe(text)
    const batches = [...sessions.values()].filter(value => value.parentID === turn.id)
    expect(batches).toHaveLength(3)
    const supplied = batches.map(value => suppliedContext(sessionContext(value.id)) as { target: { text: string }; sources: { id: string; quote: string }[] })
    expect(supplied[0]!.sources.find(source => source.id === "word")?.quote).toContain("Passage 0.")
    expect(supplied.at(-1)!.sources.find(source => source.id === "word")?.quote).toContain("Passage 29.")
    for (const [index, value] of supplied.entries()) {
      expect(value.target.text).toBe(text)
      expect(value.sources.find(source => source.id === "word-selection")?.quote).toBe(text)
      expect(batches[index]!.agent).toBe("qa")
      expect(batches[index]!.metadata.contextMode).toBe("supplied")
      expect(validateWordChatContext((batches[index]!.metadata.input as WordChatSend).context)).toBe(true)
    }
    expect(titleActions).toHaveLength(named)
    expect(sessions.get(id)?.title).toBe("Word · Draft.docx")
  } finally { modelRawOutput = undefined }
})

test("material loading supplies the persisted installation defaults without changing a conversation's scope", async () => {
  const id = await newConversation()
  writeWordChatPreferences({ defaultScope: "matter", suggestions: false })
  try {
    const material = await (await post("materials", { conversationId: id })).json() as { preferences: { defaultScope: string; suggestions: boolean } }
    expect(material.preferences).toEqual({ defaultScope: "matter", suggestions: false })
    const result = (await (await send(id, "document", "Explain the current delivery day")).json() as { turn: WordChatTurn }).turn
    expect(result.scope).toBe("document")
  } finally { writeWordChatPreferences({ defaultScope: "document", suggestions: true }) }
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
  expect(review.agent).toBe("document-reviewer")
  expect(review.metadata.contextMode).toBe("supplied")
  expect((review as unknown as { permissions: unknown }).permissions).toEqual([{ action: "*", resource: "*", effect: "deny" }])
  expect((await post("prepare", { conversationId: id, turnId: turn.id, ooxml: "not authorized" })).status).toBe(409)
  approveReview = false; amend = true
  const rejected = (await (await send(id, "document", "Rewrite delivery for next week")).json() as { turn: WordChatTurn }).turn
  amend = false; approveReview = true
  expect(rejected.proposal).toBeUndefined()
  expect(rejected.error).toContain("needs revision")
})

test("an approved JSON result from an interrupted or failed reviewer does not authorize a proposal", async () => {
  const id = await newConversation()
  amend = true
  try {
    const baseline = (await (await send(id, "document", "Rewrite delivery for Wednesday")).json() as { turn: WordChatTurn }).turn
    expect(baseline.proposal?.after).toContain("Wednesday")
    for (const outcome of ["interrupted", "failed"] as const) {
      reviewerOutcome = outcome
      const turn = (await (await send(id, "document", "Rewrite delivery for Wednesday")).json() as { turn: WordChatTurn }).turn
      expect(turn.proposal).toBeUndefined()
      expect(turn.error).toContain("needs revision")
      const review = [...sessions.values()].find(value => value.parentID === turn.id && value.metadata.type === "word-chat-review")!
      const content = messages.get(review.id)!.find(message => message.type === "assistant")!.content as { text: string }[]
      expect(JSON.parse(content[0]!.text).approved).toBe(true)
    }
  } finally { reviewerOutcome = "succeeded"; amend = false }
})

test("Stop queues behind an in-flight poll's reviewer admission and interrupts the admitted reviewer", async () => {
  const id = await newConversation()
  hold = true; amend = true
  const response = await send(id, "document", "Rewrite delivery for Wednesday")
  hold = false; amend = false
  const running = (await response.json() as { turn: WordChatTurn }).turn
  const child = sessions.get(running.id)!
  messages.get(child.id)!.push({ id: `msg_late_${next++}`, type: "assistant", content: [{ type: "text", text: structuredFixture(JSON.stringify({ answer: "Prepared a focused proposal.", citations: [], proposal: { before: context.text, after: context.text.replace("Tuesday", "Wednesday"), explanation: "Requested update.", openItems: [] } })) }], time: { created: new Date().toISOString(), completed: new Date().toISOString() } })
  child.time.idle = new Date().toISOString(); child.outcome = "succeeded"
  let enter!: () => void, release!: () => void
  const entered = new Promise<void>(resolve => { enter = resolve }), gate = new Promise<void>(resolve => { release = resolve })
  engineHook = async (value, tail) => { if (value.id === child.id && tail === "/message") { engineHook = undefined; enter(); await gate } }
  hold = true // The newly admitted reviewer remains running until Stop.
  queueReview = true
  try {
    const polling = post("poll", { conversationId: id, turnId: child.id })
    await entered
    const stopping = post("cancel", { conversationId: id, turnId: child.id })
    await Promise.resolve(); await Promise.resolve()
    expect(child.metadata.cancelled).not.toBe(true)
    release()
    const polled = (await (await polling).json() as { turn: WordChatTurn }).turn
    expect(polled.status).toBe("running")
    const stopped = (await (await stopping).json() as { turn: WordChatTurn }).turn
    expect(stopped.status).toBe("cancelled")
    const reviewers = [...sessions.values()].filter(value => value.parentID === child.id && value.metadata.type === "word-chat-review")
    expect(reviewers).toHaveLength(1)
    expect(reviewers[0]!.outcome).toBe("interrupted")
    expect(inboxes.get(reviewers[0]!.id)).toEqual([])
    const queuedStops = stoppedInputs.filter(value => value.startsWith(reviewers[0]!.id + ":"))
    expect(queuedStops[0]).toBe(`${reviewers[0]!.id}:interrupt`)
    expect(queuedStops.filter(value => value.includes(":delete:"))).toHaveLength(2)
    expect(messages.get(child.id)!.some(message => message.type === "assistant")).toBe(true)
    const before = generations
    expect(((await (await post("poll", { conversationId: id, turnId: child.id })).json()) as { turn: WordChatTurn }).turn.proposal).toBeUndefined()
    expect(generations).toBe(before)
  } finally { release(); engineHook = undefined; hold = false; queueReview = false }
})

test("an attachment changed during transcript hydration cannot publish its earlier quote", async () => {
  const name = "Late evidence.docx"
  await writeDeliverable(directory, name, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("The attached source confirms synthetic delivery evidence.")] }] })))
  const material = wordChatMaterials(auth).find(value => value.name === name)!, bytes = readFileSync(path.join(directory, name))
  const id = await newConversation()
  const baseline = (await (await send(id, "document", "Explain the attached evidence", crypto.randomUUID(), [material])).json() as { turn: WordChatTurn }).turn
  expect(baseline.sources[0]?.quote).toContain("synthetic delivery evidence")
  const requestId = crypto.randomUUID()
  engineHook = async (value, tail) => {
    if (value.parentID === id && (value.metadata.input as WordChatSend)?.requestId === requestId && tail === "/message") {
      engineHook = undefined; writeFileSync(path.join(directory, name), "Replaced while the transcript was being read.")
    }
  }
  try {
    const changed = (await (await send(id, "document", "Explain it again", requestId, [material])).json() as { turn: WordChatTurn }).turn
    expect(changed.status).toBe("error")
    expect(changed.answer).toBe("")
    expect(changed.sources).toEqual([])
    expect(changed.proposal).toBeUndefined()
    expect(changed.error).toContain("Supporting material changed")
  } finally { engineHook = undefined; writeFileSync(path.join(directory, name), bytes) }
})

test("error exits clear earlier verified quotes when evidence changes during a bad-citation or failed answer", async () => {
  const name = "Error evidence.docx"
  await writeDeliverable(directory, name, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("A synthetic attached source supports this answer.")] }] })))
  const material = wordChatMaterials(auth).find(value => value.name === name)!, bytes = readFileSync(path.join(directory, name))
  const id = await newConversation()
  const baseline = (await (await send(id, "document", "Explain the attached evidence", crypto.randomUUID(), [material])).json() as { turn: WordChatTurn }).turn
  expect(baseline.status).toBe("complete")
  const citation = { sourceId: baseline.sources[0]!.id, quote: baseline.sources[0]!.quote }
  try {
    for (const failure of ["citation", "failed"] as const) {
      const requestId = crypto.randomUUID()
      modelRawOutput = JSON.stringify({ answer: "The attached source supports this answer.", citations: [citation, ...(failure === "citation" ? [{ sourceId: "not-supplied", quote: "Unsupported quotation." }] : [])] })
      mainOutcome = failure === "failed" ? "failed" : "succeeded"
      engineHook = async (value, tail) => {
        if ((value.metadata.input as WordChatSend)?.requestId === requestId && tail === "/message") {
          engineHook = undefined; writeFileSync(path.join(directory, name), "Changed while the terminal output was being read.")
        }
      }
      const changed = (await (await send(id, "document", "Explain the attached evidence again", requestId, [material])).json() as { turn: WordChatTurn }).turn
      expect(changed.status).toBe("error")
      expect(changed.answer).toBe("")
      expect(changed.sources).toEqual([])
      expect(changed.proposal).toBeUndefined()
      expect(changed.error).toContain("Supporting material changed")
      writeFileSync(path.join(directory, name), bytes)
    }
  } finally { engineHook = undefined; modelRawOutput = undefined; mainOutcome = "succeeded"; writeFileSync(path.join(directory, name), bytes) }
})

test("an interrupted scoped answer cannot release its otherwise valid amendment for independent approval", async () => {
  const id = await newConversation()
  amend = true
  try {
    const baseline = (await (await send(id, "document", "Rewrite delivery for Wednesday")).json() as { turn: WordChatTurn }).turn
    expect(baseline.proposal).toBeDefined()
    mainOutcome = "interrupted"
    const interrupted = (await (await send(id, "document", "Rewrite delivery for Wednesday")).json() as { turn: WordChatTurn }).turn
    expect(interrupted.status).toBe("error")
    expect(interrupted.answer).toBe("")
    expect(interrupted.sources).toEqual([])
    expect(interrupted.proposal).toBeUndefined()
    expect([...sessions.values()].filter(value => value.parentID === interrupted.id && value.metadata.type === "word-chat-review")).toHaveLength(0)
  } finally { mainOutcome = "succeeded"; amend = false }
})

test("a completed answer is returned while its one native title job is still pending", async () => {
  const id = await newConversation()
  let enter!: () => void, release!: () => void
  const entered = new Promise<void>(resolve => { enter = resolve }), gate = new Promise<void>(resolve => { release = resolve })
  engineHook = async (value, tail, method, payload) => {
    if (value.parentID === id && !tail && method === "PATCH" && payload.title === "") { enter(); await gate }
  }
  try {
    const response = await send(id, "document", "Explain delivery without waiting for a conversation title")
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.status).toBe("complete")
    expect(turn.answer).toContain("Tuesday")
    await entered
    const before = titleActions.length
    expect((await post("poll", { conversationId: id, turnId: turn.id })).status).toBe(200)
    expect(titleActions).toHaveLength(before)
  } finally { release(); engineHook = undefined; await new Promise(resolve => setTimeout(resolve, 0)) }
  expect(sessions.get(id)?.title).toBe("Delivery clause review")
})

test("composer modes route questions, explicit changes and formatting without changing evidence scope", async () => {
  const id = await newConversation()
  const request = (text: string, extra: object = {}) => post("send", { conversationId: id, scope: "document", text, requestId: crypto.randomUUID(), attachments: [], context, ...extra })
  for (const [text, extra, expected] of [
    ["What formatting is used?", {}, "chat"],
    ["How can I improve the spacing?", {}, "chat"],
    ["Can you explain the heading styles?", {}, "chat"],
    ["Can you format the document better, more spacing and better headings?", {}, "amendment"],
    ["Please increase paragraph spacing without changing any words.", {}, "amendment"],
    ["Apply Heading 2 to the existing headings.", {}, "amendment"],
    ["Do not change the formatting. Explain it.", {}, "chat"],
    ["Rewrite this clause", { mode: "qa", kind: "amendment" }, "chat"],
    ["Format the document", { mode: "qa" }, "chat"],
    ["Explain this clause", { mode: "redline" }, "amendment"],
    ["Improve the layout", { mode: "redline" }, "amendment"],
    ["Prepare my requested change", { mode: "auto", kind: "amendment" }, "amendment"],
  ] as const) {
    routerAgent = expected === "chat" ? "qa" : "redliner"
    const response = await request(text, extra)
    expect(response.status, text).toBe(200)
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.kind, text).toBe(expected)
    expect(turn.scope).toBe("document")
    if (expected === "chat") expect(turn.proposal).toBeUndefined()
  }
  routerAgent = "qa"
  const before = generations
  for (const mode of ["unknown", ["qa"], { mode: "qa" }, 0, null])
    expect((await request("Rewrite this", { mode })).status).toBe(400)
  expect((await request("Rewrite this", { scope: ["matter"] })).status).toBe(400)
  expect(generations).toBe(before)
})

test("format-only proposals retain exact paragraphs, carry bounded properties and receive independent review", async () => {
  const id = await newConversation()
  const paragraphs = [{ id: "title", text: "Services agreement" }, { id: "terms", text: "The existing delivery wording stays unchanged." }]
  const before = paragraphs.map(paragraph => paragraph.text).join("\r")
  const current = { ...context, text: before, paragraphs }
  const formatting = [{ paragraph: 0, style: "Title" as const, spacingAfter: 12 }, { paragraph: 1, spacingBefore: 6, lineSpacing: 1.5 }]
  const request = (extra: object = {}) => post("send", { conversationId: id, scope: "document", text: "Format the existing heading and improve paragraph spacing, preserving all wording.", requestId: crypto.randomUUID(), attachments: [], context: current, ...extra })
  routerAgent = "redliner"
  amend = true; modelProposal = { before: before.replace(/\r/g, "\n"), after: before.replace(/\r/g, "\n"), formatting }
  try {
    const response = await request()
    expect(response.status).toBe(200)
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.kind).toBe("formatting")
    expect(turn.proposal?.formatting).toEqual(formatting)
    expect(turn.proposal?.before).toBe(before)
    expect(turn.proposal?.after).toBe(before)
    expect(turn.proposal?.target.paragraphs).toEqual(paragraphs)
    const prompt = sessionContext(turn.id)
    expect(prompt).toContain("spacingBefore")
    expect(prompt).toContain("points")
    expect(prompt).toContain("lineSpacing")
    expect(prompt).toContain("reasonable layout choices")
    expect(prompt).toContain("brief")
    expect(prompt).toContain("explicit Apply")
    expect(prompt).not.toContain("No tools, document mutation")
    const supplied = suppliedContext(prompt) as { formattingParagraphs: typeof paragraphs }
    expect(supplied.formattingParagraphs).toEqual(paragraphs)
    const reviewer = [...sessions.values()].find(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")!
    const reviewPrompt = sessionContext(reviewer.id)
    expect(reviewPrompt).toContain("Review the requested layout together with the proposed wording")
    expect(reviewPrompt).toContain("resulting paragraphs")
    // An explicit refinement keeps its already reviewed formatting target even in Q&A mode.
    const refined = (await (await request({ mode: "qa", proposal: turn.proposal })).json() as { turn: WordChatTurn }).turn
    expect(refined.kind).toBe("formatting")
    expect(refined.proposal?.target).toEqual(turn.proposal?.target)
    approveReview = false
    const refused = (await (await request()).json() as { turn: WordChatTurn }).turn
    expect(refused.proposal).toBeUndefined()
    expect(refused.error).toContain("needs revision")
    expect(refused.answer).toContain("loses a qualification")
    approveReview = true
    // Reject unsupported properties and formatting outside complete paragraph authority.
    for (const proposal of [
      { before, after: before, formatting: [{ paragraph: 0, fontSize: 30 }] },
      { before: "Services", after: "Services", formatting: [{ paragraph: 0, spacingAfter: 12 }] },
    ]) {
      modelProposal = proposal
      const count = generations
      const invalid = (await (await request()).json() as { turn: WordChatTurn }).turn
      expect(invalid.proposal).toBeUndefined()
      expect(invalid.error).toMatch(/complete paragraphs|clear passage/i)
      expect(generations - count).toBe(1)
    }
    // Every complete paragraph supplied for layout is also verifiable Word evidence,
    // including a target beyond the normal chat excerpt's initial 14k characters.
    const tail = "Schedule"
    const extended = [{ id: "long", text: "x".repeat(16_000) }, { id: "tail", text: tail }]
    modelRawOutput = JSON.stringify({ answer: "Prepare spacing for the final existing heading [word].", citations: [{ sourceId: "word", quote: tail }], proposal: { before: tail, after: tail, formatting: [{ paragraph: 0, spacingBefore: 12 }], explanation: "Separate this existing heading.", openItems: [] } })
    const extendedTurn = (await (await request({ context: { ...current, text: extended.map(paragraph => paragraph.text).join("\r"), paragraphs: extended } })).json() as { turn: WordChatTurn }).turn
    expect(extendedTurn.status).toBe("complete")
    expect(extendedTurn.proposal?.target.paragraphs).toEqual([extended[1]!])
    expect(extendedTurn.sources[0]?.quote).toBe(tail)
    modelRawOutput = undefined

  } finally { amend = false; modelProposal = undefined; modelRawOutput = undefined; approveReview = true; routerAgent = "qa" }
})

test("mixed wording and layout follow shared Auto and receive one independent review", async () => {
  const id = await newConversation()
  const before = context.text
  const after = before.replace("Tuesday", "Wednesday")
  const formatting = [{ paragraph: 0, spacingAfter: 6 }]
  routerAgent = "redliner"; amend = true
  modelProposal = { before, after, formatting }
  try {
    const request = () => post("send", { conversationId: id, scope: "document", mode: "auto", text: "Add headings and sections and make the wording more formal.", requestId: crypto.randomUUID(), attachments: [], context })
    const turn = (await (await request()).json() as { turn: WordChatTurn }).turn
    expect(turn.kind).toBe("amendment")
    expect(turn.proposal?.after).toBe(after)
    expect(turn.proposal?.formatting).toEqual(formatting)
    const prompt = sessionContext(turn.id)
    expect(prompt).toContain("requested wording and layout together")
    expect(prompt).toContain("not from the document-wide formattingParagraphs list")
    expect(prompt).not.toContain("Choose a wording or formatting-only proposal")
    const reviewer = [...sessions.values()].find(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")!
    const review = sessionContext(reviewer.id)
    expect(review).toContain("Review the requested layout together with the proposed wording")
    expect(review).toContain("invented client decision")
    expect(review).toContain("Reject silently omitted requested work")
    expect(review).toContain("remaining requested work is clearly disclosed")
    expect(suppliedContext(review).proposal).toEqual(turn.proposal)
    // Formatting references the resulting paragraphs, including inserted text.
    modelProposal = { before, after: `Delivery\r${before}`, formatting: [{ paragraph: 1, spacingAfter: 6 }] }
    const inserted = (await (await request()).json() as { turn: WordChatTurn }).turn
    expect(inserted.proposal?.formatting).toEqual([{ paragraph: 1, spacingAfter: 6 }])
    modelRawOutput = JSON.stringify({ answer: "Added the heading.", citations: [], proposal: { before, after: `Delivery\r${before}`, formatting: [{ paragraph: 2, spacingAfter: 6 }] } })
    const outOfRange = (await (await request()).json() as { turn: WordChatTurn }).turn
    expect(outOfRange.proposal).toBeUndefined()
    expect(outOfRange.answer).toBe("")
    expect(outOfRange.error).toContain("Select one clear passage")
    expect([...sessions.values()].some(session => session.parentID === outOfRange.id && session.metadata.type === "word-chat-review")).toBe(false)
    modelRawOutput = undefined
    modelProposal = { before, after, formatting }
    approveReview = false
    const rejected = (await (await request()).json() as { turn: WordChatTurn }).turn
    expect(rejected.proposal).toBeUndefined()
    expect(rejected.error).toContain("needs revision")
  } finally { amend = false; modelProposal = undefined; modelRawOutput = undefined; approveReview = true; routerAgent = "qa" }
})

test("a misplaced formatting array cannot silently become a wording-only proposal", async () => {
  const id = await newConversation(), before = context.text, after = before.replace("Tuesday", "Wednesday")
  routerAgent = "redliner"
  modelRawOutput = JSON.stringify({ answer: "Prepared the wording and spacing.", citations: [],
    proposal: { before, after, explanation: "Requested change", openItems: [] }, formatting: [{ paragraph: 0, spacingAfter: 6 }] })
  try {
    const response = await send(id, "document", "Rewrite the wording and adjust spacing")
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.proposal).toBeUndefined()
    expect(turn.error).toContain("formatting could not be included")
    expect(turn.answer).toBe("")
    expect([...sessions.values()].some(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")).toBe(false)
  } finally { routerAgent = "qa"; modelRawOutput = undefined }
})

test("reviewer issues are framed as a rejected proposal even if they use applied tense", async () => {
  const id = await newConversation()
  routerAgent = "redliner"; amend = true
  modelProposal = { before: context.text, after: context.text.replace("Tuesday", "Wednesday") }
  reviewerRawOutput = JSON.stringify({ approved: false, issues: ["The change applied to the wrong passage."] })
  try {
    const turn = (await (await send(id, "document", "Revise the passage")).json() as { turn: WordChatTurn }).turn
    expect(turn.proposal).toBeUndefined()
    expect(turn.answer).toBe("The reviewer found these issues in the proposed change:\nThe change applied to the wrong passage.")
    expect(turn.error).toContain("No change was applied.")
    const reviewer = [...sessions.values()].find(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain("hypothetical effects of the proposed change")
  } finally { routerAgent = "qa"; amend = false; modelProposal = undefined; reviewerRawOutput = undefined }
})

test("a completion claim without a proposal is not shown as an applied change", async () => {
  const id = await newConversation()
  routerAgent = "redliner"
  modelRawOutput = JSON.stringify({ answer: "Added the requested section.", citations: [] })
  try {
    const turn = (await (await send(id, "document", "Add a section")).json() as { turn: WordChatTurn }).turn
    expect(turn.proposal).toBeUndefined()
    expect(turn.answer).toBe("")
    expect(turn.error).toContain("No amendment was prepared")
  } finally { routerAgent = "qa"; modelRawOutput = undefined }
})

test("an unprepared amendment can ask a necessary clarification without claiming an edit", async () => {
  const id = await newConversation()
  routerAgent = "redliner"
  modelRawOutput = JSON.stringify({ answer: "I added the requested notice period.", citations: [], clarificationQuestions: ["Should clause 6.2 apply?", "Should Acme Ltd. use U.S. law?", "Not a question."] })
  try {
    const turn = (await (await send(id, "document", "Add a termination notice period")).json() as { turn: WordChatTurn }).turn
    expect(turn.proposal).toBeUndefined()
    expect(turn.answer).toBe("Should clause 6.2 apply?\nShould Acme Ltd. use U.S. law?")
    expect(turn.error).toContain("No amendment was prepared")
    expect(sessionContext(turn.id)).toContain("clarificationQuestions")
  } finally { routerAgent = "qa"; modelRawOutput = undefined }
})

test("independent amendment review receives first/last paragraph offsets and local-save semantics", async () => {
  const id = await newConversation()
  const paragraphs = [{ id: "first-native-id", text: "First paragraph with its original full wording." }, { id: "last-native-id", text: "Second note." }]
  const before = paragraphs.map(paragraph => paragraph.text).join("\r")
  const current = { ...context, text: before, paragraphs, saved: true }
  const amendRequest = (requestId: string) => post("send", { conversationId: id, scope: "document", kind: "amendment", text: "Insert an ordinary heading before the two paragraphs and retain both verbatim.", requestId, attachments: [], context: current })
  amend = true; modelProposal = { before, after: `Heading\r${before}` }
  try {
    const response = await amendRequest(crypto.randomUUID())
    expect(response.status).toBe(200)
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.proposal?.before).toBe(before)
    const reviewer = [...sessions.values()].find(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")!
    const prompt = sessionContext(reviewer.id)
    expect(prompt).toContain("first target paragraph")
    expect(prompt).toContain("last target paragraph")
    expect(prompt).toContain("single-paragraph target")
    expect(prompt).toContain("already deterministically validated")
    expect(prompt).toContain("local Word buffer")
    expect(prompt).toContain("confirmed matter version")
    const supplied = suppliedContext(prompt) as { proposal: NonNullable<WordChatTurn["proposal"]>; evidence: { id: string; provenance: string }[] }
    expect(supplied.proposal.target.paragraphs).toEqual(paragraphs)
    expect(supplied.proposal.target.startOffset).toBe(0)
    expect(supplied.proposal.target.endOffset).toBe(paragraphs[1]!.text.length)
    expect(supplied.proposal.target.endOffset).toBeLessThan(paragraphs[0]!.text.length)
    expect(supplied.proposal.target.saved).toBe(true)
    expect(supplied.evidence.find(source => source.id === "word")?.provenance).toBe("word-buffer")
    // Clarifying mechanical metadata must not bypass the independent content review.
    approveReview = false
    const rejected = (await (await amendRequest(crypto.randomUUID())).json() as { turn: WordChatTurn }).turn
    expect(rejected.proposal).toBeUndefined()
    expect(rejected.error).toContain("needs revision")
    expect(rejected.answer).toContain("The replacement loses a qualification.")
    // A non-boolean approval or unparseable reviewer output remains a refusal.
    approveReview = true
    for (const output of ['{"approved":"true","issues":["The approval value is invalid."]}', "The review did not return a decision."]) {
      reviewerRawOutput = output
      const invalid = (await (await amendRequest(crypto.randomUUID())).json() as { turn: WordChatTurn }).turn
      expect(invalid.proposal).toBeUndefined()
      expect(invalid.error).toContain("needs revision")
    }
  } finally { amend = false; modelProposal = undefined; approveReview = true; reviewerRawOutput = undefined }
})

test("partial-selection review receives unchanged neighboring text and boundary guidance", async () => {
  const id = await newConversation(), text = "The parties shall deliver, promptly and safely.", selected = "deliver"
  const paragraphs = [{ id: "boundary-paragraph", text }], start = text.indexOf(selected)
  const selection = { id: "boundary-selection", text: selected, paragraphs, startOffset: start, endOffset: start + selected.length, saved: false, capturedAt: context.capturedAt }
  amend = true; modelProposal = { before: selected, after: "provide notice" }
  try {
    const response = await post("send", { conversationId: id, scope: "document", kind: "amendment", text: "Change the selected wording.", requestId: crypto.randomUUID(), attachments: [], context: { ...context, text, paragraphs, selection } })
    const turn = (await response.json() as { turn: WordChatTurn }).turn
    expect(turn.proposal?.before).toBe(selected)
    const reviewer = [...sessions.values()].find(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")!
    const prompt = sessionContext(reviewer.id), supplied = suppliedContext(prompt) as { proposal: NonNullable<WordChatTurn["proposal"]> }
    expect(prompt).toContain("punctuation and spacing at both boundaries")
    expect(prompt).toContain("never expand the reviewed span silently")
    expect(supplied.proposal.target.paragraphs).toEqual(paragraphs)
    expect(supplied.proposal.target.startOffset).toBe(start)
    expect(supplied.proposal.target.endOffset).toBe(start + selected.length)
    expect(supplied.proposal.after).toBe("provide notice")
  } finally { amend = false; modelProposal = undefined }
})

test("completed amendments without a usable proposal always expose an actionable error", async () => {
  const id = await newConversation()
  modelAnswer = "Prepared the tracked amendment for review."
  try {
    for (const proposal of [undefined, { before: "Wording absent from the current buffer.", after: "Replacement wording." }, { before: context.text, after: context.text }, { before: context.text, after: undefined }]) {
      amend = Boolean(proposal); modelProposal = proposal
      const before = generations
      const response = await post("send", { conversationId: id, scope: "document", mode: "redline", kind: "amendment", text: "Prepare the requested change for review.", requestId: crypto.randomUUID(), attachments: [], context })
      expect(response.status).toBe(200)
      const turn = (await response.json() as { turn: WordChatTurn }).turn
      expect(turn.status).toBe("complete")
      expect(turn.proposal).toBeUndefined()
      expect(turn.error?.trim().length).toBeGreaterThan(0)
      expect(turn.error).toMatch(/ask|select/i)
      expect(generations - before).toBe(1)
    }
    for (const output of ["Prepared the tracked amendment for review.", "{}"] ) {
      modelRawOutput = output
      const response = await post("send", { conversationId: id, scope: "document", mode: "redline", kind: "amendment", text: "Prepare the requested change for review.", requestId: crypto.randomUUID(), attachments: [], context })
      const turn = (await response.json() as { turn: WordChatTurn }).turn
      expect(turn.status).toBe("error")
      expect(turn.proposal).toBeUndefined()
      expect(turn.error).toContain("Try a new request")
      expect(turn.answer).toBe("")
    }
  } finally { amend = false; modelProposal = undefined; modelAnswer = undefined; modelRawOutput = undefined }
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

test("multipart chat uploads read DOCX, UTF-8 text, Markdown, CSV and bounded RTF as attached evidence", async () => {
  const id = await newConversation()
  const files = [
    { name: "Upload.docx", bytes: await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("DOCX upload evidence.")] }] })), expected: "DOCX upload evidence." },
    { name: "Upload.txt", bytes: Buffer.from("UTF-8 café upload evidence.\r\nSecond line."), expected: "UTF-8 café upload evidence.\nSecond line." },
    { name: "Upload.md", bytes: Buffer.from("# Markdown evidence\n\nThe agreed wording is café."), expected: "# Markdown evidence\n\nThe agreed wording is café." },
    { name: "Upload.csv", bytes: Buffer.from('name,amount\r\n"café",42'), expected: 'name,amount\n"café",42' },
    { name: "Upload.rtf", bytes: Buffer.from(String.raw`{\rtf1\ansi\ansicpg1252{\fonttbl{\f0 Hidden Font;}}\f0 RTF caf\'e9 evidence.\par Second line.}`, "latin1"), expected: "RTF café evidence.\nSecond line." },
  ]
  for (const file of files) {
    expect(matterDocumentName(file.name)).toBe(file.name)
    const form = new FormData()
    form.set("conversationId", id); form.set("mode", "request"); form.set("file", new File([Uint8Array.from(file.bytes)], file.name))
    const response = await app.request("/word/chat/upload", { method: "POST", body: form })
    expect(response.status, file.name).toBe(200)
    const { material } = await response.json() as { material: WordChatMaterial }
    expect(material.name).toBe(file.name)
    expect(material.provenance).toBe("request-upload")
    expect(material).not.toHaveProperty("text")
    const turn = (await (await send(id, "document", "Explain the attached upload", crypto.randomUUID(), [material])).json() as { turn: WordChatTurn }).turn
    expect(turn.status).toBe("complete")
    expect(turn.sources[0]?.quote.trim()).toBe(file.expected)
    expect(turn.sources[0]?.provenance).toBe("request-upload")
    expect(prompts.at(-1)).not.toContain("Hidden Font")
  }
})

test("RTF chat uploads refuse unreadable and parser-limit files without retaining an attachment", async () => {
  const id = await newConversation()
  const files = [
    Buffer.from("This is ordinary text renamed as RTF."),
    Buffer.from(String.raw`{\rtf1\ansi ` + "{".repeat(9_000) + "Over-nested content" + "}".repeat(9_000) + "}"),
    Buffer.concat([Buffer.from(String.raw`{\rtf1 `), Buffer.alloc(8 * 1024 * 1024)]),
  ]
  for (const bytes of files) {
    const form = new FormData()
    form.set("conversationId", id); form.set("mode", "request"); form.set("file", new File([bytes], "Unreadable.rtf"))
    const response = await app.request("/word/chat/upload", { method: "POST", body: form })
    expect(response.status).toBe(413)
    expect((await response.json() as { error: string }).error).toContain("readable text")
  }
  const response = await post("materials", { conversationId: id })
  expect((await response.json() as { materials: WordChatMaterial[] }).materials.some(material => material.provenance === "request-upload")).toBe(false)
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
  const other = (await (await send(id, "document", "Rewrite delivery again")).json() as { turn: WordChatTurn }).turn
  amend = false
  const current = readCurrentVersion(directory, "Draft.docx")!
  const db = openDb(directory)
  db.run("INSERT INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)", [auth.grant.token_hash, auth.grant.document_id, auth.grant.origin, auth.grant.issued_version, Date.now()])
  const identity = { document: auth.grant.document_id, tokenHash: auth.grant.token_hash, session: "synthetic-pane-session" }
  let owner = claimWordEditLease(db, identity)
  beginWordNativeOperation(db, owner, 0, "prepare-one", undefined, { conversationId: id, turnId: turn.id })
  let duringAuthor: (() => Promise<void>) | undefined
  const preparing = new Hono()
  installWordChat(preparing, () => auth, { fetch: transport, engineURL: "http://synthetic.invalid", authorAmendment: async (_ooxml, proposal) => {
    await duringAuthor?.()
    return { proposalId: proposal.id, ooxml: "synthetic-pure-artifact", paragraphs: [proposal.after], revisions: [], author: "doc.haus" }
  } })
  const prepare = () => preparing.request("/word/chat/prepare", { method: "POST", headers: { "content-type": "application/json", "x-dochaus-word-session": owner.session, "x-dochaus-word-lease": owner.generation, "x-dochaus-word-base-version": current.current.id },
    body: JSON.stringify({ conversationId: id, turnId: turn.id, ooxml: "synthetic-original-ooxml" }) })
  // A different reviewed proposal cannot borrow the admitted operation.
  const wrong = await preparing.request("/word/chat/prepare", { method: "POST", headers: { "content-type": "application/json", "x-dochaus-word-session": owner.session, "x-dochaus-word-lease": owner.generation, "x-dochaus-word-base-version": current.current.id }, body: JSON.stringify({ conversationId: id, turnId: other.id, ooxml: "synthetic-original-ooxml" }) })
  expect(wrong.status).toBe(409)
  // Starting authority passes; only the change during the await causes refusal.
  expect((await prepare()).status).toBe(200)
  duringAuthor = async () => { db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?", ["other-turn", owner.document]) }
  expect((await prepare()).status).toBe(409)
  db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?", [turn.id, owner.document])
  duringAuthor = async () => {
    settleWordNativeOperation(db, owner, 1, "prepare-one", "not-dispatched")
    const bytes = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("A concurrent acknowledged Word save.")] }] }))
    await writeDeliverableRevision(directory, "Draft.docx", bytes, { priorSha256: current.current.sha256, keepWorkingRecord: true, wordLeaseOwner: owner, version: { origin: "word" } })
  }
  expect((await prepare()).status).toBe(409)
  clearWordEditLease(db, auth.grant.document_id); db.close()
})

test("reviewed row proposals bind host table, reject forgery and require an independent reviewer", async () => {
  const id = await newConversation()
  const paragraphs = ["Name","Amount","Fee","Ten"].map((text,i) => ({id:`cell-${i}`,text}))
  const table = {tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2)]}
  const native = {...context,text:paragraphs.map(p => p.text).join("\r"),paragraphs,table}
  const request = (extra: object = {}) => post("send", {conversationId:id,scope:"document",text:"Append the agreed row",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native,...extra})
  modelRawOutput = JSON.stringify({answer:"Prepared for review.",citations:[],rowProposal:{cells:["Tax","Five"],explanation:"Add the agreed row",openItems:[]}})
  try {
    const approved = (await (await request()).json() as {turn:WordChatTurn}).turn
    expect(approved.rowProposal?.table).toEqual(table)
    expect(approved.proposal).toBeUndefined()
    const reviewer = [...sessions.values()].find(s => s.parentID === approved.id && s.metadata.type === "word-chat-review")!
    const review = suppliedContext(sessionContext(reviewer.id))
    expect(review.proposal.table).toEqual({columns:2,firstRow:["Name","Amount"],referenceRow:["Fee","Ten"],rows:[["Name","Amount"],["Fee","Ten"]]})
    expect(sessionContext(reviewer.id)).not.toContain("cell-0")
    const forged = {...approved.rowProposal!,cells:["Other","Five"]}
    expect((await request({rowProposal:forged})).status).toBe(409)
    const refined = (await (await request({rowProposal:approved.rowProposal,context:{...native,table:undefined}})).json() as {turn:WordChatTurn}).turn
    expect(refined.rowProposal?.table).toEqual(table)
    modelRawOutput = JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{cells:["Tax","Five"],explanation:"Add",openItems:[],tableIndex:0}})
    expect((await (await request()).json() as {turn:WordChatTurn}).turn.rowProposal).toBeUndefined()
    modelRawOutput = JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{cells:["Tax","Five"],explanation:"Add",openItems:[]},proposal:{before:"Fee",after:"Cost"}})
    expect((await (await request()).json() as {turn:WordChatTurn}).turn.rowProposal).toBeUndefined()
    modelRawOutput = JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{cells:["Tax","Five"],explanation:"Add",openItems:[]}})
    approveReview = false
    expect((await (await request()).json() as {turn:WordChatTurn}).turn.rowProposal).toBeUndefined()
  } finally { modelRawOutput = undefined; approveReview = true }
})

test("formatted row cell runs normalize from model output and reach review and refinement; mismatches are refused", async () => {
  const id = await newConversation()
  const paragraphs = ["Item","Note","Alpha","Plain"].map((text,i) => ({id:`runs-${i}`,text}))
  const table = {tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2)]}
  const native = {...context,text:paragraphs.map(p => p.text).join("\r"),paragraphs,table,selection:undefined}
  const request = (extra: object = {}) => post("send", {conversationId:id,scope:"document",text:"Append a row with the first word in bold",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native,...extra})
  const output = (row: object) => { modelRawOutput = JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{explanation:"Add row",openItems:[],...row}}) }
  const turnOf = async (extra: object = {}) => (await (await request(extra)).json() as {turn:WordChatTurn}).turn
  const canonical = [[{text:"Beta",bold:true,italic:false,underline:false},{text:" due",bold:false,italic:false,underline:false}],[{text:"Later",bold:false,italic:false,underline:false}]]
  try {
    output({cells:["Beta due","Later"],cellRuns:[[{text:"Beta",bold:true},{text:""},{text:" d"},{text:"ue"}],[{text:"Later",italic:false}]]})
    const approved = await turnOf()
    if (approved.rowProposal?.kind !== "table_row_insert") throw Error("missing row insertion")
    expect(approved.rowProposal.runs).toEqual(canonical)
    const reviewer = [...sessions.values()].find(s => s.parentID === approved.id && s.metadata.type === "word-chat-review")!
    expect(suppliedContext(sessionContext(reviewer.id)).proposal.runs).toEqual(canonical)
    const refined = await turnOf({rowProposal:approved.rowProposal,context:{...native,table:undefined}})
    expect(suppliedContext(sessionContext(refined.id)).refining.cellRuns).toEqual(canonical)

    output({cells:["Beta due","Later"],cellRuns:[[{text:"Beta",bold:true}],[{text:"Later"}]]})
    const mismatched = await turnOf()
    expect(mismatched.rowProposal).toBeUndefined()
    expect(mismatched.error).toBe("This table row could not be prepared for review. Select a supported table and try again.")

    output({cells:["Beta due","Later"],cellRuns:[[{text:"Beta due"}],[{text:"Later",bold:false}]]})
    const plain = await turnOf()
    expect(plain.rowProposal?.kind).toBe("table_row_insert")
    expect(plain.rowProposal && "runs" in plain.rowProposal).toBe(false)
    const plainReviewer = [...sessions.values()].find(s => s.parentID === plain.id && s.metadata.type === "word-chat-review")!
    expect(suppliedContext(sessionContext(plainReviewer.id)).proposal).not.toHaveProperty("runs")

    output({action:"delete",cells:["Alpha","Plain"]})
    expect((await turnOf()).rowProposal?.kind).toBe("table_row_delete")
    output({action:"delete",cells:["Alpha","Plain"],cellRuns:[[{text:"Alpha",bold:true}],[{text:"Plain"}]]})
    expect((await turnOf()).rowProposal).toBeUndefined()
  } finally { modelRawOutput = undefined }
})

test("structured answer parts preserve literal brackets and verify explicit references without legacy fallback", async () => {
  const id = await newConversation(), ref = {type:"reference" as const,sourceId:"word",quote:context.text}
  const request = () => send(id,"document","Explain the current text")
  try {
    const parts = [{type:"text" as const,text:'Literal [word] and ["a", "b"] remain text. '},ref]
    modelStructuredRawOutput = JSON.stringify({answerParts:parts})
    const good = (await (await request()).json() as {turn:WordChatTurn}).turn
    expect(good.status).toBe("complete"); expect(good.answerFormat).toBe("segments-v1"); expect(good.answerParts).toEqual(parts)
    expect(good.answer).toBe('Literal [word] and ["a", "b"] remain text. '); expect(good.sources[0]?.quote).toBe(context.text)
    for (const raw of [
      {answerParts:[{...ref,sourceId:"unknown"}]}, {answerParts:[{...ref,quote:"Invented quote"}]},
      {answerParts:[{type:"text",text:"x",unexpected:true}]}, {answer:"legacy [word]",citations:[{sourceId:"word",quote:context.text}]},
      {answerParts:parts,answer:"mixed",citations:[]},
    ]) {
      modelStructuredRawOutput = JSON.stringify(raw)
      const bad = (await (await request()).json() as {turn:WordChatTurn}).turn
      expect(bad.status).toBe("error"); expect(bad.answerParts).toBeUndefined(); expect(bad.sources).toEqual([])
    }
    modelStructuredRawOutput = JSON.stringify({answerParts:parts})
    const saved = (await (await request()).json() as {turn:WordChatTurn}).turn
    const child = sessions.get(saved.id)!
    delete child.metadata.answerFormat
    const message = messages.get(child.id)!.find(value => value.type === "assistant")!
    message.content = [{type:"text",text:JSON.stringify({answer:"Legacy reference [word]",citations:[{sourceId:"word",quote:context.text}]})}]
    const old = (await (await post("poll",{conversationId:id,turnId:child.id})).json() as {turn:WordChatTurn}).turn
    expect(old.answerFormat).toBeUndefined(); expect(old.answer).toBe("Legacy reference [word]"); expect(old.sources[0]?.quote).toBe(context.text)
  } finally { modelStructuredRawOutput = undefined }
})

test("structured answers are suppressed while review is pending and when review fails", async () => {
  const id = await newConversation()
  modelStructuredRawOutput = JSON.stringify({answerParts:[{type:"text",text:"Prepared."},{type:"reference",sourceId:"word",quote:context.text}],proposal:{before:context.text,after:context.text.replace("Tuesday","Wednesday"),explanation:"Change",openItems:[]}})
  try {
    queueReview = true
    const pending = (await (await post("send",{conversationId:id,scope:"document",text:"Change the date",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context})).json() as {turn:WordChatTurn}).turn
    expect(pending.status).toBe("running"); expect(pending.answerParts).toBeUndefined(); expect(pending.sources).toEqual([])
    queueReview = false; approveReview = false
    const rejected = (await (await post("send",{conversationId:await newConversation(),scope:"document",text:"Change the date",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context})).json() as {turn:WordChatTurn}).turn
    expect(rejected.proposal).toBeUndefined(); expect(rejected.sources).toEqual([])
    expect(rejected.answerParts).toEqual([{type:"text",text:rejected.answer}]); expect(rejected.answer).toContain("reviewer")
  } finally { queueReview = false; approveReview = true; modelStructuredRawOutput = undefined }
})

test("inline formatting review binds exact Word text, preserves refinement target and refuses invalid shapes", async () => {
  const id = await newConversation(), text = "Keep the payment terms and delivery terms."
  const native = {...context,text,paragraphs:[{id:"inline-native",text}],selection:undefined}
  const request = (extra:object = {}) => post("send",{conversationId:id,scope:"document",text:"Bold payment",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native,...extra})
  const setOutput = (proposal:object, extra:object = {}) => {modelStructuredRawOutput=JSON.stringify({answerParts:[{type:"text",text:"Prepared for review."}],proposal,...extra})}
  try {
    setOutput({before:"payment",after:"payment",inlineFormatting:{bold:true,italic:false},explanation:"Emphasize",openItems:[]})
    const first = (await (await request()).json() as {turn:WordChatTurn}).turn
    expect(first.kind).toBe("formatting"); expect(first.proposal?.inlineFormatting).toEqual({bold:true,italic:false})
    expect(first.proposal?.target.startOffset).toBe(text.indexOf("payment"))
    const reviewer = [...sessions.values()].find(s => s.parentID === first.id && s.metadata.type === "word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain("false turns a property off")
    const original = first.proposal!
    setOutput({before:"payment",after:"payment",inlineFormatting:{underline:true},explanation:"Underline",openItems:[]})
    const second = (await (await request({proposal:original,context:{...native,capturedAt:"later"}})).json() as {turn:WordChatTurn}).turn
    expect(second.proposal?.target).toEqual(original.target); expect(second.proposal?.inlineFormatting).toEqual({underline:true})
    const third = (await (await request({proposal:second.proposal,context:{...native,capturedAt:"later again"}})).json() as {turn:WordChatTurn}).turn
    expect(third.proposal?.target).toEqual(original.target)
    expect((await request({proposal:original,context:{...native,text:text.replace("payment","changed"),paragraphs:[{id:"inline-native",text:text.replace("payment","changed")}]}})).status).toBe(409)
    for (const bad of [
      {before:"payment",after:"payment",inlineFormatting:{}},
      {before:"payment",after:"payment",inlineFormatting:{bold:"true"}},
      {before:"payment",after:"payment",inlineFormatting:{bold:true,offset:0}},
      {before:"payment",after:"changed",inlineFormatting:{bold:true}},
      {before:"payment",after:"payment",inlineFormatting:{bold:true},formatting:[{paragraph:0,style:"Heading1"}]},
      {before:"terms",after:"terms",inlineFormatting:{bold:true}},
    ]) {
      setOutput(bad)
      const refused = (await (await request()).json() as {turn:WordChatTurn}).turn
      expect(refused.proposal).toBeUndefined(); expect(refused.error).toBeTruthy()
    }
    setOutput({before:"payment",after:"changed"},{inlineFormatting:{bold:true}})
    expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
    const termStart = text.lastIndexOf("terms")
    const selected = {...native,selection:{id:"chosen-terms",text:"terms",paragraphs:native.paragraphs,startOffset:termStart,endOffset:termStart+5,capturedAt:native.capturedAt,saved:false}}
    setOutput({before:"terms",after:"terms",inlineFormatting:{italic:true}})
    const selectedTurn = (await (await request({context:selected})).json() as {turn:WordChatTurn}).turn
    expect(selectedTurn.proposal?.target.startOffset).toBe(termStart)
    const legacy = sessions.get(first.id)!
    delete legacy.metadata.inlineContract
    const reread = (await (await post("poll",{conversationId:id,turnId:first.id})).json() as {turn:WordChatTurn}).turn
    expect(reread.proposal).toBeUndefined()
  } finally { modelStructuredRawOutput = undefined }
})

test("row prepare uses stored review and current bytes, and rechecks the pending lease", async () => {
  const id = await newConversation(), paragraphs = ["Header","Value"].map((text,i) => ({id:`row-${i}`,text}))
  const native = {...context,text:"Header\rValue",paragraphs,table:{tableIndex:0,rows:[paragraphs]}}
  modelRawOutput = JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{cells:["New","Value"],explanation:"Add row",openItems:[]}})
  let turn: WordChatTurn
  try { turn = (await (await post("send",{conversationId:id,scope:"document",text:"Append a row",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native})).json() as {turn:WordChatTurn}).turn }
  finally { modelRawOutput = undefined }
  expect(turn.rowProposal).toBeDefined()
  const current = readCurrentVersion(directory, "Draft.docx")!, db = openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)", [auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner = claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"row-prepare-session"})
  const seq = (db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"row-prepare",undefined,{conversationId:id,turnId:turn.id})
  let changed = false
  const preparing = new Hono()
  const validator = installWordChat(preparing, () => auth, {fetch:transport,engineURL:"http://synthetic.invalid",authorTableRow:async (stored, proposal, bytes) => {
    expect(stored.table).toEqual(native.table); expect(turn.rowProposal).toEqual(proposal); expect(Buffer.from(bytes).toString()).toBe("current unsaved bytes")
    if (changed) db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other-turn",owner.document])
    return {kind:"table_row_insert",proposalId:proposal.id,sourceSha256:"source",table:proposal.table,cells:proposal.cells,author:"doc.haus",writerRequest:{source_sha256:"source",target_handle:"handle",position:"after",cells:proposal.cells,author:"doc.haus"},authoredBytes:new Uint8Array([1,2]),revisions:[]}
  }})
  const headers = {"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare = (input:object) => preparing.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn.id,...input})})
  try {
    await validator.withValidatedProposal(auth,{conversationId:id,turnId:turn.id},async () => {})
    const response = await prepare({docxBase64:Buffer.from("current unsaved bytes").toString("base64"),rowProposal:{...turn.rowProposal,cells:["Forged","Value"]}})
    expect(response.status).toBe(200)
    expect((await response.json() as {authoredBase64:string}).authoredBase64).toBe("AQI=")
    expect((await prepare({docxBase64:"%%%"})).status).toBe(413)
    changed = true
    expect((await prepare({docxBase64:Buffer.from("current unsaved bytes").toString("base64")})).status).toBe(409)
  } finally {
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn.id,owner.document])
    settleWordNativeOperation(db,owner,seq + 1,"row-prepare","not-dispatched")
    clearWordEditLease(db,owner.document); db.close()
  }
})

test("prepare returns actionable formatting failures while keeping other authoring errors private", async () => {
  const id = await newConversation()
  amend = true; modelProposal = { before: context.text, after: context.text, formatting: [{ paragraph: 0, spacingAfter: 12 }] }
  let turn: WordChatTurn
  try { turn = (await (await send(id, "document", "Increase paragraph spacing")).json() as { turn: WordChatTurn }).turn }
  finally { amend = false; modelProposal = undefined }
  expect(turn.proposal?.formatting).toEqual([{ paragraph: 0, spacingAfter: 12 }])
  const db = openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)", [auth.grant.token_hash, auth.grant.document_id, auth.grant.origin, auth.grant.issued_version, Date.now()])
  const owner = claimWordEditLease(db, { document: auth.grant.document_id, tokenHash: auth.grant.token_hash, session: "synthetic-format-session" })
  const current = readCurrentVersion(directory, "Draft.docx")!
  const seq = (db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as { native_sequence: number }).native_sequence
  beginWordNativeOperation(db, owner, seq, "prepare-format", undefined, { conversationId: id, turnId: turn.id })
  let failure: Error = new Error("Private parser detail")
  const preparing = new Hono()
  installWordChat(preparing, () => auth, { fetch: transport, engineURL: "http://synthetic.invalid", authorAmendment: async () => { throw failure } })
  const prepare = () => preparing.request("/word/chat/prepare", { method: "POST", headers: { "content-type": "application/json", "x-dochaus-word-session": owner.session, "x-dochaus-word-lease": owner.generation, "x-dochaus-word-base-version": current.current.id },
    body: JSON.stringify({ conversationId: id, turnId: turn.id, ooxml: "synthetic-original-ooxml" }) })
  try {
    const generic = await prepare()
    expect(generic.status).toBe(409)
    expect((await generic.json() as { error: string }).error).toContain("passage changed")
    const message = "That paragraph style is not available in this document. Choose an existing style in Word, then prepare the change again."
    failure = new WordFormattingError(message)
    const actionable = await prepare()
    expect(actionable.status).toBe(409)
    expect((await actionable.json() as { error: string }).error).toBe(message)
    for (const code of ["docx_protected_inline_content:private-proposal-id", "docx_protected_inline_content_changed"]) {
      failure = new Error(code)
      const refused = await prepare(), result = await refused.json() as { error: string }
      expect(refused.status).toBe(409)
      expect(result.error).toContain("Word")
      expect(result.error).not.toContain("docx_")
      expect(result.error).not.toContain("private-proposal-id")
    }
  } finally { settleWordNativeOperation(db, owner, seq + 1, "prepare-format", "not-dispatched"); clearWordEditLease(db, auth.grant.document_id); db.close() }
})

test("prepare withholds the artifact when attached evidence changes during pure authoring", async () => {
  const name = "Authoring evidence.docx"
  await writeDeliverable(directory, name, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("The source supports the requested delivery review.")] }] })))
  const material = wordChatMaterials(auth).find(value => value.name === name)!, bytes = readFileSync(path.join(directory, name))
  const id = await newConversation()
  amend = true
  let turn: WordChatTurn
  try { turn = (await (await send(id, "document", "Rewrite delivery for Wednesday", crypto.randomUUID(), [material])).json() as { turn: WordChatTurn }).turn }
  finally { amend = false }
  expect(turn.proposal).toBeDefined()
  const db = openDb(directory), current = readCurrentVersion(directory, "Draft.docx")!.current.id
  const owner = claimWordEditLease(db, { document: auth.grant.document_id, tokenHash: auth.grant.token_hash, session: "synthetic-evidence-race" })
  const seq = (db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as { native_sequence: number }).native_sequence
  beginWordNativeOperation(db, owner, seq, "prepare-evidence", undefined, { conversationId: id, turnId: turn.id })
  let changeEvidence = false, authors = 0
  const preparing = new Hono()
  installWordChat(preparing, () => auth, { fetch: transport, engineURL: "http://synthetic.invalid", authorAmendment: async (_ooxml, proposal) => {
    authors++
    await Promise.resolve()
    if (changeEvidence) writeFileSync(path.join(directory, name), "An external source replacement during authoring.")
    return { proposalId: proposal.id, ooxml: "synthetic-reviewed-artifact", paragraphs: [proposal.after], revisions: [], author: "doc.haus" }
  } })
  const prepare = () => preparing.request("/word/chat/prepare", { method: "POST", headers: { "content-type": "application/json", "x-dochaus-word-session": owner.session, "x-dochaus-word-lease": owner.generation, "x-dochaus-word-base-version": current }, body: JSON.stringify({ conversationId: id, turnId: turn.id, ooxml: "synthetic-original-ooxml" }) })
  try {
    expect((await prepare()).status).toBe(200)
    changeEvidence = true
    const changed = await prepare()
    expect(changed.status).toBe(409)
    const result = await changed.json() as { error: string; ooxml?: string }
    expect(result.error).toContain("Supporting material changed")
    expect(result.ooxml).toBeUndefined()
    expect(authors).toBe(2)
    expect(readCurrentVersion(directory, "Draft.docx")!.current.id).toBe(current)
  } finally { writeFileSync(path.join(directory, name), bytes); settleWordNativeOperation(db, owner, seq + 1, "prepare-evidence", "not-dispatched"); clearWordEditLease(db, auth.grant.document_id); db.close() }
})

test("suggestion batches retain document context as well as the selected passage", () => {
  const paragraphs = [{ id: "first", text: "Unselected prefix. First selected wording." }, { id: "second", text: "Second selected wording. Unselected suffix." }]
  const text = "First selected wording.\rSecond selected wording."
  const selection = { id: "selected", text, paragraphs, startOffset: 19, endOffset: 24, capturedAt: context.capturedAt, saved: false }
  const result = wordChatSuggestionBatches({ ...context, paragraphs, text: paragraphs.map(paragraph => paragraph.text).join("\r"), selection })
  const supplied = result.batches.flat().map(paragraph => paragraph.text).join("\r")
  expect(supplied).toBe(paragraphs.map(paragraph => paragraph.text).join("\r"))
  expect(result.batches.flat().map(paragraph => paragraph.id)).toEqual(["first", "second"])
  expect(supplied).toContain(text)
  expect(supplied).toContain("Unselected")
})

test("follow-up evidence excludes changed Word buffers while retaining unchanged conversation context", async () => {
  const id = await newConversation()
  const monday = { ...context, text: "The current Word buffer uses Monday delivery.", paragraphs: [{ id: "11", text: "The current Word buffer uses Monday delivery." }] }
  const first = await post("send", { conversationId: id, scope: "document", text: "Explain delivery", requestId: crypto.randomUUID(), attachments: [], context: monday })
  expect(first.status).toBe(200)
  const second = await send(id, "document", "Explain the current day")
  expect(second.status).toBe(200)
  const changed = suppliedContext(prompts.at(-1)!) as { compatibleDiscussion: unknown[]; sources: { quote: string }[] }
  expect(changed.compatibleDiscussion).toEqual([])
  expect(prompts.at(-1)).not.toContain("Monday")
  expect(changed.sources[0]?.quote).toContain("Tuesday")
  expect((await send(id, "document", "Explain that once more")).status).toBe(200)
  const unchanged = suppliedContext(prompts.at(-1)!) as { compatibleDiscussion: { question: string; answer: string }[] }
  expect(unchanged.compatibleDiscussion).toHaveLength(1)
  expect(unchanged.compatibleDiscussion[0]?.question).toBe("Explain the current day")
  expect(unchanged.compatibleDiscussion[0]?.answer).toContain("Tuesday")
  const visible = await post("conversation", { conversationId: id })
  expect((await visible.json() as { turns: WordChatTurn[] }).turns).toHaveLength(3)
})
test("refinement keeps its reviewed target when unrelated Word wording changes", async () => {
  const id = await newConversation(); amend = true
  const turn = (await (await send(id, "document", "Revise delivery to Wednesday")).json() as { turn: WordChatTurn }).turn
  const updated = { ...context, text: context.text + "\rAn unrelated unsaved note.", paragraphs: [...context.paragraphs, { id: "12", text: "An unrelated unsaved note." }] }
  const result = await post("send", { conversationId: id, scope: "document", text: "Keep the same target and refine the delivery change", requestId: crypto.randomUUID(), attachments: [], context: updated, proposal: turn.proposal })
  amend = false
  expect(result.status).toBe(200)
  expect((await result.json() as { turn: WordChatTurn }).turn.proposal?.target).toEqual(turn.proposal!.target)
})

test("selection identity separates follow-ups while incidental capture IDs do not", async () => {
  const id = await newConversation()
  const selected = (text: string) => ({ ...context, selection: { id: crypto.randomUUID(), text, paragraphs: context.paragraphs, startOffset: context.text.indexOf(text), endOffset: context.text.indexOf(text) + text.length, capturedAt: new Date().toISOString(), saved: false } })
  for (const [question, selectedContext] of [["Explain selected day", selected("Tuesday")], ["Explain selected delivery", selected("delivery")], ["Explain it again", selected("delivery")]] as const) {
    expect((await post("send", { conversationId: id, scope: "document", text: question, requestId: crypto.randomUUID(), attachments: [], context: selectedContext })).status).toBe(200)
    const supplied = suppliedContext(prompts.at(-1)!) as { compatibleDiscussion: { question: string }[] }
    if (question === "Explain selected delivery") expect(supplied.compatibleDiscussion).toEqual([])
    if (question === "Explain it again") expect(supplied.compatibleDiscussion.map(turn => turn.question)).toEqual(["Explain selected delivery"])
  }
})

test("authorized conversation returns its durable applied outcome and rejects replay validation", async () => {
  const id = await newConversation(); amend = true
  const result = (await (await send(id, "document", "Rewrite this for Wednesday delivery")).json() as { turn: WordChatTurn }).turn
  amend = false
  expect(result.proposal).toBeDefined()
  const validator = installWordChat(new Hono(), () => auth, { fetch: transport, engineURL: "http://synthetic.invalid" })
  expect(await validator.withValidatedProposal(auth, { conversationId: id, turnId: result.id }, async () => "admitted")).toBe("admitted")
  const db = openDb(directory)
  try { db.run("INSERT INTO word_chat_applied (document_id, conversation_id, turn_id) VALUES (?, ?, ?)", [auth.grant.document_id, id, result.id]) } finally { db.close() }
  const opened = await (await post("conversation", { conversationId: id })).json() as { turns: WordChatTurn[] }
  expect(opened.turns.find(turn => turn.id === result.id)?.proposalApplied).toBe(true)
  expect((await post("prepare", { conversationId: id, turnId: result.id, ooxml: "synthetic" })).status).toBe(409)
  await expect(validator.withValidatedProposal(auth, { conversationId: id, turnId: result.id }, async () => {})).rejects.toThrow("fresh proposal")
  await expect(validator.withValidatedProposal(auth, { conversationId: "foreign", turnId: result.id }, async () => {})).rejects.toThrow()
})


test("native amendment admission validates the reviewed turn again after the artifact lock", async () => {
  const id = await newConversation(); amend = true
  const first = (await (await send(id, "document", "Rewrite delivery Wednesday")).json() as { turn: WordChatTurn }).turn
  const second = (await (await send(id, "document", "Rewrite delivery again")).json() as { turn: WordChatTurn }).turn
  amend = false
  const token = randomBytes(32).toString("base64url"), tokenHash = createHash("sha256").update(token).digest("hex")
  const current = readCurrentVersion(directory, "Draft.docx")!
  const db = openDb(directory)
  db.run("INSERT INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)", [tokenHash, auth.grant.document_id, auth.grant.origin, current.current.id, Date.now()])
  const owner = claimWordEditLease(db, { document: auth.grant.document_id, tokenHash, session: "route-test-pane" })
  const sequence = (db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as { native_sequence: number }).native_sequence
  const routes = wordDocumentRoutes({ matter: () => auth.matter, matters: () => [auth.matter], locks: () => [auth.target.docPath], write: async () => { throw new Error("not used") }, wordChatOptions: { fetch: transport, engineURL: "http://synthetic.invalid" } })
  const headers = { "content-type": "application/json", Authorization: `Bearer ${token}`, Origin: auth.grant.origin, "x-dochaus-word-session": owner.session, "x-dochaus-word-lease": owner.generation, "x-dochaus-word-base-version": current.current.id }
  const request = (input: object) => routes.request("/word/native-operation", { method: "POST", headers, body: JSON.stringify(input) })
  try {
    const binding = { conversationId: id, turnId: first.id }
    expect((await request({ action: "begin", intent: "amendment", sequence, operationId: "first", proposal: binding })).status).toBe(200)
    expect((await request({ action: "settle", sequence: sequence + 1, operationId: "first", outcome: "confirmed" })).status).toBe(200)
    expect((await request({ action: "begin", intent: "amendment", sequence: sequence + 1, operationId: "replay", proposal: binding })).status).toBe(409)
    let release!: () => void, entered!: () => void, validated!: () => void
    const held = new Promise<void>(resolve => { release = resolve }), locked = new Promise<void>(resolve => { entered = resolve }), initial = new Promise<void>(resolve => { validated = resolve })
    const lock = withArtifactMutationLocksAsync(directory, [auth.target.docPath], async () => { entered(); await held })
    await locked
    engineHook = async (session, tail, method) => { if (session.id === second.id && tail === "/message" && method === "GET") validated() }
    const pending = request({ action: "begin", intent: "amendment", sequence: sequence + 1, operationId: "second", proposal: { conversationId: id, turnId: second.id } })
    await initial
    sessions.get(second.id)!.metadata.cancelled = true
    release(); await lock
    expect((await pending).status).toBe(409)
    expect((db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as { native_sequence: number }).native_sequence).toBe(sequence + 1)
  } finally { engineHook = undefined; clearWordEditLease(db, owner.document); db.close() }
})

test("paragraph insertion prepare uses stored reviewed proposal and existing lease fence", async () => {
  const id=await newConversation(), text=context.text, inserted="A further obligation."
  modelStructuredRawOutput=JSON.stringify({answerParts:[{type:"text",text:"Prepared."}],proposal:{before:text,after:text+"\r"+inserted,formatting:[{paragraph:1,style:"Normal"}],explanation:"Add paragraph",openItems:[]}})
  let turn:WordChatTurn
  try { turn=(await (await post("send",{conversationId:id,scope:"document",text:"Add a paragraph",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context})).json() as {turn:WordChatTurn}).turn }
  finally { modelStructuredRawOutput=undefined }
  expect(turn.proposal?.after).toBe(text+"\r"+inserted)
  expect(sessionContext(turn.id)).toContain("inserting or deleting paragraphs alone does not request formatting")
  const reviewer = [...sessions.values()].find(session => session.parentID === turn.id && session.metadata.type === "word-chat-review")!
  expect(sessionContext(reviewer.id)).toContain("Reject unrequested restyling when the instruction requires formatting preservation")
  const current=readCurrentVersion(directory,"Draft.docx")!, db=openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",[auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner=claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"insert-prepare-session"})
  const seq=(db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"insert-prepare",undefined,{conversationId:id,turnId:turn.id})
  const preparing=new Hono(); installWordChat(preparing,()=>auth,{fetch:transport,engineURL:"http://synthetic.invalid"})
  const headers={"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare=(extra:object={})=>preparing.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn.id,ooxml:"current capture",...extra})})
  try {
    const response=await prepare({proposal:{...turn.proposal,after:text+"\rForged"}})
    expect(response.status).toBe(200)
    expect(await response.json()).toEqual({kind:"paragraph_insert",proposalId:turn.proposal!.id,target:turn.proposal!.target,text:inserted,formatting:[{paragraph:1,style:"Normal"}]})
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other-turn",owner.document])
    expect((await prepare()).status).toBe(409)
  } finally {
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn.id,owner.document])
    settleWordNativeOperation(db,owner,seq+1,"insert-prepare","not-dispatched"); clearWordEditLease(db,owner.document); db.close()
  }
})

test("explicit paragraph deletion remains distinct from clearing text through review and prepare",async()=>{
  const id=await newConversation(), text=context.text
  const request=(extra:object={})=>post("send",{conversationId:id,scope:"document",text:"Remove this paragraph",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context,...extra})
  const output=(proposal:object,extra:object={})=>{modelStructuredRawOutput=JSON.stringify({answerParts:[{type:"text",text:"Prepared."}],proposal,...extra})}
  let turn:WordChatTurn
  try {
    output({before:text,after:"",paragraphAction:"delete",explanation:"Remove paragraph",openItems:[]})
    turn=(await (await request()).json() as {turn:WordChatTurn}).turn
    expect(turn.proposal?.paragraphAction).toBe("delete")
    const contract = JSON.parse(sessionContext(turn.id).split("\n").at(-2)!).responseContract
    expect(contract).toContain('set proposal.paragraphAction to "delete"')
    expect(contract).toContain("To clear wording while keeping the paragraph, omit paragraphAction")
    const reviewer=[...sessions.values()].find(session=>session.parentID===turn.id && session.metadata.type==="word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain('"paragraphAction":"delete"')
    const refined=(await (await request({proposal:turn.proposal,context:{...context,capturedAt:"later"}})).json() as {turn:WordChatTurn}).turn
    expect(refined.proposal?.target).toEqual(turn.proposal!.target); expect(refined.proposal?.paragraphAction).toBe("delete")
    expect((await request({proposal:turn.proposal,context:{...context,text:"changed",paragraphs:[{id:"11",text:"changed"}]}})).status).toBe(409)
    for (const bad of [
      {before:text,after:"",paragraphAction:"clear"}, {before:text,after:"change",paragraphAction:"delete"},
      {before:text,after:"",paragraphAction:"delete",formatting:[{paragraph:0,style:"Normal"}]},
      {before:text,after:"",paragraphAction:"delete",inlineFormatting:{bold:true}},
      {before:"Tuesday",after:"",paragraphAction:"delete"},
    ]) { output(bad); const result=(await (await request()).json() as {turn:WordChatTurn}).turn; expect(result.proposal).toBeUndefined() }
    output({before:text,after:""},{paragraphAction:"delete"})
    expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
    output({before:text,after:"",explanation:"Clear wording",openItems:[]})
    const clear=(await (await request()).json() as {turn:WordChatTurn}).turn
    expect(clear.proposal?.paragraphAction).toBeUndefined(); expect(clear.proposal?.after).toBe("")
  } finally {modelStructuredRawOutput=undefined}
  const current=readCurrentVersion(directory,"Draft.docx")!,db=openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",[auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner=claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"delete-prepare-session"})
  const seq=(db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"delete-prepare",undefined,{conversationId:id,turnId:turn.id})
  const preparing=new Hono(); installWordChat(preparing,()=>auth,{fetch:transport,engineURL:"http://synthetic.invalid"})
  const headers={"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare=()=>preparing.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn.id,ooxml:"current capture",proposal:{...turn.proposal,paragraphAction:"clear"}})})
  try {
    const result=await prepare(); expect(result.status).toBe(200)
    expect(await result.json()).toEqual({kind:"paragraph_delete",proposalId:turn.proposal!.id,target:turn.proposal!.target})
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other-turn",owner.document])
    expect((await prepare()).status).toBe(409)
  } finally {
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn.id,owner.document])
    settleWordNativeOperation(db,owner,seq+1,"delete-prepare","not-dispatched");clearWordEditLease(db,owner.document);db.close()
  }
})


test("reviewed decimal list append is server bound through preparation and rejects forged list authority",async()=>{
  const id=await newConversation(), text=context.text
  const list={id:3,format:"decimal" as const,members:[{...context.paragraphs[0]!,level:0,label:"1."}]}
  const listContext={...context,list}
  const request=(extra:object={})=>post("send",{conversationId:id,scope:"document",text:"Append an item",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:listContext,...extra})
  const output=(extra:object={})=>{modelStructuredRawOutput=JSON.stringify({answerParts:[{type:"text",text:"Prepared."}],proposal:{before:text,after:text+"\rAnother item",paragraphAction:"append_list_item",explanation:"Add",openItems:[],...extra}})}
  let turn:WordChatTurn
  try {
    output(); turn=(await (await request()).json() as {turn:WordChatTurn}).turn
    expect(turn.proposal?.list).toEqual(list)
    expect(sessionContext(turn.id)).toContain('append_list_item')
    const reviewer=[...sessions.values()].find(s=>s.parentID===turn.id && s.metadata.type==="word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain('append_list_item')
    expect((await request({proposal:turn.proposal,context:{...listContext,list:{...list,id:4}}})).status).toBe(409)
    output({list}); expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
    output({formatting:[{paragraph:1,style:"Normal"}]}); expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
  } finally {modelStructuredRawOutput=undefined}
  const current=readCurrentVersion(directory,"Draft.docx")!,db=openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",[auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner=claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"list-prepare-session"})
  const seq=(db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"list-prepare",undefined,{conversationId:id,turnId:turn.id})
  const app=new Hono(); installWordChat(app,()=>auth,{fetch:transport,engineURL:"http://synthetic.invalid"})
  const headers={"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare=()=>app.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn.id,ooxml:"current capture",proposal:{...turn.proposal,list:{...list,id:99}}})})
  try {
    const response=await prepare(); expect(response.status).toBe(200)
    expect(await response.json()).toEqual({kind:"list_insert",proposalId:turn.proposal!.id,target:turn.proposal!.target,text:"Another item",list})
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other",owner.document]); expect((await prepare()).status).toBe(409)
  } finally {db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn.id,owner.document]);settleWordNativeOperation(db,owner,seq+1,"list-prepare","not-dispatched");clearWordEditLease(db,owner.document);db.close()}
})


test("row deletion resolves only a unique exact row or a same-read selection and keeps refinement bound", async () => {
  const id = await newConversation(), paragraphs = ["Heading", "Value", "Repeat", "Two", "Repeat", "Two", "Other", "Three"].map((text,i)=>({id:`delete-${i}`,text}))
  const table = {tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2,4),paragraphs.slice(4,6),paragraphs.slice(6,8)]}
  const native = {...context,text:paragraphs.map(p=>p.text).join("\r"),paragraphs,table,selection:undefined}
  const request = (extra:object={}) => post("send",{conversationId:id,scope:"document",text:"Delete the requested row",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native,...extra})
  const output = (cells:string[], action:unknown="delete") => {modelRawOutput=JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{action,cells,explanation:"Remove row",openItems:[]}})}
  try {
    output(["Repeat","Two"])
    expect((await (await request()).json() as {turn:WordChatTurn}).turn.rowProposal).toBeUndefined()
    const crossing = {...native,selection:{id:"crossing",text:"Repeat\rTwo\rOther",paragraphs:paragraphs.slice(4,7),startOffset:0,endOffset:5,capturedAt:native.capturedAt,saved:false}}
    expect((await (await request({context:crossing})).json() as {turn:WordChatTurn}).turn.rowProposal).toBeUndefined()
    const selected = {...native,selection:{id:"selection",text:"Repeat",paragraphs:[paragraphs[4]!],startOffset:0,endOffset:6,capturedAt:native.capturedAt,saved:false}}
    const first = (await (await request({context:selected})).json() as {turn:WordChatTurn}).turn
    expect(first.rowProposal?.kind).toBe("table_row_delete"); if (first.rowProposal?.kind !== "table_row_delete") throw Error("missing delete proposal")
    expect(first.rowProposal.rowIndex).toBe(2)
    const reviewer = [...sessions.values()].find(s=>s.parentID===first.id&&s.metadata.type==="word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain("exact existing row")
    output(["Other","Three"])
    expect((await (await request({rowProposal:first.rowProposal,context:selected})).json() as {turn:WordChatTurn}).turn.rowProposal).toBeUndefined()
    output(["Other","Three"])
    const unique = (await (await request()).json() as {turn:WordChatTurn}).turn.rowProposal
    expect(unique?.kind === "table_row_delete" && unique.rowIndex).toBe(3)
  } finally {modelRawOutput=undefined}
})

test("deletion prepare uses stored target and rechecks the lease after authoring", async () => {
  const id = await newConversation(), paragraphs = ["Heading","Value","Remove","Row"].map((text,i)=>({id:`prepare-delete-${i}`,text}))
  const table = {tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2)]}
  const native = {...context,text:paragraphs.map(p=>p.text).join("\r"),paragraphs,table,selection:undefined}
  modelRawOutput=JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{action:"delete",cells:["Remove","Row"],explanation:"Remove",openItems:[]}})
  let turn: WordChatTurn
  try { turn = (await (await post("send",{conversationId:id,scope:"document",text:"Delete the row",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native})).json() as {turn:WordChatTurn}).turn }
  finally {modelRawOutput=undefined}
  if (turn.rowProposal?.kind !== "table_row_delete") throw Error("missing deletion")
  const current=readCurrentVersion(directory,"Draft.docx")!, db=openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",[auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner=claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"delete-prepare-session"})
  const seq=(db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"delete-prepare",undefined,{conversationId:id,turnId:turn.id})
  let invalidate=false
  const preparing=new Hono()
  installWordChat(preparing,()=>auth,{fetch:transport,engineURL:"http://synthetic.invalid",authorTableRowDeletion:async (stored,proposal,bytes)=>{
    expect(stored.table).toEqual(table);expect(turn.rowProposal).toEqual(proposal);expect(Buffer.from(bytes).toString()).toBe("current bytes")
    if (invalidate) db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other",owner.document])
    return {kind:"table_row_delete",proposalId:proposal.id,sourceSha256:"source",table:proposal.table,rowIndex:proposal.rowIndex,cells:proposal.cells,author:"doc.haus",writerRequest:{source_sha256:"source",target_handle:"handle",author:"doc.haus"},authoredBytes:new Uint8Array([1,2]),revisions:[]}
  }})
  const headers={"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare=()=>preparing.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn.id,docxBase64:Buffer.from("current bytes").toString("base64"),rowProposal:{...turn.rowProposal,rowIndex:0}})})
  try {
    const response=await prepare();expect(response.status).toBe(200);expect((await response.json() as {kind:string,authoredBase64:string}).kind).toBe("table_row_delete")
    invalidate=true;expect((await prepare()).status).toBe(409)
  } finally {
    db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn.id,owner.document]);settleWordNativeOperation(db,owner,seq+1,"delete-prepare","not-dispatched");clearWordEditLease(db,owner.document);db.close()
  }
})


test("large table omits deletion guidance while retaining append guidance", async () => {
  const id=await newConversation(), paragraphs=["Header","Value","x".repeat(4001),"Middle","Last","Row"].map((text,i)=>({id:`large-row-${i}`,text}))
  const table={tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2,4),paragraphs.slice(4,6)]}
  const native={...context,text:paragraphs.map(p=>p.text).join("\r"),paragraphs,table,selection:undefined}
  modelRawOutput=JSON.stringify({answer:"Prepared.",citations:[],rowProposal:{cells:["Added","Row"],explanation:"Append",openItems:[]}})
  try {
    const turn=(await (await post("send",{conversationId:id,scope:"document",text:"Append a row",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native})).json() as {turn:WordChatTurn}).turn
    expect(turn.rowProposal?.kind).toBe("table_row_insert")
    const prompt=sessionContext(turn.id)
    expect(prompt).toContain("append one row");expect(prompt).not.toContain("For deleting one existing row")
    expect(suppliedContext(prompt).table.rows).toBeUndefined()
  } finally {modelRawOutput=undefined}
})

test("paragraph merge review, refinement and stored preparation remain bound to one boundary", async () => {
  const id=await newConversation(), paragraphs=[{id:"merge-a",text:"First sentence."},{id:"merge-b",text:"Second sentence."},{id:"merge-c",text:"Third sentence."}]
  const native={...context,text:paragraphs.map(p=>p.text).join("\r"),paragraphs,selection:undefined}
  const request=(extra:object={})=>post("send",{conversationId:id,scope:"document",text:"Join the first two paragraphs",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native,...extra})
  const output=(action:unknown="merge",after="First sentence. Second sentence.")=>{modelRawOutput=JSON.stringify({answer:"Prepared.",citations:[],proposal:{before:"First sentence.\rSecond sentence.",after,paragraphAction:action,explanation:"Join",openItems:[]}})}
  let turn:WordChatTurn
  try {
    output();turn=(await (await request()).json() as {turn:WordChatTurn}).turn
    expect(turn.proposal?.paragraphAction).toBe("merge")
    const reviewer=[...sessions.values()].find(s=>s.parentID===turn.id&&s.metadata.type==="word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain("join the two complete adjacent paragraphs")
    const original=turn.proposal!
    output("delete","");expect((await (await request({proposal:original})).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
    output();const refined=(await (await request({proposal:original})).json() as {turn:WordChatTurn}).turn
    expect(refined.proposal?.target).toEqual(original.target);expect(refined.proposal?.paragraphAction).toBe("merge")
    output("merge","First sentence.Second sentence.");expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
  } finally {modelRawOutput=undefined}
  const current=readCurrentVersion(directory,"Draft.docx")!,db=openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",[auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner=claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"merge-prepare-session"})
  const seq=(db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"merge-prepare",undefined,{conversationId:id,turnId:turn!.id})
  let invalidate=false
  const preparing=new Hono()
  installWordChat(preparing,()=>auth,{fetch:transport,engineURL:"http://synthetic.invalid",authorAmendment:async (_xml,proposal,stored)=>{
    expect(proposal).toEqual(turn!.proposal!);expect(stored?.paragraphs).toEqual(paragraphs)
    if(invalidate)db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other",owner.document])
    return {kind:"paragraph_merge",proposalId:proposal.id,target:proposal.target}
  }})
  const headers={"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare=()=>preparing.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn!.id,ooxml:"current",proposal:{...turn!.proposal,after:"forged"}})})
  try {const response=await prepare();expect(response.status).toBe(200);expect((await response.json() as {kind:string}).kind).toBe("paragraph_merge");invalidate=true;expect((await prepare()).status).toBe(409)}
  finally {db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn!.id,owner.document]);settleWordNativeOperation(db,owner,seq+1,"merge-prepare","not-dispatched");clearWordEditLease(db,owner.document);db.close()}
})

test("paragraph split review and preparation keep the exact boundary and lease", async () => {
  const id=await newConversation(), paragraphs=[{id:"split-a",text:"First sentence. Second sentence."},{id:"split-b",text:"Other."}]
  const native={...context,text:paragraphs.map(p=>p.text).join("\r"),paragraphs,selection:undefined}
  const request=(extra:object={})=>post("send",{conversationId:id,scope:"document",text:"Split the first paragraph",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:native,...extra})
  const output=(action:unknown="split",after="First sentence.\rSecond sentence.")=>{modelRawOutput=JSON.stringify({answer:"Prepared.",citations:[],proposal:{before:paragraphs[0]!.text,after,paragraphAction:action,explanation:"Split",openItems:[]}})}
  let turn:WordChatTurn
  try {
    output();turn=(await (await request()).json() as {turn:WordChatTurn}).turn
    expect(turn.proposal?.paragraphAction).toBe("split")
    const reviewer=[...sessions.values()].find(s=>s.parentID===turn.id&&s.metadata.type==="word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain("replace exactly one existing space")
    const original=turn.proposal!
    output("delete","");expect((await (await request({proposal:original})).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
    output();const refined=(await (await request({proposal:original})).json() as {turn:WordChatTurn}).turn
    expect(refined.proposal?.target).toEqual(original.target);expect(refined.proposal?.paragraphAction).toBe("split")
    output("split","First sentence.\rChanged sentence.");expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
  } finally {modelRawOutput=undefined}
  const current=readCurrentVersion(directory,"Draft.docx")!,db=openDb(directory)
  db.run("INSERT OR IGNORE INTO word_grants(token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",[auth.grant.token_hash,auth.grant.document_id,auth.grant.origin,auth.grant.issued_version,Date.now()])
  const owner=claimWordEditLease(db,{document:auth.grant.document_id,tokenHash:auth.grant.token_hash,session:"split-prepare-session"})
  const seq=(db.query("SELECT native_sequence FROM word_edit_leases WHERE document_id = ?").get(owner.document) as {native_sequence:number}).native_sequence
  beginWordNativeOperation(db,owner,seq,"split-prepare",undefined,{conversationId:id,turnId:turn!.id})
  let invalidate=false
  const preparing=new Hono()
  installWordChat(preparing,()=>auth,{fetch:transport,engineURL:"http://synthetic.invalid",authorAmendment:async (_xml,proposal,stored)=>{
    expect(proposal).toEqual(turn!.proposal!);expect(stored?.paragraphs).toEqual(paragraphs)
    if(invalidate)db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",["other",owner.document])
    return {kind:"paragraph_split",proposalId:proposal.id,target:proposal.target}
  }})
  const headers={"content-type":"application/json","x-dochaus-word-session":owner.session,"x-dochaus-word-lease":owner.generation,"x-dochaus-word-base-version":current.current.id}
  const prepare=()=>preparing.request("/word/chat/prepare",{method:"POST",headers,body:JSON.stringify({conversationId:id,turnId:turn!.id,ooxml:"current",proposal:{...turn!.proposal,after:"forged"}})})
  try {const response=await prepare();expect(response.status).toBe(200);expect((await response.json() as {kind:string}).kind).toBe("paragraph_split");invalidate=true;expect((await prepare()).status).toBe(409)}
  finally {db.run("UPDATE word_edit_leases SET native_turn = ? WHERE document_id = ?",[turn!.id,owner.document]);settleWordNativeOperation(db,owner,seq+1,"split-prepare","not-dispatched");clearWordEditLease(db,owner.document);db.close()}
})

test("before-final native list proposal stays bound through review and rejects forged list and direction changes",async()=>{
  const id=await newConversation(),text=context.text
  const list={id:3,format:"native" as const,members:[{...context.paragraphs[0]!,level:2,label:"i)",siblingIndex:0}]}
  const listContext={...context,list}
  const request=(extra:object={})=>post("send",{conversationId:id,scope:"document",text:"Insert before",requestId:crypto.randomUUID(),attachments:[],mode:"redline",context:listContext,...extra})
  const output=(extra:object={})=>{modelStructuredRawOutput=JSON.stringify({answerParts:[{type:"text",text:"Prepared."}],proposal:{before:text,after:"Another item\r"+text,paragraphAction:"insert_list_item_before",explanation:"Add",openItems:[],...extra}})}
  try {
    output();const turn=(await (await request()).json() as {turn:WordChatTurn}).turn
    expect(turn.proposal?.list).toEqual(list);expect(turn.proposal?.paragraphAction).toBe("insert_list_item_before")
    expect(sessionContext(turn.id)).toContain("insert_list_item_before")
    const reviewer=[...sessions.values()].find(s=>s.parentID===turn.id && s.metadata.type==="word-chat-review")!
    expect(sessionContext(reviewer.id)).toContain("placement")
    expect((await request({proposal:turn.proposal,context:{...listContext,list:{...list,id:4}}})).status).toBe(409)
    output({list});expect((await (await request()).json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
    output({paragraphAction:"append_list_item",after:text+"\rAnother item"});const changed=await request({proposal:turn.proposal});expect(changed.status).toBe(200);expect((await changed.json() as {turn:WordChatTurn}).turn.proposal).toBeUndefined()
  } finally {modelStructuredRawOutput=undefined}
})
