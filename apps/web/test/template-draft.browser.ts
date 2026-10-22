// Run from apps/web: bun test/template-draft.browser.ts
// Real matter/library/chat/review components and Docxodus WASM; HTTP and native
// engine events are mocked. Every document and conversation below is synthetic.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { FormInfo, SessionMessageInfo } from "@opencode/client"
import type { TemplateDraft } from "../src/api/ingest/template-drafts"

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test") as typeof import("@playwright/test")
const { zipSync, strToU8 } = await import(Bun.resolveSync("fflate", path.resolve(import.meta.dir, "../../../dochaus"))) as {
  zipSync: (files: Record<string, Uint8Array>) => Uint8Array; strToU8: (text: string) => Uint8Array
}
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-template-draft-"))
const shots = process.env.TEMPLATE_DRAFT_SCREENSHOTS
if (shots) await mkdir(shots, { recursive: true })
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const sourceText = "Acme Ltd shall pay Beta LLP £4,200 on 1 January 2027."
const notice = "Either party may end this agreement on thirty days notice."
const replacement = "[insert customer] shall pay [insert supplier] [insert fee] on [insert payment date]."
const revisedNotice = "Either party may end this agreement on [insert notice period] notice."
const model = { providerID: "test", id: "model" }
const unknown: string[] = []
const sourceWrites: string[] = []
const prompts: { id: string; directory: string; text: string; agent: string }[] = []
const formReplies: unknown[] = []
const publications: unknown[] = []
let snapshotAttempts = 0
let publishAttempts = 0
let metadataFailure = false
let libraryRefreshFailure = false
let eventID = 0

function docx(first: string, second: string) {
  const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const files = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${CT}.settings+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/settings" Target="settings.xml"/></Relationships>`,
    "word/styles.xml": `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    "word/settings.xml": `<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/></w:settings>`,
    "word/document.xml": `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${first}</w:t></w:r></w:p><w:p><w:r><w:t>${second}</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Schedule of services</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  }
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, xml]) => [name, strToU8(xml)])))
}
const original = docx(sourceText, notice)
const originalCopy = new Uint8Array(original)
const drafts = new Map<string, TemplateDraft>()
type Session = { dir: string; transcript: SessionMessageInfo[]; running: boolean; forms: FormInfo[]; release?: () => void; agent: string }
const sessions = new Map<string, Session>([["ses_matter", {
  dir: "/matter", running: false, forms: [], agent: "qa", transcript: [
    { type: "user", id: "msg_other", text: "Unrelated matter instructions must stay here.", time: { created: 1 } },
    { type: "idle", id: "msg_old_idle", outcome: "succeeded", time: { created: 2 } },
  ],
}]])
const streams = new Map<string, Set<ReadableStreamDefaultController<Uint8Array>>>()
function emit(id: string, type: string, data: Record<string, unknown> = {}) {
  const dir = sessions.get(id)!.dir
  const event = { id: `evt_${++eventID}`, created: Date.now(), type, location: { directory: dir }, durable: { aggregateID: id, seq: eventID, version: 1 }, data: { sessionID: id, ...data } }
  for (const stream of streams.get(dir) ?? []) {
    try { stream.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`)) } catch { /* A reload retired this stream. */ }
  }
}
function prepare(id: string) {
  const session = sessions.get(id)!
  const draft = [...drafts.values()].find((item) => item.dir === session.dir)!
  draft.revision += 1
  draft.name = "services-template.docx"
  draft.description = "Services agreement for repeat engagements"
  draft.changes = [{ before_quote: sourceText, replacement, paragraph_before: sourceText, paragraph_after: replacement }]
  if (draft.revision > 1) draft.changes.push({ before_quote: "thirty days", replacement: "[insert notice period]", paragraph_before: notice, paragraph_after: revisedNotice })
  draft.notes = ["Review the payment terms for your intended use."]
  const number = session.transcript.length
  session.transcript.push({ type: "assistant", id: `msg_a${number}`, agent: "template-builder", model, finish: "stop", time: { created: Date.now(), completed: Date.now() }, content: [
    { type: "tool", id: `call_prepare${number}`, name: "prepare-template", time: { created: Date.now(), completed: Date.now() }, state: { status: "completed", input: { revision: draft.revision - 1 }, content: [{ type: "text", text: "Prepared for review." }], metadata: {} } },
    { type: "text", text: "Your proposed template is ready for review." },
  ] })
  session.transcript.push({ type: "idle", id: `msg_i${number}`, outcome: "succeeded", time: { created: Date.now() } })
  session.running = false
  emit(id, "session.step.ended", { assistantMessageID: `msg_a${number}`, finish: "stop" })
  emit(id, "session.execution.succeeded")
  session.release?.()
}

const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement as h } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { BrowserRouter, Route, Routes } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import MatterDetail from ${source("pages/MatterDetail.tsx")}
import Templates from ${source("pages/Templates.tsx")}
import { ToastProvider } from ${source("components/Toast.tsx")}
createRoot(document.getElementById("root")).render(h(ToastProvider, null, h(BrowserRouter, null,
  h("div", {className: "shell"}, h("div", {id: "matter-drawer-slot"}), h("main", {className: "content"}, h("div", {className: "container"},
    h(Routes, null, h(Route, {path: "/matter/:id", element: h(MatterDetail, {onSessionsChanged: () => {}})}), h(Route, {path: "/templates", element: h(Templates)}))))))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.path.endsWith(".js"))!.text()
const css = (await Promise.all([
  "styles.css", "narration.css", "narration-live.css", "styles/primitives.css", "styles/upload.css", "styles/shell.css", "styles/chat-log.css", "styles/drafting.css", "styles/drafting-items.css", "styles/run-summary.css", "styles/motion.css", "styles/design-system.css", "styles/design-shell.css", "styles/design-documents.css", "styles/design-conversation.css", "styles/design-library.css", "styles/design-review.css", "styles/design-settings.css", "styles/design-matters.css", "styles/launch-notice.css", "styles/contrast.css",
].map((file) => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const matter = { id: "matter", title: "Services matter", dir: "/matter", created_at: 1, documents: [{ id: 1, name: "Services.docx", doc_path: "/matter/Services.docx", created_at: 1, role: "source" }] }
const json = (body: unknown) => Response.json(body)
const nativeSession = (id: string, session: Session) => ({ id, projectID: "project", title: "Template conversation", agent: session.agent, model, location: { directory: session.dir }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 1, updated: 1 } })
const server = Bun.serve({
  port: 0, hostname: "127.0.0.1", idleTimeout: 0,
  async fetch(request) {
    const url = new URL(request.url)
    const route = `${request.method} ${url.pathname}`
    const dir = request.headers.get("x-opencode-directory") ?? ""
    if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (url.pathname.startsWith("/wasm/")) {
      const file = path.join(wasmDir, decodeURIComponent(url.pathname.slice(6)))
      return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response(null, { status: 404 })
    }
    if (request.method === "GET" && ["/matter/matter", "/templates"].includes(url.pathname)) return new Response(`<html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (route === "GET /ingest/matters/matter") return json(matter)
    if (route === "GET /ingest/matters") return json([matter])
    if (route === "GET /ingest/jurisdictions" || route === "GET /ingest/playbooks") return json([])
    if (route === "GET /ingest/workflows") return json({ dir: "/workflows", workflows: [] })
    if (route === "GET /ingest/templates") {
      if (libraryRefreshFailure) { libraryRefreshFailure = false; return Response.json({ error: "Library temporarily unavailable" }, { status: 503 }) }
      return json({ dir: "/templates", templates: [...drafts.values()].filter((draft) => draft.status === "published").map((draft) => ({ name: draft.publishedName, description: draft.description, placeholders: [{ text: "[insert customer]", kind: "fill" }] })) })
    }
    if (route === "GET /ingest/template-drafts") return json({ drafts: [...drafts.values()].reverse() })
    if (route === "GET /ingest/matters/matter/telemetry") return json({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(null)
    if (url.pathname.startsWith("/ingest/matters/") && url.pathname.includes("/answer-drafts/")) {
      assert.ok(url.pathname.includes("/matter/"), "library answers never use a fake matter endpoint")
      return json({ answers: {}, revision: 0 })
    }
    if (route === "GET /ingest/matters/matter/redlines") return json([])
    if (route === "GET /ingest/matters/matter/documents/amendments") return json(null)
    if (/^GET \/ingest\/matters\/matter\/documents\/(content|redlined)$/.test(route)) return new Response(original)
    if (route === "POST /ingest/matters/matter/template-drafts") {
      assert.deepEqual(await request.json(), { name: "Services.docx" })
      snapshotAttempts += 1
      if (snapshotAttempts === 1) return Response.json({ error: "Document is busy. Try again." }, { status: 409 })
      const id = `draft_${drafts.size + 1}`
      const draft: TemplateDraft = { id, dir: `/template-drafts/${id}`, source: { name: "Services.docx", matterId: "matter" }, createdAt: Date.now(), revision: 0, name: "services-template.docx", description: "", status: "draft", changes: [], notes: [] }
      drafts.set(id, draft)
      return json(draft)
    }
    const draftRoute = url.pathname.match(/^\/ingest\/template-drafts\/(draft_\d+)(?:\/(content|publish))?$/)
    if (draftRoute) {
      const draft = drafts.get(draftRoute[1])!
      if (!draftRoute[2]) return metadataFailure ? Response.json({ error: "Template draft temporarily unavailable" }, { status: 503 }) : json(draft)
      if (draftRoute[2] === "content") return new Response(url.searchParams.get("view") === "source" ? original : docx(replacement, draft.revision > 1 ? revisedNotice : notice))
      const input = await request.json() as { revision: number; name: string; description: string }
      publishAttempts += 1
      if (publishAttempts === 1) return Response.json({ error: "A template with this name already exists." }, { status: 409 })
      assert.equal(input.revision, draft.revision)
      publications.push(input)
      Object.assign(draft, { name: input.name, publishedName: input.name, description: input.description, status: "published" })
      libraryRefreshFailure = true
      return json(draft)
    }
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "GET /api/agent") return json({ data: [{ id: "qa", mode: "primary" }, { id: "draft", mode: "primary" }, { id: "template-builder", mode: "primary" }] })
    if (route === "GET /api/event") {
      let controller: ReadableStreamDefaultController<Uint8Array>
      return new Response(new ReadableStream<Uint8Array>({
        start(next) { controller = next; if (!streams.has(dir)) streams.set(dir, new Set()); streams.get(dir)!.add(next) },
        cancel() { streams.get(dir)?.delete(controller) },
      }), { headers: { "content-type": "text/event-stream" } })
    }
    if (route === "GET /api/session/active") return json({ data: Object.fromEntries([...sessions].filter(([, session]) => session.running).map(([id]) => [id, { type: "running" }])) })
    if (route === "GET /api/session") return json({ data: [...sessions].filter(([, session]) => session.dir === dir).map(([id, session]) => nativeSession(id, session)), cursor: {} })
    if (route === "GET /api/permission/request") return json({ location: { directory: dir }, data: [] })
    if (route === "POST /api/session") {
      const body = await request.json() as { location: { directory: string } }
      assert.equal(body.location.directory, dir)
      assert.ok(dir.startsWith("/template-drafts/"), "source creation never routes through matter Draft")
      const id = `ses_template${sessions.size}`
      const session: Session = { dir, transcript: [], running: false, forms: [], agent: "template-builder" }
      sessions.set(id, session)
      return json({ data: nativeSession(id, session) })
    }
    const sessionRoute = url.pathname.match(/^\/api(?:\/experimental)?\/session\/(ses_\w+)(?:\/(message|agent|model|prompt|wait|form)(?:\/([^/]+)(\/reply)?)?)?$/)
    if (sessionRoute) {
      const id = sessionRoute[1]
      const session = sessions.get(id)!
      if (!sessionRoute[2]) return json({ data: nativeSession(id, session) })
      assert.equal(dir, session.dir, "all session activity uses its exact source directory")
      if (sessionRoute[2] === "message") {
        const ordered = url.searchParams.get("order") === "desc" ? [...session.transcript].reverse() : session.transcript
        return json({ data: ordered.slice(0, Number(url.searchParams.get("limit") ?? ordered.length)), cursor: {} })
      }
      if (sessionRoute[2] === "form") {
        if (request.method === "GET") return json({ data: session.forms })
        formReplies.push(await request.json())
        session.forms = []
        prepare(id)
        return new Response(null, { status: 204 })
      }
      if (sessionRoute[2] === "agent") {
        const body = await request.json() as { agent: string }
        session.agent = body.agent
      }
      if (sessionRoute[2] === "prompt") {
        const { text } = await request.json() as { text: string }
        prompts.push({ id, directory: dir, text, agent: session.agent })
        const messageID = `msg_u${session.transcript.length}`
        session.transcript.push({ type: "user", id: messageID, text, time: { created: Date.now() } })
        session.running = true
        if (session.transcript.length === 1) {
          session.forms = [{ id: "form-intake", sessionID: id, title: "Template purpose", metadata: { kind: "question" }, fields: [{ key: "purpose", type: "string", title: "Intended use and position", description: "How will this template be used?", custom: true, options: [{ value: "neutral", label: "Neutral terms (Recommended)", description: "Reusable for either side of a services engagement." }, { value: "supplier", label: "Supplier terms", description: "Protect the supplier’s position." }] }] }]
          setTimeout(() => { emit(id, "session.step.started", { assistantMessageID: "msg_intake", agent: "template-builder", model, started: Date.now() }); emit(id, "form.created", { formID: "form-intake" }) }, 100)
        } else setTimeout(() => prepare(id), 100)
        return json({ data: { id: messageID, sessionID: id, type: "user", payload: { text }, delivery: "queue", time: { created: Date.now() } } })
      }
      if (sessionRoute[2] === "wait" && session.running) await new Promise<void>((resolve) => { session.release = resolve })
      return new Response(null, { status: 204 })
    }
    if (request.method !== "GET" && url.pathname.startsWith("/ingest/matters/")) sourceWrites.push(route)
    if (url.pathname !== "/favicon.ico") unknown.push(route)
    return new Response(null, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
async function until(condition: () => boolean, label: string) {
  const deadline = Date.now() + 10_000
  while (!condition() && Date.now() < deadline) await Bun.sleep(10)
  assert.ok(condition(), label)
}
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 950 }, reducedMotion: "reduce" })
  page.setDefaultTimeout(15_000)
  await page.addInitScript(() => localStorage.setItem("dochaus.prefs", JSON.stringify({ autoApproveDrafting: true })))
  await page.goto(`${server.url}matter/matter?view=documents&session=ses_matter`)
  await page.getByRole("button", { name: "Services.docx", exact: true }).click()
  const sourceDialog = page.getByRole("dialog", { name: "Services.docx", exact: true })
  await sourceDialog.locator(".docx-render:not(.docx-skeleton)").getByText(sourceText, { exact: true }).waitFor()
  await sourceDialog.getByLabel("More actions", { exact: true }).click()
  await sourceDialog.getByRole("button", { name: "Create reusable template" }).click()
  await page.getByText("Document is busy. Try again.", { exact: true }).waitFor()
  assert.ok(await sourceDialog.isVisible(), "snapshot errors retain the source viewer")
  await sourceDialog.getByRole("button", { name: "Create reusable template" }).click()
  await page.waitForURL("**/templates?draft=draft_1")
  await page.getByRole("button", { name: "Dismiss notification", exact: true }).click()
  await page.getByText("Document is busy. Try again.", { exact: true }).waitFor({ state: "detached" })
  await page.getByRole("region", { name: "Attached source document" }).getByText("Services.docx", { exact: true }).waitFor()
  assert.equal(prompts.length, 0, "handoff never auto-sends the instruction")
  assert.equal(await page.getByRole("button", { name: "Review proposed template" }).isDisabled(), true)
  const initial = await page.getByRole("textbox").inputValue()
  assert.ok(initial.includes("Services.docx") && !initial.includes("Unrelated matter"))
  if (shots) await page.screenshot({ path: path.join(shots, "1-source-attached.png") })
  await page.getByRole("button", { name: "View source", exact: true }).click()
  await page.getByRole("dialog", { name: "Services.docx" }).locator(".docx-render").getByText(sourceText, { exact: true }).waitFor()
  await page.getByRole("dialog", { name: "Services.docx" }).getByRole("button", { name: /Close/ }).click()
  console.log("PASS source handoff snapshots into a fresh library conversation, attaches visibly, and handles errors")

  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await until(() => prompts.length === 1, "one new source conversation is created")
  const firstSession = prompts[0].id
  await page.waitForURL(`**draft=draft_1&session=${firstSession}`)
  await page.getByRole("button", { name: "Answer", exact: true }).click()
  const question = page.getByRole("group", { name: "Intended use and position", exact: true })
  const recommended = question.getByRole("radio", { name: "Neutral terms (Recommended)", exact: true })
  const custom = question.getByRole("radio", { name: "Type your own", exact: true })
  const customAnswer = question.getByRole("textbox", { name: "Intended use and position — your own answer", exact: true })
  assert.equal(await recommended.isChecked(), false, "recommendation never auto-answers the question")
  assert.equal(await customAnswer.count(), 0, "custom textarea stays collapsed until selected")
  await recommended.check()
  if (shots) await page.screenshot({ path: path.join(shots, "5-question-options.png") })
  await custom.check()
  const answer = "Neutral supplier/customer terms for repeat services engagements."
  await customAnswer.fill(answer)
  await recommended.check()
  assert.equal(await customAnswer.count(), 0, "choosing a preset collapses the custom textarea")
  await custom.check()
  assert.equal(await customAnswer.inputValue(), answer, "switching choices retains the unsent custom text")
  await page.setViewportSize({ width: 390, height: 844 })
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true, "question cards fit mobile width")
  if (shots) await page.screenshot({ path: path.join(shots, "6-mobile-question.png") })
  await page.setViewportSize({ width: 1280, height: 950 })
  await page.reload()
  await page.getByRole("button", { name: "Answer", exact: true }).click()
  assert.equal(await customAnswer.inputValue(), answer, "unsent library answers survive reload")
  await page.getByRole("button", { name: "Send answers", exact: true }).click()
  await page.getByRole("button", { name: "Review proposed template" }).waitFor({ state: "visible" })
  await until(() => drafts.get("draft_1")!.revision === 1, "intake prepares the first candidate")
  await page.getByRole("button", { name: "Review proposed template" }).click()
  const proposal = page.getByRole("dialog", { name: "Proposed reusable template", exact: true })
  await proposal.locator(".docx-render").getByText(replacement, { exact: true }).waitFor()
  assert.equal(await proposal.locator(".docx-render table").count(), 1, "Word table renders in the proposal")
  await proposal.locator(".template-draft-changes .redline-old").getByText(sourceText, { exact: true }).waitFor()
  assert.equal(publications.length, 0, "automatic drafting approval cannot publish a candidate")
  await proposal.getByRole("button", { name: "Keep editing", exact: true }).click()
  assert.equal(publications.length, 0, "keeping edits never publishes")
  console.log("PASS matter-less intake is visible and persisted; concrete Word review precedes explicit save")

  await page.getByRole("textbox").fill("Also generalise the notice period.")
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await until(() => drafts.get("draft_1")!.revision === 2, "a follow-up prepares a revised proposal")
  await page.getByRole("button", { name: "Review proposed template" }).click()
  await proposal.locator(".docx-render").getByText(revisedNotice, { exact: true }).waitFor()
  await page.setViewportSize({ width: 1280, height: 700 })
  for (const summary of await proposal.locator(".template-draft-changes summary").all()) await summary.click()
  const saveBar = proposal.getByRole("region", { name: "Template details and saving" })
  const beforeScroll = await saveBar.boundingBox()
  assert.ok(beforeScroll)
  const scrolled = await proposal.locator(".changes-panel").evaluate((panel) => {
    panel.scrollTop = panel.scrollHeight
    return panel.scrollTop
  })
  assert.ok(scrolled > 0, "long review context actually scrolls independently")
  const afterScroll = await saveBar.boundingBox()
  assert.ok(afterScroll && Math.abs(afterScroll.y - beforeScroll.y) < 1, "template fields and saving stay pinned while reviewing changes")
  const saveBounds = await saveBar.getByRole("button", { name: "Save to library", exact: true }).boundingBox()
  assert.ok(saveBounds && saveBounds.y + saveBounds.height <= 700, "save remains inside the desktop viewport")
  if (shots) await page.screenshot({ path: path.join(shots, "2-proposal-review.png") })
  await page.setViewportSize({ width: 390, height: 844 })
  await proposal.locator(".docx-render").scrollIntoViewIfNeeded()
  const mobileFooter = await saveBar.boundingBox()
  await proposal.locator(".viewer-body").evaluate((body) => { body.scrollTop = body.scrollHeight })
  const mobileAfterScroll = await saveBar.boundingBox()
  assert.ok(mobileFooter && mobileAfterScroll && Math.abs(mobileFooter.y - mobileAfterScroll.y) < 1, "mobile document and changes scroll above the pinned footer")
  for (const control of [saveBar.getByRole("textbox", { name: "Template name", exact: true }), saveBar.getByRole("textbox", { name: "Description", exact: true }), saveBar.getByRole("button", { name: "Save to library", exact: true })]) {
    const bounds = await control.boundingBox()
    assert.ok(bounds && bounds.y >= 0 && bounds.y + bounds.height <= 844, "mobile details and save are visible without scrolling the footer")
  }
  if (shots) await page.screenshot({ path: path.join(shots, "4-mobile-document.png") })
  await proposal.getByRole("button", { name: "Save to library", exact: true }).scrollIntoViewIfNeeded()
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true, "mobile review does not overflow horizontally")
  if (shots) await page.screenshot({ path: path.join(shots, "3-mobile-review.png") })
  await page.setViewportSize({ width: 1280, height: 950 })
  await proposal.getByRole("button", { name: "Save to library", exact: true }).click()
  await proposal.getByText("A template with this name already exists.", { exact: true }).waitFor()
  assert.equal(publications.length, 0)
  await proposal.getByRole("textbox", { name: "Template name", exact: true }).fill("Standard services.docx")
  await proposal.getByRole("button", { name: "Save to library", exact: true }).click()
  await proposal.getByText("Saved to your library.", { exact: true }).waitFor()
  await page.getByText("Saved, but the library could not refresh. Reload to see the saved template.", { exact: true }).waitFor()
  assert.equal(publications.length, 1)
  assert.deepEqual(publications[0], { revision: 2, name: "Standard services.docx", description: "Services agreement for repeat engagements" })
  assert.deepEqual(formReplies, [{ answer: { purpose: answer } }])
  assert.deepEqual(original, originalCopy)
  assert.deepEqual(sourceWrites, [])
  await proposal.getByRole("button", { name: /Close/ }).click()
  await page.reload()
  await page.getByRole("region", { name: "Template review" }).getByText("Standard services.docx", { exact: true }).waitFor()
  assert.equal(snapshotAttempts, 2, "reopening a draft never re-snapshots the matter document")
  await page.getByRole("link", { name: "Back to saved templates", exact: true }).click()
  await page.getByRole("region", { name: "Template conversations" }).getByRole("link", { name: /Services.docx/ }).click()
  await page.waitForURL(`**draft=draft_1&session=${firstSession}`)
  await page.getByRole("region", { name: "Template review" }).getByText("Standard services.docx", { exact: true }).waitFor()
  console.log("PASS revised proposal, name conflict, explicit save, source preservation and published reload")

  await page.getByRole("button", { name: "Use", exact: true }).click()
  await page.getByRole("button", { name: /Services matter/ }).click()
  await page.waitForURL("**/matter/matter?view=chat&agent=draft&prompt=**")
  await page.waitForFunction(() => document.querySelector("textarea")?.value.includes("Draft a new document from the template"))
  assert.ok((await page.getByRole("textbox").inputValue()).includes('Draft a new document from the template "Standard services.docx"'))
  assert.equal(prompts.length, 2, "library Use still prefills the selected matter without starting a run")
  await page.goto(`${server.url}matter/matter?view=documents`)
  await page.getByRole("button", { name: "Services.docx", exact: true }).click()
  await sourceDialog.getByLabel("More actions", { exact: true }).click()
  await sourceDialog.getByRole("button", { name: "Create reusable template" }).click()
  await page.waitForURL("**/templates?draft=draft_2")
  await page.getByRole("textbox").waitFor()
  assert.equal(new URL(page.url()).searchParams.has("session"), false)
  await page.goto(`${server.url}templates?draft=draft_2&session=${firstSession}`)
  await page.getByText("This conversation belongs to another document. Start a fresh template conversation.").waitFor()
  assert.equal(await page.locator(".chat-card").count(), 0, "another document's session is never loaded")
  await page.getByRole("button", { name: "Start a fresh conversation" }).click()
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await until(() => prompts.length === 3, "a second source gets a second native conversation")
  assert.notEqual(prompts[2].id, firstSession)
  assert.notEqual(prompts[2].directory, prompts[0].directory)
  assert.ok(prompts.every((prompt) => prompt.agent === "template-builder" && prompt.directory !== "/matter" && !prompt.text.includes("Unrelated matter instructions")))
  await page.goto(`${server.url}templates?draft=draft_2&session=ses_matter`)
  await page.getByText("This conversation belongs to another document. Start a fresh template conversation.").waitFor()
  await page.getByRole("button", { name: "Start a fresh conversation" }).click()
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await until(() => prompts.length === 4, "explicit fresh creates a native session even when this draft has an older chat")
  assert.notEqual(prompts[3].id, prompts[2].id)
  assert.equal(prompts[3].directory, prompts[2].directory)
  console.log("PASS library Use is preserved and each source handoff has an independent session")

  metadataFailure = true
  await page.goto(`${server.url}templates?draft=draft_2`)
  await page.getByText(/Template draft temporarily unavailable/).waitFor()
  metadataFailure = false
  await page.getByRole("button", { name: "Retry", exact: true }).click()
  await page.getByRole("region", { name: "Attached source document" }).waitFor()
  assert.deepEqual(unknown, [])
  console.log("PASS handoff reload errors are actionable and retry restores the captured source")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
