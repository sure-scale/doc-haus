// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/chat-stop.browser.ts
// Uses the real matter page, chat panel and status bar; only the engine's and ingest's HTTP and event stream are mocked.
// Set CHAT_STOP_SCREENSHOTS=<directory> to save the stopped and follow-up answers.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { SessionMessageInfo } from "@opencode/client"

type Assistant = Extract<SessionMessageInfo, { type: "assistant" }>

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chat-stop-"))
const shots = process.env.CHAT_STOP_SCREENSHOTS
const sessionID = "ses_chat"
const model = { providerID: "test", id: "model" }
const transcript: SessionMessageInfo[] = [
  { type: "user", id: "msg_u1", text: "Who are the parties?", time: { created: 1 } },
  { type: "assistant", id: "msg_a1", agent: "qa", model, finish: "stop", time: { created: 2, completed: 3 }, content: [{ type: "text", text: "Acme Ltd and Beta LLP." }] },
  { type: "idle", id: "msg_i1", outcome: "succeeded", time: { created: 4 } },
]
let running = false
let stream: ReadableStreamDefaultController<Uint8Array> | undefined
let release = () => {}
let seq = 0
let routes = 0
const prompts: string[] = []
const interrupts: string[] = []
// Full transcript reads: the panel's load, resync and finalize (the prompt's own
// completion check reads only the newest two messages).
let historyReads = 0
const unknown: string[] = []

function send(type: string, data: Record<string, unknown>) {
  seq += 1
  const event = { id: `evt_${seq}`, created: Date.now(), type, location: { directory: "/matter" }, durable: { aggregateID: sessionID, seq, version: 1 }, data: { sessionID, ...data } }
  stream!.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`))
}

// One provider step as the engine stores it: the assistant message streams,
// then completes when the step ends, whether or not the turn goes on.
function store(id: string, content: Assistant["content"], end?: { finish: Assistant["finish"]; error?: Assistant["error"] }) {
  const message: Assistant = { type: "assistant", id, agent: "qa", model, time: { created: Date.now(), ...(end ? { completed: Date.now() } : {}) }, content, ...end }
  const at = transcript.findIndex(entry => entry.id === id)
  if (at === -1) {
    transcript.push(message)
    send("session.step.started", { assistantMessageID: id, agent: "qa", model, started: Date.now() })
  } else transcript[at] = message
  send(end ? "session.step.ended" : "session.tool.called", { assistantMessageID: id, finish: end?.finish })
}
const search = (query: string, result: string): Assistant["content"][number] => ({
  type: "tool", id: `call_${query.replace(/\W+/g, "_")}`, name: "search-document", time: { created: Date.now(), completed: Date.now() },
  state: { status: "completed", input: { query }, content: [{ type: "text", text: result }], metadata: { citations: [{ document: "Letter.docx", section: "1", excerpt: result }] } },
})
function end(outcome: "succeeded" | "interrupted", event: [string, Record<string, unknown>]) {
  transcript.push({ type: "idle", id: `msg_i${transcript.length}`, outcome, time: { created: Date.now() } })
  running = false
  send(...event)
  release()
}
// After acknowledging an interrupt the engine first fails the running step as
// interrupted (published as a step failure), then records the turn as interrupted.
function failRunningStep() {
  const step = transcript.findLast((entry): entry is Assistant => entry.type === "assistant" && !entry.time.completed)
  assert.ok(step, "a step is running")
  const error = { type: "aborted", message: "Step interrupted" }
  store(step.id, step.content, { finish: "error", error })
  send("session.step.failed", { assistantMessageID: step.id, error })
}

const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter, Route, Routes } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import MatterDetail from ${source("pages/MatterDetail.tsx")}
import { ToastProvider } from ${source("components/Toast.tsx")}
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(MemoryRouter, { initialEntries: ["/matter/matter?view=chat&session=${sessionID}"] },
    createElement(Routes, null, createElement(Route, { path: "/matter/:id", element: createElement(MatterDetail, { onSessionsChanged: () => {} }) })))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.path.endsWith(".js"))!.text()
const css = (await Promise.all(["styles.css", "styles/design-system.css", "styles/design-shell.css", "styles/design-conversation.css"]
  .map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const json = (body: unknown) => Response.json(body)
const session = (id: string) => ({ id, projectID: "project", agent: "qa", model, location: { directory: "/matter" }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 1, updated: 1 } })
const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  idleTimeout: 0,
  async fetch(request) {
    const url = new URL(request.url)
    const route = `${request.method} ${url.pathname}`
    if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (route === "GET /") return new Response(`<html><meta charset="utf-8"><style>body{font-family:sans-serif;margin:0}${css}</style><div id="root" style="display:flex;flex-direction:column;height:100vh"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (route === "GET /ingest/matters/matter")
      return json({ id: "matter", title: "Unpaid invoices", dir: "/matter", created_at: 1, documents: [{ id: 1, name: "Letter.docx", doc_path: "/matter/Letter.docx", created_at: 1, role: "source" }] })
    if (route === "GET /ingest/jurisdictions" || route === "GET /ingest/playbooks") return json([])
    if (route === "GET /ingest/workflows") return json({ dir: "/workflows", workflows: [] })
    if (route === "GET /api/agent") return json({ data: [{ id: "qa", description: "Answers questions.", mode: "primary" }] })
    if (route === "GET /ingest/matters/matter/telemetry")
      return json({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(null)
    if (route === `GET /ingest/matters/matter/answer-drafts/session.${sessionID}`) return json({ answers: {}, revision: 0 })
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "GET /api/event")
      return new Response(new ReadableStream<Uint8Array>({ start(controller) { stream = controller } }), { headers: { "content-type": "text/event-stream" } })
    if (route === "GET /api/session/active") return json({ data: running ? { [sessionID]: { type: "running" } } : {} })
    if (route === "GET /api/permission/request") return json({ location: { directory: "/matter" }, data: [] })
    if (route === `GET /api/session/${sessionID}/form`) return json({ data: [] })
    if (route === `GET /api/session/${sessionID}`) return json({ data: session(sessionID) })
    // Auto routes each message through a throwaway router session.
    if (route === "POST /api/session") {
      routes += 1
      return json({ data: session(`ses_router${routes}`) })
    }
    const router = url.pathname.match(/^\/api(?:\/experimental)?\/session\/(ses_router\d+)(\/\w+)?$/)
    if (router) {
      if (router[2] === "/message")
        return json({ data: [{ type: "assistant", id: `msg_route_${router[1]}`, agent: "router", model, finish: "stop", time: { created: 1, completed: 2 }, content: [{ type: "text", text: "qa" }] }], cursor: {} })
      if (router[2] === "/prompt") return json({ data: { id: "msg_route_prompt", sessionID: router[1], time: { created: 1 }, type: "user", payload: { text: "" }, delivery: "queue" } })
      return router[2] === "/interrupt" ? json({}) : new Response(null, { status: 204 })
    }
    if (route === `GET /api/session/${sessionID}/message`) {
      const desc = url.searchParams.get("order") === "desc"
      if (!desc) historyReads += 1
      const ordered = desc ? [...transcript].reverse() : transcript
      return json({ data: ordered.slice(0, Number(url.searchParams.get("limit") ?? ordered.length)), cursor: {} })
    }
    if (route === `POST /api/session/${sessionID}/agent` || route === `POST /api/session/${sessionID}/model`) return new Response(null, { status: 204 })
    if (route === `POST /api/session/${sessionID}/prompt`) {
      const { text } = (await request.json()) as { text: string }
      prompts.push(text)
      running = true
      const id = `msg_u${prompts.length + 1}`
      transcript.push({ type: "user", id, text, time: { created: Date.now() } })
      return json({ data: { id, sessionID, time: { created: Date.now() }, type: "user", payload: { text }, delivery: "queue" } })
    }
    if (route === `POST /api/experimental/session/${sessionID}/wait`) {
      if (running) await new Promise<void>(resolve => { release = resolve })
      return new Response(null, { status: 204 })
    }
    if (route === `POST /api/session/${sessionID}/interrupt`) {
      interrupts.push(sessionID)
      return json({ interrupted: true })
    }
    unknown.push(route)
    return new Response(null, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
async function until(condition: () => boolean, label: string) {
  const deadline = Date.now() + 5000
  while (!condition() && Date.now() < deadline) await Bun.sleep(10)
  assert.ok(condition(), label)
}
try {
  const page = await browser.newPage({ viewport: { width: 1100, height: 900 } })
  page.setDefaultTimeout(5000)
  await page.goto(server.url.href)
  await page.getByText("Acme Ltd and Beta LLP.").waitFor()
  await until(() => Boolean(stream), "the panel subscribes to engine events")
  const footer = page.locator(".matter-statusbar-label")
  const working = page.locator(".chat-log [role=status]")
  const live = page.locator(".msg.assistant.streaming")
  const sendButton = page.getByRole("button", { name: "Send message", exact: true })
  const stopButton = page.getByRole("button", { name: "Stop answer", exact: true })
  const reads = (text: string | RegExp) => footer.filter({ hasText: text }).waitFor()
  await reads(/^Ready$/)
  assert.equal(await stopButton.count(), 0, "no Stop while nothing runs")

  async function ask(text: string) {
    await page.getByRole("textbox").fill(text)
    await sendButton.click()
    await until(() => prompts.at(-1) === text, `"${text}" is sent`)
  }

  // A multi-step answer: a search step, then the answer starts streaming.
  await ask("Summarise the letter.")
  await stopButton.waitFor()
  assert.equal(await sendButton.count(), 0, "Stop takes Send's place while the answer runs")
  store("msg_s1", [{ type: "text", text: "I'll read the letter first." }, search("letter demand", "Payment of 4,200 is demanded.")], { finish: "tool-calls" })
  store("msg_s2", [{ type: "text", text: "The letter demands payment of 4,200 within" }])
  await live.getByText("The letter demands payment of 4,200 within").waitFor()
  await reads(/^Answer in progress/)
  if (shots) await page.screenshot({ path: path.join(shots, "1-answer-running.png") })
  console.log("PASS while an answer runs, Stop replaces Send")

  await stopButton.click()
  await until(() => interrupts.length === 1, "Stop interrupts the engine session")
  await page.getByRole("button", { name: "Stopping answer", exact: true }).waitFor()
  await working.filter({ hasText: /^Stopping$/ }).waitFor()
  console.log("PASS Stop interrupts the session once and shows it is stopping")

  // The interrupted step's failure is the stop, not an error, and does not settle the turn.
  const beforeFailure = historyReads
  failRunningStep()
  await until(() => historyReads > beforeFailure, "the panel re-reads the interrupted step")
  await Bun.sleep(300)
  await working.filter({ hasText: /^Stopping$/ }).waitFor()
  assert.equal(await page.locator(".msg-error").count(), 0, "the interrupted step is not shown as an error")
  const before = historyReads
  end("interrupted", ["session.execution.interrupted", { reason: "user" }])
  await working.waitFor({ state: "detached" })
  await reads(/^Ready$/)
  const stopped = page.locator(".msg.assistant").last()
  await stopped.getByText("The letter demands payment of 4,200 within").waitFor()
  await stopped.locator(".msg-stopped").filter({ hasText: "Stopped" }).waitFor()
  await page.getByText("Sources (1)").waitFor()
  await Bun.sleep(300)
  assert.equal(await page.locator(".msg-error").count(), 0, "a stopped answer is not shown as an error")
  assert.equal(historyReads - before, 1, "the stopped turn settles once")
  assert.equal(await page.getByRole("textbox").inputValue(), "", "the stopped question is not put back in the composer")
  assert.equal(await stopButton.count(), 0)
  await sendButton.waitFor()
  if (shots) await page.screenshot({ path: path.join(shots, "2-answer-stopped.png") })
  console.log("PASS the stopped answer keeps its partial output, is marked stopped, and the turn settles once")

  // The next question starts normally and renders live.
  await ask("When is payment due?")
  store("msg_f1", [{ type: "text", text: "Payment is due within 14 days of the letter." }], { finish: "stop" })
  await live.getByText("Payment is due within 14 days of the letter.").waitFor()
  end("succeeded", ["session.execution.succeeded", {}])
  await working.waitFor({ state: "detached" })
  await reads(/^Ready$/)
  await page.locator(".msg.assistant").last().getByText("Payment is due within 14 days of the letter.").waitFor()
  assert.equal(await page.locator(".msg-stopped").count(), 1, "only the stopped answer is marked")
  if (shots) await page.screenshot({ path: path.join(shots, "3-follow-up-answered.png") })
  console.log("PASS a follow-up question renders its answer live and settles")

  // The mark comes from the stored history, so a reload keeps it.
  await page.reload()
  await page.getByText("Payment is due within 14 days of the letter.").waitFor()
  await page.locator(".msg-stopped").filter({ hasText: "Stopped" }).waitFor()
  assert.equal(await page.locator(".msg-error").count(), 0)
  console.log("PASS after a reload the stopped answer still reads as stopped")
  assert.deepEqual(unknown, [])
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
