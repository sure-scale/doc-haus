// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/chat-live-answer.browser.ts
// Uses the real chat panel and its hooks; only the engine's HTTP and event stream are mocked.
// Set CHAT_LIVE_SCREENSHOTS=<directory> to save the live and settled answers.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { SessionMessageInfo } from "@opencode/client"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chat-live-"))
const shots = process.env.CHAT_LIVE_SCREENSHOTS
const sessionID = "ses_chat"
const model = { providerID: "test", id: "model" }
// A settled first exchange; the follow-up is answered over two steps.
const transcript: SessionMessageInfo[] = [
  { type: "user", id: "msg_u1", text: "Who are the parties?", time: { created: 1 } },
  { type: "assistant", id: "msg_a1", agent: "qa", model, finish: "stop", time: { created: 2, completed: 3 }, content: [{ type: "text", text: "Acme Ltd and Beta LLP." }] },
  { type: "idle", id: "msg_i1", outcome: "succeeded", time: { created: 4 } },
]
let running = false
let stream: ReadableStreamDefaultController<Uint8Array> | undefined
let release = () => {}
let seq = 0
const prompts: string[] = []
const reads = { messages: 0, active: 0 }
const unknown: string[] = []

function send(type: string, data: Record<string, unknown>) {
  seq += 1
  const event = { id: `evt_${seq}`, created: Date.now(), type, location: { directory: "/matter" }, durable: { aggregateID: sessionID, seq, version: 1 }, data: { sessionID, ...data } }
  stream!.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`))
}

// One provider step as the engine reports it: the step's assistant message is
// stored as completed when the step ends, whether or not the turn goes on.
function step(id: string, content: Extract<SessionMessageInfo, { type: "assistant" }>["content"], finish: "tool-calls" | "stop") {
  send("session.step.started", { assistantMessageID: id, agent: "qa", model, started: Date.now() })
  transcript.push({ type: "assistant", id, agent: "qa", model, finish, time: { created: Date.now(), completed: Date.now() }, content })
  send("session.text.ended", { assistantMessageID: id, ordinal: 0, text: "" })
  send("session.step.ended", { assistantMessageID: id, finish })
}

const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import ChatPanel from ${source("components/ChatPanel.tsx")}
createRoot(document.getElementById("root")).render(createElement(MemoryRouter, null,
  createElement(ChatPanel, {
    directory: "/matter", sessionID: ${JSON.stringify(sessionID)}, agent: "qa", available: new Set(["qa"]),
    pinned: { name: "qa", label: "Q&A", description: "Answers questions." },
    onAgentChange: () => {}, onViewDocument: () => {}, onBusyChange: (busy) => { window.busy = busy },
  })))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.path.endsWith(".js"))!.text()
const css = (await Promise.all(["styles.css", "styles/design-system.css", "styles/design-conversation.css"]
  .map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const json = (body: unknown) => Response.json(body)
const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  idleTimeout: 0,
  async fetch(request) {
    const url = new URL(request.url)
    const route = `${request.method} ${url.pathname}`
    if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (route === "GET /") return new Response(`<html><meta charset="utf-8"><style>body{font-family:sans-serif;padding:16px;max-width:760px}${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "GET /api/event")
      return new Response(new ReadableStream<Uint8Array>({ start(controller) { stream = controller } }), { headers: { "content-type": "text/event-stream" } })
    if (route === "GET /api/session/active") {
      reads.active += 1
      return json({ data: running ? { [sessionID]: { type: "running" } } : {} })
    }
    if (route === "GET /api/permission/request") return json({ location: { directory: "/matter" }, data: [] })
    if (route === `GET /api/session/${sessionID}/message`) {
      reads.messages += 1
      const ordered = url.searchParams.get("order") === "desc" ? [...transcript].reverse() : transcript
      return json({ data: ordered.slice(0, Number(url.searchParams.get("limit") ?? ordered.length)), cursor: {} })
    }
    if (route === `POST /api/session/${sessionID}/agent` || route === `POST /api/session/${sessionID}/model`) return new Response(null, { status: 204 })
    if (route === `POST /api/session/${sessionID}/prompt`) {
      const { text } = (await request.json()) as { text: string }
      prompts.push(text)
      running = true
      transcript.push({ type: "user", id: "msg_u2", text, time: { created: Date.now() } })
      return json({ data: { id: "msg_u2", sessionID, time: { created: Date.now() }, type: "user", payload: { text }, delivery: "queue" } })
    }
    if (route === `POST /api/experimental/session/${sessionID}/wait`) {
      await new Promise<void>(resolve => { release = resolve })
      return new Response(null, { status: 204 })
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
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  // The panel's slow resync poll is a 15 s interval; the page clock lets the
  // test fire it at the exact moment between two steps.
  await page.clock.install()
  await page.goto(server.url.href)
  await page.getByText("Acme Ltd and Beta LLP.").waitFor()
  await until(() => Boolean(stream), "the panel subscribes to engine events")
  await page.getByRole("textbox").fill("What date is the letter dated?")
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await until(() => prompts.length === 1, "the follow-up is sent")
  const live = page.locator(".msg.assistant.streaming")

  step("msg_s1", [
    { type: "text", text: "I'll check the letter's date." },
    { type: "tool", id: "call_search", name: "search-document", time: { created: Date.now(), completed: Date.now() }, state: {
      status: "completed", input: { query: "letter date" }, content: [{ type: "text", text: "Dated 12 October 2026." }],
      metadata: { citations: [{ document: "Letter.docx", section: "1", excerpt: "Dated 12 October 2026." }] },
    } },
  ], "tool-calls")
  await live.getByText("I'll check the letter's date.").waitFor()
  // Between steps the last stored message is a completed assistant while the
  // engine still runs the turn. The poll must not settle the turn here.
  const before = { ...reads }
  await page.clock.fastForward(15_000)
  await until(() => reads.messages > before.messages, "the resync poll reads the conversation")
  await Bun.sleep(300)
  assert.equal(await page.evaluate(() => (window as unknown as { busy: boolean }).busy), true, "the turn settled between its steps")
  assert.equal(await live.count(), 1, "the live answer left the conversation between steps")
  console.log("PASS a resync between two steps of one answer leaves the turn working")

  step("msg_s2", [{ type: "text", text: "The letter is dated 12 October 2026." }], "stop")
  await live.getByText("The letter is dated 12 October 2026.").waitFor()
  if (shots) await page.screenshot({ path: path.join(shots, "1-live-final-answer.png"), fullPage: true })
  console.log("PASS the final step's answer renders live, while the turn is still working")

  // The engine ends the turn; the answer settles at once, without a reload or a poll.
  transcript.push({ type: "idle", id: "msg_i2", outcome: "succeeded", time: { created: Date.now() } })
  running = false
  send("session.execution.succeeded", {})
  release()
  await live.waitFor({ state: "detached" })
  assert.equal(await page.evaluate(() => (window as unknown as { busy: boolean }).busy), false)
  const answers = page.locator(".msg.assistant")
  assert.equal(await answers.count(), 2)
  await answers.nth(1).getByText("I'll check the letter's date.").waitFor()
  await answers.nth(1).getByText("The letter is dated 12 October 2026.").waitFor()
  if (shots) await page.screenshot({ path: path.join(shots, "2-settled-final-answer.png"), fullPage: true })
  assert.deepEqual(unknown, [])
  console.log("PASS the engine's execution success settles the whole answer without a reload")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
