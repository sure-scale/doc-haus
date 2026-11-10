// Run from apps/web: bun test/chat-issues.browser.ts
// Focused real ChatPanel projection; engine HTTP/events are synthetic, no provider calls.
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
const sessionID = "ses_chat"
const model = { providerID: "test", id: "model" }
// A settled first exchange; the follow-up is answered over two steps, and a
// second follow-up is answered by the recovery of a printed tool call.
const transcript: SessionMessageInfo[] = [
  { type: "agent-switched", id: "msg_agent", agent: "legal-review", time: { created: 0 } },
  { type: "user", id: "msg_u1", text: "Who are the parties?", time: { created: 1 } },
  { type: "assistant", id: "msg_a1", agent: "legal-review", model, finish: "stop", time: { created: 2, completed: 3 }, content: [{ type: "text", text: "Acme Ltd and Beta LLP." }] },
  { type: "idle", id: "msg_i1", outcome: "succeeded", time: { created: 4 } },
]
let running = false
let stream: ReadableStreamDefaultController<Uint8Array> | undefined
let release = () => {}
let seq = 0
const prompts: string[] = []
const recoveries: string[] = []
const reads = { messages: 0, outcomes: 0, active: 0 }
const unknown: string[] = []

function send(type: string, data: Record<string, unknown>) {
  seq += 1
  const event = { id: `evt_${seq}`, created: Date.now(), type, location: { directory: "/matter" }, durable: { aggregateID: sessionID, seq, version: 1 }, data: { sessionID, ...data } }
  stream!.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`))
}

// One provider step as the engine reports it: the step's assistant message is
// stored as completed when the step ends, whether or not the turn goes on.
function step(id: string, content: Extract<SessionMessageInfo, { type: "assistant" }>["content"], finish: "tool-calls" | "stop") {
  send("session.step.started", { assistantMessageID: id, agent: "legal-review", model, started: Date.now() })
  transcript.push({ type: "assistant", id, agent: "legal-review", model, finish, time: { created: Date.now(), completed: Date.now() }, content })
  for (const [ordinal, part] of content.filter(part => part.type === "text").entries())
    send("session.text.ended", { assistantMessageID: id, ordinal, text: part.text })
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
    directory: "/matter", matterId: "matter", sessionID: ${JSON.stringify(sessionID)}, agent: "legal-review", available: new Set(["legal-review"]),
    pinned: { name: "legal-review", label: "Legal Review", description: "Answers questions." },
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
    if (route === `GET /api/session/${sessionID}/form`) return json({ data: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(null)
    if (route === `GET /ingest/matters/matter/answer-drafts/session.${sessionID}`) return json({ answers: {}, revision: 0 })
    if (route === `GET /api/session/${sessionID}`)
      return json({ data: { id: sessionID, projectID: "project", agent: "legal-review", model, location: { directory: "/matter" }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 1, updated: 1 } } })
    if (route === `GET /api/session/${sessionID}/message`) {
      reads.messages += 1
      // Only sendPrompt reads newest first, to confirm how the turn ended.
      if (url.searchParams.get("order") === "desc") reads.outcomes += 1
      const ordered = url.searchParams.get("order") === "desc" ? [...transcript].reverse() : transcript
      return json({ data: ordered.slice(0, Number(url.searchParams.get("limit") ?? ordered.length)), cursor: {} })
    }
    // Ingest admits the recovery prompt, and the engine runs it as a new execution.
    if (route === "POST /ingest/matters/matter/continuations/dispatch") {
      const { progressToken } = (await request.json()) as { progressToken: string }
      recoveries.push(progressToken)
      running = true
      const marker = `<dochaus-recovery kind="malformed-tool" run-id="${sessionID}" progress-token="${progressToken}" />`
      transcript.push({ type: "user", id: `msg_recovery_${progressToken.slice("malformed:".length)}`, text: `${marker}\nContinue.`, time: { created: Date.now() } })
      return json({ accepted: true, attempted: true, existing: false, executing: true, persisted: true, uncertain: false })
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
  page.setDefaultTimeout(8000)
  await page.goto(server.url.href)
  await page.getByText("Acme Ltd and Beta LLP.").waitFor()
  await until(() => Boolean(stream), "event subscription")
  const row = { area: "Cap exception", status: "deviating", severity: "high", rule: "Source §2", why: "Unless expressly agreed.", clause: "§2", redline: "" }
  const register = "```issues\n" + JSON.stringify([row]) + "\n```"
  const cases = [
    { text: "Findings.\n\n```issues\n{broken}\n```", cards: 0, literal: "{broken}" },
    { text: "Findings.\n\n" + register + "\n\nA material qualification.", cards: 1, literal: undefined },
    { text: "Example only.\n\n````markdown\n" + register + "\n````", cards: 0, literal: JSON.stringify([row]) },
  ]
  for (const [index, example] of cases.entries()) {
    await page.getByRole("textbox").fill("Review example " + index)
    await page.getByRole("button", { name: "Send message", exact: true }).click()
    await until(() => prompts.length === index + 1, "prompt accepted")
    step("msg_review_" + index, [{ type: "text", text: example.text }], "stop")
    const live = page.locator(".msg.assistant.streaming")
    await live.getByText(index === 2 ? "Example only." : "Findings.", { exact: true }).waitFor()
    assert.equal(await live.locator(".issue-card").count(), example.cards)
    if (example.literal) assert.ok((await live.innerText()).includes(example.literal))
    else {
      assert.equal(await live.locator("pre").count(), 0)
      await live.getByText("Unless expressly agreed.", { exact: true }).waitFor()
      await live.getByText("A material qualification.", { exact: true }).waitFor()
    }
    transcript.push({ type: "idle", id: "msg_idle_" + index, outcome: "succeeded", time: { created: Date.now() } })
    running = false
    send("session.execution.succeeded", {})
    release()
    await live.waitFor({ state: "detached" })
    const settled = page.locator(".msg.assistant").last()
    assert.equal(await settled.locator(".issue-card").count(), example.cards)
    if (example.literal) assert.ok((await settled.innerText()).includes(example.literal))
    await page.reload()
    await page.getByText("Acme Ltd and Beta LLP.").waitFor()
    const reloaded = page.locator(".msg.assistant").last()
    assert.equal(await reloaded.locator(".issue-card").count(), example.cards)
    if (example.literal) assert.ok((await reloaded.innerText()).includes(example.literal))
    else await reloaded.getByText("A material qualification.", { exact: true }).waitFor()
    console.log("PASS live/settled/reloaded issues case " + index)
    await until(() => Boolean(stream), "event subscription after reload")
  }
  assert.deepEqual(recoveries, [])
  assert.deepEqual(unknown, [])
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
