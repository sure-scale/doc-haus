// Run from apps/web: bun test/grounding.browser.ts
// Renders the real chat panel against a settled transcript with only the
// engine's HTTP and event stream mocked, so the grounding badge is checked in
// the surface the lawyer uses. Set GROUNDING_SCREENSHOTS=<directory> to save it.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { SessionMessageInfo } from "@opencode/client"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-grounding-"))
const shots = process.env.GROUNDING_SCREENSHOTS
const sessionID = "ses_grounding"
const model = { providerID: "test", id: "model" }

const EXCERPT = "The Supplier shall deliver the synthetic cartons within thirty days of each order."
const TERMINATION = "Either party may terminate this agreement for material breach on thirty days' written notice."
const citation = (section: string, excerpt: string) => ({ documentName: "Supply agreement.docx", docPath: "/matter/Supply agreement.docx", section, excerpt, charStart: 0, charEnd: excerpt.length })
const documents = [{ id: 1, name: "Supply agreement.docx", doc_path: "/matter/Supply agreement.docx", created_at: 1 }]

// A grounded answer, an answer that consulted the documents but produced no
// verified quotation, and an answer that links only part of what it says.
const transcript: SessionMessageInfo[] = [
  { type: "agent-switched", id: "msg_agent", agent: "qa", time: { created: 1 } },
  { type: "user", id: "msg_u1", text: "What delivery window applies, and can either side terminate for breach?", time: { created: 1 } },
  {
    type: "assistant", id: "msg_a1", agent: "qa", model, finish: "stop", time: { created: 2, completed: 3 },
    content: [
      {
        type: "tool", id: "call_cite_1", name: "cite", time: { created: 2, completed: 2 }, state: {
          status: "completed", input: { quote: EXCERPT }, content: [{ type: "text", text: "Verified verbatim against the live document." }],
          metadata: { citations: [citation("7.2", EXCERPT)] },
        },
      },
      {
        type: "tool", id: "call_cite_1b", name: "cite", time: { created: 2, completed: 2 }, state: {
          status: "completed", input: { quote: TERMINATION }, content: [{ type: "text", text: "Verified verbatim against the live document." }],
          metadata: { citations: [citation("11.4", TERMINATION)] },
        },
      },
      { type: "text", text: "The supplier must deliver the cartons within thirty days of each order [Supply agreement § 7.2]." },
      { type: "text", text: "Either party may terminate for material breach on thirty days' written notice [Supply agreement § 11.4]." },
    ],
  },
  { type: "idle", id: "msg_i1", outcome: "succeeded", time: { created: 4 } },
  { type: "user", id: "msg_u2", text: "Does the agreement cap the supplier's liability?", time: { created: 5 } },
  {
    type: "assistant", id: "msg_a2", agent: "qa", model, finish: "stop", time: { created: 6, completed: 7 },
    content: [
      {
        type: "tool", id: "call_search", name: "search-document", time: { created: 6, completed: 6 }, state: {
          status: "completed", input: { query: "liability cap" }, content: [{ type: "text", text: "No limiting clause found." }],
          metadata: { citations: [{ ...citation("9.1", TERMINATION), score: 0.72 }] },
        },
      },
      { type: "text", text: "The agreement places no cap on the supplier's liability for any loss arising under it." },
    ],
  },
  { type: "idle", id: "msg_i2", outcome: "succeeded", time: { created: 8 } },
  { type: "user", id: "msg_u3", text: "And what notice does termination require?", time: { created: 9 } },
  {
    type: "assistant", id: "msg_a3", agent: "qa", model, finish: "stop", time: { created: 10, completed: 11 },
    content: [
      {
        type: "tool", id: "call_cite_2", name: "cite", time: { created: 10, completed: 10 }, state: {
          status: "completed", input: { quote: TERMINATION }, content: [{ type: "text", text: "Verified verbatim against the live document." }],
          metadata: { citations: [citation("11.4", TERMINATION)] },
        },
      },
      { type: "text", text: "Termination for material breach requires thirty days' written notice [Supply agreement § 11.4]." },
      { type: "text", text: "That period is longer than the market standard for a supply agreement of this size." },
    ],
  },
  { type: "idle", id: "msg_i3", outcome: "succeeded", time: { created: 12 } },
]

const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import ChatPanel from ${source("components/ChatPanel.tsx")}
createRoot(document.getElementById("root")).render(createElement(MemoryRouter, null,
  createElement(ChatPanel, {
    directory: "/matter", matterId: "matter", sessionID: ${JSON.stringify(sessionID)}, agent: "qa", available: new Set(["qa"]),
    pinned: { name: "qa", label: "Q&A", description: "Answers questions." },
    documents: ${JSON.stringify(documents)},
    onAgentChange: () => {}, onViewDocument: () => {}, onBusyChange: () => {},
  })))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.path.endsWith(".js"))!.text()
const css = (await Promise.all(["styles.css", "styles/design-system.css", "styles/design-conversation.css", "components/GroundingBadge.css"]
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
    if (route === "GET /api/event") return new Response(new ReadableStream<Uint8Array>({ start() {} }), { headers: { "content-type": "text/event-stream" } })
    if (route === "GET /api/session/active") return json({ data: {} })
    if (route === "GET /api/permission/request") return json({ location: { directory: "/matter" }, data: [] })
    if (route === `GET /api/session/${sessionID}/form`) return json({ data: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(null)
    if (route === `GET /ingest/matters/matter/answer-drafts/session.${sessionID}`) return json({ answers: {}, revision: 0 })
    if (route === `GET /api/session/${sessionID}`)
      return json({ data: { id: sessionID, projectID: "project", agent: "qa", model, location: { directory: "/matter" }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 1, updated: 1 } } })
    if (route === `GET /api/session/${sessionID}/message`) {
      const ordered = url.searchParams.get("order") === "desc" ? [...transcript].reverse() : transcript
      return json({ data: ordered.slice(0, Number(url.searchParams.get("limit") ?? ordered.length)), cursor: {} })
    }
    return new Response(null, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 900, height: 1100 }, deviceScaleFactor: 2 })
  page.setDefaultTimeout(8000)
  await page.goto(server.url.href)
  const answers = page.locator(".msg.assistant")
  await answers.first().waitFor()
  assert.equal(await answers.count(), 3)

  await answers.nth(0).getByText("Grounded", { exact: true }).waitFor()
  await answers.nth(1).getByText("No verified support", { exact: true }).waitFor()
  await answers.nth(2).getByText("Limited support", { exact: true }).waitFor()

  // The disclosure names what was not linked instead of certifying it.
  await answers.nth(2).locator(".grounding > summary").click()
  await answers.nth(2).getByRole("listitem").filter({ hasText: "That period is longer than the market standard" }).waitFor()
  await answers.nth(1).locator(".grounding > summary").click()
  await answers.nth(1).getByText("No verified quotation supports this answer.").waitFor()

  if (shots) {
    // Capture the whole conversation, not the inner log's current scroll view.
    await page.addStyleTag({ content: ".chat-log{max-height:none!important;overflow:visible!important}.chat-card{height:auto!important}.composer,.chat-jump{display:none!important}" })
    await page.locator(".chat-log").evaluate(element => { element.scrollTop = 0 })
    await page.screenshot({ path: path.join(shots, "answer-grounding.png"), fullPage: true })
  }
  console.log("PASS grounded, limited and no-verified-support states render on Q&A answers")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
