// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/quick-draft.browser.ts
// The quick-draft path (#282) through the real matter page, chat panel, native
// session form card and chat actions; only the engine's and ingest's HTTP and
// event stream are mocked. Set QUICK_DRAFT_SCREENSHOTS=<directory> to save views.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { SessionMessageInfo } from "@opencode/client"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-quick-draft-"))
const shots = process.env.QUICK_DRAFT_SCREENSHOTS
const model = { providerID: "test", id: "model" }
const BANNER = "Quick draft — full drafting checks not run."
const QUICK_TEXT = "Dear client, both invoices remain unpaid [Letter.docx § 1]."
const scope = (change: Record<string, string> = {}) =>
  JSON.stringify({ route: "draft", documents: "one", legal_effect: "no", outside_law: "no", firm_material: "no", sources: "few", ...change })

// What the mocked services answer, set per scenario.
const config = { router: scope(), offer: "ask", settled: false }
// What reached them.
const routerPrompts: string[] = []
const prompts: { session: string; agent: string; text: string }[] = []
const forms: { id: string; sessionID: string; title: string; metadata: Record<string, unknown>; fields: Record<string, unknown>[] }[] = []
const pending = new Map<string, (typeof forms)[number]>()
const replies: { formID: string; answer: unknown }[] = []
const cancels: string[] = []
const deletes: string[] = []
const starts: Record<string, unknown>[] = []
const transcripts = new Map<string, SessionMessageInfo[]>()
const agents = new Map<string, string>()
const streams = new Set<ReadableStreamDefaultController<Uint8Array>>()
let sessions = 0
let seq = 0

function publish(type: string, data: Record<string, unknown>) {
  seq += 1
  const event = new TextEncoder().encode(`data: ${JSON.stringify({ id: `evt_${seq}`, created: Date.now(), type, location: { directory: "/matter" }, data })}\n\n`)
  for (const stream of streams) {
    try {
      stream.enqueue(event)
    } catch {
      streams.delete(stream)
    }
  }
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
  createElement(MemoryRouter, { initialEntries: ["/matter/matter" + location.search] },
    createElement(Routes, null, createElement(Route, { path: "/matter/:id", element: createElement(MatterDetail, { onSessionsChanged: () => {} }) })))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.path.endsWith(".js"))!.text()
const css = (await Promise.all(["styles.css", "styles/primitives.css", "styles/drafting-items.css", "styles/design-system.css", "styles/design-shell.css", "styles/design-conversation.css"]
  .map((file) => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const json = (body: unknown, status = 200) => Response.json(body, { status })
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
    if (route === "GET /ingest/preferences") return json({ quickDraftOffer: config.offer })
    if (route === "GET /api/agent")
      return json({ data: ["qa", "draft", "quick-draft"].map((id) => ({ id, description: id, mode: "primary", hidden: id === "quick-draft" })) })
    if (route === "GET /ingest/matters/matter/telemetry")
      return json({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(null)
    if (url.pathname.startsWith("/ingest/matters/matter/answer-drafts/")) return json({ answers: {}, revision: 0 })
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "POST /ingest/matters/matter/drafting-runs/start") {
      starts.push((await request.json()) as Record<string, unknown>)
      return json({ profile: "lean", mode: "started", ownerSessionID: "ses_owner", status: null })
    }
    if (route === "GET /api/event")
      return new Response(new ReadableStream<Uint8Array>({ start(controller) { streams.add(controller) } }), { headers: { "content-type": "text/event-stream" } })
    if (route === "GET /api/session/active") return json({ data: {} })
    if (route === "GET /api/permission/request") return json({ location: { directory: "/matter" }, data: [] })
    if (route === "POST /api/session") {
      const { title } = (await request.json()) as { title?: string }
      sessions += 1
      return json({ data: session(title === "doc.haus router" ? `ses_router${sessions}` : `ses_chat${sessions}`) })
    }
    const router = url.pathname.match(/^\/api(?:\/experimental)?\/session\/(ses_router\d+)(\/\w+)?$/)
    if (router) {
      if (router[2] === "/message")
        return json({ data: [{ type: "assistant", id: `msg_route_${router[1]}`, agent: "router", model, finish: "stop", time: { created: 1, completed: 2 }, content: [{ type: "text", text: config.router }] }], cursor: {} })
      if (router[2] === "/prompt") {
        routerPrompts.push(((await request.json()) as { text: string }).text)
        return json({ data: { id: "msg_route_prompt", sessionID: router[1], time: { created: 1 }, type: "user", payload: { text: "" }, delivery: "queue" } })
      }
      return router[2] === "/interrupt" ? json({}) : new Response(null, { status: 204 })
    }
    const form = url.pathname.match(/^\/api\/session\/(ses_\w+)\/form(?:\/(frm_\d+))?(\/reply)?$/)
    if (form) {
      const [, sessionID, formID, reply] = form
      if (route.startsWith("GET ")) return json({ data: [...pending.values()].filter((row) => row.sessionID === sessionID) })
      if (request.method === "POST" && !formID) {
        const body = (await request.json()) as Omit<(typeof forms)[number], "id" | "sessionID">
        const created = { ...body, id: `frm_${forms.length + 1}`, sessionID: sessionID! }
        forms.push(created)
        pending.set(created.id, created)
        publish("form.created", { form: created })
        return json({ data: created })
      }
      // A form settles once; a second answer (another tab) is refused.
      if (config.settled || !pending.has(formID!)) {
        pending.delete(formID!)
        return json({ _tag: "FormAlreadySettledError", id: formID, message: "Form was already answered." }, 409)
      }
      pending.delete(formID!)
      if (reply) {
        const { answer } = (await request.json()) as { answer: unknown }
        replies.push({ formID: formID!, answer })
        publish("form.replied", { id: formID, sessionID, answer })
      } else {
        cancels.push(formID!)
        publish("form.cancelled", { id: formID, sessionID })
      }
      return new Response(null, { status: 204 })
    }
    const chat = url.pathname.match(/^\/api(?:\/experimental)?\/session\/(ses_\w+)(\/\w+)?$/)
    if (chat) {
      const [, id, action] = chat
      const transcript = transcripts.get(id!) ?? []
      if (!action && request.method === "GET") return json({ data: session(id!) })
      if (!action && request.method === "DELETE") {
        deletes.push(id!)
        return new Response(null, { status: 204 })
      }
      if (action === "/message") {
        const ordered = url.searchParams.get("order") === "desc" ? [...transcript].reverse() : transcript
        return json({ data: ordered.slice(0, Number(url.searchParams.get("limit") ?? ordered.length)), cursor: {} })
      }
      if (action === "/agent") {
        // The engine records the switch in the transcript; turns take their agent from it.
        const { agent } = (await request.json()) as { agent: string }
        agents.set(id!, agent)
        transcripts.set(id!, [...transcript, { type: "agent-switched", id: `msg_s${transcript.length}`, agent, time: { created: Date.now() } }])
        return new Response(null, { status: 204 })
      }
      if (action === "/prompt") {
        const { text } = (await request.json()) as { text: string }
        const agent = agents.get(id!) ?? "qa"
        prompts.push({ session: id!, agent, text })
        const n = transcript.length
        transcripts.set(id!, [
          ...transcript,
          { type: "user", id: `msg_u${n}`, text, time: { created: Date.now() } },
          { type: "assistant", id: `msg_a${n}`, agent, model, finish: "stop", time: { created: Date.now(), completed: Date.now() }, content: [{ type: "text", text: agent === "quick-draft" ? QUICK_TEXT : "Answered." }] },
          { type: "idle", id: `msg_i${n}`, outcome: "succeeded", time: { created: Date.now() } },
        ])
        // The engine reports the finished turn, which settles it in the panel.
        setTimeout(() => publish("session.execution.succeeded", { sessionID: id }), 50)
        return json({ data: { id: `msg_u${n}`, sessionID: id, time: { created: Date.now() }, type: "user", payload: { text }, delivery: "queue" } })
      }
      if (action === "/interrupt") return json({ interrupted: true })
      return new Response(null, { status: 204 })
    }
    return new Response(null, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
async function until(condition: () => boolean, label: string) {
  const deadline = Date.now() + 5000
  while (!condition() && Date.now() < deadline) await Bun.sleep(10)
  assert.ok(condition(), label)
}
// A fresh browser (no remembered pick unless given) on a new chat, or on `query`.
async function open(query = "?view=chat", remembered?: string) {
  const context = await browser.newContext({ viewport: { width: 1100, height: 900 } })
  if (remembered) await context.addInitScript((value) => localStorage.setItem("dh.quickDraft.choice.matter", value), remembered)
  const page = await context.newPage()
  page.setDefaultTimeout(5000)
  await page.goto(server.url.href + query)
  await page.getByRole("textbox").waitFor()
  const send = async (text: string) => {
    await page.getByRole("textbox").fill(text)
    await page.getByRole("button", { name: "Send message", exact: true }).click()
  }
  const pick = async (label: string) => {
    await page.getByTitle("Choose assistant or workflow").click()
    await page.locator(".assistant-option", { hasText: label }).first().click()
  }
  return { context, page, send, pick, card: page.locator(".route-offer") }
}
const counts = () => ({ prompts: prompts.length, starts: starts.length, forms: forms.length })
const settle = () => Bun.sleep(300)
try {
  {
    // Auto, eligible: the offer, with Quick pre-selected, runs nothing until Start.
    const request = "Draft a short covering email to the client summarising the position on the two invoices."
    const { context, page, send, card } = await open()
    await send(request)
    await until(() => forms.length === 1, "the offer is created as a native session form")
    assert.deepEqual(forms[0]!.metadata, { kind: "route", route: "draft", request })
    assert.equal(forms[0]!.fields[0]!.default, "quick")
    assert.ok(!routerPrompts[0]!.includes("quick-draft"), "Auto never offers Quick draft as a candidate")
    assert.ok(routerPrompts[0]!.includes("- draft:"))
    await card.waitFor()
    assert.equal(await card.getByLabel(/Quick draft here in chat/).isChecked(), true)
    assert.equal(await card.getByLabel(/Full drafting run/).isChecked(), false)
    await page.locator(".msg.user", { hasText: request }).waitFor()
    await settle()
    assert.equal(prompts.length + starts.length, 0, "nothing runs before consent")
    if (shots) await page.screenshot({ path: path.join(shots, "1-offer.png") })
    console.log("PASS Auto offers an eligible request in a native form with Quick pre-selected; nothing runs before consent")

    // A reload keeps the pending offer and the request; still nothing runs.
    const sessionID = forms[0]!.sessionID
    await page.goto(`${server.url.href}?view=chat&session=${sessionID}`)
    await card.waitFor()
    assert.equal(await card.getByLabel(/Quick draft here in chat/).isChecked(), true)
    await page.locator(".msg.user", { hasText: request }).waitFor()
    assert.equal(await page.getByRole("button", { name: "Send message", exact: true }).isDisabled(), true)
    await settle()
    assert.equal(prompts.length + starts.length, 0)
    assert.equal(forms.length, 1, "a reload creates no second offer")
    if (shots) await page.screenshot({ path: path.join(shots, "2-offer-after-reload.png") })
    console.log("PASS a reload restores the pending offer and its request without running anything")

    await card.getByRole("button", { name: "Start", exact: true }).click()
    await until(() => prompts.length === 1, "Start sends the request to Quick draft")
    assert.deepEqual(replies, [{ formID: forms[0]!.id, answer: { route: "quick" } }])
    assert.deepEqual(prompts[0], { session: sessionID, agent: "quick-draft", text: request })
    await page.getByRole("note").filter({ hasText: BANNER }).waitFor()
    await page.locator(".msg.assistant", { hasText: "both invoices remain unpaid" }).waitFor()
    assert.equal(await page.evaluate(() => localStorage.getItem("dh.quickDraft.choice.matter")), "quick")
    if (shots) await page.screenshot({ path: path.join(shots, "3-quick-draft.png") })
    console.log("PASS Start dispatches the settled choice once: the quick draft answers in chat under the banner")

    await page.getByRole("button", { name: "Run full drafting cycle", exact: true }).click()
    await until(() => starts.length === 1, "the upgrade starts the full run")
    assert.deepEqual(starts[0], { request, agent: "draft", model: { providerID: "test", modelID: "model" }, workingDraft: QUICK_TEXT })
    console.log("PASS Run full drafting cycle starts the full run with the request and the quick text as its working draft")
    await context.close()
  }
  {
    // Cancel: nothing ran, so the text goes back to the composer and the empty
    // conversation made for the offer is removed.
    const before = counts()
    const request = "Draft a short file note of today's call."
    const { context, page, send, card } = await open()
    await send(request)
    await card.waitFor()
    const sessionID = forms.at(-1)!.sessionID
    await card.getByRole("button", { name: "Cancel", exact: true }).click()
    await until(() => cancels.length === 1, "Cancel cancels the native form")
    await until(() => deletes.includes(sessionID), "the empty offer conversation is removed")
    assert.equal(await page.getByRole("textbox").inputValue(), request)
    await card.waitFor({ state: "detached" })
    await settle()
    assert.deepEqual(counts(), { ...before, forms: before.forms + 1 })
    console.log("PASS Cancel returns the request to the composer and runs nothing")
    await context.close()
  }
  {
    // The lawyer's last pick on this matter is the pre-selection; Full runs in full.
    const before = counts()
    const request = "Draft a short update to the client on the invoices."
    const { context, send, card } = await open("?view=chat", "full")
    await send(request)
    await card.waitFor()
    assert.equal(forms.at(-1)!.fields[0]!.default, "full")
    assert.equal(await card.getByLabel(/Full drafting run/).isChecked(), true)
    const sessionID = forms.at(-1)!.sessionID
    await card.getByRole("button", { name: "Start", exact: true }).click()
    await until(() => starts.length === before.starts + 1, "Full starts the drafting run")
    assert.deepEqual(starts.at(-1), { request, agent: "draft", model: { providerID: "test", modelID: "model" } })
    assert.deepEqual(replies.at(-1)!.answer, { route: "full" })
    await until(() => deletes.includes(sessionID), "the empty offer conversation is removed")
    assert.equal(prompts.length, before.prompts)
    console.log("PASS the remembered pick pre-selects Full; Full starts the run with the request only")
    await context.close()
  }
  {
    // A form answered elsewhere first: this answer is refused and starts nothing.
    const before = counts()
    const { context, page, send, card } = await open()
    await send("Draft a short covering letter for the enclosed invoices.")
    await card.waitFor()
    config.settled = true
    await card.getByRole("button", { name: "Start", exact: true }).click()
    await page.getByRole("alert").waitFor()
    await settle()
    assert.deepEqual(counts(), { ...before, forms: before.forms + 1 })
    config.settled = false
    console.log("PASS an offer already answered elsewhere starts no second run")
    await context.close()
  }
  {
    // The firm turned the offer off: Auto starts the full run with no question.
    const before = counts()
    config.offer = "off"
    const { context, send } = await open()
    await send("Draft a short covering email on the invoices.")
    await until(() => starts.length === before.starts + 1, "the full run starts")
    assert.equal(forms.length, before.forms, "no offer when the firm turned it off")
    config.offer = "ask"
    console.log("PASS Off: Auto starts the full run with no question")
    await context.close()
  }
  {
    // A request with legal effect is not eligible: Full, with no question.
    const before = counts()
    config.router = scope({ legal_effect: "yes" })
    const { context, send } = await open()
    await send("Draft a short notice terminating the supply agreement.")
    await until(() => starts.length === before.starts + 1, "the full run starts")
    assert.equal(forms.length, before.forms, "no offer for a request with legal effect")
    config.router = scope()
    console.log("PASS an ineligible request starts the full run with no question")
    await context.close()
  }
  {
    // The explicit Quick draft pick is the consent: no form, straight to Quick.
    const before = counts()
    const request = "Draft a short covering email on the invoices."
    const { context, page, send, pick } = await open()
    await pick("Quick draft")
    await send(request)
    await until(() => prompts.length === before.prompts + 1, "the explicit pick sends to Quick draft")
    assert.equal(prompts.at(-1)!.agent, "quick-draft")
    assert.equal(forms.length, before.forms, "no second question after the explicit pick")
    assert.ok(routerPrompts.at(-1)!.includes("- draft:") && !routerPrompts.at(-1)!.includes("- qa:"), "the pick checks scope only")
    await page.getByRole("note").filter({ hasText: BANNER }).waitFor()
    console.log("PASS the explicit Quick draft pick is consent: no form, and the banner shows")
    await context.close()
  }
  {
    // An explicit pick outside Quick's scope explains the limit and offers Full.
    const before = counts()
    config.router = scope({ outside_law: "yes" })
    const request = "Draft a short email advising on the statutory interest due."
    const { context, page, send, pick } = await open()
    await pick("Quick draft")
    await send(request)
    const notice = page.locator(".route-offer", { hasText: "law beyond the matter's documents" })
    await notice.waitFor()
    assert.equal(await page.getByRole("textbox").inputValue(), request)
    await settle()
    assert.deepEqual(counts(), before, "nothing runs for a request outside Quick's scope")
    if (shots) await page.screenshot({ path: path.join(shots, "4-quick-limit.png") })
    await notice.getByRole("button", { name: "Run full drafting cycle", exact: true }).click()
    await until(() => starts.length === before.starts + 1, "Full starts from the notice")
    assert.equal(starts.at(-1)!.request, request)
    assert.equal(starts.at(-1)!.workingDraft, undefined)
    config.router = scope()
    console.log("PASS an explicit pick outside Quick's scope explains the limit, runs nothing, and offers Full")
    await context.close()
  }
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
