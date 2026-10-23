// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/quick-draft.browser.ts
// The quick-draft path (#282), and the client summary (#26) that shares its
// explicit-pick contract, through the real matter page, chat panel, native
// session form card and chat actions; only the engine's and ingest's HTTP and
// event stream are mocked. Set QUICK_DRAFT_SCREENSHOTS=<directory> to save views.
// Also: asking about a review-table answer (#95) never overwrites unsent text.
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
// A quick answer as the engine stores it: narration lines in the steps that
// call tools, then the form line, the draft with a placeholder, its sources and
// an open-items list that misses the placeholder.
const NARRATION = "Quick draft on your invoices — I'll pull the payment terms."
const LATER_NARRATION = "I've got the figures — now I'll verify them."
const QUICK_TEXT = "Dear [Client contact], both invoices remain unpaid [Letter.docx § 1].\n\nSources:\n- unpaid invoices [Letter.docx § 1]\n\nOpen items:\nNone."
const QUICK_ANSWER = `Form: advice letter\n\n${QUICK_TEXT}`
// The quick text as shown and carried on: code lists the draft's placeholder
// in place of the answer's "None.".
const QUICK_SHOWN = QUICK_TEXT.replace("Open items:\nNone.", "Open items:\n- [Client contact]")
const LEGEND = "Privileged & confidential"
// A client summary as the engine stores it: narration in a step that calls a
// tool, then the summary.
const SUMMARY_NARRATION = "Let me find the invoice terms first."
const SUMMARY = "What this is\n\nA letter about two unpaid invoices [Letter.docx § 1]."
// One review-table answer to ask about.
const GRID = {
  columns: [{ id: "c1", question: "Which invoices are unpaid?" }],
  cells: { "Letter.docx::c1": { answer: "Both invoices.", citation: { documentName: "Letter.docx", docPath: "/matter/Letter.docx", section: "1", excerpt: "Both invoices remain unpaid.", charStart: 0, charEnd: 28 }, status: "filled", questionHash: "unused" } },
}
const scope = (change: Record<string, string> = {}) =>
  JSON.stringify({ route: "draft", documents: "one", legal_effect: "no", outside_law: "no", firm_material: "no", sources: "few", ...change })

// What the mocked services answer, set per scenario.
const config = { router: scope(), offer: "ask", settled: false, legend: "", preferencesDown: false, hold: undefined as Promise<void> | undefined }
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
import { Link, MemoryRouter, Route, Routes } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import MatterDetail from ${source("pages/MatterDetail.tsx")}
import { ToastProvider } from ${source("components/Toast.tsx")}
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(MemoryRouter, { initialEntries: ["/matter/matter" + location.search] },
    // Stands in for the sidebar's review-table link, which keeps the current query.
    createElement(Link, { to: { search: "?view=review" } }, "Open review table"),
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
    if (route === "GET /ingest/preferences" && config.preferencesDown) return json({ error: "Settings unavailable" }, 500)
    if (route === "GET /ingest/preferences") return json({ quickDraftOffer: config.offer, agreementLegend: "", adviceLegend: config.legend })
    if (route === "GET /api/agent")
      return json({ data: ["qa", "draft", "quick-draft", "client-summary"].map((id) => ({ id, description: id, mode: "primary", hidden: id !== "qa" && id !== "draft" })) })
    if (route === "GET /ingest/matters/matter/telemetry")
      return json({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(null)
    if (route === "GET /ingest/matters/matter/grid") return json(GRID)
    if (route === "PUT /ingest/matters/matter/grid") return new Response(null, { status: 204 })
    if (route === "POST /ingest/matters/matter/grid/sources") return json({ "Letter.docx::c1": "verified" })
    if (url.pathname.startsWith("/ingest/matters/matter/answer-drafts/")) return json({ answers: {}, revision: 0 })
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "POST /ingest/matters/matter/drafting-runs/start") {
      starts.push((await request.json()) as Record<string, unknown>)
      return json({ profile: "lean", mode: "started", ownerSessionID: "ses_owner", status: null })
    }
    if (route === "GET /api/event")
      // The engine opens every stream with server.connected.
      return new Response(new ReadableStream<Uint8Array>({ start(controller) {
        streams.add(controller)
        controller.enqueue(new TextEncoder().encode(`data: ${JSON.stringify({ id: "evt_connected", created: Date.now(), type: "server.connected", data: {} })}\n\n`))
      } }), { headers: { "content-type": "text/event-stream" } })
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
        // The reply streams as text deltas while the router turn runs: its
        // acknowledgement line first, then the JSON in a later batch.
        const line = config.router.indexOf("\n")
        const delta = (text: string) => publish("session.text.delta", { sessionID: router[1], assistantMessageID: `msg_route_${router[1]}`, ordinal: 0, delta: text })
        if (line > 0) {
          delta(config.router.slice(0, line + 1))
          setTimeout(() => delta(config.router.slice(line + 1)), 30)
        }
        return json({ data: { id: "msg_route_prompt", sessionID: router[1], time: { created: 1 }, type: "user", payload: { text: "" }, delivery: "queue" } })
      }
      // A held router turn has not finished: Auto has no route yet.
      if (router[2] === "/wait") await config.hold
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
        const time = { created: Date.now(), completed: Date.now() }
        const answer: SessionMessageInfo[] =
          agent === "quick-draft"
            ? [
                { type: "assistant", id: `msg_a${n}`, agent, model, finish: "tool-calls", time, content: [
                  { type: "text", text: NARRATION },
                  { type: "tool", id: `call_${n}`, name: "search-document", time, state: { status: "completed", input: { query: "invoices" }, content: [{ type: "text", text: "Both invoices remain unpaid." }], metadata: { citations: [] } } },
                ] },
                { type: "assistant", id: `msg_c${n}`, agent, model, finish: "tool-calls", time, content: [
                  { type: "text", text: LATER_NARRATION },
                  { type: "tool", id: `call_c${n}`, name: "cite", time, state: { status: "completed", input: { quote: "Both invoices remain unpaid." }, content: [{ type: "text", text: "Verified." }] } },
                ] },
                { type: "assistant", id: `msg_b${n}`, agent, model, finish: "stop", time, content: [{ type: "text", text: QUICK_ANSWER }] },
              ]
            : agent === "client-summary"
              ? [
                  { type: "assistant", id: `msg_a${n}`, agent, model, finish: "tool-calls", time, content: [
                    { type: "text", text: SUMMARY_NARRATION },
                    { type: "tool", id: `call_${n}`, name: "search-document", time, state: { status: "completed", input: { query: "invoices" }, content: [{ type: "text", text: "Both invoices remain unpaid." }], metadata: { citations: [] } } },
                  ] },
                  { type: "assistant", id: `msg_b${n}`, agent, model, finish: "stop", time, content: [{ type: "text", text: SUMMARY }] },
                ]
              : [{ type: "assistant", id: `msg_a${n}`, agent, model, finish: "stop", time, content: [{ type: "text", text: "Answered." }] }]
        transcripts.set(id!, [
          ...transcript,
          { type: "user", id: `msg_u${n}`, text, time: { created: Date.now() } },
          ...answer,
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
    // A router reply with no acknowledgement line gets the chat's own.
    await page.locator(".msg-ack", { hasText: "Happy to draft that." }).waitFor()
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
    const quick = page.locator(".msg.assistant", { hasText: "both invoices remain unpaid" })
    await quick.waitFor()
    assert.equal(await page.evaluate(() => localStorage.getItem("dh.quickDraft.choice.matter")), "quick")
    // Only the draft is the quick text: no narration line, no form line, and
    // no legend while the firm has none.
    assert.equal(await quick.getByText(NARRATION).count(), 0, "the narration line is not quick text")
    assert.equal(await quick.getByText(LATER_NARRATION).count(), 0, "narration in a later working step is not quick text")
    assert.equal(await quick.getByText("Form: advice letter").count(), 0, "the form line is not shown")
    assert.equal(await quick.locator(".quick-draft-legend").count(), 0, "no legend while the firm's advice legend is none")
    // Code puts the placeholder the answer's own open items missed into that
    // list (its "None." replaced), and the chat does not list it again below.
    await quick.getByRole("listitem").filter({ hasText: "[Client contact]" }).waitFor()
    assert.equal(await quick.getByText("None.", { exact: true }).count(), 0, "the answer's 'None.' is replaced by its placeholder")
    assert.equal(await quick.locator(".quick-draft-open").count(), 0, "open items are listed once, not again in a note")
    if (shots) await page.screenshot({ path: path.join(shots, "3-quick-draft.png") })
    console.log("PASS Start dispatches the settled choice once: the quick draft answers in chat under the banner, its placeholder listed as an open item")

    await page.getByRole("button", { name: "Run full drafting cycle", exact: true }).click()
    await until(() => starts.length === 1, "the upgrade starts the full run")
    assert.deepEqual(starts[0], { request, agent: "draft", model: { providerID: "test", modelID: "model" }, workingDraft: QUICK_SHOWN })
    console.log("PASS Run full drafting cycle starts the full run with the request and the quick text as its working draft")
    await context.close()
  }
  {
    // Auto's acknowledgement: the router's own line shows while it is still
    // choosing, stays above the offer and the quick draft, and is never sent,
    // quick text or the upgrade's working draft.
    const ACK = "On it — I'll draft that covering email."
    const request = "Draft a short covering email to the client on the two unpaid invoices."
    const before = { ...counts(), routes: routerPrompts.length }
    let release = () => {}
    config.hold = new Promise<void>((resolve) => (release = resolve))
    config.router = `${ACK}\n${scope()}`
    const { context, page, send, card } = await open()
    await send(request)
    const ack = page.locator(".msg-ack", { hasText: ACK })
    await ack.waitFor()
    assert.equal(forms.length, before.forms, "the acknowledgement shows before Auto has chosen")
    assert.equal(await page.locator(".msg.assistant .msg-agent", { hasText: "Draft" }).count(), 0)
    if (shots) await page.screenshot({ path: path.join(shots, "ack-1-routing.png") })
    release()
    config.hold = undefined
    await card.waitFor()
    assert.equal(await page.locator(".msg-ack").count(), 1, "one acknowledgement: the router's line stays")
    await card.getByRole("button", { name: "Start", exact: true }).click()
    await until(() => prompts.length === before.prompts + 1, "Start sends the request to Quick draft")
    assert.equal(prompts.at(-1)!.text, request, "the assistant is sent the request alone")
    const quick = page.locator(".msg.assistant", { hasText: "both invoices remain unpaid" })
    await quick.waitFor()
    await ack.waitFor()
    assert.equal(await quick.getByText(ACK).count(), 0, "the acknowledgement is not quick text")
    // Above the answer, under the request, in reading order.
    const order = await page.evaluate(() => [...document.querySelectorAll(".chat-log .msg.user, .chat-log .msg-ack, .chat-log .msg.assistant")].map((el) => el.className))
    assert.deepEqual(order, ["msg user", "msg-ack", "msg assistant"])
    if (shots) await page.screenshot({ path: path.join(shots, "ack-2-quick-draft.png") })
    await page.getByRole("button", { name: "Run full drafting cycle", exact: true }).click()
    await until(() => starts.length === before.starts + 1, "the upgrade starts the full run")
    assert.deepEqual(starts.at(-1), { request, agent: "draft", model: { providerID: "test", modelID: "model" }, workingDraft: QUICK_SHOWN })
    assert.ok(routerPrompts.slice(before.routes).every((prompt) => !prompt.includes(ACK)), "routing history carries no acknowledgement")
    console.log("PASS Auto shows the router's acknowledgement while choosing; it stays above the offer and quick draft and is never sent, quick text or working draft")
    await context.close()
  }
  {
    // An answer: the acknowledgement stays above it once the turn settles and
    // the new conversation reloads; Stop while Auto is still choosing removes
    // the request, its acknowledgement and sends nothing.
    const ACK = "Got it — checking which invoices are unpaid."
    config.router = `${ACK}\n${JSON.stringify({ route: "qa" })}`
    const before = counts()
    const { context, page, send } = await open()
    await send("Which invoices are unpaid?")
    await page.locator(".msg.assistant", { hasText: "Answered." }).waitFor()
    await page.getByRole("button", { name: "Edit message" }).waitFor()
    await page.locator(".msg-ack", { hasText: ACK }).waitFor()
    assert.deepEqual(prompts.at(-1), { session: prompts.at(-1)!.session, agent: "qa", text: "Which invoices are unpaid?" })
    console.log("PASS an answered request keeps its acknowledgement above the answer after it settles")

    let release = () => {}
    config.hold = new Promise<void>((resolve) => (release = resolve))
    config.router = `Got it — looking at the payment dates.\n${JSON.stringify({ route: "qa" })}`
    await send("When were they due?")
    await page.locator(".msg-ack", { hasText: "looking at the payment dates" }).waitFor()
    await page.getByRole("button", { name: "Stop answer", exact: true }).click()
    release()
    config.hold = undefined
    await page.locator(".msg.user", { hasText: "When were they due?" }).waitFor({ state: "detached" })
    assert.equal(await page.getByRole("textbox").inputValue(), "When were they due?")
    assert.equal(await page.locator(".msg-ack", { hasText: "looking at the payment dates" }).count(), 0)
    await settle()
    assert.equal(counts().prompts, before.prompts + 1, "a request stopped while Auto chose is never sent")
    assert.equal(await page.locator(".msg-error").count(), 0)
    config.router = scope()
    console.log("PASS Stop while Auto chooses removes the request and its acknowledgement and sends nothing")

    // The abandoned router call resolves after Stop: its route must not
    // reach the chip, which keeps what the last answered request routed to.
    const chipOf = async () => (await page.getByRole("button", { name: /^Auto/ }).first().innerText()).trim()
    const chipBefore = await chipOf()
    config.hold = new Promise<void>((resolve) => (release = resolve))
    config.router = scope()
    await send("Draft a short letter to the tenant.")
    await page.getByRole("button", { name: "Stop answer", exact: true }).click()
    release()
    config.hold = undefined
    await page.locator(".msg.user", { hasText: "Draft a short letter" }).waitFor({ state: "detached" })
    await settle()
    await new Promise((resolve) => setTimeout(resolve, 300))
    assert.equal(await chipOf(), chipBefore, "a route that resolves after Stop never changes the chip")
    assert.equal(counts().prompts, before.prompts + 1, "the stopped draft request is never sent")
    console.log("PASS a router call that resolves after Stop leaves the Auto chip unchanged")
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
    // Automatic: the firm's setting is the consent, so Auto starts the quick
    // draft with no question, under its banner and upgrade.
    const before = counts()
    config.offer = "auto"
    const request = "Draft a short covering email to the client on the invoices."
    const { context, page, send } = await open()
    await send(request)
    await until(() => prompts.length === before.prompts + 1, "Automatic sends the request to Quick draft")
    assert.equal(prompts.at(-1)!.agent, "quick-draft")
    assert.equal(prompts.at(-1)!.text, request)
    await page.getByRole("note").filter({ hasText: BANNER }).waitFor()
    await page.getByRole("button", { name: "Run full drafting cycle", exact: true }).waitFor()
    await settle()
    assert.equal(forms.length, before.forms, "no question under Automatic")
    assert.equal(starts.length, before.starts, "no full run under Automatic")
    if (shots) await page.screenshot({ path: path.join(shots, "6-automatic-quick-draft.png") })
    config.offer = "ask"
    console.log("PASS Automatic: an eligible request starts a quick draft with no question; the banner and upgrade show")
    await context.close()
  }
  {
    // Automatic leaves eligibility alone: legal effect still runs in full.
    const before = counts()
    config.offer = "auto"
    config.router = scope({ legal_effect: "yes" })
    const { context, send } = await open()
    await send("Draft a short notice terminating the supply agreement.")
    await until(() => starts.length === before.starts + 1, "the full run starts")
    await settle()
    assert.equal(forms.length, before.forms, "no question under Automatic")
    assert.equal(prompts.length, before.prompts, "no quick draft for a request with legal effect")
    config.router = scope()
    config.offer = "ask"
    console.log("PASS Automatic: a request with legal effect starts the full run with no question")
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
    // A firm advice legend is applied by code to an advice letter's quick draft.
    config.legend = LEGEND
    const before = counts()
    const { context, page, send, pick } = await open()
    await pick("Quick draft")
    await send("Draft a short email advising the client on the invoices.")
    await until(() => prompts.length === before.prompts + 1, "the explicit pick sends to Quick draft")
    const quick = page.locator(".msg.assistant", { hasText: "both invoices remain unpaid" })
    await quick.waitFor()
    assert.equal(await quick.locator(".quick-draft-legend").textContent(), LEGEND)
    assert.equal(await quick.getByText("Form: advice letter").count(), 0)
    config.legend = ""
    if (shots) await page.screenshot({ path: path.join(shots, "5-quick-legend.png") })
    console.log("PASS the firm's advice legend heads an advice letter's quick draft; the model's form line is not shown")
    await context.close()
  }
  {
    // A one-document matter can only need few sources, whatever the router said.
    const before = counts()
    config.router = scope({ sources: "many" })
    const { context, send, card } = await open()
    await send("Draft a short email to the client listing what the letter asks for.")
    await until(() => forms.length === before.forms + 1, "the offer is made on a one-document matter")
    await card.waitFor()
    await card.getByRole("button", { name: "Cancel", exact: true }).click()
    await card.waitFor({ state: "detached" })
    await settle()
    assert.equal(starts.length, before.starts, "no full run without the lawyer's choice")
    config.router = scope()
    console.log("PASS a one-document matter is offered the quick draft even when the router judged many sources")
    await context.close()
  }
  {
    // Auto cannot choose (the router never gives a valid route): the lawyer is
    // told, nothing runs, and the request is back in the composer to resend.
    const before = counts()
    // Its acknowledgement streams, but the route is off the list.
    config.router = `On it — I'll draft that note.\n${JSON.stringify({ route: "drafting" })}`
    let release = () => {}
    config.hold = new Promise<void>((resolve) => (release = resolve))
    const request = "Draft a short note to the client about the invoices."
    const { context, page, send } = await open()
    await send(request)
    await page.locator(".msg-ack", { hasText: "I'll draft that note." }).waitFor()
    release()
    config.hold = undefined
    await page.getByText("Auto could not choose an assistant.", { exact: false }).waitFor()
    assert.equal(await page.getByRole("textbox").inputValue(), request)
    assert.equal(await page.locator(".msg-ack").count(), 0, "nothing was handed over, so no acknowledgement stays")
    await settle()
    assert.deepEqual(counts(), before, "nothing runs when Auto cannot choose")
    config.router = scope()
    console.log("PASS when Auto cannot choose, the lawyer is told, nothing runs and the request returns to the composer")
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
    assert.equal(await page.locator(".msg-ack").count(), 0, "an explicit pick shows no acknowledgement")
    if (shots) await page.screenshot({ path: path.join(shots, "4-quick-limit.png") })
    await notice.getByRole("button", { name: "Run full drafting cycle", exact: true }).click()
    await until(() => starts.length === before.starts + 1, "Full starts from the notice")
    assert.equal(starts.at(-1)!.request, request)
    assert.equal(starts.at(-1)!.workingDraft, undefined)
    config.router = scope()
    console.log("PASS an explicit pick outside Quick's scope explains the limit, runs nothing, and offers Full")
    await context.close()
  }
  {
    // Client summary: an explicit pick sent straight to its agent, never routed,
    // carrying the firm's advice legend and none of its narration.
    assert.ok(routerPrompts.length > 0 && routerPrompts.every((prompt) => !prompt.includes("client-summary")), "Auto never offers Client summary as a candidate")
    const before = { ...counts(), routes: routerPrompts.length }
    config.legend = LEGEND
    const request = "Explain the letter for the client in plain English."
    const { context, page, send } = await open()
    await page.getByTitle("Choose assistant or workflow").click()
    const option = page.locator(".assistant-option", { hasText: "Client summary" })
    await option.waitFor()
    assert.match(await option.innerText(), /plain English for your client/)
    await option.click()
    await send(request)
    await until(() => prompts.length === before.prompts + 1, "the request is sent")
    assert.equal(prompts.at(-1)!.agent, "client-summary")
    assert.equal(prompts.at(-1)!.text, request)
    const answer = page.locator(".msg.assistant", { hasText: "A letter about two unpaid invoices" })
    await answer.waitFor()
    await settle()
    assert.equal(routerPrompts.length, before.routes, "an explicit pick is not routed")
    assert.equal(await page.locator(".msg-ack").count(), 0, "an explicit pick shows no acknowledgement")
    assert.equal(starts.length, before.starts)
    assert.equal(forms.length, before.forms)
    assert.equal(await answer.locator(".quick-draft-legend").innerText(), LEGEND)
    assert.ok(!(await answer.innerText()).includes(SUMMARY_NARRATION), "narration is not part of the summary")
    config.legend = ""
    console.log("PASS Client summary is a picker entry sent to its own agent, never routed, with the firm's legend and no narration")
    await context.close()
  }
  {
    // Settings that cannot be read leave the summary without its legend, and say so.
    config.preferencesDown = true
    const { context, page, send, pick } = await open()
    await pick("Client summary")
    await send("Explain the letter for the client in plain English.")
    const answer = page.locator(".msg.assistant", { hasText: "A letter about two unpaid invoices" })
    await answer.waitFor()
    await answer.getByRole("note").filter({ hasText: "Firm legend could not be loaded." }).waitFor()
    assert.equal(await answer.locator(".quick-draft-legend").count(), 0)
    config.preferencesDown = false
    console.log("PASS a client summary whose firm legend could not be loaded says so")
    await context.close()
  }
  {
    // Asking about a review answer seats the question in an empty composer of
    // the open chat, and keeps text the lawyer had not sent.
    const before = { ...counts(), routes: routerPrompts.length }
    const { context, page } = await open()
    const composer = page.getByRole("textbox")
    const ask = async () => {
      await page.getByRole("link", { name: "Open review table" }).click()
      const table = page.getByRole("dialog", { name: /^Review table/ })
      await table.getByRole("button", { name: /^Both invoices\./ }).click()
      await page.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
      await table.waitFor({ state: "detached" })
    }
    const unsent = "Check whether the second invoice was disputed."
    await composer.fill(unsent)
    await ask()
    await page.getByText("Your unsent message was kept", { exact: false }).waitFor()
    assert.equal(await composer.inputValue(), unsent, "unsent composer text survives Ask about this answer")
    await composer.fill("")
    await ask()
    await page.waitForFunction(() => (document.querySelector("textarea") as HTMLTextAreaElement | null)?.value.includes("Question: Which invoices are unpaid?"))
    assert.match(await composer.inputValue(), /Answer: Both invoices\.\n.*Cited source: Letter\.docx § 1/s)
    await settle()
    assert.deepEqual({ ...counts(), routes: routerPrompts.length }, before, "the question is prefilled, never sent")
    console.log("PASS asking about a review answer keeps unsent composer text, and seats the question in an empty composer without sending")
    await context.close()
  }
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
