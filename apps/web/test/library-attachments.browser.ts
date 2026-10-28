// Run from apps/web: PLAYWRIGHT_CHROMIUM_PATH=<chrome> bun test/library-attachments.browser.ts
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
const { chromium } = await import("@playwright/test")
const temp = await mkdtemp(path.join(os.tmpdir(), "dochaus-library-chat-"))
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
const entry = path.join(temp, "harness.tsx")
await Bun.write(entry, `
import { createElement as h } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter, Route, Routes } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import { ToastProvider } from ${source("components/Toast.tsx")}
import Playbooks from ${source("pages/Playbooks.tsx")}
import Templates from ${source("pages/Templates.tsx")}
import Skills from ${source("pages/Skills.tsx")}
import Workflows from ${source("pages/Workflows.tsx")}
import Agents from ${source("pages/Agents.tsx")}
const name = new URLSearchParams(location.search).get("page")
createRoot(document.getElementById("root")).render(h(ToastProvider, null, h(MemoryRouter, {initialEntries: ['/' + name]}, h(Routes, null, h(Route, {path: '/playbooks', element: h(Playbooks)}), h(Route, {path: '/templates', element: h(Templates)}), h(Route, {path: '/skills', element: h(Skills)}), h(Route, {path: '/workflows', element: h(Workflows)}), h(Route, {path: '/agents', element: h(Agents)})))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.path.endsWith(".js"))!.text()
const json = (body: unknown) => Response.json(body)
const prompts: { text: string; directory: string | null }[] = []
let failNext = false
let removals = 0
let failure: "dispatch" | "settle" | undefined
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", idleTimeout: 0, async fetch(request) {
  const route = `${request.method} ${new URL(request.url).pathname}`
  if (route === "GET /") return new Response('<div id="root"></div><script type="module" src="/harness.js"></script>', { headers: { "content-type": "text/html" } })
  if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (route === "GET /ingest/playbooks/library") return json({ dir: "/fixture/playbooks", playbooks: [] })
  if (route === "GET /ingest/templates") return json({ dir: "/fixture/templates", templates: [] })
  if (route === "GET /ingest/template-drafts") return json({ drafts: [] })
  if (route === "GET /ingest/skills") return json({ dir: "/fixture/skills", skills: [] })
  if (route === "GET /ingest/workflows") return json({ dir: "/fixture/workflows", workflows: [] })
  if (route === "GET /ingest/agents") return json({ dir: "/fixture/agents", agents: [] })
  if (route.startsWith("POST /ingest/library/")) { const file = (await request.formData()).get("file") as File; return json({ id: "fixture-id", name: file.name, source: `.chat-attachments/${file.name}.txt`, original: `.chat-attachments/${file.name}` }) }
  if (route.startsWith("DELETE /ingest/library/")) { removals++; return new Response(null, { status: 204 }) }
  if (route === "GET /ingest/engine/config") { if (failNext) { failNext = false; return new Response("unavailable", { status: 503 }) }; return json({ model: "test/model", small_model: "test/model" }) }
  if (route === "GET /api/event") return new Response(new ReadableStream({ start() {} }), { headers: { "content-type": "text/event-stream" } })
  if (route === "GET /api/session/active") return json({ data: {} })
  if (route === "GET /api/permission/request") return json({ data: [] })
  if (route === "POST /api/session") return json({ data: { id: "ses_test", location: { directory: request.headers.get("x-opencode-directory") } } })
  if (/^POST \/api\/session\/ses_test\/(agent|model)$/.test(route)) return new Response(null, { status: 204 })
  if (route === "POST /api/session/ses_test/prompt") { const { text } = await request.json() as { text: string }; prompts.push({ text, directory: request.headers.get("x-opencode-directory") }); if (failure === "dispatch") return new Response("lost", { status: 503 }); return json({ data: { id: "msg_1", sessionID: "ses_test", time: { created: 1 }, type: "user", payload: { text }, delivery: "queue" } }) }
  if (route === "POST /api/experimental/session/ses_test/wait") return failure === "settle" ? new Response("unavailable", { status: 503 }) : new Response(null, { status: 204 })
  if (route === "GET /api/session/ses_test/message") return json({ data: [{ type: "idle", id: "msg_i", outcome: "succeeded", time: { created: 1 } }], cursor: {} })
  if (route === "GET /api/session/ses_test/form") return json({ data: [] })
  if (route === "GET /api/session/ses_test") return json({ data: { id: "ses_test", location: { directory: request.headers.get("x-opencode-directory") } } })
  return new Response(null, { status: 404 })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  for (const [index, kind] of ["playbooks", "templates", "skills", "workflows", "agents"].entries()) {
    const page = await browser.newPage()
    page.setDefaultTimeout(8000)
    await page.goto(`${server.url.href}?page=${kind}`)
    await page.getByRole("button", { name: "Attach", exact: true }).waitFor()
    if (kind === "playbooks") {
      assert.equal(await page.getByRole("button", { name: "Create playbook", exact: true }).count(), 0)
      await page.locator('input[type=file][accept=".pdf,.docx,.txt,.md,.markdown"]').setInputFiles({ name: "discard.txt", mimeType: "text/plain", buffer: Buffer.from("discard") })
      await page.getByRole("button", { name: "Remove discard.txt" }).click()
      await page.getByRole("button", { name: "Remove discard.txt" }).waitFor({ state: "detached" })
      assert.equal(removals, 1)
    }
    await page.locator('input[type=file][accept=".pdf,.docx,.txt,.md,.markdown"]').setInputFiles({ name: `${kind}.txt`, mimeType: "text/plain", buffer: Buffer.from("firm source") })
    await page.getByRole("button", { name: `Remove ${kind}.txt` }).waitFor()
    await page.reload()
    await page.getByRole("button", { name: `Remove ${kind}.txt` }).waitFor()
    await page.getByRole("textbox", { name: "Message" }).fill("Use this guide")
    if (kind === "playbooks") failNext = true
    await page.getByRole("textbox", { name: "Message" }).press("Enter")
    if (kind === "playbooks") {
      await page.getByRole("button", { name: "Send message" }).waitFor({ state: "visible" })
      await page.waitForTimeout(300)
      assert.equal(await page.getByRole("textbox", { name: "Message" }).inputValue(), "Use this guide")
      await page.getByRole("button", { name: `Remove ${kind}.txt` }).waitFor()
      await page.getByRole("button", { name: "Send message" }).click()
    }
    const deadline = Date.now() + 8000
    while (prompts.length < index + 1 && Date.now() < deadline) await Bun.sleep(20)
    assert.equal(prompts.length, index + 1, `prompt sent for ${kind}`)
    const sent = prompts.at(-1)!
    assert.equal(sent.directory, `/fixture/${kind}`)
    assert.ok(sent.text.includes(`.chat-attachments/${kind}.txt.txt`))
    assert.ok(sent.text.includes("Use this guide"))
    await page.close()
    console.log(`PASS ${kind} attachment is sent in its library session`)
  }
  for (const mode of ["settle", "dispatch"] as const) {
    const page = await browser.newPage()
    page.setDefaultTimeout(8000)
    await page.goto(`${server.url.href}?page=playbooks`)
    await page.getByRole("button", { name: "Attach", exact: true }).waitFor()
    await page.locator('input[type=file][accept=".pdf,.docx,.txt,.md,.markdown"]').setInputFiles({ name: `${mode}.txt`, mimeType: "text/plain", buffer: Buffer.from("source") })
    const chip = page.getByRole("button", { name: `Remove ${mode}.txt` })
    await chip.waitFor()
    failure = mode
    await page.getByRole("textbox", { name: "Message" }).fill("Use this source")
    await page.getByRole("textbox", { name: "Message" }).press("Enter")
    if (mode === "settle") {
      await chip.waitFor({ state: "detached" })
      await page.getByRole("button", { name: "Send message" }).waitFor({ state: "visible" })
    } else {
      await page.getByRole("button", { name: "Send message" }).waitFor({ state: "visible" })
      await chip.click()
      await chip.waitFor({ state: "detached" })
    }
    assert.equal(removals, 1, `${mode}: dispatched source remains available`)
    failure = undefined
    await page.close()
    console.log(`PASS ${mode} failure retains dispatched source`)
  }
} finally { await browser.close(); server.stop(true); await rm(temp, { recursive: true, force: true }) }
