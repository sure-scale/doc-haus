// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/checklist-refinement.browser.ts
// Uses the real header/action hook; only HTTP is mocked. No live run is touched.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Route } from "playwright"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-checklist-ui-"))
const initial = {
  ...blockedGate,
  stages: { ...blockedGate.stages, "checklist-check": { status: "blocked", attempts: 1, started_at: null, finished_at: null } },
  checklist_refinement: { available: true, parents: [
    { parent_id: "req-a", check: "State the agreed payment terms.", target: { document: "d1", role: "body", treatment: "operative" }, finding: { status: "unsatisfied", reason: "The payment date was not confirmed." }, scopes: ["target_body"] },
    { parent_id: "req-b", check: "Address each requested amendment in the draft.", target: { document: "d1", role: "body", treatment: "operative" }, finding: { status: "indeterminate", reason: "The earlier inspection did not cover the whole draft." }, scopes: ["target_body"] },
  ] },
}
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import RunHeader from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/drafting/RunHeader.tsx"))}
import useRunActions from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/drafting/useRunActions.ts"))}
function Harness() {
  const [status, setStatus] = useState(${JSON.stringify(initial)})
  const actions = useRunActions("matter-check", status, setStatus)
  window.runActions = actions
  return createElement(RunHeader, { status, actions, own: true, onOpen: () => {}, onShowDrafts: () => {} })
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://checklist.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs[0].text()
const css = await Bun.file(path.resolve(import.meta.dir, "../src/styles/drafting.css")).text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><style>body{font-family:sans-serif}.run-note{color:#555}${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const calls: { url: string; body: unknown }[] = []
  let release = () => {}
  let refuse = true
  await page.route("http://checklist.invalid/**", async (route: Route) => {
    calls.push({ url: route.request().url(), body: route.request().postDataJSON() })
    await new Promise<void>(resolve => { release = resolve })
    return route.fulfill({ status: refuse ? 409 : 202, contentType: "application/json", body: JSON.stringify(refuse
      ? { error: "The active drafting run changed. Reload before refining its checklist." }
      : { profile: "lean", mode: "queued", request_id: "accepted-check", status: { ...initial, pending_requests: 1, checklist_refinement: { available: false, parents: [] } } }) })
  })
  await page.goto(server.url.href)
  const check = page.getByRole("button", { name: "Check the full draft", exact: true })
  await check.waitFor({ state: "visible" })
  assert.equal(await check.isDisabled(), true)
  assert.equal(await page.getByRole("radio").count(), 2)
  assert.equal(await page.getByRole("radio").nth(0).isChecked(), false)
  await page.getByRole("radio", { name: /Address each requested amendment/ }).check()
  assert.equal(await check.isEnabled(), true)
  await check.click()
  await page.getByRole("button", { name: "Requesting check…", exact: true }).waitFor()
  assert.equal(await page.getByRole("radio").nth(1).isDisabled(), true)
  assert.equal(calls.length, 1)
  // A second synchronous call cannot bypass the button's pending state.
  assert.equal(await page.evaluate(() => (window as unknown as { runActions: { refine(id: string): Promise<boolean> } }).runActions.refine("req-b")), false)
  assert.equal(calls.length, 1)
  assert.deepEqual(calls[0]!.body, { ownerSessionID: initial.owner_session_id, parents: [{ parent_id: "req-b", scope: "target_body" }] })
  assert.ok(calls[0]!.url.endsWith("/matters/matter-check/lean/checklist/refine"))
  release()
  await page.getByRole("alert").filter({ hasText: "The active drafting run changed." }).waitFor()
  assert.equal(await page.getByRole("radio").nth(1).isChecked(), true)
  assert.equal(await check.isEnabled(), true)
  refuse = false
  await check.click()
  await page.getByRole("button", { name: "Requesting check…", exact: true }).waitFor()
  assert.equal(calls.length, 2)
  release()
  await check.waitFor({ state: "hidden" })
  await page.getByText("Request queued", { exact: true }).waitFor()
  assert.equal(await page.getByRole("alert").count(), 0)
  console.log("PASS explicit requirement selection, owner request, pending guard, native error/retry and accepted queued status")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
