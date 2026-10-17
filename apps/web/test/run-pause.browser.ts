// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/run-pause.browser.ts
// Uses the real header/action hook; only HTTP is mocked. No live run is touched.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-pause-ui-"))
const working = start.status
const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
const paused = {
  ...blockedGate,
  stop_reason: "Paused at your request.",
  paused_at: "2026-10-17T09:00:00.000Z",
  stages: { ...blockedGate.stages, "d1.apply": pending, finalize: pending },
}
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import RunHeader from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/drafting/RunHeader.tsx"))}
import useRunActions from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/drafting/useRunActions.ts"))}
function Harness() {
  const [status, setStatus] = useState(${JSON.stringify(working)})
  const actions = useRunActions("matter-pause", status, setStatus)
  return createElement(RunHeader, { status, actions, own: true, onOpen: () => {} })
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://pause.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs[0].text()
const css = await Bun.file(path.resolve(import.meta.dir, "../src/styles/drafting.css")).text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><style>body{font-family:sans-serif}${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const calls: { url: string; method: string; body: unknown }[] = []
  let release = () => {}
  await page.route("http://pause.invalid/**", async route => {
    calls.push({ url: route.request().url(), method: route.request().method(), body: route.request().postDataJSON() })
    await new Promise<void>(resolve => { release = resolve })
    return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ profile: "lean", mode: "paused", status: paused }) })
  })
  await page.goto(server.url.href)
  const pause = page.getByRole("button", { name: "Pause drafting", exact: true })
  await pause.waitFor({ state: "visible" })
  assert.equal(await page.getByRole("button", { name: "Continue drafting", exact: true }).count(), 0)
  await pause.click()
  // The button is held while the request is in flight; the run never reads as resuming.
  const pausing = page.getByRole("button", { name: "Pausing…", exact: true })
  await pausing.waitFor()
  assert.equal(await pausing.isDisabled(), true)
  assert.equal(await page.getByText("Resuming…").count(), 0)
  assert.equal(calls.length, 1)
  assert.ok(calls[0]!.url.endsWith("/matters/matter-pause/lean/pause"))
  assert.equal(calls[0]!.method, "POST")
  assert.deepEqual(calls[0]!.body, { ownerSessionID: working.owner_session_id })
  release()
  await page.getByRole("button", { name: "Continue drafting", exact: true }).waitFor()
  await page.getByText("You paused drafting", { exact: true }).waitFor()
  assert.equal(await pause.count(), 0)
  assert.equal(await page.getByText("Resuming…").count(), 0)
  assert.equal(await page.getByRole("alert").count(), 0)
  console.log("PASS Pause shows only while working, posts the owner, holds the button in flight and settles to a paused run with Continue")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
