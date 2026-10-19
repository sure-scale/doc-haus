// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/run-pause.browser.ts
// Uses the real header, action hook and composer; only HTTP is mocked. No live run is touched.
// Set RUN_PAUSE_SCREENSHOTS=<directory> to save the paused, confirming and stopped headers.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-pause-ui-"))
const shots = process.env.RUN_PAUSE_SCREENSHOTS
const working = start.status
const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
const paused = {
  ...blockedGate,
  stop_reason: "Paused at your request.",
  paused_at: "2026-10-17T09:00:00.000Z",
  stages: { ...blockedGate.stages, "d1.apply": pending, finalize: pending },
}
const { paused_at: _, ...unpaused } = paused
const stopped = { ...unpaused, status: "failed", stopped: true, stop_reason: "Stopped at your request.", revision: paused.revision + 1, updated_at: "2026-10-17T09:05:00.000Z" }
const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement, useRef, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import RunHeader from ${source("components/drafting/RunHeader.tsx")}
import useRunActions from ${source("components/drafting/useRunActions.ts")}
import Composer from ${source("components/chat/Composer.tsx")}
import { leanRunTakesUpdates } from ${source("components/drafting/model.ts")}
// The composer is wired as ChatPanel wires it: an open run takes updates, any
// other message is a new request. Sends are recorded, never posted.
function Harness() {
  const [status, setStatus] = useState(${JSON.stringify(working)})
  const [input, setInput] = useState("")
  const inputRef = useRef(null)
  const actions = useRunActions("matter-pause", status, setStatus)
  const updating = leanRunTakesUpdates(status)
  return createElement(MemoryRouter, null,
    createElement(RunHeader, { status, actions, own: true, onOpen: () => {}, onShowDrafts: () => {} }),
    createElement(Composer, {
      available: new Set(), agent: "auto", setResolvedAgent: () => {}, onAgentChange: () => {}, inputRef,
      workflows: [], runWorkflow: async () => {}, composerPlaceholder: "Ask about this matter", input, setInput,
      onSend: async () => { window.sent = [...(window.sent ?? []), { text: input, as: updating ? "update" : "request" }] },
      busy: false, draftingUpdate: updating,
    }))
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://pause.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs[0].text()
const css = (await Promise.all(["styles.css", "styles/design-system.css", "styles/drafting.css", "styles/design-conversation.css"]
  .map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><meta charset="utf-8"><style>body{font-family:sans-serif;padding:16px;max-width:760px}${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const calls: { url: string; method: string; body: unknown }[] = []
  let release = () => {}
  // The first Stop is refused as ingest refuses a changed run; the next is held, then settles.
  const stops = [{ status: 409, body: { error: "The active drafting run changed. Reload before stopping it.", code: "lean_run_changed" } }]
  await page.route("http://pause.invalid/**", async route => {
    calls.push({ url: route.request().url(), method: route.request().method(), body: route.request().postDataJSON() })
    const stop = route.request().url().endsWith("/lean/stop")
    const refusal = stop ? stops.shift() : undefined
    if (refusal) return route.fulfill({ status: refusal.status, contentType: "application/json", body: JSON.stringify(refusal.body) })
    await new Promise<void>(resolve => { release = resolve })
    const body = stop ? { profile: "lean", mode: "stopped", status: stopped } : { profile: "lean", mode: "paused", status: paused }
    return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
  })
  await page.goto(server.url.href)
  const pause = page.getByRole("button", { name: "Pause drafting", exact: true })
  await pause.waitFor({ state: "visible" })
  assert.equal(await page.getByRole("button", { name: "Continue drafting", exact: true }).count(), 0)
  assert.equal(await page.getByRole("button", { name: "Stop draft", exact: true }).count(), 0)
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
  const resume = page.getByRole("button", { name: "Continue drafting", exact: true })
  await resume.waitFor()
  await page.getByText("You paused drafting", { exact: true }).waitFor()
  assert.equal(await pause.count(), 0)
  assert.equal(await page.getByText("Resuming…").count(), 0)
  assert.equal(await page.getByRole("alert").count(), 0)
  console.log("PASS Pause shows only while working, posts the owner, holds the button in flight and settles to a paused run with Continue")

  // A paused run offers Stop beside Continue; the composer still updates this run.
  const stopDraft = page.getByRole("button", { name: "Stop draft", exact: true })
  await stopDraft.waitFor()
  await page.getByRole("button", { name: "Send drafting update", exact: true }).waitFor()
  if (shots) await page.screenshot({ path: path.join(shots, "1-paused.png"), fullPage: true })
  const confirm = page.getByRole("group", { name: /^Stop this draft\?/ })
  const keep = page.getByRole("button", { name: "Keep draft", exact: true })

  // Stop asks in place; Keep draft and Escape both withdraw the question and nothing is sent.
  await stopDraft.click()
  await confirm.waitFor()
  assert.equal(await keep.evaluate(button => button === document.activeElement), true)
  assert.equal(await resume.count(), 0)
  if (shots) await page.screenshot({ path: path.join(shots, "2-confirm.png"), fullPage: true })
  await keep.click()
  await confirm.waitFor({ state: "detached" })
  await resume.waitFor()
  assert.equal(await stopDraft.evaluate(button => button === document.activeElement), true)
  await stopDraft.click()
  await confirm.waitFor()
  await page.keyboard.press("Escape")
  await confirm.waitFor({ state: "detached" })
  assert.equal(calls.length, 1)

  // A refusal is shown and the question stays, so the lawyer can retry or keep the draft.
  await stopDraft.click()
  await confirm.getByRole("button", { name: "Stop draft", exact: true }).click()
  await page.getByRole("alert").filter({ hasText: "The active drafting run changed. Reload before stopping it." }).waitFor()
  await confirm.waitFor()
  assert.equal(calls.length, 2)

  // Confirmed: the owner is posted to the stop route and both choices are held in flight.
  await confirm.getByRole("button", { name: "Stop draft", exact: true }).click()
  const stopping = page.getByRole("button", { name: "Stopping…", exact: true })
  await stopping.waitFor()
  assert.equal(await stopping.isDisabled(), true)
  assert.equal(await keep.isDisabled(), true)
  assert.equal(await page.getByText("Resuming…").count(), 0)
  assert.equal(calls.length, 3)
  for (const call of calls.slice(1)) {
    assert.ok(call.url.endsWith("/matters/matter-pause/lean/stop"))
    assert.equal(call.method, "POST")
    assert.deepEqual(call.body, { ownerSessionID: working.owner_session_id })
  }
  release()

  // Stopped: the header says so and offers nothing to continue; the composer sends a new request.
  await page.getByRole("heading", { name: "Draft stopped", exact: true }).waitFor()
  await page.getByText("You stopped this draft", { exact: true }).waitFor()
  await confirm.waitFor({ state: "detached" })
  for (const name of ["Continue drafting", "Stop draft", "Pause drafting"]) assert.equal(await page.getByRole("button", { name, exact: true }).count(), 0)
  assert.equal(await page.getByRole("alert").count(), 0)
  assert.equal(await page.getByRole("button", { name: "Send drafting update", exact: true }).count(), 0)
  if (shots) await page.screenshot({ path: path.join(shots, "3-stopped.png"), fullPage: true })
  await page.getByRole("textbox", { name: "Message" }).fill("Draft a second supply note.")
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  assert.deepEqual(await page.evaluate(() => (window as unknown as { sent?: unknown[] }).sent), [{ text: "Draft a second supply note.", as: "request" }])
  assert.equal(calls.length, 3)
  console.log("PASS Stop asks in place, keeps a refused stop retryable, posts the owner once confirmed and settles to a stopped run whose composer starts a new request")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
