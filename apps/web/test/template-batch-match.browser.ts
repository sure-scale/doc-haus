// Real TemplateBatchDialog in Chromium with isolated in-memory ingest and a
// mocked engine: dropping a spreadsheet whose headings differ from the
// template's fields, then confirming the suggested matches.
// Run from apps/web: bun test/template-batch-match.browser.ts
import assert from "node:assert/strict"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Route } from "@playwright/test"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-template-batch-"))
const evidence = path.join(process.env.DOCHAUS_EVIDENCE_DIR ?? os.tmpdir(), "template-batch-match")
await mkdir(evidence, { recursive: true })
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import TemplateBatchDialog from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/TemplateBatchDialog.tsx"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
const template = {
  name: "Synthetic services letter",
  description: "Fixture template",
  placeholders: [],
}
createRoot(document.getElementById("root")).render(
  createElement(MemoryRouter, null, createElement(TemplateBatchDialog, { template, matter: "one", onClose: () => {} })),
)
`)
const build = await Bun.build({
  entrypoints: [entry],
  target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://ingest.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) },
})
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const css = await build.outputs.find((output) => output.path.endsWith(".css"))?.text() ?? ""
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: (request) => {
  const pathname = new URL(request.url).pathname
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/harness.css") return new Response(css, { headers: { "content-type": "text/css" } })
  return new Response('<html><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>', { headers: { "content-type": "text/html" } })
} })

type Field = { key: string; label: string; kind: "fill" | "optional" }
type Sheet = { fields: Field[]; columns: string[]; mapping: Record<string, string | null>; rows: { row: number; values: string[] }[] }
const fields: Field[] = [
  { key: "[Client Name]", label: "Client Name", kind: "fill" },
  { key: "[Contract Title]", label: "Contract Title", kind: "fill" },
  { key: "[Notice Date]", label: "Notice date", kind: "fill" },
  { key: "[optional: Non-solicit]", label: "Include: Non-solicit", kind: "optional" },
]
const columns = ["Recipient", "Agreement", "Notice", "Include: Non-solicit"]
const rows = [
  { row: 2, values: ["Example Holdings Ltd", "Synthetic services agreement", "2026-12-01", "yes"] },
  { row: 3, values: ["Second Client Ltd", "Synthetic support agreement", "2027-01-15", "no"] },
  { row: 4, values: ["Third Client Ltd", "Synthetic goods agreement", "2027-02-20", "yes"] },
  { row: 5, values: ["Fourth Client Ltd", "Synthetic works agreement", "2027-03-25", "no"] },
]
// Field by field, what the rules left open.
const unmatchedMapping = (): Record<string, string | null> => ({
  "[Client Name]": null, "[Contract Title]": null, "[Notice Date]": null, "[optional: Non-solicit]": "Include: Non-solicit",
})
let sheet: Sheet = { fields, columns, rows, mapping: unmatchedMapping() }
let config: { model?: string; small_model?: string } = { model: "fixture/primary", small_model: "fixture/fast" }

// Engine: each matching session, its prompt, whether it was interrupted and deleted.
type Session = { id: string; agent?: string; model?: unknown; prompt?: string; interrupted: boolean; deleted: boolean }
const sessions: Session[] = []
const replies: string[] = []
let holdWait = false
const heldWaits: Route[] = []
const otherEngineCalls: string[] = []
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 860 } })
  page.setDefaultTimeout(8000)
  await page.route("http://ingest.invalid/**", (route) => {
    const request = route.request()
    const url = new URL(request.url())
    const reply = (body: unknown, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) })
    if (url.pathname === "/matters/one" && request.method() === "GET")
      return reply({ id: "one", title: "First matter", dir: "/synthetic/one", documents: [] })
    if (url.pathname === "/templates/batch/read" && request.method() === "POST")
      return reply({
        template: { name: "Synthetic services letter", sha256: "fixture" },
        fields: sheet.fields,
        unnamedBlanks: 0,
        columns: sheet.columns,
        rows: sheet.rows,
        dropped: { blank: 0, header: 0 },
        mapping: sheet.mapping,
        remembered: false,
        identifyingKey: "[Client Name]",
        source: "synthetic.csv",
      })
    return reply({ error: `Unexpected ${request.method()} ${url.pathname}` }, 404)
  })
  await page.route("http://ingest.invalid/engine/config", (route) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(config) }))
  await page.route("http://engine.invalid/**", (route) => {
    const request = route.request()
    const url = new URL(request.url())
    const key = `${request.method()} ${url.pathname}`
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (key === "POST /api/session") {
      const body = request.postDataJSON() as { id: string; agent?: string; model?: unknown }
      sessions.push({ id: body.id, agent: body.agent, model: body.model, interrupted: false, deleted: false })
      return json({ data: { id: body.id, location: { directory: "/synthetic/one" } } })
    }
    const session = [...sessions].reverse().find((item) => url.pathname.includes(`/session/${item.id}`))
    if (session) {
      if (key.endsWith("/prompt")) {
        session.prompt = (request.postDataJSON() as { text: string }).text
        return json({ data: { id: "msg_match", type: "user", sessionID: session.id, time: { created: 1 }, payload: { text: "match" }, delivery: "queue" } })
      }
      if (key.endsWith("/wait")) {
        if (holdWait) { heldWaits.push(route); return }
        return route.fulfill({ status: 204 })
      }
      if (key.endsWith("/interrupt")) { session.interrupted = true; return json({ interrupted: true }) }
      if (key.startsWith("GET") && key.endsWith("/message")) return json({ data: [
        { type: "assistant", id: "msg_reply", agent: "template-match", model: { providerID: "fixture", id: "fast" }, finish: "stop", time: { created: 1, completed: 2 }, content: [{ type: "text", text: replies.shift() ?? "" }] },
        { type: "idle", id: "msg_idle", outcome: "succeeded", time: { created: 3 } },
      ], cursor: {} })
      if (key === `DELETE /api/session/${session.id}`) { session.deleted = true; return route.fulfill({ status: 204 }) }
    }
    otherEngineCalls.push(key)
    return route.fulfill({ status: 404, body: `Unexpected ${key}` })
  })

  await page.goto(server.url.href)
  const file = page.locator('input[type="file"]')
  const select = (name: string) => page.getByRole("combobox", { name: new RegExp(`^${name}`) })
  const suggested = (name: string) => select(name).locator("xpath=preceding-sibling::label/span[contains(@class,'template-batch-suggested')]")
  const restart = page.getByRole("button", { name: "Use another spreadsheet" })
  const csv = new File([Buffer.from("Recipient,Agreement,Notice,Include: Non-solicit\r\nExample Holdings Ltd,Synthetic services agreement,2026-12-01,yes\r\n")], "synthetic.csv", { type: "text/csv" })

  // A sheet whose headings differ from the field names. The Match step shows
  // at once; the model's suggestion arrives after it.
  holdWait = true
  await file.setInputFiles({ name: "synthetic.csv", mimeType: "text/csv", buffer: Buffer.from(await csv.arrayBuffer()) })
  await expect(page.getByRole("heading", { name: "Match 3 fields" })).toBeVisible()
  await expect(page.getByRole("status")).toHaveText("Checking the headings for likely matches…")
  await expect.poll(() => sessions.length).toBe(1)
  assert.equal(sessions[0]!.agent, "template-match")
  assert.deepEqual(sessions[0]!.model, { providerID: "fixture", id: "fast" }, "matching runs on the fast model")
  const prompt = sessions[0]!.prompt!
  for (const text of ['"[Client Name]"', '"[Contract Title]"', '"[Notice Date]"', '["Recipient","Agreement","Notice"]', "Example Holdings Ltd"])
    assert.ok(prompt.includes(text), `the prompt carries ${text}`)
  assert.ok(!prompt.includes("Include: Non-solicit"), "a column a rule already used is not offered again")
  assert.ok(!prompt.includes("Fourth Client Ltd"), "only the first rows are sent")
  assert.equal(otherEngineCalls.length, 0, "the match job makes no other engine call")

  replies.push(JSON.stringify({ matches: { "[Client Name]": "Recipient", "[Contract Title]": "Agreement", "[Notice Date]": "Notice" } }))
  holdWait = false
  await expect.poll(() => heldWaits.length).toBe(1)
  await heldWaits.shift()!.fulfill({ status: 204 })
  await expect(select("Client Name")).toHaveValue("Recipient")
  await expect(select("Contract Title")).toHaveValue("Agreement")
  await expect(select("Notice date")).toHaveValue("Notice")
  await expect(page.locator(".template-batch-suggested")).toHaveCount(3)
  await expect(page.getByRole("status")).toHaveCount(0)
  await expect.poll(() => sessions[0]?.deleted).toBe(true)
  assert.equal(sessions[0]!.interrupted, false, "a finished call is not interrupted")
  await page.screenshot({ path: path.join(evidence, "suggested-matches.png"), fullPage: true })
  console.log("PASS the Match step shows first, the prompt carries the open fields and first rows, and the reply pre-selects the columns as suggested")

  // Changing a suggested field keeps the lawyer's choice and removes its mark.
  await select("Contract Title").selectOption("")
  await expect(suggested("Contract Title")).toHaveCount(0)
  await expect(page.locator(".template-batch-suggested")).toHaveCount(2)
  await page.screenshot({ path: path.join(evidence, "changed-match.png"), fullPage: true })
  console.log("PASS changing a suggested field clears its Suggested mark and keeps the chosen column")

  // An unusable reply changes nothing: no selection, no mark, no error.
  await restart.click()
  sheet = { fields, columns, rows, mapping: unmatchedMapping() }
  replies.push("These headings look similar.")
  await file.setInputFiles({ name: "synthetic.csv", mimeType: "text/csv", buffer: Buffer.from("Recipient,Agreement\r\nx,y\r\n") })
  await expect(page.getByRole("heading", { name: "Match 3 fields" })).toBeVisible()
  await expect.poll(() => sessions[1]?.deleted).toBe(true)
  for (const name of ["Client Name", "Contract Title", "Notice date"]) await expect(select(name)).toHaveValue("")
  await expect(page.locator(".template-batch-suggested")).toHaveCount(0)
  await expect(page.getByRole("alert")).toHaveCount(0)
  console.log("PASS a reply outside the contract leaves every field for the lawyer, with no error and the session deleted")

  // With no configured model the step is exactly the pre-existing flow.
  await restart.click()
  sheet = { fields, columns, rows, mapping: unmatchedMapping() }
  config = {}
  await file.setInputFiles({ name: "synthetic.csv", mimeType: "text/csv", buffer: Buffer.from("Recipient,Agreement\r\nx,y\r\n") })
  await expect(page.getByRole("heading", { name: "Match 3 fields" })).toBeVisible()
  await page.waitForTimeout(300)
  assert.equal(sessions.length, 2, "no model is called without a configured model")
  await expect(page.getByRole("status")).toHaveCount(0)
  await expect(page.locator(".template-batch-suggested")).toHaveCount(0)
  console.log("PASS without a configured model no match session is created and the Match step is unchanged")

  // A sheet every rule matches opens the preview and never calls a model.
  await restart.click()
  config = { model: "fixture/primary", small_model: "fixture/fast" }
  sheet = { fields, columns, rows, mapping: {
    "[Client Name]": "Recipient", "[Contract Title]": "Agreement", "[Notice Date]": "Notice", "[optional: Non-solicit]": "Include: Non-solicit",
  } }
  await file.setInputFiles({ name: "synthetic.csv", mimeType: "text/csv", buffer: Buffer.from("Recipient,Agreement\r\nx,y\r\n") })
  await expect(page.getByRole("heading", { name: /^Match \d+ field/ })).toHaveCount(0)
  await page.waitForTimeout(300)
  assert.equal(sessions.length, 2, "a fully matched sheet calls no model")
  console.log("PASS a sheet the rules already match skips the Match step and calls no model")

  // A field the lawyer changes while the suggestion is still running is never
  // overwritten, and gets no Suggested mark.
  await page.reload()
  sheet = { fields, columns, rows, mapping: unmatchedMapping() }
  holdWait = true
  await file.setInputFiles({ name: "synthetic.csv", mimeType: "text/csv", buffer: Buffer.from("Recipient,Agreement\r\nx,y\r\n") })
  await expect(page.getByRole("heading", { name: "Match 3 fields" })).toBeVisible()
  await expect.poll(() => sessions.length).toBe(3)
  await select("Client Name").selectOption("")
  replies.push(JSON.stringify({ matches: { "[Client Name]": "Recipient", "[Contract Title]": "Agreement", "[Notice Date]": "Notice" } }))
  holdWait = false
  await expect.poll(() => heldWaits.length).toBe(1)
  await heldWaits.shift()!.fulfill({ status: 204 })
  await expect(select("Contract Title")).toHaveValue("Agreement")
  await expect(select("Client Name")).toHaveValue("")
  await expect(suggested("Client Name")).toHaveCount(0)
  await expect(page.locator(".template-batch-suggested")).toHaveCount(2)
  console.log("PASS a field changed while the job runs keeps the lawyer's choice and is not marked as suggested")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
