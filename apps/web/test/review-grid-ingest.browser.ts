// Run from apps/web: PLAYWRIGHT_CHROMIUM_PATH=<chromium> bun test/review-grid-ingest.browser.ts
// Real ReviewGrid, real isolated ingest, two browser tabs; no model calls.
import assert from "node:assert/strict"
import { mkdtemp, mkdir, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-grid-browser-ingest-"))
const config = path.join(root, "config")
await mkdir(config)
await Bun.write(path.join(config, "opencode.json"), "{}")
const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => Response.json({}) })
const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
const port = probe.port!
probe.stop(true)
const base = `http://127.0.0.1:${port}`
const request = (endpoint: string, input?: unknown, method = "GET") => fetch(`${base}${endpoint}`, {
  method,
  headers: { "Content-Type": "application/json" },
  ...(input === undefined ? {} : { body: JSON.stringify(input) }),
  signal: AbortSignal.timeout(15000),
})
const child = Bun.spawn([process.execPath, "run", "src/server.ts"], {
  cwd: path.resolve(import.meta.dir, "../../../services/ingest"),
  env: {
    ...process.env, DOCHAUS_DATA_ROOT: root, WORKSPACE_ROOT: undefined, DOCHAUS_RUNTIME_DIR: undefined,
    DOCHAUS_DIR: config, INGEST_HOST: "127.0.0.1", INGEST_PORT: String(port),
    OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "",
  },
  stdout: Bun.file(path.join(root, "ingest.log")), stderr: Bun.file(path.join(root, "ingest.log")),
})
let harness: ReturnType<typeof Bun.serve> | undefined
let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined
try {
  const deadline = Date.now() + 20000
  while (!(await request("/matters").catch(() => undefined))?.ok) {
    if (Date.now() > deadline || child.exitCode !== null) throw new Error("Isolated ingest did not start")
    await Bun.sleep(50)
  }
  const created = await request("/matters", { title: "Synthetic review" }, "POST")
  assert.equal(created.status, 200)
  const matter = await created.json() as { id: string; dir: string }
  const endpoint = `/matters/${matter.id}/grid`
  const seed = await request(endpoint, { version: 0, columns: [{ id: "q1", question: "Original question?" }], cells: {} }, "PUT")
  assert.equal(seed.status, 200)
  const expectedInitial = await seed.json()
  const diskPath = path.join(matter.dir, ".dochaus", "grid.json")
  const originalBytes = await readFile(diskPath)
  const entry = path.join(root, "harness.tsx")
  await Bun.write(entry, `
import { createElement, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ReviewGrid from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/ReviewGrid.tsx"))}
function Harness() {
  const [open, setOpen] = useState(true)
  return open ? createElement(ReviewGrid, {
    matterId: ${JSON.stringify(matter.id)}, title: "Synthetic review", directory: ${JSON.stringify(matter.dir)},
    documents: [], onClose: () => setOpen(false),
  }) : null
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
  const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: base, VITE_OPENCODE_URL: `http://127.0.0.1:${engine.port}` }) } })
  assert.equal(build.success, true, build.logs.join("\n"))
  const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
  harness = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
    ? new Response(script, { headers: { "content-type": "text/javascript" } })
    : new Response('<html><div id="root"></div><script type="module" src="/harness.js"></script></html>', { headers: { "content-type": "text/html" } }) })
  browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
  const context = await browser.newContext()
  const first = await context.newPage()
  const second = await context.newPage()
  first.setDefaultTimeout(10000)
  second.setDefaultTimeout(10000)
  const puts: string[] = []
  for (const page of [first, second]) page.on("request", req => { if (req.url() === base + endpoint && req.method() === "PUT") puts.push(req.url()) })
  await first.goto(harness.url.href)
  await first.getByRole("button", { name: "Edit question: Original question?" }).waitFor()
  await second.goto(harness.url.href)
  await second.getByRole("button", { name: "Edit question: Original question?" }).waitFor()
  await first.getByText("This review table is open in another window.").waitFor()
  await second.getByText("This review table is open in another window.").waitFor()
  assert.deepEqual(puts, [], "opening both tables must not issue a save")
  assert.deepEqual(await request(endpoint).then(res => res.json()), expectedInitial)
  assert.deepEqual(await readFile(diskPath), originalBytes, "opening must not modify the file")
  console.log("PASS two windows display the advisory notice and opening is read-only through ingest")

  first.once("dialog", dialog => dialog.accept("First window's question?"))
  const firstSave = first.waitForResponse(response => response.url() === base + endpoint && response.request().method() === "PUT")
  await first.getByRole("button", { name: "Edit question: Original question?" }).click()
  assert.equal((await firstSave).status(), 200)
  await first.getByRole("button", { name: "Edit question: First window's question?" }).waitFor()
  second.once("dialog", dialog => dialog.accept("Second window's question?"))
  const staleSave = second.waitForResponse(response => response.url() === base + endpoint && response.request().method() === "PUT")
  await second.getByRole("button", { name: "Edit question: Original question?" }).click()
  assert.equal((await staleSave).status(), 409)
  await second.getByText("Your changes are still shown here and have not been saved.", { exact: false }).waitFor()
  const pending = second.getByRole("button", { name: "Edit question: Second window's question?" })
  assert.equal(await pending.isVisible(), true)
  assert.equal(await pending.isDisabled(), true)
  const stored = await request(endpoint).then(res => res.json()) as { version: number; columns: { question: string }[] }
  assert.equal(stored.version, 2)
  assert.equal(stored.columns[0]!.question, "First window's question?")
  assert.equal(JSON.parse(await readFile(diskPath, "utf8")).columns[0].question, "First window's question?")
  console.log("PASS first save persists and stale second save keeps its local edit visible without overwriting ingest")
  await second.close({ runBeforeUnload: false })
  await first.getByText("This review table is open in another window.").waitFor({ state: "detached" })
  console.log("PASS the notice clears when the other window closes")
  await context.close()
} finally {
  await browser?.close()
  harness?.stop(true)
  child.kill("SIGTERM")
  await child.exited
  engine.stop(true)
  await rm(root, { recursive: true, force: true })
}
