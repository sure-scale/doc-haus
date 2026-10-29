// Real tool metadata -> folded StepsPanel -> application-owned confirmation card.
// Local scripted ingest only; no model or engine requests.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { toTurns } from "../src/components/chat/turn-model"
import type { Chronology, ChronologyEvent } from "../src/api/ingest/chronology"
import type { Part } from "../src/api/session-view"
const { chromium, expect } = (await import(
  process.env.PLAYWRIGHT_MODULE ?? "@playwright/test"
)) as typeof import("@playwright/test")
function steps(metadata: Record<string, unknown>, tool = "propose-chronology-changes") {
  const state: Extract<Part, { type: "tool" }>["state"] = {
    status: "completed",
    input: {},
    title: "Chronology proposal staged",
    output: "A correction is staged for the lawyer. Saved events are unchanged.",
    metadata,
  }
  return toTurns(
    [
      {
        info: { id: "message", role: "assistant" },
        parts: [{ type: "tool", id: "call", sessionID: "chat", messageID: "message", tool, state }],
      },
    ],
    "/fixture",
  )[0]!.steps
}
const metadata = {
  chronologyMatterId: "fixture",
  chronologyProposalIds: ["date-proposal"],
  chronologyRevision: 4,
  chronologyLink: "/matter/fixture?view=chronology&proposal=date-proposal",
}
const projected = steps(metadata)[0]!
assert.equal(projected.kind, "tool")
assert.deepEqual(
  projected.kind === "tool" ? projected.chronologyResult : undefined,
  { matterId: "fixture", proposalIds: ["date-proposal"] },
  "actual completed tool metadata yields confirmation identities",
)
assert.equal(
  steps({ ...metadata, chronologyProposalIds: [3] })[0]?.kind === "tool" &&
    "chronologyResult" in steps({ ...metadata, chronologyProposalIds: [3] })[0]!,
  false,
  "malformed proposal IDs cannot create a confirmation card",
)
assert.equal(
  steps(metadata, "read-chronology")[0]?.kind === "tool" &&
    "chronologyResult" in steps(metadata, "read-chronology")[0]!,
  false,
  "read tool metadata cannot create a confirmation card",
)
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chronology-chat-"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter, Route, Routes } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import StepsPanel from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/chat/StepsPanel.tsx"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
const steps = ${JSON.stringify(steps(metadata))}
if (location.search.includes('cross')) steps[0].chronologyResult.matterId = 'another-matter'
createRoot(document.getElementById('root')).render(createElement(MemoryRouter, { initialEntries: ['/matter/fixture'] }, createElement(Routes, null, createElement(Route, { path: '/matter/:id', element: createElement(StepsPanel, { steps, busy: false, answered: true }) }))))
`,
)
const build = await Bun.build({
  entrypoints: [entry],
  target: "browser",
  define: {
    "import.meta.env": JSON.stringify({
      VITE_INGEST_URL: "http://chronology.invalid",
      VITE_OPENCODE_URL: "http://engine.invalid",
    }),
  },
})
assert.equal(build.success, true, build.logs.join("\n"))
const js = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const css = (await build.outputs.find((output) => output.path.endsWith(".css"))?.text()) ?? ""
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch: (request) => {
    const name = new URL(request.url).pathname
    if (name === "/app.js") return new Response(js, { headers: { "content-type": "text/javascript" } })
    if (name === "/app.css") return new Response(css, { headers: { "content-type": "text/css" } })
    return new Response(
      '<html><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/app.css"><div id="root"></div><script type="module" src="/app.js"></script></html>',
      { headers: { "content-type": "text/html" } },
    )
  },
})
const current: ChronologyEvent = {
  id: "delivery",
  title: "Delivery date in saved account",
  detail: "The lawyer preserved this disputed delivery date.",
  date: { original: "14 June 2026", precision: "exact", start: "2026-06-14" },
  accounts: [
    {
      id: "account",
      title: "Delivery account",
      detail: "",
      date: { original: "14 June 2026", precision: "exact", start: "2026-06-14" },
      evidence: [
        {
          id: "evidence",
          sourceId: "source",
          documentName: "Dispatch record.md",
          version: "v2",
          excerpt: "The goods arrived at the reception desk on 16 June 2026.",
          section: "Arrival",
          charStart: 0,
          charEnd: 59,
          verified: true,
          freshness: "stale",
          provenance: "source",
        },
      ],
      privilege: "uncertain",
      significance: "relevant",
      people: [],
      status: "reported",
      gaps: [],
    },
  ],
  origin: "generated",
  reviewed: true,
  opened: true,
  hidden: true,
  key: true,
  disputed: true,
  dispute: "Delivery timing disputed.",
  note: "Keep this lawyer note.",
  lawyerEdited: true,
  createdAt: 1,
  updatedAt: 1,
  previousIds: [],
}
const proposed = {
  ...structuredClone(current),
  title: "Arrival reported on 16 June",
  detail: "The revised dispatch record reports arrival two days later.",
  date: { original: "16 June 2026", precision: "exact" as const, start: "2026-06-16" },
  reviewed: false,
}
let state: Chronology = {
  schemaVersion: 1,
  revision: 4,
  created: true,
  events: [current],
  coverage: [],
  sourceDates: {},
  suggestion: { status: "relevant", reason: "Delivery sequence" },
  framing: "Neutral",
  gaps: [],
  history: [],
  proposals: [
    {
      id: "date-proposal",
      kind: "amendment",
      targetIds: ["delivery"],
      baseEventHashes: {},
      events: [proposed],
      reason: "A new source version reports another arrival date.",
      recommendation: "Keep competing accounts until the lawyer resolves them.",
      status: "pending",
      origin: "chat",
      createdAt: 2,
    },
  ],
}
let reads = 0,
  decisions = 0,
  stale = true
const browser = await chromium.launch({
  headless: true,
  ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}),
})
try {
  const page = await browser.newPage({ viewport: { width: 760, height: 900 } })
  page.setDefaultTimeout(8000)
  await page.route("http://chronology.invalid/**", (route) => {
    const request = route.request(),
      url = new URL(request.url())
    const reply = (value: unknown, status = 200) =>
      route.fulfill({ status, contentType: "application/json", body: JSON.stringify(value) })
    if (request.method() === "GET" && url.pathname === "/matters/fixture/chronology") {
      reads++
      return reply(state)
    }
    if (
      request.method() === "POST" &&
      url.pathname === "/matters/fixture/chronology/proposals/date-proposal/decision"
    ) {
      decisions++
      const input = request.postDataJSON() as { expectedRevision: number; decision: string }
      if (stale) {
        stale = false
        state.revision++
        state.events[0]!.detail = "A concurrent lawyer correction is now the current saved entry."
        return reply({ error: "Stale confirmation" }, 409)
      }
      assert.equal(input.expectedRevision, state.revision)
      assert.equal(input.decision, "accept")
      state.events = state.proposals[0]!.events.map((event) => ({ ...event, reviewed: true }))
      state.proposals[0]!.status = "accepted"
      state.revision++
      return reply(state)
    }
    throw new Error(`Unexpected request ${request.method()} ${url.pathname}`)
  })
  await page.goto(server.url.href)
  const card = page.getByRole("region", { name: "Chronology proposals awaiting your decision" })
  await expect(card).toBeVisible()
  await expect(card).toContainText("A new source version reports another arrival date.")
  assert.equal(await page.locator("details.steps").getAttribute("open"), null, "ordinary steps are folded")
  assert.equal(
    await card.evaluate((element) => !!element.closest("details")),
    false,
    "confirmation card stays outside folded tool steps",
  )
  assert.equal(decisions, 0, "displaying model output never applies it")
  await expect(card.getByRole("button", { name: "Accept this change" })).toHaveCount(0)
  await card.getByRole("button", { name: "Inspect exact change and sources" }).click()
  await expect(card).toContainText("14 June 2026")
  await expect(card).toContainText("16 June 2026")
  await expect(card).toContainText("The goods arrived at the reception desk on 16 June 2026.")
  await expect(card).toContainText("Hidden")
  await expect(card).toContainText("Lawyer corrections")
  await expect(card).toContainText("Quotation verified · stale")
  await card.getByRole("button", { name: "Accept this change" }).click()
  await expect(card.getByRole("alert")).toContainText("Review the reloaded entry before deciding again")
  await expect(card).toContainText("A concurrent lawyer correction")
  assert.equal(decisions, 1, "stale confirmation reloads without replay")
  assert.equal(state.events[0]!.title, "Delivery date in saved account")
  assert.equal(reads, 2)
  await card.getByRole("button", { name: "Accept this change" }).click()
  await expect(card).toContainText("Accepted by the local operator.")
  assert.equal(decisions, 2)
  assert.equal(state.events[0]!.reviewed, true)
  const previousReads = reads
  await page.goto(`${server.url.href}?cross=1`)
  await expect(card).toHaveCount(0)
  await page.waitForTimeout(100)
  assert.equal(reads, previousReads, "cross-matter metadata does not read another matter or display apply controls")
  console.log(
    "Chronology chat browser passed: real metadata projection, card outside folded steps, current/proposed evidence, hidden/freshness, explicit acceptance, stale reload without retry, cross-matter rejection.",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
