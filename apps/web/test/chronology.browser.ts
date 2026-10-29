// Chromium workbench smoke with scripted local ingest and no provider.
// Run from apps/web: bun test/chronology.browser.ts
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Chronology, ChronologyAction, ChronologyEvent, ChronologyProposal } from "../src/api/ingest/chronology"
const { chromium, expect } = (await import(
  process.env.PLAYWRIGHT_MODULE ?? "@playwright/test"
)) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chronology-ui-"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import Workbench from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/chronology/ChronologyWorkbench.tsx"))}
import Sidebar from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/Sidebar.tsx"))}
import { setDocumentsCount } from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/matter/documents-drawer.ts"))}
import { MemoryRouter } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-shell.css"))}
setDocumentsCount(3)
const workbench = createElement(Workbench, { matterId: 'fixture', title: 'Harbour equipment matter', documents: [{id: 1,name: 'Carrier note.md',doc_path: '/fixture/Carrier note.md',created_at: 0}], onState: () => {}, onAsk: context => { window.__context = context }, onClose: () => {} })
createRoot(document.getElementById('root')).render(createElement(MemoryRouter, { initialEntries: ['/matter/fixture?view=chronology'] }, createElement('div', { className: 'shell shell-matter' }, createElement(Sidebar, { onOpenSettings: () => {}, sessionsVersion: 0 }), createElement('main', { className: 'content' }, createElement('div', { className: 'container matter-container' }, createElement('div', { className: 'matter-body' }, workbench))))))
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
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const css = (await build.outputs.find((output) => output.path.endsWith(".css"))?.text()) ?? ""
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch: (request) => {
    const name = new URL(request.url).pathname
    if (name === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (name === "/harness.css") return new Response(css, { headers: { "content-type": "text/css" } })
    return new Response(
      '<html><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>',
      { headers: { "content-type": "text/html" } },
    )
  },
})
const quote = "The carrier collected the sealed equipment on 6 May 2026."
function event(id: string, title: string, date: string): ChronologyEvent {
  return {
    id,
    title,
    detail: "The source reports collection; arrival and signed receipt remain separate.",
    date: { original: date, precision: "exact", start: date },
    accounts: [
      {
        id: `account-${id}`,
        title,
        detail: "",
        date: { original: date, precision: "exact", start: date },
        evidence: [
          {
            id: `evidence-${id}`,
            sourceId: "source-1",
            documentName: "Carrier note.md",
            version: "v1",
            excerpt: quote,
            section: "Collection",
            charStart: 10,
            charEnd: 10 + quote.length,
            verified: true,
            freshness: "current",
            provenance: "source",
          },
        ],
        privilege: "uncertain",
        significance: "relevant",
        people: ["Carrier"],
        status: "reported",
        gaps: [],
      },
    ],
    origin: "generated",
    reviewed: false,
    opened: false,
    hidden: false,
    key: false,
    disputed: false,
    note: "",
    lawyerEdited: false,
    createdAt: 1,
    updatedAt: 1,
    previousIds: [],
  }
}
let state: Chronology = {
  schemaVersion: 1,
  revision: 0,
  created: true,
  events: [
    event("collection", "Equipment collected", "2026-05-06"),
    { ...event("arrival", "Equipment arrived", "2026-05-08"), reviewed: true, opened: true },
  ],
  coverage: [
    {
      sourceId: "source-1",
      documentName: "Carrier note.md",
      version: "v1",
      status: "complete",
      uploaded: true,
      textExtracted: true,
      sections: 2,
      processedSections: [0, 1],
      completedKeys: ["one", "two"],
      events: 2,
      issues: [],
      suppliedChronology: false,
    },
  ],
  proposals: [],
  sourceDates: {},
  suggestion: { status: "relevant", reason: "Dates and delivery accounts make a chronology useful." },
  history: [],
  framing: "Neutral",
  gaps: [],
}
let runCalls = 0,
  decisions = 0,
  forceConflict = false
const browser = await chromium.launch({
  headless: true,
  ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}),
})
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } })
  page.setDefaultTimeout(10000)
  const errors: string[] = []
  page.on("pageerror", (error) => errors.push(error.message))
  await page.route("http://chronology.invalid/**", (route) => {
    const request = route.request(),
      url = new URL(request.url())
    const reply = (body: unknown, status = 200) =>
      route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) })
    if (url.pathname === "/matters/fixture" && request.method() === "GET")
      return reply({
        id: "fixture",
        title: "Harbour equipment matter",
        dir: "/fixture",
        created_at: 0,
        documents: [{ id: 1, name: "Carrier note.md", doc_path: "/fixture/Carrier note.md", created_at: 0 }],
      })
    if (url.pathname.endsWith("/documents/text"))
      return reply({ text: `# Source\n\n${quote}\n\nArrival was reported two days later.` })
    if (url.pathname.endsWith("/documents/versions"))
      return reply({ name: "Carrier note.md", document: null, versions: [], events: [], working_version: null })
    if (url.pathname.endsWith("/chronology/preflight"))
      return reply({
        sources: [
          {
            sourceId: "source-1",
            documentName: "Carrier note.md",
            version: "v1",
            kind: "source",
            extraction: "readable",
            estimatedSections: 2,
          },
        ],
        model: { providerID: "scripted", modelID: "fixture" },
        framing: "Unknown",
        framingKnown: false,
        estimatedSections: 2,
        workExplanation: "Bounded extraction over the source record.",
      })
    if (url.pathname.endsWith("/chronology") && request.method() === "GET") return reply(state)
    const input = request.postDataJSON() as {
      expectedRevision: number
      action?: ChronologyAction
      privilege?: string
      decision?: string
    }
    if (url.pathname.endsWith("/chronology/actions")) {
      if (forceConflict || input.expectedRevision !== state.revision) {
        forceConflict = false
        state.revision++
        return reply({ error: "The chronology changed elsewhere. Reload and review your change." }, 409)
      }
      const action = input.action!,
        target = "eventId" in action ? state.events.find((item) => item.id === action.eventId) : undefined
      if (target) {
        if (action.type === "open") target.opened = true
        if (action.type === "review") target.reviewed = action.reviewed
        if (action.type === "note") target.note = action.note
        if (action.type === "key") target.key = action.key
        if (action.type === "hide") target.hidden = action.hidden
        if (action.type === "edit") {
          target.title = action.title
          target.detail = action.detail
          target.date = action.date
          target.lawyerEdited = true
          target.reviewed = false
        }
      }
      state.revision++
      return reply(state)
    }
    if (url.pathname.endsWith("/chronology/run")) {
      assert.ok(["mixed", "stop", "all-cleared"].includes(input.privilege!))
      runCalls++
      state.revision++
      return reply(state)
    }
    if (/\/chronology\/proposals\/.+\/decision$/.test(url.pathname)) {
      const proposal = state.proposals.find((item) => url.pathname.includes(item.id))!
      decisions++
      if (input.decision === "accept") {
        state.events.push(...proposal.events.map((item) => ({ ...item, reviewed: true })))
        proposal.status = "accepted"
      }
      if (input.decision === "keep") proposal.status = "kept"
      if (input.decision === "defer") proposal.status = "deferred"
      state.revision++
      return reply(state)
    }
    return reply({ error: `Unexpected ${request.method()} ${url.pathname}` }, 404)
  })
  await page.route("http://engine.invalid/**", (route) =>
    route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ data: [] }) }),
  )
  await page.goto(server.url.href)
  const detail = page.getByRole("complementary", { name: "Event and evidence" })
  await expect(page.getByRole("heading", { name: "Fact chronology", exact: true })).toBeVisible()
  await expect(detail).toHaveCount(0)
  await page.getByRole("combobox", { name: "Sort events", exact: true }).selectOption("newest")
  await expect(page.locator(".chronology-event-table tbody tr").first()).toHaveAttribute("data-event-id", "arrival")
  await page.getByRole("combobox", { name: "Sort events", exact: true }).selectOption("oldest")
  await expect(page.locator(".chronology-event-table tbody tr").first()).toHaveAttribute("data-event-id", "collection")
  await page.getByRole("searchbox", { name: "Search events and quotations" }).fill("sealed equipment")
  await page.getByRole("button", { name: "Equipment collected", exact: true }).focus()
  await page.keyboard.press("Enter")
  await expect(detail).toBeVisible()
  await expect.poll(() => state.events[0]!.opened).toBe(true)
  assert.equal(state.events[0]!.reviewed, false, "opening clears New without review")
  await detail.getByRole("button", { name: "View in document" }).click()
  const viewer = page.getByRole("dialog", { name: "Carrier note.md", exact: true })
  await expect(viewer).toBeVisible()
  await expect(viewer.locator("mark.citation-passage")).toHaveText(quote)
  await page.keyboard.press("Escape")
  await expect(viewer).toHaveCount(0)
  await expect(detail.getByRole("button", { name: "View in document" })).toBeFocused()
  await expect(page.getByRole("searchbox")).toHaveValue("sealed equipment")
  await detail.getByRole("button", { name: "Mark reviewed", exact: true }).click()
  await detail.locator(".chronology-more-actions > summary").click()
  await detail.getByRole("textbox", { name: "Optional review note" }).fill("Confirm the carrier name with the client.")
  await detail.getByRole("button", { name: "Save note" }).click()
  await detail.getByRole("button", { name: "Mark as key event" }).click()
  await expect.poll(() => state.events[0]!.key).toBe(true)
  await detail.getByRole("button", { name: "Ask doc.haus", exact: true }).click()
  assert.match(
    await page.evaluate(() => (window as unknown as { __context: string }).__context),
    /Return to chronology entry/,
  )
  await detail.getByRole("button", { name: "Hide event", exact: true }).click()
  await expect.poll(() => state.events[0]!.hidden).toBe(true)
  await detail.getByRole("button", { name: "Back to events" }).click()
  await expect(page.getByRole("searchbox")).toBeFocused()
  await expect(page.getByRole("button", { name: "Equipment collected", exact: true })).toHaveCount(0)
  await page.getByRole("button", { name: /^Filter/ }).click()
  await page.getByRole("combobox", { name: /^Visibility/ }).selectOption("only")
  await page.locator('[data-event-id="collection"] .chronology-event-title').click()
  await detail.locator(".chronology-more-actions > summary").click()
  await detail.getByRole("button", { name: "Restore event", exact: true }).click()
  await detail.getByRole("button", { name: "Edit / correct date" }).click()
  const edit = page.getByRole("dialog", { name: "Edit event", exact: true })
  await edit.getByRole("textbox", { name: "Event", exact: true }).fill("Equipment collected and sealed")
  forceConflict = true
  await edit.getByRole("button", { name: "Confirm event" }).click()
  await expect(edit.getByRole("alert")).toContainText("changed elsewhere")
  assert.equal(state.events[0]!.title, "Equipment collected", "stale action not applied or replayed")
  await edit.getByRole("button", { name: "Confirm event" }).click()
  await expect(edit).toHaveCount(0)
  assert.equal(state.events[0]!.reviewed, false)
  await detail.getByRole("button", { name: "Back to events" }).click()
  await page.getByRole("button", { name: "Clear", exact: true }).click()
  await page.reload()
  await page.locator('[data-event-id="collection"] .chronology-event-title').click()
  await detail.locator(".chronology-more-actions > summary").click()
  await expect(detail.getByRole("textbox", { name: "Optional review note" })).toHaveValue(
    "Confirm the carrier name with the client.",
  )
  await detail.getByRole("button", { name: "Back to events" }).click()
  await page.getByRole("button", { name: "Refresh from sources" }).click()
  const refresh = page.getByRole("dialog", { name: "Refresh chronology", exact: true })
  await expect(refresh).toContainText("Carrier note.md")
  await expect(refresh.getByRole("button", { name: "Generate refresh proposals" })).toBeDisabled()
  await refresh.getByRole("radio", { name: /^Mixed/ }).check()
  await refresh.getByRole("button", { name: "Generate refresh proposals" }).click()
  await expect(refresh).toHaveCount(0)
  assert.equal(runCalls, 1)
  await page.getByRole("button", { name: "Refresh from sources" }).click()
  await expect(refresh.getByRole("radio", { name: /^Mixed/ })).not.toBeChecked()
  await refresh.getByRole("button", { name: "Cancel", exact: true }).click()
  assert.equal(runCalls, 1)
  const proposal: ChronologyProposal = {
    id: "proposal-new",
    kind: "addition",
    targetIds: [],
    baseEventHashes: {},
    events: [event("receipt", "Signed receipt recorded", "2026-05-09")],
    reason: "A new source reports a distinct receipt.",
    recommendation: "Keep collection, arrival and receipt separately.",
    status: "pending",
    origin: "refresh",
    createdAt: 10,
  }
  state.proposals.push(proposal)
  await page.reload()
  await page.getByRole("button", { name: "1 Updates", exact: true }).click()
  await expect(page.getByRole("region", { name: "Review chronology updates" })).toContainText("Signed receipt recorded")
  await page.getByRole("button", { name: "Skip for later", exact: true }).click()
  assert.equal(state.proposals[0]!.status, "deferred")
  await page.getByRole("button", { name: "Accept this change", exact: true }).click()
  await expect.poll(() => decisions).toBe(2)
  assert.equal(state.events.find((item) => item.id === "receipt")!.reviewed, true)
  await page.getByRole("button", { name: "Back to saved events" }).click()
  const evidenceDirectory = process.env.CHRONOLOGY_EVIDENCE_DIR
  if (evidenceDirectory)
    await page.screenshot({ path: path.join(evidenceDirectory, "implementation-desktop.png"), fullPage: false })
  for (const width of [768, 375]) {
    await page.setViewportSize({ width, height: 900 })
    await page.getByRole("button", { name: "Equipment arrived", exact: true }).scrollIntoViewIfNeeded()
    const scrollPosition = () =>
      page.evaluate(() => ({
        page: document.scrollingElement?.scrollTop ?? 0,
        workbench: document.querySelector(".chronology-workbench")?.scrollTop ?? 0,
        table: document.querySelector(".chronology-table-scroll")?.scrollLeft ?? 0,
      }))
    const beforeSelection = await scrollPosition()
    await page.getByRole("button", { name: "Equipment arrived", exact: true }).click()
    await expect(detail).toBeVisible()
    await expect(detail.getByRole("button", { name: "Back to events" })).toBeInViewport()
    await expect(detail.getByRole("button", { name: "Back to events" })).toBeInViewport()
    await expect(page.getByRole("searchbox")).not.toBeVisible()
    assert.ok(
      await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
      `page fits ${width}px`,
    )
    if (evidenceDirectory)
      await page.screenshot({ path: path.join(evidenceDirectory, `implementation-${width}.png`), fullPage: false })
    await detail.getByRole("button", { name: "Back to events" }).click()
    await expect(page.getByRole("searchbox")).toBeVisible()
    await expect.poll(scrollPosition).toEqual(beforeSelection)
    assert.ok(
      await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
      `table is locally scrollable at ${width}px`,
    )
    if (width === 375) {
      const surfaces = page.locator(".sidebar-surfaces")
      await surfaces.getByRole("link", { name: "Review table", exact: true }).focus()
      const navBounds = await surfaces.boundingBox()
      const reviewBounds = await surfaces.getByRole("link", { name: "Review table", exact: true }).boundingBox()
      assert.ok(
        navBounds &&
          reviewBounds &&
          reviewBounds.x >= navBounds.x &&
          reviewBounds.x + reviewBounds.width <= navBounds.x + navBounds.width + 1,
        "keyboard focus reveals the complete navigation control",
      )
      assert.ok(
        await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
        "navigation scroll stays inside the shell",
      )
    }
  }
  // A narrow evidence column must keep the complete review popup reachable.
  await page.setViewportSize({ width: 1050, height: 900 })
  await page.locator('[data-event-id="arrival"] .chronology-event-title').click()
  await detail.locator(".chronology-more-actions > summary").click()
  const panelBounds = await detail.boundingBox()
  const popupBounds = await detail.locator(".chronology-more-actions-content").boundingBox()
  assert.ok(
    panelBounds &&
      popupBounds &&
      popupBounds.x >= panelBounds.x &&
      popupBounds.x + popupBounds.width <= panelBounds.x + panelBounds.width + 1,
    "review popup stays inside evidence column",
  )
  await detail.getByRole("textbox", { name: "Optional review note" }).focus()
  await expect(detail.getByRole("textbox", { name: "Optional review note" })).toBeFocused()
  await detail.locator(".chronology-more-actions > summary").click()
  await detail.getByRole("button", { name: "Back to events" }).click()
  await page.emulateMedia({ reducedMotion: "reduce" })
  assert.deepEqual(errors, [])
  console.log(
    "Chronology browser passed: source/focus return, review/note/star/hide, stale save, reopen, fresh privilege, proposals and 375/768/1440px (scripted ingest; no model).",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
