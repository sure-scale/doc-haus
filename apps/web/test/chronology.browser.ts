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
import MatterHeader from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/matter/MatterHeader.tsx"))}
import { setDocumentsCount } from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/matter/documents-drawer.ts"))}
import { MemoryRouter } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-shell.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/modal-frost.css"))}
setDocumentsCount(3)
const workbench = createElement(Workbench, { matterId: 'fixture', title: 'Harbour equipment matter', documents: [{id: 1,name: 'Carrier note.md',doc_path: '/fixture/Carrier note.md',created_at: 0}], onState: () => {}, onAsk: context => { window.__context = context.text }, onClose: () => {} })
const noop = () => {}, done = async () => {}
const header = createElement(MatterHeader, {
  matter: { id: 'fixture', title: 'Harbour equipment matter', dir: '/fixture', created_at: 0 }, jurisdictions: [], compact: true,
  editingRef: false, setEditingRef: noop, draftRef: '', setDraftRef: noop, commitRef: done,
  editingClient: false, setEditingClient: noop, draftClient: '', setDraftClient: noop, commitClient: done,
  renaming: false, setRenaming: noop, draftTitle: '', setDraftTitle: noop, commitRename: done,
  patch: done, onEditConflicts: noop,
})
createRoot(document.getElementById('root')).render(createElement(MemoryRouter, { initialEntries: ['/matter/fixture?view=chronology'] }, createElement('div', { className: 'shell shell-matter' }, createElement(Sidebar, { onOpenSettings: () => {}, sessionsVersion: 0 }), createElement('main', { className: 'content' }, createElement('div', { className: 'container matter-container' }, header, createElement('div', { className: 'matter-body' }, workbench))))))
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
      if (forceConflict) {
        forceConflict = false
        state.revision++
        return reply({ error: "The chronology changed elsewhere. Reload and review your change." }, 409)
      }
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
  const detail = page.locator('.chronology-evidence-panel[aria-label="Event and evidence"]')
  const canvasGeometry = () =>
    page.evaluate(() => ({
      boxes: [".chronology-main", ".chronology-timeline", ".chronology-event-table"].map((selector) => {
        const box = document.querySelector(selector)!.getBoundingClientRect()
        return { x: box.x, y: box.y, width: box.width, height: box.height }
      }),
      pageScroll: document.scrollingElement?.scrollTop ?? 0,
      workbenchScroll: document.querySelector(".chronology-workbench")!.scrollTop,
      tableScroll: document.querySelector(".chronology-table-scroll")!.scrollLeft,
    }))
  await expect(page.getByRole("heading", { name: "Fact chronology", exact: true })).toBeVisible()
  const zoomBox = await page.locator(".chronology-timeline-controls").boundingBox()
  const chartBox = await page.locator(".chronology-swimlane-chart").boundingBox()
  assert.ok(zoomBox && chartBox && zoomBox.y + zoomBox.height <= chartBox.y, "zoom controls sit above the timeline")
  assert.ok(
    Math.abs(zoomBox!.x + zoomBox!.width - chartBox!.x - chartBox!.width) < 2,
    "zoom controls align to its right edge",
  )
  await expect(detail).toHaveCount(0)
  const timelinePopup = page.getByRole("dialog", { name: "Timeline events", exact: true })
  const collectionPoint = page.getByRole("button", { name: "2026-05-06: Equipment collected", exact: true })
  const beforePreview = {
    revision: state.revision,
    opened: state.events[0]!.opened,
    reviewed: state.events[0]!.reviewed,
  }
  const beforePreviewCanvas = await canvasGeometry()
  await collectionPoint.hover()
  await expect(timelinePopup).toContainText("2026-05-06")
  await expect(timelinePopup).toContainText("Equipment collected")
  await expect(timelinePopup).toContainText("Exact date")
  await expect(detail).toHaveCount(0)
  assert.deepEqual(await canvasGeometry(), beforePreviewCanvas, "hover help does not resize or scroll the chronology")
  await page.keyboard.press("Escape")
  await expect(timelinePopup).toHaveCount(0)
  await collectionPoint.focus()
  await expect(timelinePopup).toContainText("Equipment collected")
  await expect(collectionPoint).toHaveAttribute("aria-expanded", "true")
  await page.keyboard.press("Escape")
  await expect(timelinePopup).toHaveCount(0)
  await expect(collectionPoint).toBeFocused()
  assert.deepEqual(
    { revision: state.revision, opened: state.events[0]!.opened, reviewed: state.events[0]!.reviewed },
    beforePreview,
    "hover, focus and Escape preview the event without opening or reviewing it",
  )
  const arrivalPoint = page.getByRole("button", { name: "2026-05-08: Equipment arrived", exact: true })
  await arrivalPoint.hover()
  await timelinePopup.getByRole("button", { name: /Equipment arrived/ }).click()
  await expect(timelinePopup).toHaveCount(0)
  await expect(detail.getByRole("heading", { name: "Equipment arrived", exact: true })).toBeVisible()
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  await expect(detail).toHaveCount(0)
  await expect(page.locator('[data-event-id="arrival"] .chronology-event-title')).toBeFocused()
  await collectionPoint.click()
  await expect(detail.getByRole("heading", { name: "Equipment collected", exact: true })).toBeVisible()
  await expect.poll(() => state.events[0]!.opened).toBe(true)
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  await expect(detail).toHaveCount(0)
  await expect(page.locator('[data-event-id="collection"] .chronology-event-title')).toBeFocused()
  // Restore this fixture's unopened baseline for the review workflow below.
  state.events[0]!.opened = false
  // Each searchable filter must narrow actual events and restore its trigger,
  // rather than merely searching labels in an unconnected dropdown.
  const inspection = event("inspection-filter", "Independent inspection completed", "2026-05-10")
  inspection.accounts[0]!.people = ["Independent surveyor"]
  inspection.accounts[0]!.evidence[0]!.documentName = "Inspection report.md"
  inspection.accounts[0]!.evidence[0]!.sourceId = "source-inspection"
  state.events.push(inspection)
  state.revision++
  await page.reload()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(3)
  await page.getByRole("button", { name: /^Filter/ }).click()
  const filterColumns = await page.locator(".chronology-filter-primary > *").evaluateAll((fields) =>
    fields.map((field) => {
      const box = field.getBoundingClientRect()
      return { x: box.x, y: box.y }
    }),
  )
  assert.equal(filterColumns.length, 3)
  assert.ok(
    filterColumns.every((field) => Math.abs(field.y - filterColumns[0]!.y) < 20),
    "desktop filters use one primary row",
  )
  assert.ok(filterColumns[0]!.x < filterColumns[1]!.x && filterColumns[1]!.x < filterColumns[2]!.x)
  const documentFilter = page.getByRole("button", { name: "Document", exact: true })
  const filterControls = await Promise.all([
    documentFilter.boundingBox(),
    page.getByRole("button", { name: "Person / organisation", exact: true }).boundingBox(),
    page.getByRole("textbox", { name: "From date", exact: true }).boundingBox(),
    page.getByRole("textbox", { name: "To date", exact: true }).boundingBox(),
  ])
  assert.ok(
    filterControls.every((box) => box && Math.abs(box.y - filterControls[0]!.y) < 2),
    "filter controls share a baseline",
  )
  await documentFilter.click()
  const documentPicker = page.getByRole("group", { name: "Document", exact: true })
  await expect(documentPicker.getByRole("textbox", { name: "Search documents" })).toBeFocused()
  await documentPicker.getByRole("textbox", { name: "Search documents" }).fill("inspection")
  await expect(documentPicker.getByRole("radio")).toHaveCount(1)
  await documentPicker.getByRole("radio", { name: "Inspection report.md", exact: true }).click()
  await expect(documentPicker).toHaveCount(0)
  await expect(documentFilter).toBeFocused()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(1)
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveAttribute("data-event-id", "inspection-filter")
  await documentFilter.click()
  await documentPicker.getByRole("button", { name: "All documents", exact: true }).click()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(3)
  const personFilter = page.getByRole("button", { name: "Person / organisation", exact: true })
  await personFilter.click()
  const personPicker = page.getByRole("group", { name: "Person / organisation", exact: true })
  await personPicker.getByRole("textbox", { name: "Search people and organisations" }).fill("unlisted participant")
  await expect(personPicker.getByText("No match.", { exact: true })).toBeVisible()
  await page.keyboard.press("Escape")
  await expect(personPicker).toHaveCount(0)
  await expect(personFilter).toBeFocused()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(3)
  await personFilter.click()
  await personPicker.getByRole("textbox", { name: "Search people and organisations" }).fill("surveyor")
  await expect(personPicker.getByRole("radio")).toHaveCount(1)
  await personPicker.getByRole("radio", { name: "Independent surveyor", exact: true }).click()
  await expect(personFilter).toBeFocused()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(1)
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveAttribute("data-event-id", "inspection-filter")
  await personFilter.click()
  await personPicker.getByRole("button", { name: "All people", exact: true }).click()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(3)
  state.events = state.events.filter((item) => item.id !== inspection.id)
  state.revision++
  await page.reload()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(2)
  const savedFilters = page.getByRole("group", { name: "Filter saved events", exact: true })
  await savedFilters.getByRole("button", { name: /^Not opened/ }).click()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(1)
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveAttribute("data-event-id", "collection")
  await savedFilters.getByRole("button", { name: /^All events/ }).click()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(2)
  const sourcesButton = page.getByRole("button", { name: /^Sources/ })
  await sourcesButton.click()
  const sources = page.getByRole("dialog", { name: "Sources and coverage", exact: true })
  await expect(sources).toContainText("Carrier note.md")
  await expect(sources.getByText("File", { exact: true })).toBeVisible()
  await expect(sources.getByText("Text", { exact: true })).toBeVisible()
  await expect(sources.getByText("Recorded reading", { exact: true })).toBeVisible()
  await expect(sources).toContainText("2 saved entries")
  await sources.getByText("How source processing works", { exact: true }).click()
  await expect(sources).toContainText("Display filters change only the events shown")
  await page.keyboard.press("Escape")
  await expect(sources).toHaveCount(0)
  await expect(sourcesButton).toBeFocused()
  await expect(page.locator(".chronology-event-table tbody tr").first()).toHaveAttribute("data-event-id", "arrival")
  await expect(page.getByRole("columnheader", { name: /Date/ })).toHaveAttribute("aria-sort", "descending")
  await page.getByRole("button", { name: "Sort by Date", exact: true }).click()
  await expect(page.locator(".chronology-event-table tbody tr").first()).toHaveAttribute("data-event-id", "collection")
  await page.getByRole("searchbox", { name: "Search events and quotations" }).fill("sealed equipment")
  await page.getByRole("button", { name: "Equipment collected", exact: true }).focus()
  const desktopCanvas = await canvasGeometry()
  await page.keyboard.press("Enter")
  await expect(detail).toBeVisible()
  assert.deepEqual(
    await canvasGeometry(),
    desktopCanvas,
    "evidence overlays the desktop canvas without reflow or scroll",
  )
  const evidenceSurface = await detail.evaluate((panel) => {
    const overlay = panel.closest(".chronology-detail-overlay")!
    const background = getComputedStyle(overlay)
    const surface = getComputedStyle(panel)
    const bounds = overlay.getBoundingClientRect()
    return {
      blur: background.backdropFilter,
      panelBackground: surface.backgroundColor,
      panelBlur: surface.backdropFilter,
      coversWorkspace: bounds.x === 0 && bounds.y === 0 && bounds.width === innerWidth && bounds.height === innerHeight,
    }
  })
  assert.equal(evidenceSurface.blur, "blur(6px)", "shared frost blurs the workspace to the left of the sheet")
  assert.equal(evidenceSurface.panelBackground, "rgb(255, 255, 255)", "evidence sheet stays solid white")
  assert.equal(evidenceSurface.panelBlur, "none")
  assert.equal(evidenceSurface.coversWorkspace, true, "frost covers the navigation as well as the chronology")
  const evidencePanelBox = (await detail.boundingBox())!
  assert.ok(
    evidencePanelBox.height >= 700,
    "short evidence fills the desktop overlay, keeping its footer near the bottom",
  )
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
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  await expect(page.getByRole("searchbox")).toBeFocused()
  await expect(page.getByRole("button", { name: "Equipment collected", exact: true })).toHaveCount(0)
  await page.getByRole("button", { name: /^Filter/ }).click()
  await page.getByRole("combobox", { name: /^Visibility/ }).selectOption("only")
  await page.locator('[data-event-id="collection"] .chronology-event-title').click()
  await detail.locator(".chronology-more-actions > summary").click()
  await detail.getByRole("button", { name: "Restore event", exact: true }).click()
  await detail.getByRole("button", { name: "Edit / correct date" }).click()
  const edit = page.getByRole("dialog", { name: "Edit event", exact: true })
  const expectSharedBackdrop = async () => {
    const backdrop = await page.locator(".chronology-dialog[open]").evaluate((dialog) => {
      const reference = document.createElement("dialog")
      reference.className = "review-modal"
      document.body.append(reference)
      const shared = getComputedStyle(reference, "::backdrop")
      const actual = getComputedStyle(dialog, "::backdrop")
      const result = {
        background: actual.backgroundColor,
        blur: actual.backdropFilter,
        sharedBackground: shared.backgroundColor,
        sharedBlur: shared.backdropFilter,
      }
      reference.remove()
      return result
    })
    assert.equal(backdrop.background, backdrop.sharedBackground, "chronology uses the shared pale modal backdrop")
    assert.equal(backdrop.blur, backdrop.sharedBlur, "chronology uses the shared modal blur")
    assert.notEqual(backdrop.blur, "none", "modal background is blurred")
  }
  await expectSharedBackdrop()
  await edit.getByRole("textbox", { name: "Event", exact: true }).fill("Equipment collected and sealed")
  forceConflict = true
  await edit.getByRole("button", { name: "Confirm event" }).click()
  await expect(edit.getByRole("alert")).toContainText("changed elsewhere")
  assert.equal(state.events[0]!.title, "Equipment collected", "stale action not applied or replayed")
  await edit.getByRole("button", { name: "Confirm event" }).click()
  await expect(edit).toHaveCount(0)
  assert.equal(state.events[0]!.reviewed, false)
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  await page.getByRole("button", { name: "Clear", exact: true }).click()
  await page.reload()
  await page.locator('[data-event-id="collection"] .chronology-event-title').click()
  await detail.locator(".chronology-more-actions > summary").click()
  await expect(detail.getByRole("textbox", { name: "Optional review note" })).toHaveValue(
    "Confirm the carrier name with the client.",
  )
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  await page.locator(".matter-actions-menu > summary").click()
  const menuItems = await page.locator(".matter-actions-menu-items > button").evaluateAll((buttons) =>
    buttons.map((button) => {
      const box = button.getBoundingClientRect()
      return { x: box.x, y: box.y, right: box.right, bottom: box.bottom, width: box.width }
    }),
  )
  assert.equal(menuItems.length, 2)
  assert.ok(menuItems[1]!.y >= menuItems[0]!.bottom, "overflow actions stack vertically")
  assert.ok(Math.abs(menuItems[0]!.x - menuItems[1]!.x) < 1, "overflow actions share a left edge")
  assert.ok(Math.abs(menuItems[0]!.width - menuItems[1]!.width) < 1, "overflow actions fill one column")
  await page.getByRole("button", { name: "Check for source changes" }).click()
  const refresh = page.getByRole("dialog", { name: "Refresh chronology", exact: true })
  await expectSharedBackdrop()
  await expect(refresh).toContainText("Carrier note.md")
  await expect(refresh.getByRole("button", { name: "Check sources and suggest updates" })).toBeDisabled()
  await refresh.getByRole("radio", { name: /^Mixed/ }).check()
  await refresh.getByRole("button", { name: "Check sources and suggest updates" }).click()
  await expect(refresh).toHaveCount(0)
  assert.equal(runCalls, 1)
  await page.locator(".matter-actions-menu > summary").click()
  await page.getByRole("button", { name: "Check for source changes" }).click()
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
  await page.getByRole("button", { name: "Updates 1", exact: true }).click()
  await expect(page.getByRole("region", { name: "Review chronology updates" })).toContainText("Signed receipt recorded")
  await expect(page.getByRole("region", { name: "Current saved entry", exact: true })).toHaveCount(0)
  await page.getByRole("button", { name: "Update to review", exact: true }).click()
  const updatePicker = page.getByRole("group", { name: "Update to review", exact: true })
  await updatePicker.getByRole("textbox", { name: "Search updates", exact: true }).fill("receipt")
  await expect(updatePicker.getByRole("radio")).toHaveCount(1)
  await updatePicker.getByRole("radio").click()
  await expect(updatePicker).toHaveCount(0)
  forceConflict = true
  await page.getByRole("button", { name: "Skip for later", exact: true }).click()
  const updatesDialog = page.getByRole("dialog", { name: "Review chronology updates", exact: true })
  await expect(updatesDialog.getByRole("alert")).toContainText("The chronology changed elsewhere")
  assert.equal(state.proposals[0]!.status, "pending", "conflicting decisions preserve the pending change")
  assert.equal(decisions, 0, "conflicting decisions are not replayed")
  await updatesDialog.getByRole("button", { name: "Reload latest", exact: true }).click()
  await expect(updatesDialog.getByRole("alert")).toHaveCount(0)
  await page.getByRole("button", { name: "Skip for later", exact: true }).click()
  await expect.poll(() => state.proposals[0]!.status).toBe("deferred")
  await page.getByRole("button", { name: "Accept this change", exact: true }).click()
  await expect.poll(() => decisions).toBe(2)
  assert.equal(state.events.find((item) => item.id === "receipt")!.reviewed, true)
  await page.getByRole("button", { name: "Back to saved events" }).click()
  for (const kind of ["amendment", "conflict", "removal"] as const) {
    const proposed = event(`fixture-${kind}`, `${kind} collection proposal`, "2026-05-07")
    state.proposals.push({
      ...proposal,
      id: `review-${kind}`,
      kind,
      targetIds: ["collection"],
      events: kind === "removal" ? [] : [proposed],
      status: "pending",
    })
  }
  state.revision++
  await page.reload()
  await page.getByRole("button", { name: "Updates 3", exact: true }).click()
  const reviewSurface = page.getByRole("dialog", { name: "Review chronology updates", exact: true })
  for (const purpose of ["Update this event", "Review a competing account", "Hide this event"]) {
    await expect(reviewSurface.getByRole("heading", { name: purpose, exact: true })).toBeVisible()
    await expect(reviewSurface.getByRole("region", { name: "Current saved entry", exact: true })).toContainText(
      "Equipment collected",
    )
    if (purpose === "Hide this event") {
      await expect(reviewSurface.getByRole("region", { name: "Proposed result", exact: true })).toHaveCount(0)
      await expect(reviewSurface).toContainText("This does not delete the record")
    } else await expect(reviewSurface.getByRole("region", { name: "Proposed result", exact: true })).toBeVisible()
    await reviewSurface.getByRole("button", { name: "Keep current", exact: true }).click()
  }
  await reviewSurface.getByRole("button", { name: "Back to saved events", exact: true }).click()
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
    const beforeCanvas = await canvasGeometry()
    await page.getByRole("button", { name: "Equipment arrived", exact: true }).click()
    await expect(detail).toBeVisible()
    await expect(detail.getByRole("button", { name: "Close event evidence" })).toBeInViewport()
    assert.deepEqual(await canvasGeometry(), beforeCanvas, `evidence does not reposition the ${width}px canvas`)
    for (const control of [
      detail.getByRole("button", { name: "Close event evidence" }),
      detail.getByRole("button", { name: "Mark as key event" }),
      detail.getByRole("button", { name: "View in document", exact: true }).first(),
    ]) {
      const bounds = await control.boundingBox()
      assert.ok(bounds && bounds.height >= 44, `evidence controls retain 44px targets at ${width}px`)
    }
    assert.ok(
      await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
      `page fits ${width}px`,
    )
    if (evidenceDirectory)
      await page.screenshot({ path: path.join(evidenceDirectory, `implementation-${width}.png`), fullPage: false })
    await detail.getByRole("button", { name: "Close event evidence" }).click()
    await expect(page.getByRole("searchbox")).toBeVisible()
    await expect.poll(scrollPosition).toEqual(beforeSelection)
    assert.deepEqual(await canvasGeometry(), beforeCanvas, `closing evidence does not reposition the ${width}px canvas`)
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
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  // Dense same-day events across source kinds remain selectable by cluster,
  // rather than overlapping transparent hit areas that select a neighbour.
  state.events = Array.from({ length: 24 }, (_, index) => {
    const item = event(
      `dense-${index}`,
      `Independent occurrence ${index + 1}`,
      `2026-05-${String(6 + (index % 5)).padStart(2, "0")}`,
    )
    item.accounts[0]!.evidence[0]!.sourceKind = ["email", "agreement", "file note"][index % 3]
    return item
  })
  state.events.push(event("far-date", "Later follow-up", "2026-12-20"))
  state.revision++
  await page.reload()
  const clusterPoint = page.getByRole("button", { name: /clustered events near/ }).first()
  await expect(clusterPoint).toBeVisible()
  assert.ok(
    await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight + 1),
    "dense desktop record stays inside the application viewport",
  )
  assert.ok(
    await page.locator(".chronology-workbench").evaluate((element) => element.scrollHeight > element.clientHeight),
    "long desktop record scrolls inside the workbench",
  )
  await clusterPoint.hover()
  await expect(timelinePopup).toContainText("8 nearby events")
  await expect(timelinePopup.locator(".chronology-marker-popup-event")).toHaveCount(8)
  await expect(timelinePopup).toContainText("Independent occurrence")
  await expect(timelinePopup).toContainText("Choose an event to view its evidence.")
  const eventTargets = await timelinePopup
    .locator(".chronology-marker-popup-event")
    .evaluateAll((buttons) => buttons.map((button) => button.getBoundingClientRect().height))
  assert.ok(
    eventTargets.every((height) => height >= 44),
    "every clustered event has its own 44px action",
  )
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(25)
  await page.keyboard.press("Escape")
  await expect(timelinePopup).toHaveCount(0)
  await page.setViewportSize({ width: 375, height: 900 })
  await page.getByRole("button", { name: "2026-12-20: Later follow-up", exact: true }).hover()
  await expect(timelinePopup).toContainText("Later follow-up")
  const tooltipBounds = (await timelinePopup.boundingBox())!
  assert.ok(
    tooltipBounds.x >= 0 &&
      tooltipBounds.y >= 0 &&
      tooltipBounds.x + tooltipBounds.width <= 376 &&
      tooltipBounds.y + tooltipBounds.height <= 901,
    "hover help at the right timeline edge stays inside a narrow viewport",
  )
  await page.keyboard.press("Escape")
  await expect(timelinePopup).toHaveCount(0)
  await page.setViewportSize({ width: 1050, height: 900 })
  await clusterPoint.click()
  const cluster = page.getByRole("group", { name: "Clustered events", exact: true })
  await expect(cluster.getByRole("button").first()).toBeFocused()
  await page.keyboard.press("Escape")
  await expect(cluster).toHaveCount(0)
  await expect(clusterPoint).toBeFocused()
  await clusterPoint.click()
  const firstClusterEvent = cluster.getByRole("button").first()
  const chosenTitle = await firstClusterEvent.locator(".chronology-marker-popup-title").innerText()
  await firstClusterEvent.click()
  await expect(timelinePopup).toHaveCount(0)
  await expect(detail.getByRole("heading", { name: chosenTitle, exact: true })).toBeVisible()
  await expect(page.locator(".chronology-event-table tbody tr")).toHaveCount(25)
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  await expect(detail).toHaveCount(0)
  await expect(
    page.locator(".chronology-event-table").getByRole("button", { name: chosenTitle, exact: true }),
  ).toBeFocused()
  const overlappingTargets = () =>
    page.locator(".chronology-timeline-point").evaluateAll((points) => {
      const bounds = points.map((point) => ({
        lane: [...point.classList].find((name) => name.startsWith("lane-")),
        box: point.getBoundingClientRect(),
      }))
      return bounds.some((a, i) =>
        bounds.some(
          (b, j) => j > i && a.lane === b.lane && a.box.left < b.box.right - 1 && a.box.right > b.box.left + 1,
        ),
      )
    })
  await expect.poll(overlappingTargets).toBe(false)
  await page.getByRole("button", { name: "Zoom in timeline", exact: true }).click()
  await expect.poll(overlappingTargets).toBe(false)
  await page.getByRole("button", { name: "Fit", exact: true }).click()
  await page
    .getByRole("button", { name: /clustered events near/ })
    .first()
    .click()
  await page.getByRole("searchbox", { name: "Search events and quotations" }).fill("absent phrase for empty view")
  await expect(page.getByRole("group", { name: "Clustered events", exact: true })).toHaveCount(0)
  await expect(page.getByRole("searchbox")).toBeFocused()
  await expect(page.getByText("No events match these display filters.", { exact: true })).toBeVisible()
  const manual = event("manual-undated", "Undated site observation", "2026-05-06")
  manual.date = { original: "Date not supplied", precision: "undated" }
  manual.accounts = []
  manual.origin = "manual"
  state.events.push(manual)
  state.revision++
  await page.reload()
  await page.getByRole("searchbox").fill("")
  await page.setViewportSize({ width: 768, height: 900 })
  const collapse = page.getByRole("button", { name: "Collapse sidebar", exact: true })
  if (await collapse.isVisible()) await collapse.click()
  await page.getByRole("button", { name: "Expand sidebar", exact: true }).click()
  await expect(collapse).toBeVisible()
  await expect
    .poll(() =>
      page.locator(".chronology-date-grid > span").evaluateAll((labels) => {
        const visible = labels
          .filter((label) => getComputedStyle(label).display !== "none")
          .map((label) => label.getBoundingClientRect())
          .sort((left, right) => left.left - right.left)
        return visible.length >= 2 && visible.every((box, index) => !index || box.left >= visible[index - 1]!.right + 3)
      }),
    )
    .toBe(true)
  assert.ok(
    await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight + 1),
    "dense tablet record stays bounded at the shell's 701px breakpoint, above its status bar",
  )
  await page.getByRole("button", { name: "Independent occurrence 2", exact: true }).click()
  await expect(detail.getByRole("button", { name: "Close event evidence" })).toBeInViewport()
  await detail.getByRole("button", { name: "Close event evidence" }).click()
  state.run = {
    id: "synthetic-running-preview",
    mode: "initial",
    status: "running",
    privilege: "mixed",
    startedAt: Date.now(),
    sourceKeys: ["source-1:v1"],
    completedKeys: [],
    errors: [],
  }
  state.revision++
  await page.reload()
  await expect(page.locator(".chronology-run")).toBeVisible()
  assert.ok(
    (await page.locator(".chronology-run > div:first-of-type").boundingBox())!.width >= 200,
    "tablet progress copy retains a readable line width beside or above its actions",
  )
  assert.equal(
    await page.locator(".chronology-run strong").evaluate((el) => {
      const range = document.createRange()
      range.selectNodeContents(el)
      return range.getClientRects().length
    }),
    1,
    "Building chronology does not become a narrow word column",
  )
  delete state.run
  state.revision++
  await page.reload()
  await page.setViewportSize({ width: 1440, height: 1000 })
  const undated = page.locator(".chronology-undated-events")
  await expect(undated.locator("summary")).toContainText("not placed on timeline")
  await undated.locator("summary").click()
  await expect(undated).toContainText("Source processing is shown separately under Sources.")
  await undated.getByRole("button", { name: /Undated site observation/ }).click()
  await expect(detail.getByText("Source not supplied. This entry was added by hand.", { exact: true })).toBeVisible()
  await page.setViewportSize({ width: 1440, height: 600 })
  await expect(detail.getByRole("button", { name: "Mark reviewed", exact: true })).toBeInViewport()
  await expect(detail.getByRole("button", { name: "Close event evidence" })).toBeInViewport()
  await page.emulateMedia({ reducedMotion: "reduce" })
  assert.deepEqual(errors, [])
  console.log(
    "Chronology browser passed: actionable hover/focus/Escape event lists, source coverage/review filters, searchable source/person filters, source/focus return, review/note/star/hide, stale save, reopen, fresh privilege, proposals and 375/768/1440px (scripted ingest; no model).",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
