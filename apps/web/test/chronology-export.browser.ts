// Real workbench → preview → binary download, with scripted source state and no model.
// Run from apps/web: bun test/chronology-export.browser.ts
import assert from "node:assert/strict"
import { mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { strFromU8, unzipSync } from "fflate"
import type {
  Chronology,
  ChronologyEvent,
  ChronologyExportRequest,
  ChronologyExportSnapshot,
} from "../src/api/ingest/chronology"
import { chronologyDocx } from "../../../services/ingest/src/chronology-docx"
const { chromium, expect } = await import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chronology-export-"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import Workbench from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/chronology/ChronologyWorkbench.tsx"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/primitives.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/modal-frost.css"))}
createRoot(document.getElementById('root')).render(createElement(Workbench, { matterId: 'fixture', title: 'Equipment inspection', documents: [], onState: () => {}, onAsk: () => {}, onClose: () => {}, onDocumentsChanged: () => { window.__documentsRefreshed = (window.__documentsRefreshed || 0) + 1 } }))
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
assert.ok(build.success, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const css = (await build.outputs.find((output) => output.path.endsWith(".css"))?.text()) ?? ""
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch(request) {
    const name = new URL(request.url).pathname
    if (name === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (name === "/harness.css") return new Response(css, { headers: { "content-type": "text/css" } })
    return new Response(
      '<html><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>',
      { headers: { "content-type": "text/html" } },
    )
  },
})
function event(id: string, title: string, hidden = false): ChronologyEvent {
  return {
    id,
    title,
    hidden,
    detail: "Reported by the inspector; acceptance remains disputed.",
    date: { original: "May 2026", start: "2026-05", precision: "month" },
    accounts: [],
    origin: "manual",
    reviewed: false,
    opened: true,
    key: false,
    disputed: true,
    dispute: "The operator gives a different account.",
    note: "Check the access record before relying on this entry.",
    lawyerEdited: false,
    createdAt: 1,
    updatedAt: 1,
    previousIds: [],
  }
}
const collection = event("collection", "Collection reported")
collection.origin = "generated"
collection.date = { original: "21 May 2026", start: "2026-05-21", precision: "exact" }
collection.accounts = [
  {
    id: "collection-account",
    title: "Collection recorded by the warehouse operator",
    detail: "The equipment was collected on 21 May 2026.",
    date: collection.date,
    attribution: "Warehouse operator",
    privilege: "not-privileged",
    significance: "relevant",
    people: [],
    status: "reported",
    gaps: [],
    evidence: [
      {
        id: "collection-evidence",
        documentName: "Equipment handover record.pdf",
        sourceId: "handover-record",
        version: "source-v1",
        excerpt: "The equipment was collected on 21 May 2026.",
        section: "Collection record",
        charStart: 0,
        charEnd: 43,
        verified: true,
        freshness: "current",
        provenance: "source",
      },
    ],
  },
]
const state: Chronology = {
  schemaVersion: 1,
  revision: 3,
  created: true,
  events: [event("inspection", "Inspection reported"), collection, event("hidden", "Inspection correction", true)],
  coverage: [
    {
      sourceId: "handover-record",
      documentName: "Equipment handover record.pdf",
      version: "source-v1",
      status: "complete",
      uploaded: true,
      textExtracted: true,
      sections: 1,
      processedSections: [0],
      completedKeys: ["collection-record"],
      events: 1,
      issues: [],
      suppliedChronology: false,
    },
  ],
  proposals: [],
  sourceDates: {},
  suggestion: { status: "relevant", reason: "Dated inspection accounts" },
  history: [],
  framing: "Neutral",
  gaps: [],
}
const requests: ChronologyExportRequest[] = []
let conflict = false
let downloaded: Buffer = Buffer.alloc(0)
const browser = await chromium.launch({
  headless: true,
  ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}),
})
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true })
  const errors: string[] = []
  page.on("pageerror", (error) => errors.push(error.message))
  await page.route("http://chronology.invalid/**", async (route) => {
    const pathname = new URL(route.request().url()).pathname
    const json = (body: unknown, status = 200) =>
      route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) })
    if (pathname.endsWith("/chronology")) return json(state)
    if (pathname.endsWith("/export/preview") || pathname.endsWith("/chronology/export")) {
      const input = route.request().postDataJSON() as ChronologyExportRequest
      requests.push(input)
      if (conflict || input.expectedRevision !== state.revision) {
        conflict = false
        state.revision++
        return json({ error: "The chronology changed. Refresh the export preview before downloading." }, 409)
      }
      const selected = state.events.filter(
        (event) =>
          (input.includeHidden || !event.hidden) &&
          (input.scope === "all-visible" || input.eventIds!.includes(event.id)),
      )
      const snapshot: ChronologyExportSnapshot = {
        schemaVersion: 1,
        matterId: "fixture",
        title: "Equipment inspection",
        snapshotAt: Date.now(),
        revision: state.revision,
        scope: input.scope,
        includeHidden: input.includeHidden,
        totalEvents: 3,
        events: selected,
        coverage: state.coverage,
        excludedHidden: selected.some((event) => event.hidden) ? 0 : 1,
        pendingProposals: 1,
        warnings: ["1 pending proposal is excluded from this snapshot."],
      }
      if (pathname.endsWith("/preview")) return json(snapshot)
      downloaded = await chronologyDocx(snapshot)
      return route.fulfill({
        status: 200,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers: {
          "X-DocHaus-Export-Name": "Chronology-fixture.docx",
          "Access-Control-Expose-Headers": "X-DocHaus-Export-Name",
        },
        body: downloaded,
      })
    }
    return json({ error: `Unexpected ${pathname}` }, 404)
  })
  await page.route("http://engine.invalid/**", (route) =>
    route.fulfill({ contentType: "application/json", body: "[]" }),
  )
  await page.goto(server.url.href)
  await expect(page.getByRole("heading", { name: "Fact chronology", exact: true })).toBeVisible()
  const open = async () => {
    await page.locator(".matter-actions-menu > summary").click()
    await page.getByRole("button", { name: "Export chronology to Word", exact: true }).click()
  }
  await open()
  const dialog = page.getByRole("dialog", { name: "Export chronology to Word" })
  await expect(dialog.getByText("2 of 3 saved entries", { exact: false })).toBeVisible()
  await expect(dialog.getByRole("radio", { name: "All entries (2)", exact: true })).toBeChecked()
  await expect(dialog.getByRole("radio", { name: "Filtered entries (2)", exact: true })).toBeVisible()
  await expect(dialog).toContainText("Hidden entries are excluded from both options unless included below.")
  await expect(dialog).toContainText("1 hidden entry excluded")
  assert.equal(requests.at(-1)!.includeHidden, false)
  await expect(dialog.locator("tbody tr")).toHaveCount(2)
  await expect(dialog.getByText("Inspection correction", { exact: false })).toHaveCount(0)
  await expect(dialog).toContainText("1 pending proposal is excluded")
  await dialog.locator("details").first().locator("summary").click()
  await expect(dialog).toContainText("Check the access record before relying on this entry.")
  await expect(dialog).toContainText("Source not supplied")
  await dialog.getByRole("checkbox", { name: "Include hidden entries", exact: true }).check()
  await expect(dialog.locator("tbody tr")).toHaveCount(3)
  await expect(dialog.getByRole("radio", { name: "All entries (3)", exact: true })).toBeChecked()
  await expect(dialog.getByRole("radio", { name: "Filtered entries (3)", exact: true })).toBeVisible()
  await expect(dialog).toContainText("Hidden entry explicitly included")
  await dialog.locator(".chronology-export-actions").getByRole("button", { name: "Close", exact: true }).click()
  await page.getByRole("searchbox").fill("Inspection")
  await open()
  await dialog.getByRole("radio", { name: "Filtered entries (1)", exact: true }).check()
  await expect(dialog.locator("tbody tr")).toHaveCount(1)
  await expect(dialog.getByRole("radio", { name: "All entries (2)", exact: true })).toBeVisible()
  assert.deepEqual(requests.at(-1)!.eventIds, ["inspection"])
  await dialog.getByRole("checkbox", { name: "Include hidden entries", exact: true }).check()
  await expect(dialog.locator("tbody tr")).toHaveCount(2)
  await expect(dialog.getByRole("radio", { name: "Filtered entries (2)", exact: true })).toBeChecked()
  await expect(dialog.getByRole("radio", { name: "All entries (3)", exact: true })).toBeVisible()
  assert.deepEqual([...requests.at(-1)!.eventIds!].sort(), ["hidden", "inspection"])
  conflict = true
  await dialog.getByRole("button", { name: "Save and download Word" }).click()
  await expect(dialog.getByRole("alert")).toContainText("The chronology changed")
  await expect(dialog.getByRole("button", { name: "Save and download Word" })).toBeDisabled()
  await dialog.getByRole("button", { name: "Refresh preview" }).click()
  await expect(dialog.getByRole("button", { name: "Save and download Word" })).toBeEnabled()
  const downloadPromise = page.waitForEvent("download")
  await dialog.getByRole("button", { name: "Save and download Word" }).click()
  const download = await downloadPromise
  assert.equal(download.suggestedFilename(), "Chronology-fixture.docx")
  assert.deepEqual(await readFile((await download.path())!), downloaded)
  assert.equal(
    await page.evaluate(() => (window as unknown as { __documentsRefreshed: number }).__documentsRefreshed),
    1,
  )
  const zip = unzipSync(downloaded)
  const xml = strFromU8(zip["word/document.xml"]!)
  assert.ok(xml.includes("Inspection correction"))
  assert.ok(!xml.includes("Collection reported"))
  assert.ok(xml.includes("May 2026 · Month only"), "Word table uses the compact date presentation")
  assert.ok(xml.includes("Date: 2026-05; May 2026; Date precision: month"), "Word notes retain the full date")
  assert.ok(strFromU8(zip["docProps/custom.xml"]!).includes("chronology-snapshot-v1"))
  if (process.env.CHRONOLOGY_EXPORT_QA_DIR)
    await Bun.write(path.join(process.env.CHRONOLOGY_EXPORT_QA_DIR, "chronology.docx"), downloaded)
  await expect(dialog.getByRole("status")).toContainText("Saved Chronology-fixture.docx")
  await dialog.locator(".chronology-export-actions").getByRole("button", { name: "Close", exact: true }).click()
  for (const width of [1440, 768, 375]) {
    await page.setViewportSize({ width, height: 900 })
    await open()
    await expect(dialog.getByRole("button", { name: "Save and download Word" })).toBeEnabled()
    const geometry = await dialog.evaluate((element) => ({
      width: element.getBoundingClientRect().width,
      scroll: element.scrollWidth,
      client: element.clientWidth,
      viewport: innerWidth,
    }))
    assert.ok(geometry.width <= geometry.viewport, `dialog fits at ${width}`)
    assert.ok(geometry.scroll <= geometry.client + 1, `no dialog overflow at ${width}`)
    const notesGap = await dialog.evaluate((element) => {
      const table = element.querySelector(".chronology-export-table-scroll")!.getBoundingClientRect()
      const notes = element.querySelector(".chronology-export-notes h3")!.getBoundingClientRect()
      return notes.top - table.bottom
    })
    assert.ok(notesGap >= 16 && notesGap <= 24, `evidence notes have a clear section gap at ${width}`)
    if (process.env.CHRONOLOGY_EXPORT_QA_DIR)
      await page.screenshot({ path: path.join(process.env.CHRONOLOGY_EXPORT_QA_DIR, `export-${width}.png`) })
    if (width === 375) {
      const table = dialog.getByRole("region", { name: "Exported chronology entries" })
      await table.scrollIntoViewIfNeeded()
      assert.ok(
        await table.evaluate((element) => element.scrollWidth <= element.clientWidth + 1),
        "phone preview fits without a hidden horizontal source column",
      )
      const dateWidth = await dialog
        .locator("th")
        .first()
        .evaluate((element) => element.getBoundingClientRect().width)
      assert.ok(dateWidth >= 80, "phone date column keeps a readable width")
      await expect(dialog.getByRole("columnheader", { name: "Source", exact: true })).toHaveCount(0)
      const source = dialog.locator(".chronology-export-source-inline").nth(1)
      await expect(source).toBeVisible()
      await expect(source).toContainText("Equipment handover record.pdf — Collection record")
      assert.ok(
        await source.evaluate((element) => {
          const source = element.getBoundingClientRect()
          const table = element.closest(".chronology-export-table-scroll")!.getBoundingClientRect()
          return source.left >= table.left && source.right <= table.right
        }),
        "source is beneath its event and within the phone preview width",
      )
      await table.focus()
      await expect(table).toBeFocused()
      if (process.env.CHRONOLOGY_EXPORT_QA_DIR)
        await page.screenshot({ path: path.join(process.env.CHRONOLOGY_EXPORT_QA_DIR, "export-375-table.png") })
    } else {
      await expect(dialog.getByRole("columnheader", { name: "Source", exact: true })).toBeVisible()
      await expect(dialog.locator(".chronology-export-source-inline").first()).toBeHidden()
    }
    if (width !== 375)
      await dialog.locator(".chronology-export-actions").getByRole("button", { name: "Close", exact: true }).click()
  }
  await page.keyboard.press("Escape")
  await expect(dialog).toHaveCount(0)
  assert.deepEqual(errors, [])
  console.log(
    "Chronology export: scope, hidden opt-in, qualifications, stale revision, exact DOCX download, keyboard and responsive checks passed.",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
