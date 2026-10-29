// Focused EventForm + native dialog + shared picker regression; scripted preferences/previews, no model.
// Run from apps/web: bun test/chronology-forms.browser.ts
import assert from "node:assert/strict"
import { existsSync } from "node:fs"
import { mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { ChronologyAction } from "../src/api/ingest/chronology"

const { chromium, expect } = (await import(
  process.env.PLAYWRIGHT_MODULE ?? "@playwright/test"
)) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chronology-forms-"))
const source = path.resolve(import.meta.dir, "../src")
const imports = [...(await readFile(path.join(source, "main.tsx"), "utf8")).matchAll(/import "(\.\/[^"\n]+\.css)"/g)]
const css =
  (await Promise.all(imports.map(([, file]) => readFile(path.resolve(source, file!), "utf8")))).join("\n") +
  (await readFile(path.join(source, "components/chronology/chronology.css"), "utf8"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement as h, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import EventForm from ${JSON.stringify(path.join(source, "components/chronology/EventForm.tsx"))}
function account(id, title, documentName) {
 return { id, title, detail: "A source reports this occurrence.", date: { original: "6 May 2026", precision: "exact", start: "2026-05-06" }, evidence: [{id: "evidence-"+id, documentName, sourceId: "source-"+id, version: "fixture-version", excerpt: "The sealed equipment was collected on 6 May 2026.", section: "Collection", charStart: 0, charEnd: "The sealed equipment was collected on 6 May 2026.".length, verified: true, freshness: "current", provenance: "source"}], privilege: "uncertain", significance: "relevant", people: ["Carrier"], status: "reported", gaps: [] }
}
const collection = { id: "collection", title: "Equipment collected", detail: "Collection and signed receipt are separate occurrences.", date: { original: "6 May 2026", precision: "exact", start: "2026-05-06", time: "10:30", timezone: "Ship's time" }, accounts: [account("carrier", "Carrier collection account", "Carrier note.md"), account("receipt", "Signed receipt account", "Receipt.md")], origin: "generated", reviewed: true, opened: true, hidden: false, key: false, disputed: false, note: "Check the signed receipt.", lawyerEdited: false, createdAt: 1, updatedAt: 1, previousIds: [] }
const arrival = { ...collection, id: "arrival", title: "Equipment arrived", date: { original: "8 May 2026", precision: "exact", start: "2026-05-08" }, accounts: [account("depot", "Depot arrival account", "Depot note.md")], note: "" }
const events = [collection, arrival]
window.__formWrites = []; window.__formPreviews = []
function Harness() {
 const [mode, setMode] = useState(null)
 const [savedDate, setSavedDate] = useState(collection.date.start)
 return h("main", null,
  h("button", { onClick: () => setMode("edit") }, "Open edit"),
  h("button", { onClick: () => setMode("merge") }, "Open merge"),
  h("button", { onClick: () => setMode("split") }, "Open split"),
  h("output", { id: "saved-date" }, savedDate),
  mode && h(EventForm, { key: mode, mode, event: collection, events, busy: false, onClose: () => setMode(null),
   onSave: async action => { window.__formWrites.push(structuredClone(action)); if (action.type === "edit") setSavedDate(action.date.start); return true },
   onPreview: async action => {
    window.__formPreviews.push(structuredClone(action))
    if (action.type === "merge") return [{ ...collection, id: "merged", title: action.title, detail: action.detail, date: action.date, accounts: events.filter(item => action.eventIds.includes(item.id)).flatMap(item => item.accounts), reviewed: false, previousIds: action.eventIds }]
    return action.parts.map((part, index) => ({ ...collection, ...part, id: "split-"+index, accounts: collection.accounts.filter(account => part.accountIds.includes(account.id)), reviewed: false, previousIds: [collection.id] }))
   }
  }))
}
createRoot(document.getElementById("root")).render(h(Harness))
`,
)
const build = await Bun.build({
  entrypoints: [entry],
  target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://chronology-forms.invalid" }) },
})
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch(request) {
    const pathname = new URL(request.url).pathname
    if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (pathname === "/product.css") return new Response(css, { headers: { "content-type": "text/css" } })
    return new Response(
      '<!doctype html><html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/product.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>',
      { headers: { "content-type": "text/html" } },
    )
  },
})
const installedChrome = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
const executablePath =
  process.env.PLAYWRIGHT_CHROMIUM_PATH ?? (existsSync(installedChrome) ? installedChrome : undefined)
const browser = await chromium.launch({ headless: true, ...(executablePath ? { executablePath } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 1100, height: 900 }, reducedMotion: "reduce" })
  page.setDefaultTimeout(7000)
  const errors: string[] = []
  page.on("pageerror", (error) => errors.push(error.message))
  await page.route("http://chronology-forms.invalid/**", (route) => {
    assert.equal(new URL(route.request().url()).pathname, "/preferences", "only drafting preferences are requested")
    return route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify({ dateFormat: "day-month-year" }),
    })
  })
  const writes = () => page.evaluate(() => (window as unknown as { __formWrites: ChronologyAction[] }).__formWrites)
  const previews = () =>
    page.evaluate(() => (window as unknown as { __formPreviews: ChronologyAction[] }).__formPreviews)
  await page.goto(server.url.href)
  await page.getByRole("button", { name: "Open edit", exact: true }).click()
  const edit = page.getByRole("dialog", { name: "Edit event", exact: true })
  const date = edit.getByRole("textbox", { name: "Date", exact: true })
  const confirm = edit.getByRole("button", { name: "Confirm event", exact: true })
  await expect(date).toHaveAttribute("placeholder", "dd/mm/yyyy")
  await expect(date).toHaveValue("06/05/2026")
  await date.fill("07052026")
  await expect(date).toHaveValue("07/05/2026")
  await expect(confirm).toBeEnabled()
  for (const invalid of ["0805", "31022026", ""]) {
    await date.fill(invalid)
    await expect(date).toHaveAttribute("aria-invalid", "true")
    await expect(confirm).toBeDisabled()
    await edit.locator("form").evaluate((form) => (form as HTMLFormElement).requestSubmit())
    assert.equal((await writes()).length, 0, "invalid/incomplete date cannot save the retained ISO")
    await expect(page.locator("#saved-date")).toHaveText("2026-05-06")
  }
  await date.fill("07052026")
  await expect(confirm).toBeEnabled()
  await edit
    .locator("summary")
    .filter({ hasText: /^Date details$/ })
    .click()
  const timezone = edit.getByRole("button", { name: "Timezone", exact: true })
  await expect(timezone).toContainText("Ship's time (as recorded)")
  await timezone.click()
  const zoneSearch = edit.getByRole("textbox", { name: "Search timezones", exact: true })
  await expect(zoneSearch).toBeFocused()
  assert.ok(
    await zoneSearch.evaluate((node) => Boolean(node.closest("dialog[open]"))),
    "search portal stays in the native dialog",
  )
  await zoneSearch.fill("UTC")
  await page.keyboard.press("Escape")
  await expect(zoneSearch).toHaveCount(0)
  await expect(edit).toBeVisible()
  await expect(timezone).toBeFocused()
  await timezone.click()
  await zoneSearch.fill("UTC")
  const utc = edit.getByRole("radio", { name: "UTC", exact: true })
  await utc.focus()
  await page.keyboard.press("Space")
  await expect(zoneSearch).toHaveCount(0)
  await expect(timezone).toBeFocused()
  await expect(timezone).toContainText("UTC")
  await confirm.click()
  await expect(edit).toHaveCount(0)
  const correction = (await writes())[0]!
  assert.equal(correction.type, "edit")
  if (correction.type !== "edit") throw new Error("Expected correction")
  assert.equal(correction.date.start, "2026-05-07")
  assert.equal(correction.date.timezone, "UTC")
  assert.equal(correction.date.time, "10:30")
  await expect(page.locator("#saved-date")).toHaveText("2026-05-07")
  console.log(
    "PASS saved dd/mm/yyyy preference, slash insertion, invalid/incomplete save rejection, timezone search/native focus/Escape and valid correction",
  )

  await page.getByRole("button", { name: "Open merge", exact: true }).click()
  const merge = page.getByRole("dialog", { name: "Merge events", exact: true })
  await expect(merge.getByRole("textbox", { name: "Event", exact: true })).toHaveCount(0)
  await expect(merge.getByRole("textbox", { name: "Description", exact: true })).toHaveCount(0)
  await expect(merge.getByRole("combobox", { name: "Precision", exact: true })).toHaveCount(0)
  const mergeWith = merge.getByRole("button", { name: "Merge with", exact: true })
  await mergeWith.click()
  const mergeSearch = merge.getByRole("textbox", { name: "Search events to merge", exact: true })
  await expect(mergeSearch).toBeFocused()
  await mergeSearch.fill("arrived")
  const target = merge.getByRole("radio", { name: "8 May 2026 · Equipment arrived", exact: true })
  await target.focus()
  await page.keyboard.press("Space")
  await expect(mergeSearch).toHaveCount(0)
  await expect(mergeWith).toBeFocused()
  await merge.getByRole("button", { name: "Preview changes", exact: true }).click()
  await expect(merge.getByRole("region", { name: "Resulting events" })).toContainText("3 accounts")
  assert.equal((await previews()).length, 1)
  assert.equal((await writes()).length, 1, "preview does not apply a merge")
  await merge
    .locator("summary")
    .filter({ hasText: /^Adjust resulting event$/ })
    .click()
  await merge.getByRole("textbox", { name: "Event", exact: true }).fill("Collection and arrival reported")
  await expect(merge.getByRole("region", { name: "Resulting events" })).toHaveCount(0)
  await expect(merge.getByRole("button", { name: "Confirm merge", exact: true })).toHaveCount(0)
  await merge.getByRole("button", { name: "Preview changes", exact: true }).click()
  assert.equal((await previews()).length, 2)
  await merge.getByRole("button", { name: "Confirm merge", exact: true }).click()
  await expect(merge).toHaveCount(0)
  const merged = (await writes())[1]!
  assert.equal(merged.type, "merge")
  if (merged.type !== "merge") throw new Error("Expected merge")
  assert.deepEqual(merged.eventIds, ["collection", "arrival"])
  assert.equal(merged.title, "Collection and arrival reported")
  console.log(
    "PASS minimal merge, searchable event selection, mandatory preview and changed-field preview invalidation",
  )

  await page.getByRole("button", { name: "Open split", exact: true }).click()
  const split = page.getByRole("dialog", { name: "Split event", exact: true })
  await split.getByRole("textbox", { name: "First event", exact: true }).fill("Equipment collection")
  await split.getByRole("textbox", { name: "Second event", exact: true }).fill("Signed receipt")
  const carrier = split.getByRole("checkbox", { name: /Carrier collection account/ })
  const receipt = split.getByRole("checkbox", { name: /Signed receipt account/ })
  await expect(carrier).toBeChecked()
  await expect(receipt).not.toBeChecked()
  await receipt.check()
  await carrier.uncheck()
  await split.getByRole("button", { name: "Preview changes", exact: true }).click()
  await expect(split.getByRole("region", { name: "Resulting events" }).locator("article")).toHaveCount(2)
  assert.equal((await writes()).length, 2, "split preview does not save")
  await split.getByRole("button", { name: "Confirm split", exact: true }).click()
  await expect(split).toHaveCount(0)
  const separated = (await writes())[2]!
  assert.equal(separated.type, "split")
  if (separated.type !== "split") throw new Error("Expected split")
  assert.deepEqual(separated.parts[0]!.accountIds, ["receipt"])
  assert.deepEqual(separated.parts[1]!.accountIds, ["carrier"])
  assert.deepEqual(separated.parts.flatMap((part) => part.accountIds).sort(), ["carrier", "receipt"])
  assert.equal(separated.parts[0]!.title, "Equipment collection")
  assert.equal(separated.parts[1]!.title, "Signed receipt")
  assert.equal(separated.parts[1]!.date.start, "2026-05-06")
  assert.deepEqual(errors, [])
  console.log(
    "PASS split titles/source allocation with every original account retained; no browser errors (scripted previews, no model)",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
