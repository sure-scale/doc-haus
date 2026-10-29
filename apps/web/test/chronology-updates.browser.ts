// Focused review-layout regression: independent synthetic proposals; no model or saved matter writes.
import assert from "node:assert/strict"
import { existsSync } from "node:fs"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Chronology, ChronologyAccount, ChronologyEvent, ChronologyProposal } from "../src/api/ingest/chronology"
const { chromium, expect } = await import("@playwright/test")
const date = { original: "8 April 2037", precision: "exact" as const, start: "2037-04-08" }
function account(id: string, quote: string): ChronologyAccount {
  return {
    id,
    title: "Calibration unit received",
    detail: "The warehouse recorded a single receipt.",
    date,
    evidence: [
      {
        id: `evidence-${id}`,
        documentName: `${id}.md`,
        sourceId: id,
        version: "synthetic-version",
        excerpt: quote,
        section: "Receiving",
        charStart: 0,
        charEnd: quote.length,
        verified: true,
        freshness: "current",
        provenance: "source",
      },
    ],
    privilege: "cleared",
    significance: "relevant",
    people: [],
    status: "reported",
    gaps: [],
  }
}
function event(id: string, accounts: ChronologyAccount[]): ChronologyEvent {
  return {
    id,
    title: "Calibration unit received",
    detail: "The warehouse recorded one calibration unit arriving on 8 April 2037.",
    date,
    accounts,
    origin: "generated",
    reviewed: true,
    opened: true,
    hidden: false,
    key: true,
    disputed: false,
    note: "Check the packing seal with the client.",
    lawyerEdited: true,
    createdAt: 1,
    updatedAt: 2,
    previousIds: [],
  }
}
const first = account("warehouse-register", "The calibration unit arrived on 8 April 2037."),
  second = account("carrier-confirmation", "Our van delivered the calibration unit to the warehouse on 8 April 2037."),
  extra = account(
    "additional-receiving-note",
    "The receiving supervisor confirms that the single calibration unit arrived on 8 April 2037.",
  )
const saved = event("saved-receipt", [first, second]),
  evidenceOnly = { ...structuredClone(saved), accounts: [first, second, extra] }
const changed = {
  ...structuredClone(saved),
  title: "Calibration case delivered",
  detail: "The proposed account places arrival on the following day.",
  date: { ...date, original: "9 April 2037", start: "2037-04-09" },
}
const disputed = {
  ...structuredClone(changed),
  disputed: true,
  dispute: "Carrier and warehouse dates differ.",
  accounts: [
    first,
    {
      ...extra,
      id: "competing-account",
      date: changed.date,
      privilege: "uncertain" as const,
      evidence: [
        {
          ...extra.evidence[0]!,
          id: "competing-evidence",
          freshness: "stale" as const,
          excerpt: "The driver records arrival on 9 April 2037.",
        },
      ],
    },
  ],
}
const future = event("future-check", [
  {
    ...account("inspection-letter", "A joint inspection is booked for 18 April 2037; it has not yet occurred."),
    status: "scheduled",
    date: { ...date, original: "18 April 2037", start: "2037-04-18" },
  },
])
future.title = "Joint inspection appointment"
future.date = future.accounts[0]!.date
future.detail = "A future joint inspection is booked."
future.reviewed = false
future.lawyerEdited = false
const other = {
  ...event("saved-handover", [account("handover-note", "The unit was handed to the warehouse on 8 April 2037.")]),
  title: "Warehouse handover",
}
function proposal(
  id: string,
  kind: ChronologyProposal["kind"],
  events: ChronologyEvent[],
  targetIds: string[] = [saved.id],
): ChronologyProposal {
  return {
    id,
    kind,
    targetIds,
    baseEventHashes: {},
    events,
    reason: "The additional source reports the receiving sequence.",
    recommendation: "Check the passage and retain competing accounts where they differ.",
    status: "pending",
    origin: "refresh",
    createdAt: 3,
  }
}
const hiddenA = { ...saved, id: "hidden-a", hidden: true },
  hiddenB = { ...other, id: "hidden-b", hidden: true }
const fixtures: ChronologyProposal[] = [
  proposal("evidence-only", "amendment", [evidenceOnly]),
  proposal("changed-fields", "amendment", [changed]),
  proposal("conflicting-account", "conflict", [disputed]),
  proposal("scheduled-addition", "addition", [future], []),
  proposal("hide-record", "removal", []),
  proposal(
    "merge-records",
    "merge",
    [{ ...saved, id: "merged", accounts: [...saved.accounts, ...other.accounts], previousIds: [saved.id, other.id] }],
    [saved.id, other.id],
  ),
  proposal("split-record", "split", [saved, other]),
  proposal("partial-accounts", "amendment", [{ ...saved, accounts: [extra] }]),
  proposal(
    "hidden-merge",
    "merge",
    [{ ...saved, id: "hidden-merged", hidden: false, key: false }],
    [hiddenA.id, hiddenB.id],
  ),
]
const state: Chronology = {
  schemaVersion: 1,
  revision: 1,
  created: true,
  events: [saved, other, hiddenA, hiddenB],
  coverage: [],
  proposals: fixtures,
  sourceDates: {},
  suggestion: { status: "relevant", reason: "Synthetic review" },
  history: [],
  framing: "Invented calibration equipment matter",
  gaps: [],
}
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-updates-review-")),
  entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement as h, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ProposalReview from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/chronology/ProposalReview.tsx"))}
import Surface from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/chronology/ChronologySurface.tsx"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/primitives.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/modal-frost.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/chronology/chronology.css"))}
const initial = ${JSON.stringify(state)};
window.__decisions=[];window.__sources=[];window.__asks=[]
function Harness(){const [open,setOpen]=useState(false),[revision,setRevision]=useState(1);return h('main',null,h('button',{onClick:()=>setOpen(true)},'Open updates'),open&&h(Surface,{title:'Review chronology updates',onClose:()=>setOpen(false)},h(ProposalReview,{chronology:{...initial,revision},busy:false,onClose:()=>setOpen(false),onEvidence:e=>window.__sources.push(e),onAsk:(e,p)=>window.__asks.push({e,p}),onDecision:async(p,d)=>{window.__decisions.push({id:p.id,decision:d});p.status=d==='accept'?'accepted':d==='keep'?'kept':'deferred';setRevision(r=>r+1);return true},onBatch:async(ids)=>{window.__decisions.push({batch:ids});return true},onDiscard:async()=>{window.__decisions.push({discard:true});return true}})))}
createRoot(document.getElementById('root')).render(h(Harness))
`,
)
const build = await Bun.build({
  entrypoints: [entry],
  target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://updates-fixture.invalid" }) },
})
assert.equal(build.success, true, build.logs.join("\n"))
const js = await build.outputs.find((o) => o.kind === "entry-point")!.text(),
  css = await build.outputs.find((o) => o.path.endsWith(".css"))!.text()
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch(request) {
    const name = new URL(request.url).pathname
    return name === "/harness.js"
      ? new Response(js, { headers: { "content-type": "text/javascript" } })
      : name === "/harness.css"
        ? new Response(css, { headers: { "content-type": "text/css" } })
        : new Response(
            '<!doctype html><html><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>',
            { headers: { "content-type": "text/html" } },
          )
  },
})
const chrome = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  executablePath = process.env.PLAYWRIGHT_CHROMIUM_PATH ?? (existsSync(chrome) ? chrome : undefined),
  browser = await chromium.launch({ headless: true, ...(executablePath ? { executablePath } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, reducedMotion: "reduce" }),
    errors: string[] = []
  page.on("pageerror", (e) => errors.push(e.message))
  page.setDefaultTimeout(7000)
  await page.goto(server.url.href)
  await page.getByRole("button", { name: "Open updates", exact: true }).click()
  const dialog = page.getByRole("dialog", { name: "Review chronology updates", exact: true }),
    changes = dialog.getByRole("region", { name: "What changes", exact: true }),
    footer = dialog.locator(".update-review-decision-area"),
    scroll = dialog.locator(".update-review-scroll")
  await expect(changes).toContainText("No change to event wording or date")
  await expect(changes).toContainText("Added supporting evidence")
  await expect(changes.getByText(extra.evidence[0]!.excerpt, { exact: true })).toBeVisible()
  for (const quote of await dialog.getByText(first.evidence[0]!.excerpt, { exact: true }).all())
    await expect(quote).not.toBeVisible()
  const full = dialog.locator("details").filter({ has: page.locator("summary", { hasText: "See full comparison" }) })
  await expect(full).not.toHaveAttribute("open", "")
  for (const width of [1440, 768, 375]) {
    await page.setViewportSize({ width, height: 900 })
    await expect(footer.getByRole("button", { name: "Accept this change", exact: true })).toBeInViewport()
    const start = (await footer.boundingBox())!
    await full.locator(":scope > summary").click()
    await scroll.evaluate((e) => (e.scrollTop = e.scrollHeight))
    const end = (await footer.boundingBox())!
    assert.ok(Math.abs(start.y - end.y) < 1, "Decision footer stays in its own layout row while full evidence scrolls")
    const bounds = (await scroll.boundingBox())!
    assert.ok(bounds.y + bounds.height <= end.y + 1, "Scrollable content cannot render behind decisions")
    await expect(footer.getByRole("button", { name: "Keep current", exact: true })).toBeInViewport()
    assert.ok((await dialog.boundingBox())!.width <= width)
    if (process.env.CHRONOLOGY_EVIDENCE_DIR)
      await page.screenshot({
        path: path.join(process.env.CHRONOLOGY_EVIDENCE_DIR, `updates-review-expanded-${width}.png`),
      })
    await full.locator(":scope > summary").click()
    await scroll.evaluate((e) => (e.scrollTop = 0))
  }
  await page.setViewportSize({ width: 1440, height: 900 })
  await changes.getByRole("button", { name: "Inspect source", exact: true }).click()
  assert.equal(await page.evaluate(() => (window as any).__sources[0].id), extra.evidence[0]!.id)
  await dialog.getByRole("button", { name: "Next update", exact: true }).click()
  await expect(changes).toContainText(changed.title)
  await expect(changes).toContainText(changed.detail)
  await expect(changes).toContainText("9 April 2037")
  await expect(changes).not.toContainText("No change to event wording or date")
  await dialog.getByRole("button", { name: "Next update", exact: true }).click()
  await expect(dialog.getByText("Conflict / dispute", { exact: false }).first()).toBeVisible()
  await expect(dialog.getByText(/Stale or unavailable support/).first()).toBeVisible()
  await expect(dialog.getByText(/Privilege: lawyer review needed/).first()).toBeVisible()
  await expect(changes).toContainText("9 April 2037")
  await dialog.getByRole("button", { name: "Next update", exact: true }).click()
  await expect(dialog.getByText("Scheduled", { exact: false }).first()).toBeVisible()
  await expect(dialog).toContainText(future.title)
  await dialog.getByRole("button", { name: "Next update", exact: true }).click()
  await expect(dialog).toContainText("history")
  await expect(dialog.getByRole("region", { name: "Proposed result", exact: true })).toHaveCount(0)
  for (const kind of ["merge", "split"]) {
    await dialog.getByRole("button", { name: "Next update", exact: true }).click()
    await expect(dialog).toContainText(kind === "merge" ? "Combine these events" : "Separate this event")
    const comparison = dialog
      .locator("details")
      .filter({ has: page.locator("summary", { hasText: "See full comparison" }) })
    await comparison.locator(":scope > summary").click()
    await expect(dialog.getByRole("region", { name: "Current saved entry", exact: true })).toBeVisible()
    await expect(dialog.getByRole("region", { name: "Proposed result", exact: true })).toBeVisible()
  }
  await dialog.getByRole("button", { name: "Next update", exact: true }).click()
  await expect(changes).toContainText("absent or changed")
  await dialog.getByText("See full comparison", { exact: true }).click()
  const partial = dialog.getByRole("region", { name: "Proposed result", exact: true })
  await expect(partial).toContainText("Proposed content")
  await expect(partial).not.toContainText("If you accept")
  await expect(dialog).toContainText("retained")
  await dialog.getByRole("button", { name: "Next update", exact: true }).click()
  await expect(changes.getByText(/This entry is hidden/).first()).toBeVisible()
  await dialog.getByLabel("More update actions", { exact: true }).click()
  await expect(dialog.getByText("Choose which to add as unreviewed.", { exact: true })).toBeVisible()
  await expect(dialog.getByText("Keep your saved chronology unchanged.", { exact: true })).toBeVisible()
  await page.keyboard.press("Escape")
  await expect(dialog).toBeVisible()
  assert.deepEqual(
    await page.evaluate(() => (window as any).__decisions),
    [],
    "Viewing, expanding and navigating never applies a decision",
  )
  await footer.getByRole("button", { name: "Skip for later", exact: true }).click()
  assert.deepEqual(await page.evaluate(() => (window as any).__decisions), [{ id: "hidden-merge", decision: "defer" }])
  await page.keyboard.press("Escape")
  await expect(dialog).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Open updates", exact: true })).toBeFocused()
  assert.deepEqual(errors, [])
  console.log(
    "Chronology Updates browser passed evidence-only, field changes, conflict, scheduled addition, removal, merge/split, separate footer scrolling and1440/768/375px; no model.",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
