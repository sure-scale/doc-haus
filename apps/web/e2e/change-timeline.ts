// Change timeline card (#297) against a separately started Vite server. The
// runner starts a real ingest service on DOCHAUS_E2E_INGEST_PORT with a
// temporary data root and settings directory and uploads four synthetic Word
// sources: the marked-up supply agreement
// (dochaus/test/support/counterparty-markup-fixture.ts), a long clause marked up
// by three fictional reviewers, a file whose text box shifts the reader's
// paragraph numbers, and a file with list numbers, a changed table and row and
// two changed sections (the last three built here). It drives the card on the
// painted page: every marked paragraph's card against the markup route's list,
// the hover delay, the pointer crossing into the card, Escape, placement below,
// above, right and left, a long timeline scrolled by wheel and keyboard,
// counterparty text shown as text, touch at 400 px, a dark colour scheme, the
// comments-only fallback, drags and clicks that must not open or pin the card,
// Tab and Shift+Tab out of the card, closing when its line scrolls away, the
// marks' accessible description, changed containers and sections, and long
// list numbers. Nothing reaches a model: the sources' embeddings use the local
// model (DOCHAUS_MODEL_CACHE_DIR, else the per-user cache), browser engine
// traffic is intercepted and unexpected requests fail the run. The JSON summary
// goes beneath DOCHAUS_EVIDENCE_DIR, screenshots beneath DOCHAUS_SCREENSHOT_DIR
// (default: the evidence directory).
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises"
import os, { tmpdir } from "node:os"
import path from "node:path"
import type { Browser, BrowserContextOptions, Locator, Page } from "@playwright/test"
import type { DocumentMarkup, MarkupComment, MarkupLocation, MarkupRevision } from "../src/api/ingest/markup"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const { zipSync, strToU8 } = (await import(Bun.resolveSync("fflate", path.resolve(import.meta.dir, "../../../dochaus")))) as {
  zipSync: (files: Record<string, Uint8Array>, options?: { mtime?: Date }) => Uint8Array; strToU8: (text: string) => Uint8Array
}

const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
const ingestPort = Number(process.env.DOCHAUS_E2E_INGEST_PORT)
if (!baseURL || !evidence || !Number.isInteger(ingestPort) || ingestPort < 1024)
  throw new Error("Set PLAYWRIGHT_BASE_URL, DOCHAUS_EVIDENCE_DIR and DOCHAUS_E2E_INGEST_PORT (the port Vite's /ingest proxy targets).")
const shots = process.env.DOCHAUS_SCREENSHOT_DIR || evidence
await mkdir(evidence, { recursive: true })
await mkdir(shots, { recursive: true })

const SUPPLY = COUNTERPARTY_MARKUP.name
const BUSY = "change-timeline-busy.docx"
const TEXT_BOX = "change-timeline-text-box.docx"
const STRUCTURES = "change-timeline-structures.docx"
const HINT = "Change history available. Press Enter to read it."
const FLAG = "Contains text that looks like an instruction — doc.haus did not follow it."
const DESKTOP = { width: 1440, height: 1000 }
const LAPTOP = { width: 1440, height: 760 }
const PHONE = { width: 400, height: 860 }
// The card's hover delay, and the gap and window margin its placement keeps.
const REST = 300, GAP = 8, MARGIN = 8
// A placement case is set up only when every side before the expected one
// misses, and the expected one fits, by at least this many pixels.
const SLACK = 24

// ---- Synthetic sources -------------------------------------------------------
// Every name, date and figure below is invented.
const [ROWAN, PRIYA, TOMASZ] = ["Rowan Achterberg", "Priya Halvorsen", "Tomasz Ellery"] as const
const UNTRUSTED = `<img src=x onerror="window.__timelineXss=1"><b>bold</b>`
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
const xml = (value: string) => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;")
let annotation = 100
const text = (value: string) => `<w:r><w:t xml:space="preserve">${xml(value)}</w:t></w:r>`
// Word leaves w:date out when the reviewer's clock was not recorded.
const by = (author: string, date?: string) => `w:author="${xml(author)}"${date ? ` w:date="${date}"` : ""}`
const inserted = (author: string, date: string | undefined, value: string) => `<w:ins w:id="${annotation++}" ${by(author, date)}>${text(value)}</w:ins>`
const deleted = (author: string, date: string | undefined, value: string) =>
  `<w:del w:id="${annotation++}" ${by(author, date)}><w:r><w:delText xml:space="preserve">${xml(value)}</w:delText></w:r></w:del>`
const paragraph = (content: string, properties = "") => `<w:p>${properties && `<w:pPr>${properties}</w:pPr>`}${content}</w:p>`
const commented = (id: number, content: string, replies: number[] = []) =>
  `<w:commentRangeStart w:id="${id}"/>${content}<w:commentRangeEnd w:id="${id}"/>${[id, ...replies].map(ref => `<w:r><w:commentReference w:id="${ref}"/></w:r>`).join("")}`
const on = (day: number, hour = 9) => `2026-09-${String(day).padStart(2, "0")}T${String(hour).padStart(2, "0")}:00:00Z`

type SyntheticComment = { id: number; author: string; date: string; text: string; parent?: number }
const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
// `numbering` is the numbering part's content; `sectionTail` ends the body's own section properties.
function syntheticDocx(paragraphs: string[], comments: SyntheticComment[], extra: { numbering?: string; sectionTail?: string } = {}) {
  const WM = "application/vnd.openxmlformats-officedocument.wordprocessingml"
  const OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const paraId = (id: number) => (0x30000001 + id).toString(16).toUpperCase()
  const initials = (author: string) => author.split(" ").map(part => part[0]).join("")
  // The reader resolves list numbers only with a styles part beside the numbering part.
  const listed = extra.numbering !== undefined
  const parts: Record<string, string> = {
    "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${WM}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${WM}.comments+xml"/>` +
      `<Override PartName="/word/commentsExtended.xml" ContentType="${WM}.commentsExtended+xml"/>` +
      (listed ? `<Override PartName="/word/numbering.xml" ContentType="${WM}.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="${WM}.styles+xml"/>` : "") + "</Types>",
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${OD}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${OD}/comments" Target="comments.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>' +
      (listed ? `<Relationship Id="rId3" Type="${OD}/numbering" Target="numbering.xml"/><Relationship Id="rId4" Type="${OD}/styles" Target="styles.xml"/>` : "") + "</Relationships>",
    "word/document.xml": `<w:document ${NS}><w:body>${paragraphs.join("")}<w:sectPr>${PAGE}${extra.sectionTail ?? ""}</w:sectPr></w:body></w:document>`,
    "word/comments.xml": `<w:comments ${NS}>${comments.map(comment =>
      `<w:comment w:id="${comment.id}" ${by(comment.author, comment.date)} w:initials="${initials(comment.author)}"><w:p w14:paraId="${paraId(comment.id)}">${text(comment.text)}</w:p></w:comment>`).join("")}</w:comments>`,
    "word/commentsExtended.xml": `<w15:commentsEx ${NS}>${comments.map(comment =>
      `<w15:commentEx w15:paraId="${paraId(comment.id)}"${comment.parent === undefined ? "" : ` w15:paraIdParent="${paraId(comment.parent)}"`} w15:done="0"/>`).join("")}</w15:commentsEx>`,
    ...listed ? {
      "word/numbering.xml": `<w:numbering ${NS}>${extra.numbering}</w:numbering>`,
      "word/styles.xml": `<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    } : {},
  }
  return zipSync(Object.fromEntries(Object.entries(parts).map(([name, body]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`)])), { mtime: new Date(1980, 0, 1) })
}

// Clause 1 carries sixteen dated changes by three reviewers, written in an order
// unlike their dates, one undated change and a comment thread; clause 2 a
// comment written as HTML. Forty plain paragraphs follow, then two medium
// clauses for placement: one whose first change opens it (the start of a line)
// and one right-aligned whose last change ends it (the end of a line), then
// plain paragraphs so either can be scrolled anywhere in the window.
// Plain text separates every change, so the reader keeps each one on its own.
function busyDocx() {
  const clauseOne = [
    text("1. Fees and payment. The Customer shall pay "), inserted(PRIYA, on(14), "each undisputed"),
    text(" invoice "), deleted(ROWAN, on(3), "in full"),
    text(" within "), inserted(TOMASZ, on(22), "forty"),
    text(" days "), deleted(PRIYA, on(8), "of the invoice date"),
    text(" of receipt, "), inserted(ROWAN, on(27), "by bank transfer"),
    text(" to the account "), deleted(TOMASZ, on(1), "notified by email"),
    text(" the Supplier names "), inserted(PRIYA, on(18), "in writing"),
    text(". Late amounts "), deleted(ROWAN, on(11), "shall"),
    text(" carry interest at "), commented(0, text("the agreed rate"), [1]),
    text(" "), inserted(TOMASZ, on(5), "of two per cent a year"),
    text(" from the due date "), deleted(PRIYA, on(25), "until paid"),
    text(". Either party may raise "), inserted(ROWAN, on(16), "a good-faith"),
    text(" dispute "), deleted(TOMASZ, on(29), "orally or"),
    text(" in writing "), inserted(PRIYA, on(10), "before the due date"),
    text(", and "), deleted(ROWAN, on(20), "both"),
    text(" the parties "), inserted(TOMASZ, on(7), "shall meet within ten days"),
    text(" to resolve it "), deleted(PRIYA, on(12), "promptly"),
    text(" and "), inserted(ROWAN, undefined, "without delay"), text("."),
  ].join("")
  const filler = (from: number, count: number) => Array.from({ length: count }, (_, index) =>
    paragraph(text(`${from + index}. Filler clause ${from + index}. This synthetic paragraph has no tracked changes; it only makes the page long enough to scroll the marked clauses into place.`)))
  return syntheticDocx([
    paragraph(clauseOne),
    paragraph(text("2. Dispute notices. ") + commented(2, text("A dispute notice must name the invoice concerned."))),
    ...filler(3, 40),
    paragraph(inserted(PRIYA, on(4, 13), "Service levels.") + text(" The Supplier shall meet the agreed service levels ") + deleted(ROWAN, on(9, 13), "in all material respects") +
      text(" and report on them ") + inserted(TOMASZ, on(15, 13), "each month") + text(", ") + deleted(PRIYA, on(19, 13), "quarterly") + text(", ") +
      inserted(ROWAN, on(26, 13), "with a short summary") + text(".")),
    paragraph(text("The parties may vary these service levels ") + inserted(TOMASZ, on(6, 13), "in writing") + text(" ") + deleted(PRIYA, on(17, 13), "or orally") +
      text(" by agreement; a variation takes effect ") + inserted(ROWAN, on(24, 13), "on the first day of") + text(" the month after ") +
      deleted(TOMASZ, on(28, 13), "its") + text(" ") + inserted(PRIYA, on(30, 13), "signature."), '<w:jc w:val="right"/>'),
    ...filler(45, 12),
  ], [
    { id: 0, author: PRIYA, date: "2026-09-23T15:30:00Z", text: "Is two per cent a year acceptable to the Customer's treasury team?" },
    { id: 1, author: ROWAN, date: "2026-09-24T08:45:00Z", text: "Treasury agreed two per cent in its note of 24 September.", parent: 0 },
    { id: 2, author: TOMASZ, date: "2026-09-26T11:00:00Z", text: UNTRUSTED },
  ])
}

// The reader counts a VML text box's paragraph; the page paints it inline, so
// every reader paragraph number after the box is one ahead of the page's.
function textBoxDocx() {
  const box = `<w:r><w:pict><v:shape style="width:160pt;height:40pt"><v:textbox><w:txbxContent>${paragraph(text("Boxed note") + inserted(PRIYA, on(2), " added"))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`
  return syntheticDocx([
    paragraph(text("1. Opening clause.") + inserted(ROWAN, on(1), " Agreed.")),
    paragraph(text("2. A clause with a text box: ") + box),
    paragraph(commented(0, text("3. After the box."), [1]) + inserted(ROWAN, on(3), " Tail added.")),
    paragraph(text("4. Last clause.")),
  ], [
    { id: 0, author: TOMASZ, date: "2026-09-09T10:00:00Z", text: "Is the boxed note part of the clause?" },
    { id: 1, author: PRIYA, date: "2026-09-10T12:00:00Z", text: "No, it is a drafting aid only.", parent: 0 },
  ])
}

// A clause with a short list number and one whose number is too long to name a
// place; a table whose properties changed, with an inserted row; a section
// break whose page setup changed; and a final section of two paragraphs whose
// page setup changed too. Each container or section change has its own author.
const [INES, MATTEO, ODILE, BASTIEN] = ["Ines Calloway", "Matteo Varga", "Odile Brandt", "Bastien Kerr"] as const
const LONG_LABEL = "Supplementary Provision 1"
function structuresDocx() {
  const list = (id: number, format: string) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
    `<w:lvlText w:val="${xml(format)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`
  const numbered = (num: number) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${num}"/></w:numPr>`
  const cell = (value: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph(text(value))}</w:tc>`
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"].map(side => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`).join("")
  const pageChange = (author: string, date: string) => `<w:sectPrChange w:id="${annotation++}" ${by(author, date)}><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:sectPrChange>`
  const table = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${borders}</w:tblBorders>` +
    `<w:tblPrChange w:id="${annotation++}" ${by(MATTEO, on(12, 10))}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
    `<w:tr>${cell("Rate card")}${cell("Fees as listed in the order.")}</w:tr>` +
    `<w:tr><w:trPr><w:ins w:id="${annotation++}" ${by(INES, on(14, 11))}/></w:trPr>${cell("Late fee")}${cell("Two per cent a month on overdue sums.")}</w:tr></w:tbl>`
  return syntheticDocx([
    paragraph(text("Payment terms apply ") + inserted(ROWAN, on(2), "from the start date") + text("."), numbered(1)),
    paragraph(text("Records are kept ") + inserted(PRIYA, on(4), "for six years") + text("."), numbered(2)),
    table,
    paragraph(text("End of the first section."), `<w:sectPr>${PAGE}${pageChange(ODILE, on(16, 12))}</w:sectPr>`),
    paragraph(text("The second section starts here.")),
    paragraph(text("The second section ends here.")),
  ], [], {
    numbering: list(1, "%1.") + list(2, LONG_LABEL.replace(/1$/, "%1")) +
      '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>',
    sectionTail: pageChange(BASTIEN, on(18, 15)),
  })
}

// ---- Isolated ingest and engine fixture ----------------------------------------
const results: { step: string; status: "passed" | "failed"; error?: string }[] = []
const pageErrors: string[] = []
const consoleErrors: string[] = []
const unexpected: string[] = []
const dialogs: string[] = []
const engineRequests: { method: string; path: string }[] = []
const ingestRequests: { method: string; path: string }[] = []
const stubbed: string[] = []
const measured: Record<string, unknown> = {}

// The service needs the engine only for its model list (readiness); anything
// else that reaches it is recorded.
const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: (request) => {
  const url = new URL(request.url)
  if (request.method === "GET" && url.pathname === "/api/model")
    return Response.json({ data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
  unexpected.push(`engine fixture ${request.method} ${url.pathname}`)
  return new Response(null, { status: 404 })
} })

const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-change-timeline-e2e-"))
const dataRoot = path.join(temporary, "data")
const settings = path.join(temporary, "settings")
await mkdir(path.join(dataRoot, "workspace", ".preferences"), { recursive: true })
await mkdir(settings, { recursive: true })
await writeFile(path.join(settings, "opencode.json"), JSON.stringify({ model: "fixture/primary", small_model: "fixture/fast" }))
await writeFile(path.join(dataRoot, "workspace", ".preferences", "preferences.json"), JSON.stringify({ firm: "Synthetic Firm" }))
const ingestURL = `http://127.0.0.1:${ingestPort}`
const api = (route: string, init?: RequestInit) => fetch(`${ingestURL}${route}`, { ...init, signal: AbortSignal.timeout(120_000) })
let ingest: ReturnType<typeof Bun.spawn> | undefined

async function portFree(port: number) {
  try { const server = Bun.listen({ hostname: "127.0.0.1", port, socket: { data() {} } }); server.stop(true); return true }
  catch { return false }
}
// Only explicit isolation settings reach the service: no inherited workspace,
// runtime, engine address or credentials from the caller's environment.
async function startIngest() {
  if (!await portFree(ingestPort)) throw new Error(`Port ${ingestPort} is occupied; choose a free DOCHAUS_E2E_INGEST_PORT.`)
  const log = Bun.file(path.join(temporary, "ingest.log"))
  ingest = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, "../../../services/ingest"), stdout: log, stderr: log,
    env: {
      PATH: process.env.PATH ?? "", HOME: process.env.HOME ?? temporary, TMPDIR: process.env.TMPDIR ?? tmpdir(),
      DOCHAUS_DATA_ROOT: dataRoot, WORKSPACE_ROOT: path.join(dataRoot, "workspace"), DOCHAUS_RUNTIME_DIR: path.join(dataRoot, "runtime"), DOCHAUS_DIR: settings, INGEST_HOST: "127.0.0.1", INGEST_PORT: String(ingestPort),
      OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "",
      // The local embedding model's weights are the same everywhere; reuse the download.
      DOCHAUS_MODEL_CACHE_DIR: process.env.DOCHAUS_MODEL_CACHE_DIR || path.join(os.homedir(), ".cache", "doc-haus", "models"),
    },
  })
  const deadline = Date.now() + 30_000
  while (!await fetch(`${ingestURL}/matters`).then(response => response.ok).catch(() => false)) {
    if (ingest.exitCode !== null || Date.now() > deadline) throw new Error(`Isolated ingest did not start; see ${log.name}`)
    await Bun.sleep(100)
  }
}
async function stopIngest() {
  if (!ingest) return
  ingest.kill("SIGTERM")
  await ingest.exited
  ingest = undefined
}

// A failed step is recorded with a screenshot and the run goes on: every later
// step opens its document afresh. Only a failed upload stops the run.
let active: Page | undefined
async function step(name: string, run: () => Promise<void>, fatal = false) {
  try { await run(); results.push({ step: name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) {
    results.push({ step: name, status: "failed", error: error instanceof Error ? error.message : String(error) })
    console.log(`FAIL ${name}\n  ${error instanceof Error ? error.message.split("\n").slice(0, 6).join("\n  ") : String(error)}`)
    await active?.screenshot({ path: path.join(shots, `failed-step-${name.split(" ")[0]}.png`) }).catch(() => undefined)
    if (fatal) throw error
  }
}
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(shots, `${name}.png`) })

// ---- Page helpers --------------------------------------------------------------
type Rect = { left: number; top: number; right: number; bottom: number }
type Box = { x: number; y: number; width: number; height: number }
type View = { width: number; height: number }
type Side = "below" | "above" | "right" | "left"
type Entry = { author: string; date: string | null; kind: string }
type Painted = { ordinal: number; section: string | null; comments: number[]; marks: number[]; text: string }

const cardOf = (page: Page) => page.locator("div.change-timeline[role=dialog]")
const settle = (page: Page) => page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))))
// A point no card ever covers (it keeps 8 px inside the window) and no mark sits on.
const leave = async (page: Page) => {
  await page.mouse.move(2, 2)
  await expect(cardOf(page)).toBeHidden()
}
async function hoverOpen(page: Page, target: Locator) {
  await target.hover()
  await expect(cardOf(page)).toBeVisible()
}
async function cardBox(page: Page): Promise<Box> {
  const box = await cardOf(page).boundingBox()
  if (!box) throw new Error("The card has no box.")
  return box
}
async function lineBoxes(target: Locator): Promise<Rect[]> {
  const rects = await target.evaluate(node => [...node.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0)
    .map(rect => ({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom })))
  if (!rects.length) throw new Error("The mark has no line box.")
  return rects
}
const inside = (box: Box, view: View, margin = 0) =>
  box.x >= margin - 0.5 && box.y >= margin - 0.5 && box.x + box.width <= view.width - margin + 0.5 && box.y + box.height <= view.height - margin + 0.5
// The viewer's document pane as it shows on screen: below the viewer's toolbar,
// within its scrolling body (the page's own scroll box on a wide screen, the
// body holding the page and Their changes on a narrow one) and the window.
type Pane = { top: number; bottom: number; toolbarBottom: number }
async function paneOf(page: Page): Promise<Pane> {
  const bar = await page.locator(".viewer-panel > .viewer-bar").boundingBox()
  const body = await page.locator(".viewer-panel > .viewer-body").boundingBox()
  if (!bar || !body) throw new Error("The viewer has no toolbar or body box.")
  const toolbarBottom = bar.y + bar.height
  return { toolbarBottom, top: Math.max(toolbarBottom, body.y), bottom: Math.min(body.y + body.height, page.viewportSize()!.height) }
}
// An open card never covers the viewer's toolbar and ends within the pane's
// visible part (1 px for rounding).
const inPane: Record<string, { card: Box } & Pane> = {}
async function expectInPane(page: Page, label: string) {
  const box = await cardBox(page), pane = await paneOf(page)
  inPane[label] = { card: box, ...pane }
  expect(box.y, `${label}: the card's top ${box.y} is not above the toolbar's bottom ${pane.toolbarBottom} or the pane's top ${pane.top}`).toBeGreaterThanOrEqual(pane.top - 1)
  expect(box.y + box.height, `${label}: the card's bottom ${box.y + box.height} is within the pane's visible bottom ${pane.bottom}`).toBeLessThanOrEqual(pane.bottom + 1)
}
// The screenshot of an open card, once it is checked to sit in the pane.
async function cardShot(page: Page, name: string) {
  await expectInPane(page, name)
  await shot(page, name)
}
const besides = (side: Side, box: Box, line: Rect) => ({
  below: box.y >= line.bottom - 0.5,
  above: box.y + box.height <= line.top + 0.5,
  right: box.x >= line.right - 0.5,
  left: box.x + box.width <= line.left + 0.5,
})[side]
// How many pixels the card has to spare on each side of a line box (negative:
// how far it misses); left and right also need the pane's height, below and
// above the window's width. Vertically the card keeps to the pane.
function spare(line: Rect, size: View, view: View, pane: Pane): Record<Side, number> {
  const across = view.width - 2 * MARGIN - size.width, tall = pane.bottom - pane.top - 2 * MARGIN - size.height
  return {
    below: Math.min(pane.bottom - MARGIN - line.bottom - GAP - size.height, across),
    above: Math.min(line.top - GAP - MARGIN - pane.top - size.height, across),
    right: Math.min(view.width - MARGIN - line.right - GAP - size.width, tall),
    left: Math.min(line.left - GAP - MARGIN - size.width, tall),
  }
}
// The side the card must take: every earlier side clearly misses, it clearly fits.
function clearSide(room: Record<Side, number>): Side | undefined {
  for (const side of ["below", "above", "right", "left"] as const) {
    if (room[side] >= SLACK) return side
    if (room[side] > -SLACK) return undefined
  }
  return undefined
}

// Scrolls the viewer's page so the chosen line box of `target` is centred on
// window height `y`, as near as the scroll range allows, and returns it.
async function centreLine(page: Page, doc: Locator, target: Locator, pick: "first" | "last", y: number) {
  const line = async () => { const rects = await lineBoxes(target); return pick === "first" ? rects[0] : rects[rects.length - 1] }
  for (let attempt = 0; attempt < 4; attempt++) {
    const current = await line()
    const delta = (current.top + current.bottom) / 2 - y
    if (Math.abs(delta) < 1) break
    await doc.evaluate((node, by) => { node.scrollTop += by }, delta)
    await settle(page)
  }
  return line()
}
// Rests the mouse on the middle of a line box, which must be the mark's own pixels.
async function restOn(page: Page, target: Locator, line: Rect) {
  const x = (line.left + line.right) / 2, y = (line.top + line.bottom) / 2
  const hit = await target.evaluate((node, [px, py]) => { const found = document.elementFromPoint(px, py); return Boolean(found && node.contains(found)) }, [x, y] as const)
  if (!hit) throw new Error(`Setup: the mark is not the element under (${x}, ${y}).`)
  await page.mouse.move(x, y, { steps: 2 })
  await expect(cardOf(page)).toBeVisible()
}
// The card's own size, opened with room to spare below a line near the top.
async function naturalSize(page: Page, doc: Locator, target: Locator, pick: "first" | "last"): Promise<View> {
  const top = (await doc.boundingBox())!.y
  const line = await centreLine(page, doc, target, pick, top + 120)
  await restOn(page, target, line)
  const body = cardOf(page).locator(".change-timeline-body")
  if (await body.evaluate(node => node.scrollHeight > node.clientHeight + 1)) throw new Error("Setup: the card scrolls even with room below; it is no medium card.")
  const box = await cardBox(page)
  await leave(page)
  return { width: box.width, height: box.height }
}
// A short window's height: its pane too short for a card of `size` above or
// below a line at the pane's mid-height, tall enough for it beside the line.
// The viewer's toolbar and margins are measured at the current size.
async function shortHeight(page: Page, size: View) {
  const view = page.viewportSize()!, pane = await paneOf(page)
  return Math.max(360, Math.ceil(view.height - (pane.bottom - pane.top) + size.height + 60))
}
// Opens the card at `line` and checks the side it took. The expected side is
// first checked again with the card's real size, so a wrong premise reads as a
// setup failure rather than a placement fault.
async function expectPlacement(page: Page, label: string, target: Locator, line: Rect, expected: Side, margin = 0) {
  await restOn(page, target, line)
  const view = page.viewportSize()!
  const box = await cardBox(page), pane = await paneOf(page)
  const room = spare(line, box, view, pane)
  const placement = await cardOf(page).getAttribute("data-placement")
  measured[label] = { view, pane, line, card: box, room, placement }
  if (clearSide(room) !== expected) throw new Error(`Setup: with the card's real size the clear side is ${clearSide(room) ?? "ambiguous"}, not ${expected} (card took ${placement}): ${JSON.stringify(room)}`)
  await expect(cardOf(page)).toHaveAttribute("data-placement", expected)
  expect(inside(box, view, margin), `card ${JSON.stringify(box)} inside ${view.width}×${view.height} with ${margin} px margins`).toBe(true)
  expect(besides(expected, box, line), `card ${JSON.stringify(box)} sits ${expected} of the line ${JSON.stringify(line)}`).toBe(true)
}

// The page's paragraphs as the card counts them: each section's p and h1–h6 in
// document order, numbered from 1; the section each sits in, the comment ids
// painted in each, and the indices (among all .change-timeline-mark elements)
// of the marks it offers. A changed table or row is its first paragraph's mark.
const PARAGRAPH = "p, h1, h2, h3, h4, h5, h6"
async function paintedParagraphs(render: Locator): Promise<{ rows: Painted[]; marks: number }> {
  return render.evaluate((root, selector) => {
    const seen = new Set<Element>()
    const paragraphs = [...root.querySelectorAll("div[data-section-index]")]
      .flatMap(section => [...section.querySelectorAll(selector)])
      .filter(node => !seen.has(node) && Boolean(seen.add(node)))
    const marks = [...root.querySelectorAll(".change-timeline-mark")]
    return {
      marks: marks.length,
      rows: paragraphs.map((node, index) => ({
        ordinal: index + 1,
        section: node.closest("div[data-section-index]")?.getAttribute("data-section-index") ?? null,
        comments: [...new Set([...node.querySelectorAll("[data-comment-id], a.comment-marker")]
          .map(found => found.getAttribute("data-comment-id") ?? /^comment-ref-(\d+)$/.exec(found.id)?.[1])
          .filter((id): id is string => id != null).map(Number))],
        marks: marks.flatMap((mark, at) => ((mark.closest(selector) ?? mark.querySelector(selector)) === node ? [at] : [])),
        text: (node.textContent ?? "").slice(0, 60),
      })),
    }
  }, PARAGRAPH)
}
// Painted paragraph `ordinal` (from 1), counted as paintedParagraphs counts.
const paragraphAt = (render: Locator, ordinal: number) => render.locator(`div[data-section-index] :is(${PARAGRAPH})`).nth(ordinal - 1)
async function markFor(render: Locator, ordinal: number) {
  const row = (await paintedParagraphs(render)).rows.find(found => found.ordinal === ordinal)
  if (!row?.marks.length) throw new Error(`Painted paragraph ${ordinal} offers no card.`)
  return render.locator(".change-timeline-mark").nth(row.marks[0])
}

async function readMarkup(name: string) {
  const response = await api(`/matters/${file.matter}/documents/markup?name=${encodeURIComponent(name)}`)
  expect(response.status, `the markup route reads ${name}`).toBe(200)
  return await response.json() as DocumentMarkup
}
const authorName = (name: string) => name.trim() || "No author named"
const when = (value: string | null) => {
  if (!value) return null
  const time = Date.parse(value)
  return Number.isNaN(time) ? value : new Date(time).toISOString()
}
const entryKey = (entry: Entry) => `${entry.author} | ${when(entry.date) ?? "undated"} | ${entry.kind}`

// What the reader lists for one painted paragraph: body revisions with a
// location there (unless the document's revisions are withheld), and the whole
// thread of every listed body comment whose id is painted in it. A section's
// own change, which the reader places at the section's first paragraph, belongs
// on the card of the section's last.
function expectedEntries(markup: DocumentMarkup, row: Painted, revisionsShown: boolean, rows: Painted[]): Entry[] {
  const listed = new Map(markup.comments.filter(comment => !comment.scope).map(comment => [comment.id, comment]))
  const root = (comment: MarkupComment) => {
    let at = comment
    for (let hops = 0; at.parent !== null && listed.has(at.parent) && hops < listed.size; hops++) at = listed.get(at.parent)!
    return at.id
  }
  const shown = (place: MarkupLocation) => {
    if (place.container !== "section" || place.paragraph === null) return place.paragraph
    const section = rows[place.paragraph - 1]?.section
    return rows.findLast(found => found.section === section)?.ordinal ?? place.paragraph
  }
  const threads = new Set(row.comments.filter(id => listed.has(id)).map(id => root(listed.get(id)!)))
  return [
    ...(revisionsShown ? markup.revisions.filter(revision => revision.scope === "body" && revision.locations.some(place => shown(place) === row.ordinal))
      .map(revision => ({ author: authorName(revision.author), date: revision.date, kind: revision.kind })) : []),
    ...[...listed.values()].filter(comment => threads.has(root(comment))).map(comment => ({ author: authorName(comment.author), date: comment.date, kind: "comment" })),
  ]
}

type CardRead = { place: string | null; undatedLabel: string | null; nodes: { author: string; datetime: string | null; afterUndated: boolean; kinds: (string | null)[] }[] }
async function readCard(card: Locator): Promise<CardRead> {
  return card.evaluate((root) => {
    const list = root.querySelector(".change-timeline-body ol.change-timeline-list")
    const undated = list?.querySelector("li.change-timeline-undated") ?? null
    // A node's own author and time, not an action's (a reply names whom it answers).
    const own = (node: Element, selector: string) => [...node.querySelectorAll(selector)].find(found => !found.closest(".change-timeline-action"))
    return {
      place: root.querySelector("p.change-timeline-place")?.textContent?.trim() ?? null,
      undatedLabel: undated?.textContent?.trim() ?? null,
      nodes: [...(list?.querySelectorAll("li.change-timeline-node") ?? [])].map(node => ({
        author: own(node, ".change-timeline-author")?.textContent?.trim() ?? "",
        datetime: own(node, "time[datetime]")?.getAttribute("datetime") ?? null,
        afterUndated: undated ? Boolean(undated.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING) : false,
        kinds: [...node.querySelectorAll("li.change-timeline-action")].map(action => action.getAttribute("data-kind")),
      })),
    }
  })
}
function checkCard(read: CardRead, expected: Entry[], label: string) {
  const shown = read.nodes.flatMap(node => node.kinds.map(kind => ({ author: node.author, date: node.datetime, kind: kind ?? "(no data-kind)" })))
  expect(shown.map(entryKey).sort(), `${label}: the card's (author, date, kind) entries equal the reader's`).toEqual(expected.map(entryKey).sort())
  const dated = read.nodes.filter(node => node.datetime !== null)
  const times = dated.map(node => Date.parse(node.datetime!))
  expect(times.every(time => !Number.isNaN(time)), `${label}: every <time datetime> parses: ${dated.map(node => node.datetime).join(", ")}`).toBe(true)
  expect(times.every((time, index) => index === 0 || time >= times[index - 1]), `${label}: oldest first: ${dated.map(node => node.datetime).join(", ")}`).toBe(true)
  const stamps = dated.map(node => `${node.author} | ${when(node.datetime)}`)
  expect(stamps.length - new Set(stamps).size, `${label}: one node per author and time: ${stamps.join("; ")}`).toBe(0)
  const firstUndated = read.nodes.findIndex(node => node.datetime === null)
  if (firstUndated < 0) return
  expect(read.nodes.slice(firstUndated).every(node => node.datetime === null), `${label}: undated nodes come after every dated one`).toBe(true)
  expect(read.undatedLabel ?? "", `${label}: undated nodes sit under "Date not recorded"`).toContain("Date not recorded")
  expect(read.nodes.every(node => node.afterUndated === (node.datetime === null)), `${label}: only undated nodes follow the "Date not recorded" item`).toBe(true)
}

// Hovers the first mark of every paragraph the reader has entries for and
// compares the card with them. Every .change-timeline-mark must sit in such a
// paragraph, one per paragraph, as a Tab stop announcing its card. A changed
// table or row is hovered on its first paragraph. Each card is placed by
// paragraph number unless `places` names the paragraph's list number.
async function matchReader(page: Page, render: Locator, markup: DocumentMarkup, label: string, revisionsShown: boolean, places: Record<number, RegExp> = {}) {
  const { rows, marks } = await paintedParagraphs(render)
  const all = render.locator(".change-timeline-mark")
  expect(rows.reduce((sum, row) => sum + row.marks.length, 0), `${label}: every .change-timeline-mark sits in a painted paragraph`).toBe(marks)
  const attributes = await all.evaluateAll(nodes => nodes.map(node => [node.getAttribute("tabindex"), node.getAttribute("aria-haspopup"), node.getAttribute("aria-expanded")].join(" ")))
  expect(attributes.filter(found => !/^0 dialog (true|false)$/.test(found)), `${label}: marks are Tab stops with aria-haspopup=dialog and aria-expanded`).toEqual([])
  const expected = rows.map(row => ({ ...row, expected: expectedEntries(markup, row, revisionsShown, rows) }))
  expect(expected.filter(row => row.marks.length).map(row => row.ordinal), `${label}: the paragraphs offering a card are those the reader has entries for`)
    .toEqual(expected.filter(row => row.expected.length).map(row => row.ordinal))
  expect(expected.filter(row => row.marks.length > 1).map(row => row.ordinal), `${label}: one Tab stop per paragraph`).toEqual([])
  const checked: { paragraph: number; entries: number; nodes: number; mark: string }[] = []
  for (const row of expected.filter(found => found.expected.length && found.marks.length)) {
    await leave(page)
    const mark = all.nth(row.marks[0])
    const tag = await mark.evaluate((node, selector) => `${node.tagName.toLowerCase()}${node.closest(selector) ? "" : " (container)"}`, PARAGRAPH)
    await hoverOpen(page, tag.endsWith("(container)") ? paragraphAt(render, row.ordinal) : mark)
    await expect(cardOf(page).locator("p.change-timeline-place")).toHaveText(places[row.ordinal] ?? new RegExp(`^Paragraph ${row.ordinal}(?!\\d)`))
    const read = await readCard(cardOf(page))
    checkCard(read, row.expected, `${label} paragraph ${row.ordinal}`)
    checked.push({ paragraph: row.ordinal, entries: row.expected.length, nodes: read.nodes.length, mark: tag })
    await cardShot(page, `reader-match-${label}-paragraph-${String(row.ordinal).padStart(2, "0")}`)
  }
  await leave(page)
  return checked
}

// ---- Run -------------------------------------------------------------------------
const file = { dir: "", matter: "" }
let browser: Browser | undefined
const origin = new URL(baseURL).origin
async function openPage(options: BrowserContextOptions, collapsedRail = false) {
  const context = await browser!.newContext({ reducedMotion: "reduce", ...options })
  await context.addInitScript((collapsed) => {
    sessionStorage.setItem("dh.launch-notice.v1", "accepted")
    // The rail starts collapsed at phone width, as on a first visit from a phone.
    if (collapsed) localStorage.setItem("dh.sidebar", "1")
  }, collapsedRail)
  // Engine traffic is answered here; ingest traffic continues through Vite's
  // /ingest proxy to the isolated service.
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    if (url.origin !== origin) { unexpected.push(`${request.method()} ${request.url()}`); return route.abort() }
    const reply = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    // The release update check runs git in the checkout; it is not part of this flow.
    if (request.method() === "GET" && url.pathname === "/ingest/updates") {
      stubbed.push("GET /ingest/updates")
      return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    }
    if (url.pathname.startsWith("/ingest/")) {
      ingestRequests.push({ method: request.method(), path: url.pathname.replace(/\/matters\/[^/]+/, "/matters/:id") })
      return route.continue()
    }
    if (!url.pathname.startsWith("/engine/")) return route.continue()
    const pathname = url.pathname.slice("/engine".length)
    engineRequests.push({ method: request.method(), path: pathname })
    if (request.method() === "GET" && pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (request.method() === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (request.method() === "GET" && pathname === "/api/session") return reply({ data: [], cursor: { next: null, previous: null } })
    if (request.method() === "GET" && pathname === "/api/agent") return reply({ location: { directory: file.dir }, data: [] })
    // An idle event stream: left open, as the engine would hold it between events.
    if (request.method() === "GET" && pathname === "/api/event") return
    unexpected.push(`engine ${request.method()} ${pathname}${url.search}`)
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: "not stubbed" }) })
  })
  const page = await context.newPage()
  page.setDefaultTimeout(20_000)
  page.on("pageerror", (error: Error) => pageErrors.push(error.message))
  page.on("console", (message: { type: () => string; text: () => string }) => message.type() === "error" && consoleErrors.push(message.text()))
  page.on("dialog", dialog => { dialogs.push(`${dialog.type()}: ${dialog.message()}`); void dialog.dismiss() })
  return page
}

// Opens a document from the matter's Documents list on a fresh load, and waits
// for the page to be painted and the markup reading to offer its marks.
async function openDocument(page: Page, name: string) {
  active = page
  await page.goto(`${baseURL!.replace(/\/$/, "")}/matter/${file.matter}`)
  const documents = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
  await expect(documents).toBeVisible()
  if (await documents.getAttribute("aria-expanded") !== "true") await documents.click()
  await page.getByRole("button", { name, exact: true }).click()
  const viewer = page.getByRole("dialog", { name, exact: true })
  await expect(viewer).toBeVisible()
  // The first open loads the Word runtime; later renders are quick.
  const render = viewer.locator(".docx-render:not(.docx-skeleton)")
  await expect(render).toBeVisible({ timeout: 30_000 })
  await expect(render.locator(".change-timeline-mark").first()).toBeAttached({ timeout: 30_000 })
  await page.mouse.move(2, 2)
  return { viewer, render, doc: viewer.locator(".viewer-doc") }
}
const fortyFive = (render: Locator) => render.locator("div[data-section-index] ins").filter({ hasText: /^forty-five$/ })

// Real Tab presses from the viewer's top, as a keyboard user moves through the
// page (a scripted focus() is not :focus-visible); returns the presses taken.
async function tabTo(page: Page, viewer: Locator, target: Locator, limit = 60) {
  await viewer.focus()
  for (let presses = 1; presses <= limit; presses++) {
    await page.keyboard.press("Tab")
    if (await target.evaluate(node => node === document.activeElement)) return presses
  }
  throw new Error(`Tab did not reach the mark within ${limit} presses.`)
}
// The accessible description Chromium itself computes for the index-th mark,
// as assistive technology receives it.
async function chromiumDescription(page: Page, index: number) {
  const cdp = await page.context().newCDPSession(page)
  try {
    const { root } = await cdp.send("DOM.getDocument", { depth: 0 })
    const { nodeIds } = await cdp.send("DOM.querySelectorAll", { nodeId: root.nodeId, selector: ".change-timeline-mark" })
    const { nodes } = await cdp.send("Accessibility.getPartialAXTree", { nodeId: nodeIds[index], fetchRelatives: false })
    return String(nodes[0]?.description?.value ?? "").replace(/\s+/g, " ").trim()
  } finally {
    await cdp.detach()
  }
}
// Whether the open card lists `revision`: a node by its author at its time with an action of its kind.
async function lists(page: Page, revision: MarkupRevision) {
  const read = await readCard(cardOf(page))
  return read.nodes.some(node => node.author === authorName(revision.author) && when(node.datetime) === when(revision.date) && node.kinds.includes(revision.kind))
}

try {
  await startIngest()
  browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
  const page = await openPage({ viewport: DESKTOP })

  await step("1 upload the supply agreement, the busy clause, the text-box file and the structures file as sources", async () => {
    const matter = await (await api("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic Change Timeline" }) })).json() as { id: string; dir: string }
    expect(matter.dir.startsWith(dataRoot)).toBe(true)
    Object.assign(file, { matter: matter.id, dir: matter.dir })
    for (const [name, bytes] of [[SUPPLY, counterpartyMarkupDocx()], [BUSY, busyDocx()], [TEXT_BOX, textBoxDocx()], [STRUCTURES, structuresDocx()]] as const) {
      const form = new FormData(); form.append("file", new File([bytes], name))
      expect((await api(`/matters/${matter.id}/documents`, { method: "POST", body: form })).status, `upload ${name}`).toBe(200)
    }
    const busy = await readMarkup(BUSY)
    measured.busyReader = { revisions: busy.revisions.length, comments: busy.comments.length }
    expect(busy.revisions.filter(revision => !revision.date).length, "precondition: the busy clause has one undated change").toBe(1)
  }, true)

  await step("2 hover forty-five: after a rest the card opens below, fully on screen", async () => {
    await page.setViewportSize(DESKTOP)
    const { viewer, render } = await openDocument(page, SUPPLY)
    const target = fortyFive(render)
    const [line] = await lineBoxes(target)
    // A pass across the mark without resting opens nothing.
    await page.mouse.move((line.left + line.right) / 2, (line.top + line.bottom) / 2, { steps: 3 })
    await page.mouse.move(2, 2, { steps: 3 })
    await page.waitForTimeout(REST * 2)
    await expect(cardOf(page)).toBeHidden()
    await hoverOpen(page, target)
    const card = cardOf(page)
    await expect(card).toHaveAccessibleName("Change history")
    await expect(card.locator("h2.change-timeline-title")).toHaveText("Change history")
    expect(await card.getAttribute("aria-modal"), "the card is not modal").not.toBe("true")
    // Inside the viewer's modal dialog, so assistive technology can reach it.
    await expect(viewer.locator("div.change-timeline")).toHaveCount(1)
    await expect(card).toHaveAttribute("data-placement", "below")
    await expect(card.locator("p.change-timeline-place")).toHaveText(/^Paragraph 1(?!\d)/)
    await expect(card.locator("button.change-timeline-close")).toBeHidden()
    const [hovered] = await lineBoxes(target), box = await cardBox(page)
    measured.hoverBelow = { line: hovered, card: box }
    expect(inside(box, DESKTOP), `card ${JSON.stringify(box)} inside the window`).toBe(true)
    expect(besides("below", box, hovered), `card ${JSON.stringify(box)} below ${JSON.stringify(hovered)}`).toBe(true)
    await cardShot(page, "card-below")
  })

  await step("3a every card's authors, dates and kinds match the reader (supply agreement)", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, SUPPLY)
    measured.readerSupply = await matchReader(page, render, await readMarkup(SUPPLY), "supply-agreement", true)
  })

  await step("3b every card's authors, dates and kinds match the reader (busy clause)", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, BUSY)
    measured.readerBusy = await matchReader(page, render, await readMarkup(BUSY), "busy", true)
  })

  await step("4 a click does not pin it; the pointer crosses into the card and it stays; leaving both closes it", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, SUPPLY)
    const target = fortyFive(render)
    await hoverOpen(page, target)
    await target.click()
    await page.waitForTimeout(REST * 2)
    await expect(cardOf(page)).toBeVisible()
    const [line] = await lineBoxes(target), box = await cardBox(page)
    await page.mouse.move((line.left + line.right) / 2, (line.top + line.bottom) / 2)
    await page.mouse.move(box.x + box.width / 2, box.y + Math.min(box.height / 2, 80), { steps: 10 })
    await page.waitForTimeout(800)
    await expect(cardOf(page)).toBeVisible()
    await cardShot(page, "pointer-in-card-stays-open")
    // Closing once the pointer leaves also shows the click did not pin it.
    await page.mouse.move(2, 2, { steps: 4 })
    await expect(cardOf(page)).toBeHidden({ timeout: 3_000 })
  })

  await step("5 Escape closes the card and leaves the viewer open; a second Escape closes the viewer", async () => {
    await page.setViewportSize(DESKTOP)
    const { viewer, render } = await openDocument(page, SUPPLY)
    // With focus inside the viewer, as after opening it, Escape reaches the viewer's own handler too.
    await viewer.focus()
    await hoverOpen(page, fortyFive(render))
    await page.keyboard.press("Escape")
    await expect(cardOf(page)).toBeHidden()
    await expect(viewer).toBeVisible()
    await shot(page, "escape-closes-card-viewer-stays")
    await page.mouse.move(2, 2)
    await page.waitForTimeout(REST * 2)
    await page.keyboard.press("Escape")
    await expect(viewer).toBeHidden()
  })

  await step("6a near the window's bottom the card opens above", async () => {
    await page.setViewportSize(DESKTOP)
    const { doc, render } = await openDocument(page, BUSY)
    const target = render.locator("ins").filter({ hasText: /^Service levels\.$/ })
    const size = await naturalSize(page, doc, target, "first")
    const docBox = (await doc.boundingBox())!
    const line = await centreLine(page, doc, target, "first", docBox.y + docBox.height - 24)
    const room = spare(line, size, DESKTOP, await paneOf(page))
    if (clearSide(room) !== "above") throw new Error(`Setup: no clear above case: ${JSON.stringify({ line, size, room })}`)
    await expectPlacement(page, "placementAbove", target, line, "above")
    await cardShot(page, "card-above")
  })

  await step("6b in a short window, at a line's start, the card opens right", async () => {
    await page.setViewportSize(DESKTOP)
    const { doc, render } = await openDocument(page, BUSY)
    const target = render.locator("ins").filter({ hasText: /^Service levels\.$/ })
    const size = await naturalSize(page, doc, target, "first")
    const view = { width: DESKTOP.width, height: await shortHeight(page, size) }
    await page.setViewportSize(view)
    await settle(page)
    const pane = await paneOf(page)
    const line = await centreLine(page, doc, target, "first", (pane.top + pane.bottom) / 2)
    const room = spare(line, size, view, pane)
    if (clearSide(room) !== "right") throw new Error(`Setup: no clear right case: ${JSON.stringify({ view, pane, line, size, room })}`)
    await expectPlacement(page, "placementRight", target, line, "right")
    await cardShot(page, "card-right")
  })

  await step("6c in a narrower short window, at a line's end, the card opens left", async () => {
    await page.setViewportSize(DESKTOP)
    const { viewer, doc, render } = await openDocument(page, BUSY)
    const target = render.locator("ins").filter({ hasText: /^signature\.$/ })
    const size = await naturalSize(page, doc, target, "last")
    // Their changes take the window's right-hand side, so no line ends near its
    // edge while they are open.
    const toggle = viewer.locator(".viewer-markup-toggle")
    if (await toggle.getAttribute("aria-expanded") === "true") await toggle.click()
    await expect(viewer.locator("#document-markup")).toHaveCount(0)
    const height = await shortHeight(page, size)
    const tried: unknown[] = []
    let chosen: { view: View; line: Rect } | undefined
    // Widths stay above the 720 px narrow layout.
    for (const width of [1280, 1100, 1024, 960, 900, 860, 820, 780, 740]) {
      const view = { width, height }
      await page.setViewportSize(view)
      await settle(page)
      const pane = await paneOf(page)
      const line = await centreLine(page, doc, target, "last", (pane.top + pane.bottom) / 2)
      const room = spare(line, size, view, pane)
      tried.push({ width, line, room })
      if (clearSide(room) === "left") { chosen = { view, line }; break }
    }
    if (!chosen) throw new Error(`Setup: no width gives a clear left case: ${JSON.stringify({ size, tried })}`)
    await expectPlacement(page, "placementLeft", target, chosen.line, "left")
    await cardShot(page, "card-left")
  })

  await step("7 a long timeline scrolls inside the card by wheel and the card stays open", async () => {
    await page.setViewportSize(LAPTOP)
    const { doc, render } = await openDocument(page, BUSY)
    await hoverOpen(page, await markFor(render, 1))
    const card = cardOf(page), body = card.locator(".change-timeline-body")
    const metrics = await body.evaluate(node => ({ scrollHeight: node.scrollHeight, clientHeight: node.clientHeight }))
    measured.longTimeline = { ...metrics, card: await cardBox(page), placement: await card.getAttribute("data-placement") }
    expect(metrics.scrollHeight, "the timeline is taller than the card").toBeGreaterThan(metrics.clientHeight)
    expect(inside(await cardBox(page), LAPTOP), "the card stays on screen").toBe(true)
    await expect(card).toContainText("Date not recorded")
    const area = (await body.boundingBox())!
    await page.mouse.move(area.x + area.width / 2, area.y + area.height / 2, { steps: 10 })
    const pageTop = await doc.evaluate(node => node.scrollTop)
    await page.mouse.wheel(0, 240)
    await expect.poll(() => body.evaluate(node => node.scrollTop)).toBeGreaterThan(0)
    await page.waitForTimeout(REST * 2)
    await expect(card).toBeVisible()
    expect(await doc.evaluate(node => node.scrollTop), "the page under the card did not scroll").toBe(pageTop)
    await cardShot(page, "long-timeline-scrolled")
  })

  await step("8 focus opens the card, Enter moves into it, arrows scroll it, Escape closes it and returns focus", async () => {
    await page.setViewportSize(LAPTOP)
    const { viewer, render } = await openDocument(page, BUSY)
    const mark = await markFor(render, 1)
    measured.tabPressesToMark = await tabTo(page, viewer, mark)
    const card = cardOf(page), body = card.locator(".change-timeline-body")
    await expect(card).toBeVisible()
    await expect(mark).toHaveAttribute("aria-expanded", "true")
    await cardShot(page, "keyboard-focus-open")
    await page.keyboard.press("Enter")
    await expect(body).toBeFocused()
    await expect(card).toBeVisible()
    const top = () => body.evaluate(node => node.scrollTop)
    const start = await top()
    await page.keyboard.press("ArrowDown")
    await expect.poll(top).toBeGreaterThan(start)
    const afterArrow = await top()
    await page.keyboard.press("PageDown")
    await expect.poll(top).toBeGreaterThan(afterArrow)
    await cardShot(page, "keyboard-in-card-scrolled")
    await page.keyboard.press("Escape")
    await expect(card).toBeHidden()
    await expect(mark).toBeFocused()
    await expect(mark).toHaveAttribute("aria-expanded", "false")
    await expect(viewer).toBeVisible()
    // Focus coming back to the mark does not open the card again.
    await page.waitForTimeout(REST * 2)
    await expect(card).toBeHidden()
  })

  await step("9a a comment written as HTML shows as text and runs nothing", async () => {
    await page.setViewportSize(DESKTOP)
    const markup = await readMarkup(BUSY)
    const comment = markup.comments.find(found => found.text.includes("__timelineXss"))
    expect(comment?.text ?? "", "precondition: the reader lists the comment with its markup intact").toContain("<img")
    const { render } = await openDocument(page, BUSY)
    const flagged = () => page.evaluate(() => (window as unknown as { __timelineXss?: unknown }).__timelineXss)
    const beforeCard = await flagged()
    await hoverOpen(page, await markFor(render, 2))
    const card = cardOf(page)
    await expect(card.locator('li.change-timeline-action[data-kind="comment"]')).toContainText(comment!.text)
    expect(await card.locator("img, b, script").count(), "no element made from the comment's text").toBe(0)
    const afterCard = await flagged()
    measured.untrusted = { beforeCard: beforeCard ?? null, afterCard: afterCard ?? null }
    expect(beforeCard, "window.__timelineXss was set before any card opened (by the page render)").toBeUndefined()
    expect(afterCard, "window.__timelineXss was set once the card opened").toBeUndefined()
    await cardShot(page, "untrusted-text-as-text")
  })

  await step("9b the planted deletion and comment carry the instruction flag in paragraph 6's card", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, SUPPLY)
    await hoverOpen(page, await markFor(render, 6))
    const card = cardOf(page)
    await expect(card.locator("p.change-timeline-place")).toHaveText(/^Paragraph 6(?!\d)/)
    await expect(card.locator('li.change-timeline-action[data-kind="delete"]').filter({ hasText: COUNTERPARTY_MARKUP.plantedDeletion.trim() })).toContainText(FLAG)
    await expect(card.locator('li.change-timeline-action[data-kind="comment"]').filter({ hasText: COUNTERPARTY_MARKUP.plantedComment })).toContainText(FLAG)
    await cardShot(page, "planted-instruction-flag")
  })

  await step("10 at 400 px a tap opens the card on screen with Close; an outside tap and Close each close it", async () => {
    const phone = await openPage({ viewport: PHONE, hasTouch: true, isMobile: true }, true)
    const { render } = await openDocument(phone, SUPPLY)
    const target = fortyFive(render), card = cardOf(phone), close = card.locator("button.change-timeline-close")
    const overflow = () => phone.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)
    await target.tap()
    await expect(card).toBeVisible()
    expect(["below", "above"], "anchored above or below the tapped text").toContain(await card.getAttribute("data-placement"))
    const box = await cardBox(phone)
    measured.phone = { card: box, placement: await card.getAttribute("data-placement") }
    expect(inside(box, PHONE, MARGIN), `card ${JSON.stringify(box)} inside 400×860 with 8 px margins`).toBe(true)
    await expect(close).toBeVisible()
    await expect(close).toHaveAccessibleName("Close")
    await expect(close).toBeInViewport()
    expect(await overflow(), "no horizontal page overflow with the card open").toBeLessThanOrEqual(1)
    await cardShot(phone, "phone-400px-touch")
    // An outside tap on plain, inert page: away from the card, every mark and every control.
    const spot = await phone.evaluate(([avoid]) => {
      const busy = "div.change-timeline, .change-timeline-mark, ins, del, .rev-format-change, .rev-para-mark-del, .comment-highlight, a, button, input, select, textarea, label, summary, details, [role=button], [role=radio], [role=checkbox], [role=tab], [contenteditable], [tabindex]:not([tabindex='-1'])"
      for (let y = 16; y < innerHeight - 16; y += 12) for (const x of [200, 40, 360]) {
        if (y >= avoid.y - 12 && y <= avoid.y + avoid.height + 12) continue
        const found = document.elementFromPoint(x, y)
        if (found && !found.closest(busy)) return { x, y, element: `${found.tagName.toLowerCase()}.${found.className}` }
      }
      return null
    }, [box] as const)
    if (!spot) throw new Error("Setup: no inert spot outside the card to tap.")
    measured.phoneOutsideTap = spot
    await phone.touchscreen.tap(spot.x, spot.y)
    await expect(card).toBeHidden()
    await target.tap()
    await expect(card).toBeVisible()
    await close.tap()
    await expect(card).toBeHidden()
    expect(await overflow(), "no horizontal page overflow").toBeLessThanOrEqual(1)
    await shot(phone, "phone-400px-closed")
    await phone.context().close()
  })

  await step("11 in a dark colour scheme the card stays readable", async () => {
    const dark = await openPage({ viewport: DESKTOP, colorScheme: "dark" })
    const { render } = await openDocument(dark, SUPPLY)
    await hoverOpen(dark, fortyFive(render))
    // Computed colours as sRGB bytes, whatever notation the style uses.
    const colours = await cardOf(dark).evaluate((card) => {
      const canvas = document.createElement("canvas"); canvas.width = canvas.height = 1
      const context = canvas.getContext("2d", { willReadFrequently: true })!
      const srgb = (colour: string) => { context.clearRect(0, 0, 1, 1); context.fillStyle = colour; context.fillRect(0, 0, 1, 1); return [...context.getImageData(0, 0, 1, 1).data] }
      const author = card.querySelector(".change-timeline-author"), style = getComputedStyle(card)
      return { css: { background: style.backgroundColor, text: style.color, author: author && getComputedStyle(author).color },
        background: srgb(style.backgroundColor), text: srgb(style.color), author: author ? srgb(getComputedStyle(author).color) : null }
    })
    const luminance = ([r, g, b]: number[]) => {
      const [x, y, z] = [r, g, b].map(value => { const s = value / 255; return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4 })
      return 0.2126 * x + 0.7152 * y + 0.0722 * z
    }
    const contrast = (one: number[], two: number[]) => { const [high, low] = [luminance(one), luminance(two)].sort((a, b) => b - a); return (high + 0.05) / (low + 0.05) }
    const opaque = colours.background[3] === 255
    const ratios = opaque ? { text: contrast(colours.text, colours.background), author: colours.author && contrast(colours.author, colours.background) } : null
    measured.dark = { ...colours, ratios, note: opaque ? undefined : "the card's own background is translucent; contrast not measured" }
    if (ratios) {
      expect(ratios.text, "card text contrast").toBeGreaterThanOrEqual(4.5)
      if (ratios.author !== null) expect(ratios.author, "author contrast").toBeGreaterThanOrEqual(4.5)
    }
    expect(inside(await cardBox(dark), DESKTOP)).toBe(true)
    await cardShot(dark, "dark-scheme-stays-light")
    await dark.context().close()
  })

  await step("12 where the page's paragraphs cannot match the reader's, only comment threads show", async () => {
    await page.setViewportSize(DESKTOP)
    const markup = await readMarkup(TEXT_BOX)
    const tail = markup.revisions.find(revision => revision.scope === "body" && revision.text.includes("Tail added"))
    expect(tail, "precondition: the reader lists the change after the text box").toBeTruthy()
    const { render } = await openDocument(page, TEXT_BOX)
    const after = (await paintedParagraphs(render)).rows.find(row => row.text.includes("3. After the box."))
    expect(after, "the paragraph after the box is painted").toBeTruthy()
    expect(tail!.locations.map(place => place.paragraph), "precondition: the text box shifts the reader's paragraph numbers").not.toContain(after!.ordinal)
    measured.commentsOnly = await matchReader(page, render, markup, "text-box", false)
    // The change before the box, which would line up, is withheld too.
    await render.locator("ins").filter({ hasText: /Agreed\./ }).hover()
    await page.waitForTimeout(REST * 3)
    await expect(cardOf(page)).toBeHidden()
    await hoverOpen(page, await markFor(render, after!.ordinal))
    await expect(cardOf(page).locator("li.change-timeline-action:not([data-kind=comment])")).toHaveCount(0)
    await cardShot(page, "comments-only-gate")
  })

  await step("13 a drag across a mark with the button held opens no card and keeps the selection", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, SUPPLY)
    const target = fortyFive(render)
    // The word "Supplier" in paragraph 1: plain text, outside every mark.
    const start = await paragraphAt(render, 1).evaluate((paragraph) => {
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const at = node.textContent?.indexOf("Supplier") ?? -1
        if (at < 0) continue
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + "Supplier".length)
        // Just inside the word's first letter, so the selection starts before it.
        const rect = range.getBoundingClientRect()
        const point = { x: rect.left + 1, y: rect.top + rect.height / 2 }
        const found = document.elementFromPoint(point.x, point.y)
        return { ...point, plain: Boolean(found && paragraph.contains(found) && !found.closest(".change-timeline-mark, ins, del, .rev-format-change, .comment-highlight, a.comment-marker")) }
      }
      return null
    })
    if (!start?.plain) throw new Error(`Setup: no plain text "Supplier" in paragraph 1: ${JSON.stringify(start)}`)
    const [line] = await lineBoxes(target)
    const over = { x: (line.left + line.right) / 2, y: (line.top + line.bottom) / 2 }
    await page.mouse.move(start.x, start.y)
    await page.mouse.down()
    await page.mouse.move(over.x, over.y, { steps: 12 })
    // A rest on the mark well past the hover delay, with the button held.
    await page.waitForTimeout(REST * 3)
    await expect(cardOf(page)).toBeHidden()
    await page.mouse.up()
    await page.waitForTimeout(REST * 2)
    await expect(cardOf(page)).toBeHidden()
    const selection = await page.evaluate(() => ({ text: document.getSelection()?.toString() ?? "", collapsed: document.getSelection()?.isCollapsed ?? true }))
    measured.dragSelect = { start, over, selection }
    expect(selection.collapsed, "the drag left a selection").toBe(false)
    expect(selection.text.replace(/\s+/g, " "), "the selection runs from Supplier into the mark").toMatch(/Supplier shall deliver the Goods within ?thirty ?fo/)
    await shot(page, "drag-select-no-card")
    // Control: the same rest without a held button opens the card.
    await page.evaluate(() => document.getSelection()?.removeAllRanges())
    await leave(page)
    await page.mouse.move(over.x, over.y, { steps: 2 })
    await expect(cardOf(page)).toBeVisible()
  })

  await step("14 a click on a clause's mark does not pin its card: hovering another clause switches the card", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, SUPPLY)
    const first = await markFor(render, 1), second = await markFor(render, 2)
    const place = cardOf(page).locator("p.change-timeline-place")
    await hoverOpen(page, first)
    await expect(place).toHaveText(/^Paragraph 1(?!\d)/)
    await first.click()
    await expect(first).toBeFocused()
    await page.waitForTimeout(REST * 2)
    await expect(place).toHaveText(/^Paragraph 1(?!\d)/)
    await second.hover()
    await expect(place).toHaveText(/^Paragraph 2(?!\d)/)
    await expect(second).toHaveAttribute("aria-expanded", "true")
    await expect(first).toHaveAttribute("aria-expanded", "false")
    // Focus stayed on the clicked mark; it no longer holds that clause's card.
    await expect(first).toBeFocused()
    await cardShot(page, "click-then-hover-switches")
  })

  await step("15 Shift+Tab from the card returns to its mark with the card open; Tab from the card moves to the next clause's mark and opens its card", async () => {
    await page.setViewportSize(LAPTOP)
    const { viewer, render } = await openDocument(page, BUSY)
    const offering = (await paintedParagraphs(render)).rows.filter(row => row.marks.length)
    const at = offering.findIndex(row => row.ordinal === 43)
    const here = offering[at], next = offering[at + 1]
    if (!here || !next) throw new Error("Setup: paragraph 43 and a later paragraph must offer cards.")
    const all = render.locator(".change-timeline-mark")
    const mark = all.nth(here.marks[0]), following = all.nth(next.marks[0])
    const card = cardOf(page), body = card.locator(".change-timeline-body"), place = card.locator("p.change-timeline-place")
    const presses = await tabTo(page, viewer, mark)
    await expect(place).toHaveText(/^Paragraph 43(?!\d)/)
    await page.keyboard.press("Enter")
    await expect(body).toBeFocused()
    await page.keyboard.press("Shift+Tab")
    await expect(mark).toBeFocused()
    await page.waitForTimeout(REST * 2)
    await expect(card).toBeVisible()
    await expect(place).toHaveText(/^Paragraph 43(?!\d)/)
    await expect(mark).toHaveAttribute("aria-expanded", "true")
    await cardShot(page, "shift-tab-back-to-mark")
    await page.keyboard.press("Enter")
    await expect(body).toBeFocused()
    await page.keyboard.press("Tab")
    await expect(following).toBeFocused()
    await expect(place).toHaveText(new RegExp(`^Paragraph ${next.ordinal}(?!\\d)`))
    await expect(following).toHaveAttribute("aria-expanded", "true")
    await expect(mark).toHaveAttribute("aria-expanded", "false")
    measured.tabOut = { presses, from: here.ordinal, to: next.ordinal, focused: await page.evaluate(() => `${document.activeElement?.tagName.toLowerCase()} "${document.activeElement?.textContent?.slice(0, 40)}"`) }
    await cardShot(page, "tab-out-next-clause")
  })

  await step("16 a card opened by keyboard closes once its line scrolls out of view and is never drawn over the viewer's toolbar", async () => {
    await page.setViewportSize(LAPTOP)
    const { viewer, doc, render } = await openDocument(page, BUSY)
    const mark = await markFor(render, 1), card = cardOf(page)
    await tabTo(page, viewer, mark)
    await expect(card).toBeVisible()
    const pane = (await doc.boundingBox())!, bar = (await viewer.locator(".viewer-bar").boundingBox())!
    const samples: { scrollTop: number; line: Rect; card: Box | null }[] = []
    let lastOpen: Buffer | undefined
    for (let index = 0; index < 80; index++) {
      const [line] = await lineBoxes(mark)
      const box = await card.isVisible() ? await card.boundingBox() : null
      samples.push({ scrollTop: await doc.evaluate(node => node.scrollTop), line, card: box })
      if (box) {
        expect(box.y, `sample ${index}: the card's top ${box.y} is below the toolbar's bottom ${bar.y + bar.height}`).toBeGreaterThanOrEqual(bar.y + bar.height - 0.5)
        await expectInPane(page, `scroll-away sample ${index}`)
        lastOpen = await page.screenshot()
      }
      if (line.bottom <= pane.y) break
      await doc.evaluate(node => { node.scrollTop += 16 })
      await settle(page)
    }
    const [gone] = await lineBoxes(mark)
    expect(gone.bottom, "the mark's line has left the document pane").toBeLessThanOrEqual(pane.y)
    await expect(card).toBeHidden({ timeout: 2_000 })
    const open = samples.filter(sample => sample.card)
    measured.scrollAway = { pane: { top: pane.y }, toolbarBottom: bar.y + bar.height, samples: samples.length, openSamples: open.length, closestOpenLineTop: open.at(-1)?.line.top, closestOpenCardTop: open.at(-1)?.card?.y }
    expect(open.length, "the card followed its line for several scroll steps before closing").toBeGreaterThanOrEqual(2)
    if (lastOpen) await writeFile(path.join(shots, "scroll-away-last-open.png"), lastOpen)
    await shot(page, "scroll-away-closed")
  })

  await step("17 a mark's accessible description is the hint, and the place line too while its card is open", async () => {
    await page.setViewportSize(DESKTOP)
    const { render } = await openDocument(page, SUPPLY)
    const row = (await paintedParagraphs(render)).rows.find(found => found.ordinal === 1)
    if (!row?.marks.length) throw new Error("Setup: paragraph 1 offers no card.")
    const index = row.marks[0], mark = render.locator(".change-timeline-mark").nth(index)
    const closed = await chromiumDescription(page, index)
    expect(closed, "Chromium's description while the card is closed").toBe(HINT)
    await expect(mark).toHaveAccessibleDescription(HINT)
    await hoverOpen(page, mark)
    const placeLine = (await cardOf(page).locator("p.change-timeline-place").textContent() ?? "").replace(/\s+/g, " ").trim()
    expect(placeLine).toMatch(/^Paragraph 1 · /)
    const open = await chromiumDescription(page, index)
    expect(open, "Chromium's description while the card is open").toBe(`${HINT} ${placeLine}`)
    await expect(mark).toHaveAccessibleDescription(`${HINT} ${placeLine}`)
    await cardShot(page, "described-by-open")
    await leave(page)
    const after = await chromiumDescription(page, index)
    expect(after, "the place line leaves the description once the card closes").toBe(HINT)
    measured.describedBy = { mark: await mark.evaluate(node => node.tagName.toLowerCase()), closed, open, after }
  })

  await step("18 a changed table and row open their card, section changes show on the section's last paragraph, and every card matches the reader", async () => {
    await page.setViewportSize(DESKTOP)
    const markup = await readMarkup(STRUCTURES)
    const onlyBy = (author: string) => {
      const found = markup.revisions.filter(revision => revision.author === author)
      if (found.length !== 1) throw new Error(`Setup: expected one change by ${author}, found ${found.length}.`)
      return found[0]!
    }
    const [table, row, sectionBreak, finalSection] = [onlyBy(MATTEO), onlyBy(INES), onlyBy(ODILE), onlyBy(BASTIEN)]
    const placeOf = (revision: MarkupRevision) => revision.locations[0]!
    expect([placeOf(table).container, placeOf(row).container, placeOf(finalSection).container], "precondition: the reader places the table, row and final section changes by container").toEqual(["table", "row", "section"])
    const { render } = await openDocument(page, STRUCTURES)
    const { rows } = await paintedParagraphs(render)
    const card = cardOf(page), place = card.locator("p.change-timeline-place")
    const tableAt = placeOf(table).paragraph!, rowAt = placeOf(row).paragraph!
    // The table's second cell in its unchanged first row has no mark of its own;
    // hovering it opens the table's card, placed at the table's first paragraph.
    expect(rows[tableAt]!.marks, "precondition: the table's second cell has no mark").toEqual([])
    await hoverOpen(page, paragraphAt(render, tableAt + 1))
    await expect(place).toHaveText(new RegExp(`^Paragraph ${tableAt}(?!\\d)`))
    expect(await lists(page, table), "the table's card lists the table change").toBe(true)
    await cardShot(page, "table-change-card")
    await leave(page)
    // The inserted row's second cell opens the row's card.
    await hoverOpen(page, paragraphAt(render, rowAt + 1))
    await expect(place).toHaveText(new RegExp(`^Paragraph ${rowAt}(?!\\d)`))
    expect(await lists(page, row), "the row's card lists the row insertion").toBe(true)
    await cardShot(page, "row-insert-card")
    await leave(page)
    // The section break's change is on the paragraph that ends its section.
    const breakAt = placeOf(sectionBreak).paragraph!
    expect(rows.findLast(found => found.section === rows[breakAt - 1]!.section)?.ordinal, "precondition: the break's paragraph ends its section").toBe(breakAt)
    await hoverOpen(page, paragraphAt(render, breakAt))
    await expect(place).toHaveText(new RegExp(`^Paragraph ${breakAt}(?!\\d)`))
    expect(await lists(page, sectionBreak), "the section break's card lists its page setup change").toBe(true)
    await cardShot(page, "section-break-change-card")
    await leave(page)
    // The final section's change: listed at its first paragraph, shown on its last.
    const first = placeOf(finalSection).paragraph!
    const last = rows.findLast(found => found.section === rows[first - 1]!.section)!.ordinal
    expect(last, "precondition: the final section has more than one paragraph").toBeGreaterThan(first)
    await paragraphAt(render, first).hover()
    await page.waitForTimeout(REST * 3)
    await expect(card, "the section's first paragraph offers no card").toBeHidden()
    await hoverOpen(page, paragraphAt(render, last))
    await expect(place).toHaveText(new RegExp(`^Paragraph ${last}(?!\\d)`))
    expect(await lists(page, finalSection), "the section's last paragraph lists the section change").toBe(true)
    await cardShot(page, "section-change-card")
    measured.containers = { table: tableAt, row: rowAt, sectionBreak: breakAt, finalSection: { listedAt: first, shownOn: last } }
    measured.readerStructures = await matchReader(page, render, markup, "structures", true, { 1: /^Clause 1(?!\d)/ })
  })

  await step("19 a list number over 16 characters is shown as \"Paragraph N\"; a short one names the clause", async () => {
    await page.setViewportSize(DESKTOP)
    const markup = await readMarkup(STRUCTURES)
    const label = (ordinal: number) => markup.revisions.flatMap(revision => revision.locations).find(found => found.paragraph === ordinal)?.label
    expect([label(1), label(2)], "precondition: the reader reads both list numbers").toEqual(["1.", LONG_LABEL])
    expect(LONG_LABEL.length).toBeGreaterThan(16)
    const { render } = await openDocument(page, STRUCTURES)
    await expect(paragraphAt(render, 2), "the page prints the long list number").toContainText(LONG_LABEL)
    const place = cardOf(page).locator("p.change-timeline-place")
    await hoverOpen(page, await markFor(render, 1))
    await expect(place).toHaveText(/^Clause 1 · /)
    await cardShot(page, "short-label-clause")
    await leave(page)
    await hoverOpen(page, await markFor(render, 2))
    await expect(place).toHaveText(/^Paragraph 2 · /)
    await expect(place).not.toContainText(LONG_LABEL)
    await cardShot(page, "long-label-fallback")
  })

  await step("20 no page errors, unexpected requests or browser dialogs", async () => {
    expect(pageErrors).toEqual([])
    expect(unexpected).toEqual([])
    expect(dialogs).toEqual([])
  })
}
finally {
  const summary = (rows: { method: string; path: string }[]) =>
    Object.entries(Object.groupBy(rows, row => `${row.method} ${row.path}`)).map(([endpoint, found]) => ({ endpoint, count: found!.length }))
  await writeFile(path.join(evidence, "change-timeline-results.json"), JSON.stringify({
    results, pageErrors, consoleErrors, unexpected, dialogs, stubbed: [...new Set(stubbed)],
    engineEndpoints: summary(engineRequests), ingestEndpoints: summary(ingestRequests), measured: { ...measured, inPane },
  }, null, 2) + "\n")
  await browser?.close().catch(() => undefined)
  await stopIngest().catch(error => console.error(error))
  engine.stop(true)
  await rm(temporary, { recursive: true, force: true })
  if (results.some(row => row.status === "failed")) process.exitCode = 1
}
