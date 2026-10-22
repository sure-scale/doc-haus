// Working Word review acceptance against a separately started Vite server.
// The runner starts a real ingest service on DOCHAUS_E2E_INGEST_PORT with a
// temporary data root and settings directory. Its legal generation calls reach
// a scripted local engine-protocol fixture, never a provider. Browser engine
// traffic is intercepted; unexpected requests fail the run. Screenshots and
// the JSON summary are written only beneath DOCHAUS_EVIDENCE_DIR.
import { createHash } from "node:crypto"
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import type { Browser } from "@playwright/test"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"

type Bytes = Uint8Array
type XmlNode = { nodeType: number; nodeValue: string | null; parentNode: XmlNode | null; childNodes: ArrayLike<XmlNode> }
type XmlElement = XmlNode & {
  namespaceURI: string | null; localName: string | null; textContent: string | null
  attributes: ArrayLike<{ namespaceURI: string | null; localName: string | null; name: string; value: string }>
  getElementsByTagNameNS: (namespace: string, name: string) => ArrayLike<XmlElement>
}

const repository = path.resolve(import.meta.dir, "../../..")
const ingestDirectory = path.join(repository, "services/ingest")
// Reuse the archive/XML packages the Word runtime already ships with.
const { zipSync, unzipSync, strFromU8, strToU8 } = await import(Bun.resolveSync("fflate", path.join(repository, "dochaus"))) as {
  zipSync: (files: Record<string, Bytes>) => Bytes; unzipSync: (bytes: Bytes) => Record<string, Bytes>
  strFromU8: (bytes: Bytes) => string; strToU8: (text: string) => Bytes
}
const { DOMParser } = await import(Bun.resolveSync("@xmldom/xmldom", path.join(repository, "dochaus"))) as {
  DOMParser: new () => { parseFromString: (xml: string, type: string) => { documentElement: XmlElement } }
}
const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")

const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
const ingestPort = Number(process.env.DOCHAUS_E2E_INGEST_PORT)
if (!baseURL || !evidence || !Number.isInteger(ingestPort) || ingestPort < 1024)
  throw new Error("Set PLAYWRIGHT_BASE_URL, DOCHAUS_EVIDENCE_DIR and DOCHAUS_E2E_INGEST_PORT (the port Vite's /ingest proxy targets).")
await mkdir(evidence, { recursive: true })

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PT = "http://powertools.codeplex.com/2011"
const NAME = "synthetic-services-agreement.docx"
// A second deliverable, opened only to check that switching documents clears the first one's Prepare result.
const OTHER = "synthetic-supply-schedule.docx"
const FIRST = "Extend the delivery period to seven days, raise the service credit to three percent and say the services are performed."
const SECOND = "Extend the payment period to forty-five days."
const THIRD = "Raise the liability cap to two hundred thousand pounds."
// Every legal-pass answer is scripted against the synthetic document below.
const plans: Record<string, { before_quote: string; replacement: string; target_context: string; reason: string }[]> = {
  [FIRST]: [
    { before_quote: "within five days", replacement: "within seven days", target_context: "numbered delivery obligation", reason: "Requested delivery period." },
    { before_quote: "Two percent", replacement: "Three percent", target_context: "service credit table value", reason: "Requested service credit." },
    // Same numbered paragraph as the delivery period: stale once that is accepted.
    { before_quote: "deliver the services", replacement: "perform the services", target_context: "numbered delivery obligation", reason: "Requested service wording." },
  ],
  [SECOND]: [
    { before_quote: "within thirty days", replacement: "within forty-five days", target_context: "numbered payment obligation", reason: "Requested payment period." },
  ],
  [THIRD]: [
    { before_quote: "One hundred thousand pounds", replacement: "Two hundred thousand pounds", target_context: "liability cap table value", reason: "Requested liability cap." },
  ],
}
// Open items the legal pass reports beside its proposals.
const openItems: Record<string, string[]> = {
  [FIRST]: ["Confirm whether the revised service credit applies to invoices already issued."],
  [THIRD]: ["Confirm with the client whether the cap applies per claim or in aggregate."],
}
const replacements = Object.fromEntries(Object.values(plans).flat().map(row => [row.before_quote, row.replacement]))

const results: { step: string; status: "passed" | "failed"; error?: string }[] = []
const pageErrors: string[] = []
const consoleErrors: string[] = []
const unexpected: string[] = []
const engineRequests: { method: string; path: string }[] = []
const stubbed: string[] = []
const fixtureRequests: { method: string; path: string; directory: string | null; instructions?: string }[] = []
const hashes: Record<string, string> = {}
const hash = (bytes: Bytes) => createHash("sha256").update(bytes).digest("hex")

// A minimal synthetic package: formatted runs, a decimal numbered list and a table.
function syntheticDocx() {
  const run = (text: string, props = "") => `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`
  const paragraph = (runs: string, props = "") => `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ""}${runs}</w:p>`
  const item = (runs: string) => paragraph(runs, `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`)
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${paragraph(run(text))}</w:tc>`
  const row = (...cells: string[]) => `<w:tr>${cells.map(cell).join("")}</w:tr>`
  const body = [
    paragraph(run("SYNTHETIC SERVICES AGREEMENT", `<w:b/><w:sz w:val="32"/>`)),
    paragraph(run("This agreement is ") + run("synthetic fixture data", `<w:i/><w:color w:val="1F4E79"/>`) + run(" for acceptance testing.")),
    item(run("The ") + run("Supplier", `<w:b/>`) + run(" shall deliver the services within five days.")),
    item(run("The Customer shall pay each invoice within thirty days.")),
    item(run("Either party may terminate on ninety days notice.")),
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>${
      row("Term", "Value") + row("Service credit", "Two percent") + row("Liability cap", "One hundred thousand pounds")}</w:tbl>`,
    paragraph(run("Signed", `<w:b/><w:u w:val="single"/>`) + run(" for and on behalf of the parties.")),
  ].join("")
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`,
    "word/numbering.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style><w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/></w:style></w:styles>`,
    "word/settings.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/></w:settings>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  }
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, xml]) => [name, strToU8(xml)])))
}

// Structural view of a Word package. Paragraph and cell text read accepted text
// (w:t, never w:delText), so a tracked redline reports its proposed result.
function inspect(bytes: Bytes) {
  const parts = unzipSync(bytes)
  const root = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]), "application/xml").documentElement
  const all = (node: XmlElement, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
  const text = (node: XmlElement) => all(node, "t").map(value => value.textContent ?? "").join("")
  const canonical = (node: XmlNode | undefined): string => {
    if (!node) return "null"
    if (node.nodeType === 3 || node.nodeType === 4) return JSON.stringify(node.nodeValue)
    if (node.nodeType !== 1) return ""
    const element = node as XmlElement
    const attributes = Array.from(element.attributes).filter(value => value.namespaceURI !== "http://www.w3.org/2000/xmlns/" && value.namespaceURI !== PT)
      .map(value => [value.namespaceURI ?? "", value.localName ?? value.name, value.value]).sort()
    return JSON.stringify([element.namespaceURI ?? "", element.localName, attributes, Array.from(element.childNodes).map(canonical).filter(Boolean)])
  }
  const paragraphs = all(root, "p")
  return {
    parts: Object.keys(parts).sort(),
    numbering: parts["word/numbering.xml"] ? hash(parts["word/numbering.xml"]) : null,
    paragraphs: paragraphs.map(text),
    paragraphXml: paragraphs.map(node => canonical(node)),
    numbered: paragraphs.filter(node => all(node, "numPr").length).map(node => ({ text: text(node), numPr: canonical(all(node, "numPr")[0]) })),
    cells: all(root, "tc").map(text),
    tables: all(root, "tbl").length,
    runs: all(root, "r").map(node => ({ text: text(node), rPr: canonical(all(node, "rPr")[0]) })),
    insertions: all(root, "ins").length,
    deletions: all(root, "del").length,
  }
}
const amended = (value: string, applied: string[]) => applied.reduce((text, quote) => text.replace(quote, replacements[quote]), value)
// Only the requested words change. Numbering, table cells, untouched run
// properties and every sibling paragraph stay exactly as seeded.
function expectPreserved(seeded: ReturnType<typeof inspect>, actual: ReturnType<typeof inspect>, applied: string[]) {
  expect(actual.parts).toEqual(seeded.parts)
  expect(actual.numbering).toBe(seeded.numbering)
  expect(actual.paragraphs).toEqual(seeded.paragraphs.map(text => amended(text, applied)))
  expect(actual.numbered).toEqual(seeded.numbered.map(row => ({ ...row, text: amended(row.text, applied) })))
  expect(actual.numbered.map(row => row.numPr)).toEqual(Array(3).fill(seeded.numbered[0].numPr))
  expect(actual.tables).toBe(1)
  expect(actual.cells).toEqual(seeded.cells.map(text => amended(text, applied)))
  for (const formatted of ["SYNTHETIC SERVICES AGREEMENT", "synthetic fixture data", "Supplier", "Signed"])
    expect(actual.runs.filter(run => run.text === formatted)).toEqual(seeded.runs.filter(run => run.text === formatted))
  seeded.paragraphs.forEach((text, index) => {
    if (amended(text, applied) === text) expect(actual.paragraphXml[index]).toBe(seeded.paragraphXml[index])
  })
}

// Scripted engine protocol: the model catalog for setup readiness and the two
// legal-generation passes. Anything else is recorded and refused.
const fixture = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: async (request) => {
  const url = new URL(request.url)
  const directory = request.headers.get("x-opencode-directory")
  if (request.method === "GET" && url.pathname === "/api/model") {
    fixtureRequests.push({ method: "GET", path: url.pathname, directory })
    return Response.json({ data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
  }
  if (request.method === "POST" && url.pathname === "/api/dochaus/generate") {
    const { prompt } = await request.json() as { prompt: string }
    const reply = (value: unknown) => Response.json({ data: { text: JSON.stringify(value), model: { providerID: "fixture", id: "scripted" }, elapsed_ms: 1, usage: null } })
    const draft = prompt.match(/<client_instructions encoding="json">([\s\S]*?)<\/client_instructions>/)
    if (draft && prompt.includes("<draft_view")) {
      const instructions = JSON.parse(draft[1]) as string
      fixtureRequests.push({ method: "POST", path: url.pathname, directory, instructions })
      if (!plans[instructions]) { unexpected.push(`fixture instructions ${instructions}`); return new Response(null, { status: 500 }) }
      return reply({ amendments: plans[instructions], open_items: openItems[instructions] ?? [] })
    }
    const candidates = JSON.parse(prompt.match(/<candidates>([\s\S]*)<\/candidates>/)?.[1] ?? "[]") as { amendment_id: string; target_handle: string }[]
    const ids = JSON.parse(prompt.match(/<frozen_amendments>([\s\S]*?)<\/frozen_amendments>/)?.[1] ?? "[]") as { amendment_id: string }[]
    fixtureRequests.push({ method: "POST", path: `${url.pathname}#locator`, directory })
    return reply({ selections: ids.map(({ amendment_id }) => ({ amendment_id, target_handle: candidates.find(row => row.amendment_id === amendment_id)!.target_handle })) })
  }
  unexpected.push(`fixture ${request.method} ${url.pathname}`)
  return new Response(null, { status: 404 })
} })

const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-working-word-e2e-"))
const dataRoot = path.join(temporary, "data")
const settings = path.join(temporary, "settings")
await mkdir(path.join(dataRoot, "workspace", ".preferences"), { recursive: true })
await mkdir(settings, { recursive: true })
await writeFile(path.join(settings, "opencode.json"), JSON.stringify({ model: "fixture/primary", small_model: "fixture/fast" }))
await writeFile(path.join(dataRoot, "workspace", ".preferences", "preferences.json"), JSON.stringify({ attorney: "Synthetic Reviewing Lawyer", firm: "Synthetic Firm" }))
const ingestURL = `http://127.0.0.1:${ingestPort}`
const api = (route: string, init?: RequestInit) => fetch(`${ingestURL}${route}`, { ...init, signal: AbortSignal.timeout(60_000) })
let ingest: ReturnType<typeof Bun.spawn> | undefined
let starts = 0

async function portFree(port: number) {
  try { const server = Bun.listen({ hostname: "127.0.0.1", port, socket: { data() {} } }); server.stop(true); return true }
  catch { return false }
}
// Only explicit isolation settings reach the service: no inherited workspace,
// runtime, engine address or credentials from the caller's environment.
async function startIngest() {
  if (!await portFree(ingestPort)) throw new Error(`Port ${ingestPort} is occupied; choose a free DOCHAUS_E2E_INGEST_PORT.`)
  const log = Bun.file(path.join(temporary, `ingest-${++starts}.log`))
  ingest = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: ingestDirectory, stdout: log, stderr: log,
    env: {
      PATH: process.env.PATH ?? "", HOME: process.env.HOME ?? temporary, TMPDIR: process.env.TMPDIR ?? tmpdir(),
      // Mirror scripts/data-paths.sh: the launcher exports both derived roots.
      DOCHAUS_DATA_ROOT: dataRoot, WORKSPACE_ROOT: path.join(dataRoot, "workspace"), DOCHAUS_RUNTIME_DIR: path.join(dataRoot, "runtime"), DOCHAUS_DIR: settings, INGEST_HOST: "127.0.0.1", INGEST_PORT: String(ingestPort),
      OPENCODE_URL: `http://127.0.0.1:${fixture.port}`, OPENCODE_SERVER_PASSWORD: "", DOCHAUS_MODEL_CACHE_DIR: path.join(temporary, "models"),
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
  const deadline = Date.now() + 10_000
  while (!await portFree(ingestPort)) {
    if (Date.now() > deadline) throw new Error(`Port ${ingestPort} stayed occupied after ingest stopped.`)
    await Bun.sleep(100)
  }
}

let capture: ((name: string) => Promise<unknown>) | undefined
async function step(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ step: name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) {
    results.push({ step: name, status: "failed", error: error instanceof Error ? error.message : String(error) })
    console.log(`FAIL ${name}`)
    await capture?.(`failed-${name.split(" ")[0]}`).catch(() => undefined)
    throw error
  }
}

const file = { path: "", dir: "", matter: "" }
let browser: Browser | undefined
try {
  await startIngest()
  browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
  const context = await browser!.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const origin = new URL(baseURL).origin
  // Engine traffic is answered here. The Word flow's ingest traffic continues
  // through Vite's /ingest proxy to the isolated service.
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
    if (url.pathname.startsWith("/ingest/")) return route.continue()
    if (!url.pathname.startsWith("/engine/")) return route.continue()
    const pathname = url.pathname.slice("/engine".length)
    engineRequests.push({ method: request.method(), path: pathname })
    if (request.method() === "GET" && pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    // Startup connection checks are unrelated to Word review and must never
    // contact a provider. Mirror the explicit synthetic health result used by
    // the other web browser runners.
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
  const dialog = page.getByRole("dialog", { name: NAME, exact: true })
  const cards = dialog.locator(".changes-panel .change-card")
  const shot = (name: string) => page.screenshot({ path: path.join(evidence, `${name}.png`) })
  capture = shot
  const current = async () => hash(await Bun.file(file.path).bytes())
  const pending = async () => await (await api(`/matters/${file.matter}/redlines?name=${encodeURIComponent(NAME)}`)).json() as { id: number; new_text: string; stale?: boolean }[]
  const render = dialog.locator(".docx-render")
  const status = dialog.locator("#document-edit-request").getByRole("status")
  const open = dialog.getByRole("list", { name: "Open items", exact: true }).getByRole("listitem")
  const acceptAll = dialog.getByRole("button", { name: "Accept all", exact: true })
  const staleReason = "The paragraph this change targets has changed since it was prepared. Prepare a fresh change list."
  const exportRefusal = "A proposal no longer resolves against this Word revision."
  const staleNotice = "1 stale proposal is left out of the tracked changes and the redline download. Reject it to clear it; prepare a fresh change list if the change is still needed."
  const working = () => readWorkingDocx(file.dir, file.path)?.source_sha256
  const resolved = (action: "accept" | "reject") => page.waitForResponse((response: { url: () => string }) => new RegExp(`/redlines/\\d+/${action}$`).test(new URL(response.url()).pathname))
  async function openDocument(navigate = true) {
    if (navigate) await page.goto(`${baseURL!.replace(/\/$/, "")}/matter/${file.matter}`)
    // The drawer's open state persists per browser; open it only when closed.
    const documents = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
    await expect(documents).toBeVisible()
    if (await documents.getAttribute("aria-expanded") !== "true") await documents.click()
    await page.getByRole("button", { name: NAME, exact: true }).click()
    await expect(dialog).toBeVisible()
    // The first open loads the Word runtime; later renders are quick.
    await expect(dialog.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
  }
  let seeded: ReturnType<typeof inspect>

  await step("1 seed synthetic DOCX with a numbered list and a table through the isolated ingest", async () => {
    const matter = await (await api("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic Working Word Review" }) })).json() as { id: string; dir: string }
    expect(matter.dir.startsWith(dataRoot)).toBe(true)
    const form = new FormData(); form.append("file", new File([syntheticDocx()], NAME))
    expect((await api(`/matters/${matter.id}/deliverables`, { method: "POST", body: form })).status).toBe(200)
    Object.assign(file, { matter: matter.id, dir: matter.dir, path: path.join(matter.dir, NAME) })
    const other = new FormData(); other.append("file", new File([syntheticDocx()], OTHER))
    expect((await api(`/matters/${matter.id}/deliverables`, { method: "POST", body: other })).status).toBe(200)
    hashes.seeded = await current()
    seeded = inspect(await Bun.file(file.path).bytes())
    expect(seeded.numbered).toHaveLength(3)
    expect(seeded.cells).toEqual(["Term", "Value", "Service credit", "Two percent", "Liability cap", "One hundred thousand pounds"])
  })

  await step("2 open the Word document in the matter viewer and render every page", async () => {
    await openDocument()
    await expect(render).toContainText("The Supplier shall deliver the services within five days.")
    await expect(render.locator("table")).toHaveCount(1)
    await expect(render.locator("td, th").filter({ hasText: "Two percent" })).toHaveCount(1)
    await shot("02-viewer-open")
    await render.screenshot({ path: path.join(evidence, "02-document-render.png") })
  })

  // The line counts the batch's proposals once the change list has refetched;
  // the open items sit in their own list.
  async function prepare(instructions: string, message: string) {
    const request = dialog.getByRole("button", { name: "Request edits", exact: true })
    if (await request.getAttribute("aria-expanded") !== "true") await request.click()
    await dialog.getByLabel("Describe the changes you want to review", { exact: true }).fill(instructions)
    await dialog.getByRole("button", { name: "Prepare changes", exact: true }).click()
    await expect(status).toHaveText(message, { timeout: 60_000 })
    await expect(open).toHaveText(openItems[instructions] ?? [])
    await expect(render).toBeVisible()
  }

  await step("3 request edits and prepare three proposals without changing the Word file", async () => {
    await prepare(FIRST, "3 proposed changes ready for review.")
    await expect(dialog.locator(".changes-count")).toHaveText("3 pending changes")
    await expect(cards).toHaveCount(3)
    await expect(cards.nth(0)).toContainText("within seven days")
    await expect(cards.nth(1)).toContainText("Three percent")
    await expect(cards.nth(2)).toContainText("perform the services")
    await expect(cards.nth(0)).toContainText("Synthetic Reviewing Lawyer")
    await expect(acceptAll).toBeEnabled()
    expect(await current()).toBe(hashes.seeded)
    expect(working()).toBeUndefined()
    // The tracked-changes export carries every proposal over the intact structure.
    const [download] = await Promise.all([page.waitForEvent("download"), (async () => { await dialog.getByRole("button", { name: "More download options", exact: true }).click(); await dialog.getByRole("button", { name: "Download redline", exact: true }).click() })()])
    expect(download.suggestedFilename()).toBe("synthetic-services-agreement (tracked changes).docx")
    await download.saveAs(path.join(evidence, "03-tracked-changes.docx"))
    const redline = inspect(await Bun.file(path.join(evidence, "03-tracked-changes.docx")).bytes())
    expect(redline.insertions).toBeGreaterThanOrEqual(2)
    expect(redline.deletions).toBeGreaterThanOrEqual(2)
    expect(redline.paragraphs).toEqual(seeded.paragraphs.map(text => amended(text, ["within five days", "Two percent", "deliver the services"])))
    expect(redline.numbered.map(row => row.numPr)).toEqual(seeded.numbered.map(row => row.numPr))
    expect(redline.cells).toEqual(seeded.cells.map(text => amended(text, ["Two percent"])))
    expect(await current()).toBe(hashes.seeded)
    await shot("03-prepared")
  })

  await step("4 accept card 1; the card on another paragraph stays acceptable and the same-paragraph card goes stale", async () => {
    const accepted = resolved("accept")
    await cards.nth(0).getByRole("button", { name: "Accept", exact: true }).click()
    expect((await accepted).status()).toBe(200)
    await expect(cards).toHaveCount(2)
    const credit = cards.filter({ hasText: "Three percent" })
    const sibling = cards.filter({ hasText: "perform the services" })
    await expect(credit).not.toContainText(staleReason)
    await expect(credit.getByRole("button", { name: "Accept", exact: true })).toBeEnabled()
    await expect(sibling).toContainText(staleReason)
    await expect(sibling.getByRole("button", { name: "Accept", exact: true })).toBeDisabled()
    await expect(sibling.getByRole("button", { name: "Reject", exact: true })).toBeEnabled()
    await expect(acceptAll).toBeDisabled()
    await expect(render).toContainText("The Supplier shall deliver the services within seven days.")
    // One of the batch is left to review; the stale one is the stale notice's. Open items stay.
    await expect(status).toHaveText("1 proposed change ready for review.")
    await expect(open).toHaveText(openItems[FIRST])
    hashes.firstAccepted = await current()
    expect(hashes.firstAccepted).not.toBe(hashes.seeded)
    expect(working()).toBe(hashes.firstAccepted)
    expect((await pending()).map(row => [row.new_text, row.stale])).toEqual([["Three percent", undefined], ["perform the services", true]])
    await shot("04-stale-card")
    const second = resolved("accept")
    await credit.getByRole("button", { name: "Accept", exact: true }).click()
    expect((await second).status()).toBe(200)
    await expect(cards).toHaveCount(1)
    await expect(cards.nth(0)).toContainText("perform the services")
    await expect(cards.nth(0)).toContainText(staleReason)
    await expect(render.locator("td, th").filter({ hasText: "Three percent" })).toHaveCount(1)
    // Nothing fresh is left from the batch, so the line goes; the open items do not.
    await expect(status).toHaveCount(0)
    await expect(open).toHaveText(openItems[FIRST])
    hashes.accepted = await current()
    expect(hashes.accepted).not.toBe(hashes.firstAccepted)
    expect(working()).toBe(hashes.accepted)
    expect((await pending()).map(row => row.stale)).toEqual([true])
    await shot("04-other-paragraph-accepted")
  })

  await step("5 API refuses the stale card, a hashless revision upload and a deliverable overwrite", async () => {
    const [stale] = await pending()
    const form = () => { const body = new FormData(); body.append("file", new File([syntheticDocx()], NAME)); return body }
    expect((await api(`/matters/${file.matter}/redlines/${stale.id}/accept`, { method: "POST" })).status).toBe(409)
    expect(await current()).toBe(hashes.accepted)
    expect((await api(`/matters/${file.matter}/documents/revisions`, { method: "POST", body: form() })).status).toBe(409)
    expect(await current()).toBe(hashes.accepted)
    expect((await api(`/matters/${file.matter}/deliverables`, { method: "POST", body: form() })).status).toBe(409)
    expect(await current()).toBe(hashes.accepted)
    expect(working()).toBe(hashes.accepted)
    expect((await pending()).map(row => [row.id, row.stale])).toEqual([[stale.id, true]])
  })

  await step("6 with only a stale proposal pending, Download returns the accepted Word file and the redline export leaves the stale change out", async () => {
    // No pending proposal can render as tracked changes, so only Download is offered.
    await expect(dialog.getByRole("button", { name: "Download redline", exact: true })).toHaveCount(0)
    await expect(dialog.getByRole("status").filter({ hasText: "stale proposal" })).toHaveText(staleNotice)
    await expect(dialog.getByText(exportRefusal, { exact: false })).toHaveCount(0)
    const content = page.waitForResponse((response: { url: () => string }) => new URL(response.url()).pathname.endsWith("/documents/content"))
    const [download] = await Promise.all([page.waitForEvent("download"), dialog.getByRole("button", { name: "Download", exact: true }).click()])
    expect((await content).status()).toBe(200)
    expect(download.suggestedFilename()).toBe(NAME)
    await download.saveAs(path.join(evidence, "06-download.docx"))
    const bytes = await Bun.file(path.join(evidence, "06-download.docx")).bytes()
    expect(hash(bytes)).toBe(hashes.accepted)
    expect(hash(bytes)).toBe(working())
    expectPreserved(seeded, inspect(bytes), ["within five days", "Two percent"])
    // The redline export no longer refuses: it omits the stale row, which leaves
    // nothing to track, so it is the accepted file without the stale change.
    const exported = await api(`/matters/${file.matter}/documents/redlined?name=${encodeURIComponent(NAME)}`)
    expect(exported.status).toBe(200)
    const redlined = await exported.bytes()
    await writeFile(path.join(evidence, "06-redlined-export.docx"), redlined)
    const tracked = inspect(redlined)
    expect(tracked.insertions).toBe(0)
    expect(tracked.deletions).toBe(0)
    expect(strFromU8(unzipSync(redlined)["word/document.xml"])).not.toContain("perform the services")
    expect(tracked.paragraphs).toContain("The Supplier shall deliver the services within seven days.")
    expect(hash(redlined)).toBe(hashes.accepted)
    expectPreserved(seeded, tracked, ["within five days", "Two percent"])
    // Downloading resolves nothing: the stale row stays pending and stale.
    expect(await current()).toBe(hashes.accepted)
    expect((await pending()).map(row => row.stale)).toEqual([true])
    await expect(cards).toHaveCount(1)
    await expect(cards.nth(0)).toContainText(staleReason)
    await shot("06-stale-excluded-download")
  })

  await step("7 restart ingest on the same root; reload keeps the hash, stale card and open items; a fresh Prepare and Accept succeed", async () => {
    await stopIngest()
    await startIngest()
    await page.reload()
    await openDocument(false)
    expect(await current()).toBe(hashes.accepted)
    expect(working()).toBe(hashes.accepted)
    await expect(render).toContainText("within seven days")
    await expect(cards).toHaveCount(1)
    await expect(cards.nth(0)).toContainText(staleReason)
    await expect(cards.nth(0).getByRole("button", { name: "Accept", exact: true })).toBeDisabled()
    // The last Prepare's open items are stored with the matter, not in the page.
    await expect(open).toHaveText(openItems[FIRST])
    await shot("07-after-restart")
    await prepare(SECOND, "1 proposed change ready for review.")
    expect(await current()).toBe(hashes.accepted)
    await expect(cards).toHaveCount(2)
    const fresh = cards.filter({ hasText: "within forty-five days" })
    await expect(fresh).not.toContainText(staleReason)
    await expect(fresh.getByRole("button", { name: "Accept", exact: true })).toBeEnabled()
    await expect(acceptAll).toBeDisabled()
    const accepted = resolved("accept")
    await fresh.getByRole("button", { name: "Accept", exact: true }).click()
    expect((await accepted).status()).toBe(200)
    await expect(cards).toHaveCount(1)
    await expect(cards.nth(0)).toContainText("perform the services")
    await expect(render).toContainText("The Customer shall pay each invoice within forty-five days.")
    await expect(status).toHaveCount(0)
    hashes.reaccepted = await current()
    expect(hashes.reaccepted).not.toBe(hashes.accepted)
    expect(working()).toBe(hashes.reaccepted)
    await shot("07-fresh-accepted")
  })

  await step("8 Reject leaves the bytes unchanged; with nothing pending, Download returns the exact Word file", async () => {
    await prepare(THIRD, "1 proposed change ready for review.")
    await expect(cards).toHaveCount(2)
    for (const text of ["Two hundred thousand pounds", "perform the services"]) {
      const rejected = resolved("reject")
      await cards.filter({ hasText: text }).getByRole("button", { name: "Reject", exact: true }).click()
      expect((await rejected).status()).toBe(200)
      await expect(cards.filter({ hasText: text })).toHaveCount(0)
      expect(await current()).toBe(hashes.reaccepted)
      expect(working()).toBe(hashes.reaccepted)
    }
    await expect(dialog.locator(".changes-panel")).toHaveCount(0)
    expect(await pending()).toEqual([])
    await expect(status).toHaveCount(0)
    await expect(open).toHaveText(openItems[THIRD])
    await expect(render).toContainText("One hundred thousand pounds")
    const [download] = await Promise.all([page.waitForEvent("download"), dialog.getByRole("button", { name: "Download", exact: true }).click()])
    expect(download.suggestedFilename()).toBe(NAME)
    await download.saveAs(path.join(evidence, "08-download.docx"))
    const bytes = await Bun.file(path.join(evidence, "08-download.docx")).bytes()
    expect(hash(bytes)).toBe(hashes.reaccepted)
    expectPreserved(seeded, inspect(bytes), ["within five days", "Two percent", "within thirty days"])
    await shot("08-rejected-and-downloaded")
  })

  await step("9 switching documents shows only each document's own open items; switching back restores them", async () => {
    await expect(open).toHaveText(openItems[THIRD])
    // Close the modal before switching files; focus cannot reach the drawer
    // behind an open document. The same keyboard path works on every viewport.
    await dialog.getByRole("button", { name: "Close", exact: true }).press("Enter")
    await page.getByRole("button", { name: OTHER, exact: true }).press("Enter")
    const other = page.getByRole("dialog", { name: OTHER, exact: true })
    await expect(other.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
    await expect(dialog).toHaveCount(0)
    await expect(other.getByRole("list", { name: "Open items", exact: true })).toHaveCount(0)
    await expect(other.getByRole("button", { name: "Request edits", exact: true })).toHaveAttribute("aria-expanded", "false")
    await expect(other.locator("#document-edit-request")).toHaveCount(0)
    await shot("09-switched-document")
    await other.getByRole("button", { name: "Close", exact: true }).press("Enter")
    await page.getByRole("button", { name: NAME, exact: true }).press("Enter")
    await expect(dialog.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
    await expect(open).toHaveText(openItems[THIRD])
  })

  await step("no page errors, unexpected requests or unscripted generation", async () => {
    expect(pageErrors).toEqual([])
    expect(unexpected).toEqual([])
    const generations = fixtureRequests.filter(row => row.path.startsWith("/api/dochaus/generate"))
    expect(generations.map(row => row.instructions ?? "locator")).toEqual([FIRST, "locator", SECOND, "locator", THIRD, "locator"])
    expect(generations.every(row => row.directory === file.dir)).toBe(true)
  })
}
finally {
  await writeFile(path.join(evidence, "working-word-review-results.json"), JSON.stringify({
    results, pageErrors, consoleErrors, unexpected, stubbed: [...new Set(stubbed)],
    engineEndpoints: Object.entries(Object.groupBy(engineRequests, row => `${row.method} ${row.path}`)).map(([endpoint, rows]) => ({ endpoint, count: rows!.length })),
    fixtureRequests, hashes,
  }, null, 2) + "\n")
  await browser?.close().catch(() => undefined)
  await stopIngest().catch(error => console.error(error))
  fixture.stop(true)
  await rm(temporary, { recursive: true, force: true })
}
