// Dashboard visual and interaction regression against an already-running web
// preview. Every API request is intercepted with synthetic data relative to a
// pinned clock; unknown requests are refused before any network I/O.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
type Page = import("@playwright/test").Page
type Route = import("@playwright/test").Route
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const origin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: Record<string, unknown>[] = []
const screens: Record<string, unknown>[] = []
const allStates: State[] = []

// Thursday 29 October 2026, 10:00 in London (GMT). Fixture times sit between
// 08:00 and 16:00 UTC, the same calendar day in London either side of the
// clock change.
const NOW = Date.UTC(2026, 9, 29, 10)
const DAY = 86_400_000
const TODAY = Date.UTC(2026, 9, 29)
const at = (daysAgo: number, hour: number, minute = 0) => TODAY - daysAgo * DAY + hour * 3_600_000 + minute * 60_000
const iso = (time: number) => new Date(time).toISOString()

// Deterministic pseudo-random numbers, so every run draws the same practice.
function rng(seed: number) {
  return () => {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

type Matter = { id: string; title: string; reference?: string; client?: string; jurisdictions: string[]; dir: string; created_at: number; sample?: boolean }
type Session = { id: string; parentID?: string; projectID: string; title: string; cost: number; tokens: { input: number; output: number; reasoning: number; cache: { read: number; write: number } }; time: { created: number; updated: number }; location: { directory: string }; metadata: Record<string, unknown> }

const MATTERS: Omit<Matter, "dir" | "created_at">[] = [
  { id: "m-alder", title: "Alder House · Studio lease renewal", reference: "AH-2026-014", client: "Alder House Studio", jurisdictions: ["EW"] },
  { id: "m-northbank", title: "Northbank Works · Supply agreement", reference: "NB-2026-008", client: "Northbank Works", jurisdictions: ["EW"] },
  { id: "m-fernlea", title: "Fernlea Bakery · Shareholders' agreement", reference: "FB-2026-003", client: "Fernlea Bakery Co.", jurisdictions: ["EW"] },
  { id: "m-quarry", title: "Quarry Lane Cycles · Franchise terms", reference: "QL-2026-021", client: "Quarry Lane Cycles", jurisdictions: ["SC"] },
  { id: "m-marlowe", title: "Marlowe & Pike Design · Master services agreement and statement of work", reference: "MP-2026-011", client: "Marlowe & Pike Design", jurisdictions: ["EW"] },
  { id: "m-tidewater", title: "Tidewater Robotics · Mutual NDA set", reference: "TR-2026-030", client: "Tidewater Robotics", jurisdictions: ["EW"] },
  { id: "m-copperleaf", title: "Copperleaf Gardens · Employment contracts", jurisdictions: ["EW"] },
  { id: "m-sample", title: "Kestrel Books · Distribution agreement", client: "Kestrel Books", jurisdictions: ["EW"], sample: true },
]

function tokensFor(cost: number, random: () => number) {
  const scale = 0.7 + random() * 0.6
  return {
    input: Math.round(cost * 160_000 * scale),
    output: Math.round(cost * 22_000 * scale),
    reasoning: Math.round(cost * 14_000 * scale),
    cache: { read: Math.round(cost * 520_000 * scale), write: Math.round(cost * 36_000 * scale) },
  }
}

// A populated practice: about a year of conversations, quiet at first, a
// holiday lull in August and a busy autumn; weekdays busier than weekends and
// Tuesday the busiest. The last fortnight is spelled out so the week card has
// clear high, middle and low days.
function practice() {
  const random = rng(20261029)
  const sessions: (Session & { matter: string })[] = []
  let serial = 0
  const add = (matter: string, created: number, cost: number, title: string, children = 0, ids: string[] = []) => {
    const id = `ses_${String(++serial).padStart(4, "0")}`
    const share = children ? 0.45 + random() * 0.2 : 1
    const parent = Math.round(cost * share * 100) / 100
    sessions.push({ matter, id, projectID: "fixture-project", title, cost: parent, tokens: tokensFor(parent, random), time: { created, updated: created + 25 * 60_000 }, location: { directory: "" }, metadata: {} })
    let left = Math.round((cost - parent) * 100) / 100
    for (let index = 0; index < children; index++) {
      const childCost = index === children - 1 ? left : Math.round(left * (0.3 + random() * 0.4) * 100) / 100
      left = Math.round((left - childCost) * 100) / 100
      const childId = ids[index] ?? `ses_${String(++serial).padStart(4, "0")}`
      sessions.push({ matter, id: childId, parentID: id, projectID: "fixture-project", title: `Stage ${index + 1}`, cost: childCost, tokens: tokensFor(childCost, random), time: { created: created + (index + 1) * 4 * 60_000, updated: created + (index + 2) * 4 * 60_000 }, location: { directory: "" }, metadata: {} })
    }
    return id
  }
  const weekday = [1.0, 1.3, 1.05, 1.0, 0.8, 0.12, 0.06] // Monday first
  const pool = ["m-alder", "m-northbank", "m-fernlea", "m-quarry", "m-marlowe", "m-tidewater", "m-copperleaf"]
  const topics = ["Clause review", "Summarise the heads of terms", "Compare against the playbook", "Draft a cover email", "Check the definitions", "Notice period question"]
  for (let daysAgo = 364; daysAgo >= 14; daysAgo--) {
    const date = new Date(TODAY - daysAgo * DAY)
    const month = date.getUTCMonth()
    const streak = date.getUTCMonth() === 6 && date.getUTCDate() >= 6 && date.getUTCDate() <= 21 // a 16-day closing push in July
    const base = daysAgo > 300 ? 0.12 : month === 7 ? 0.18 : month >= 8 ? 0.72 : 0.48
    const chance = streak ? 1 : base * weekday[(date.getUTCDay() + 6) % 7]!
    if (random() > chance) continue
    const count = 1 + Math.floor(random() * (month >= 8 || streak ? 4 : 2.4))
    for (let index = 0; index < count; index++) {
      const matter = pool[Math.floor(random() * (daysAgo > 200 ? 3 : pool.length))]!
      add(matter, at(daysAgo, 8 + Math.floor(random() * 7), Math.floor(random() * 60)), Math.round((0.08 + random() * 1.9) * 100) / 100, topics[Math.floor(random() * topics.length)]!, random() < 0.35 ? 1 + Math.floor(random() * 3) : 0)
    }
  }
  // The sample matter, added and explored about six weeks ago.
  add("m-sample", at(40, 9), 0.42, "Explore the sample")
  add("m-sample", at(39, 11), 0.27, "What does clause 7 do?")
  // An older drafting run, outside the 30-day window.
  const quarryOwner = add("m-quarry", at(55, 9), 4.6, "Draft the franchise terms", 3, ["ses_quarry_d", "ses_quarry_r", "ses_quarry_f"])
  // The last fortnight, day by day: [daysAgo, matter, total cost, children].
  const fortnight: [number, string, number, number, string?][] = [
    [13, "m-northbank", 1.2, 0], [13, "m-tidewater", 0.8, 0],
    [10, "m-northbank", 2.1, 1], [10, "m-copperleaf", 1.4, 0], [10, "m-alder", 0.6, 0],
    [9, "m-northbank", 3.2, 3, "northbank"], [9, "m-quarry", 1.1, 0],
    [8, "m-fernlea", 0.9, 0],
    // Thursday 22 October has nothing: the current streak starts the day after.
    [6, "m-marlowe", 2.4, 1], [6, "m-tidewater", 1.6, 0],
    [5, "m-tidewater", 0.45, 0],
    [4, "m-copperleaf", 0.3, 0],
    [3, "m-alder", 5.8, 4, "alder"], [3, "m-marlowe", 3.9, 1], [3, "m-tidewater", 2.2, 1, "tidewater"],
    [2, "m-marlowe", 3.1, 2, "marlowe"], [2, "m-fernlea", 2.7, 0],
    [1, "m-fernlea", 6.4, 3, "fernlea"], [1, "m-alder", 2.9, 1],
    [0, "m-copperleaf", 1.75, 1, "copperleaf"],
  ]
  const owners: Record<string, string> = { quarry: quarryOwner }
  const stageIds: Record<string, string[]> = { quarry: ["ses_quarry_d", "ses_quarry_r", "ses_quarry_f"] }
  for (const [daysAgo, matter, cost, children, run] of fortnight) {
    const ids = run ? Array.from({ length: children }, (_, index) => `ses_${run}_${index + 1}`) : []
    const id = add(matter, at(daysAgo, daysAgo === 0 ? 8 : 9 + Math.floor(random() * 6), daysAgo === 0 ? 30 : 0), cost, run ? "Drafting run" : topics[Math.floor(random() * topics.length)]!, children, ids)
    if (run) { owners[run] = id; stageIds[run] = ids }
  }
  const firstSession = new Map<string, number>()
  for (const session of sessions) firstSession.set(session.matter, Math.min(firstSession.get(session.matter) ?? Infinity, session.time.created))
  const matters: Matter[] = MATTERS.map((row) => ({ ...row, dir: `/fixture/matters/${row.id.slice(2)}`, created_at: (firstSession.get(row.id) ?? at(30, 9)) - 2 * 3_600_000 }))
  for (const session of sessions) session.location.directory = matters.find((row) => row.id === session.matter)!.dir
  return { matters, sessions, owners, stageIds }
}

const stage = (status: string, sessions: string[], started: number, finished: number | null, attempts = 1) => ({ status, attempts, session_ids: sessions, started_at: iso(started), finished_at: finished === null ? null : iso(finished) })

// A lean run's status in the shape isLeanDraftingStatus accepts.
function run(options: { status: "running" | "blocked" | "failed" | "completed"; owner: string; stages: string[]; started: number; updated: number; documents?: { title: string; kind?: "body" | "combined" }[]; questions?: string[]; stop?: string }) {
  const [drafting = [], review = [], finalize = []] = [options.stages.slice(0, 1), options.stages.slice(1, -1), options.stages.slice(-1)]
  const done = options.status === "completed"
  const documents = options.documents ?? []
  return {
    schema: "dochaus.lean.state/2", profile: "lean", revision: 7, status: options.status,
    stop_reason: options.stop ?? null, owner_session_id: options.owner,
    model: { providerID: "fixture", modelID: "primary" },
    created_at: iso(options.started), updated_at: iso(options.updated),
    research: null,
    stages: {
      drafting: stage("done", drafting, options.started, options.started + 600_000),
      "review-legal": stage(options.status === "running" ? "running" : options.status === "failed" ? "failed" : options.status === "blocked" ? "blocked" : "done", review, options.started + 600_000, options.status === "running" ? null : options.started + 1_200_000, options.status === "failed" ? 3 : 1),
      finalize: stage(done ? "done" : "pending", done ? finalize : [], options.started + 1_200_000, done ? options.started + 1_500_000 : null),
      "package-review": stage(done ? "done" : "pending", [], options.started + 1_500_000, done ? options.updated - 60_000 : null),
      gate: stage(done ? "done" : "pending", [], options.updated - 60_000, done ? options.updated : null),
    },
    in_flight: options.status === "running", stopped: false, pending_requests: 0, open_items: [],
    questions: (options.questions ?? []).map((question, index) => ({ id: `q${index + 1}`, stage: "review-legal", question, context: "The instructions do not settle this point.", status: "open", kind: "instruction" })),
    gate: done ? { status: "pass", delivery: "review_draft", blocking: [], disclosed: [], items: [] } : null,
    documents: done ? { status: "completed", documents: [
      // A combined deliverable keeps its clean body row beside the one combined file.
      ...documents.map((row, index) => ({ kind: "body", document: `d${index + 1}`, title: row.title, path: `.dochaus/lean/out/d${index + 1}.docx`, sha256: "0".repeat(64), deliverable: `${row.title}.docx` })),
      ...documents.flatMap((row, index) => row.kind === "combined" ? [{ kind: "combined", document: `d${index + 1}`, title: row.title, path: `${row.title}.docx`, sha256: "3".repeat(64), deliverable: `${row.title}.docx` }] : []),
      ...documents.filter((row) => (row.kind ?? "body") === "body").map((row, index) => ({ kind: "redline", document: `d${index + 1}`, title: row.title, path: `.dochaus/lean/out/d${index + 1}-redline.docx`, sha256: "1".repeat(64), deliverable: `${row.title} - redline.docx` })),
      { kind: "drafting_notes", document: null, title: null, path: ".dochaus/lean/out/notes.docx", sha256: "2".repeat(64), deliverable: "Drafting notes.docx" },
    ] } : null,
    plan: { documents: documents.map((row, index) => ({ id: `d${index + 1}`, title: row.title, purpose: "Synthetic fixture document." })) },
    memo: null, combined_output: null, request: "Synthetic fixture request.",
  }
}

function populated() {
  const { matters, sessions, owners, stageIds } = practice()
  const owned = (key: string) => ({ owner: owners[key]!, stages: stageIds[key]! })
  const created = (key: string) => sessions.find((row) => row.id === owners[key])!.time.created
  const drafting: Record<string, unknown> = {
    "m-alder": run({ status: "completed", ...owned("alder"), started: created("alder"), updated: created("alder") + 3_000_000, documents: [{ title: "Lease renewal" }, { title: "Side letter" }] }),
    "m-northbank": run({ status: "completed", ...owned("northbank"), started: created("northbank"), updated: created("northbank") + 2_400_000, documents: [{ title: "Supply agreement" }] }),
    "m-fernlea": run({ status: "completed", ...owned("fernlea"), started: created("fernlea"), updated: created("fernlea") + 2_700_000, documents: [{ title: "Shareholders' agreement", kind: "combined" }] }),
    "m-quarry": run({ status: "completed", ...owned("quarry"), started: created("quarry"), updated: created("quarry") + 2_100_000, documents: [{ title: "Franchise terms" }] }),
    "m-marlowe": run({ status: "blocked", ...owned("marlowe"), started: created("marlowe"), updated: created("marlowe") + 1_300_000, questions: ["Should the statement of work cap change requests?", "Which party owns the design files on termination?"] }),
    "m-tidewater": run({ status: "failed", ...owned("tidewater"), started: created("tidewater"), updated: created("tidewater") + 1_500_000, stop: "The review stage did not produce a valid output after three tries." }),
    "m-copperleaf": run({ status: "running", ...owned("copperleaf"), started: created("copperleaf"), updated: created("copperleaf") + 600_000 }),
  }
  const documents: Record<string, number> = { "m-alder": 9, "m-northbank": 6, "m-fernlea": 4, "m-quarry": 5, "m-marlowe": 7, "m-tidewater": 3, "m-copperleaf": 8, "m-sample": 3 }
  return { matters, sessions, drafting, documents }
}

// A new lawyer's first days: one matter of their own, the sample, or both.
const FIRST: Matter = { id: "m-harbour", title: "Harbour Lane Bakery · Supplier terms", client: "Harbour Lane Bakery", jurisdictions: ["EW"], dir: "/fixture/matters/harbour", created_at: at(1, 9) }
const SAMPLE: Matter = { id: "m-sample", title: "Kestrel Books · Engagement letter", client: "Kestrel Books", jurisdictions: ["EW"], dir: "/fixture/matters/sample", created_at: at(0, 8), sample: true }
const chat = (matter: Matter, id: string, created: number, title: string, cost = 0.18): Session & { matter: string } =>
  ({ matter: matter.id, id, projectID: "fixture-project", title, cost, tokens: tokensFor(cost, rng(id.length)), time: { created, updated: created + 12 * 60_000 }, location: { directory: matter.dir }, metadata: {} })
function early(options: { first?: number; sample?: boolean; conversations?: number; drafting?: boolean }) {
  const matters = [...(options.first === undefined ? [] : [FIRST]), ...(options.sample ? [SAMPLE] : [])]
  const sessions = [
    ...(options.sample ? [chat(SAMPLE, "ses_sample_1", at(0, 8, 10), "What are the key points in these materials?")] : []),
    ...Array.from({ length: options.conversations ?? 0 }, (_, index) => chat(FIRST, `ses_first_${index + 1}`, at(0, 8, 20 + index * 15), index ? "Which terms favour the supplier?" : "What are the key points in these materials?")),
    ...(options.drafting ? [{ ...chat(FIRST, "ses_first_run", at(0, 9, 30), "Drafting run", 0.9), metadata: { dochaus_role: "lean-drafting-owner" } }] : []),
  ]
  const drafting: Record<string, unknown> = options.drafting ? { [FIRST.id]: run({ status: "running", owner: "ses_first_run", stages: ["ses_first_run_1"], started: at(0, 9, 30), updated: at(0, 9, 40) }) } : {}
  return { matters, sessions, drafting, documents: { [FIRST.id]: options.first ?? 0, [SAMPLE.id]: 3 } as Record<string, number> }
}

type Kind = "new" | "populated" | "error"
type Data = ReturnType<typeof populated>
function fixture(kind: Kind, layout?: unknown, data: Data = kind === "populated" ? populated() : { matters: [], sessions: [], drafting: {}, documents: {} }) {
  return {
    kind, layout, ...data, samplePosts: 0,
    unexpected: [] as string[], errors: [] as string[], sessionDirs: [] as string[], requests: [] as { method: string; path: string }[],
  }
}
type State = ReturnType<typeof fixture>

const jurisdictions = [{ code: "EW", name: "England and Wales", citationStyle: "oscola" }, { code: "SC", name: "Scotland", citationStyle: "scotland" }]
const templates = [
  { name: "Consulting agreement.docx", description: "A starting point for professional services engagements.", placeholders: [{ text: "Client name", kind: "text" }] },
  { name: "Mutual confidentiality agreement.docx", description: "Two-way confidentiality terms for early discussions.", placeholders: [] },
  { name: "Commercial lease renewal.docx", description: "Renewal terms for a commercial unit.", placeholders: [] },
]
const playbooks = [
  { name: "playbook-commercial-services", description: "Preferred positions for services engagements.", content: "# Services agreements", version: "1", last_reviewed: "2026-10-20", builtin: false, updated_at: NOW, matters: [] },
  { name: "playbook-leases", description: "Preferred positions for commercial leases.", content: "# Leases", version: "1", last_reviewed: "2026-09-14", builtin: false, updated_at: NOW, matters: [] },
]
const clients = ["Alder House Studio", "Northbank Works", "Fernlea Bakery Co.", "Quarry Lane Cycles"].map((name, index) => ({ id: index + 1, name, created_at: NOW - (200 - index * 30) * DAY }))
const word = { available: true, origin: null, manifestUrl: null, platform: "macos", wordInstalled: true, canSetup: true, configured: false, registered: false, trusted: false, running: false, status: "not_configured" }

async function open(state: State, width: number, height: number, timezoneId = "Europe/London") {
  allStates.push(state)
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId, locale: "en-GB" })
  await context.clock.setFixedTime(new Date(NOW))
  await context.addInitScript(({ layout, jurisdiction }) => {
    localStorage.setItem("dh.launch-notice.v1", new Date().toDateString())
    localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1")
    if (jurisdiction) localStorage.setItem("dochaus.prefs", JSON.stringify({ defaultJurisdictions: ["EW"] }))
    // Seed the saved layout once; a reload keeps what the page saved.
    if (layout && !sessionStorage.getItem("dashboard-fixture-seeded")) { localStorage.setItem("dh.dashboard.v2", JSON.stringify(layout)); sessionStorage.setItem("dashboard-fixture-seeded", "1") }
  }, { layout: state.layout ?? null, jurisdiction: state.kind === "populated" })
  await context.route("**/*", async (route: Route) => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const method = request.method()
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === origin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (method === "OPTIONS") return route.fulfill({ status: 204, headers })
    state.requests.push({ method, path: pathname })
    const reply = (body: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(body) })
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (method === "POST" && pathname === "/sample-matter") { state.samplePosts++; return reply({ matter: SAMPLE }) }
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    const populatedState = state.kind === "populated"
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    // A new install finished onboarding with only the model step: details and
    // research were skipped, so the app opens normally and nothing else is set.
    if (pathname === "/setup") return reply(populatedState ? { version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" } : { version: 1, ready: true, completed: ["models"], skipped: ["details", "research"], next: "done" })
    if (pathname === "/preferences") return reply(populatedState ? { ...DEFAULT_DRAFTING, attorney: "Alex Example", firm: "Example Partners LLP", practiceAreas: ["Commercial contracts", "Property"], houseStyle: "Arial 11pt, numbered clauses." } : DEFAULT_DRAFTING)
    if (pathname === "/matters") return state.kind === "error" ? reply({ error: "The document service is temporarily unavailable." }, 503) : reply(state.matters)
    const draftingMatch = pathname.match(/^\/matters\/([^/]+)\/drafting-status$/)
    if (draftingMatch) return reply(state.drafting[draftingMatch[1]!] ?? null)
    const matterMatch = pathname.match(/^\/matters\/([^/]+)$/)
    const detail = matterMatch && state.matters.find((row) => row.id === matterMatch[1])
    if (detail) return reply({ ...detail, documents: Array.from({ length: state.documents[detail.id] ?? 0 }, (_, index) => ({ id: `doc-${index + 1}`, name: `Document ${index + 1}.pdf` })) })
    if (pathname === "/word/setup") return reply(word)
    if (pathname === "/templates") return reply({ dir: "/fixture/templates", templates: populatedState ? templates : [] })
    if (pathname === "/playbooks/library") return reply({ dir: "/fixture/playbooks", playbooks: populatedState ? playbooks : [] })
    if (pathname === "/playbooks") return reply(populatedState ? playbooks : [])
    if (pathname === "/clients") return reply(populatedState ? clients : [])
    if (pathname === "/jurisdictions") return reply(jurisdictions)
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    const provider = { id: "fixture", name: "Fixture Models", integrationID: "fixture", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [provider] })
    if (pathname === "/api/model") return reply({ location: { directory: "/fixture" }, data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
    if (pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture", methods: [], connections: [{ type: "credential", id: "fixture-account", label: "Fixture", method: "key" }] }] })
    if (pathname === "/api/session/active") return reply({ data: {} })
    if (pathname === "/api/session") {
      // listSessions sends the matter folder as `directory` (and the client's
      // x-opencode-directory header); answer each matter with its own sessions.
      const directory = url.searchParams.get("directory") ?? request.headers()["x-opencode-directory"]
      if (!directory) { state.unexpected.push(`${method} ${pathname} without a directory`); return route.abort() }
      state.sessionDirs.push(directory)
      const data = state.sessions.filter((row) => row.location.directory === directory).map(({ matter: _, ...row }) => row).sort((a, b) => b.time.updated - a.time.updated)
      return reply({ data, cursor: { previous: null, next: null } })
    }
    if (pathname === "/api/permission/request") return reply({ location: { directory: "/fixture" }, data: [] })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: `data: ${JSON.stringify({ id: "evt_fixture", type: "server.connected", data: {} })}\n\n` })
    state.unexpected.push(`${method} ${pathname}`)
    return route.abort()
  })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", (error: Error) => state.errors.push(error.message))
  return { page, context }
}

async function capture(page: Page, name: string, options: { fullPage?: boolean } = {}) {
  await page.evaluate(async () => { await document.fonts.ready })
  const metrics = await page.evaluate(() => ({
    overflowX: document.documentElement.scrollWidth - document.documentElement.clientWidth,
    dialogs: Array.from(document.querySelectorAll<HTMLElement>('dialog[open], [role="dialog"][aria-modal="true"]')).map(node => { const r = node.getBoundingClientRect(); return { name: node.getAttribute("aria-label") ?? node.getAttribute("aria-labelledby"), left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: innerWidth, height: innerHeight } }),
    // Text cut off inside a card: an element whose content is wider than its box.
    clipped: Array.from(document.querySelectorAll<HTMLElement>(".dash-card *")).filter(node => node.children.length === 0 && !node.closest(".sr-only") && node.textContent?.trim() && node.scrollWidth > node.clientWidth + 1 && getComputedStyle(node).overflow !== "visible").map(node => ({ text: node.textContent!.trim().slice(0, 80), card: node.closest(".dash-card")?.querySelector("h3")?.textContent })),
  }))
  await page.screenshot({ path: path.join(evidence!, `${name}.png`), fullPage: options.fullPage ?? metrics.dialogs.length === 0 })
  screens.push({ name, ...metrics })
  expect(metrics.overflowX, `${name}: document horizontal overflow`).toBeLessThanOrEqual(1)
  for (const dialog of metrics.dialogs) { expect(dialog.left, `${name}: dialog left`).toBeGreaterThanOrEqual(-1); expect(dialog.right, `${name}: dialog right`).toBeLessThanOrEqual(dialog.width + 1); expect(dialog.top, `${name}: dialog top`).toBeGreaterThanOrEqual(-1); expect(dialog.bottom, `${name}: dialog bottom`).toBeLessThanOrEqual(dialog.height + 1) }
}
async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); console.log(`FAIL ${name}: ${error instanceof Error ? error.message : String(error)}`) }
}
const clean = (state: State) => { expect(state.unexpected, "unexpected requests").toEqual([]); expect(state.errors, "page errors").toEqual([]) }
const card = (page: Page, title: string) => page.locator("section.dash-card").filter({ has: page.getByRole("heading", { name: title, exact: true, level: 3 }) })
// The widget menu opener is a <summary>, which has no ARIA button role.
const menuButton = (page: Page, title: string) => page.locator(`summary[aria-label="Options for ${title}"]`)
// What a pointer at the centre of a menu item would hit, recorded for the
// stacking check; the action itself is taken from the keyboard so the flow
// still runs when card content paints over the open menu.
const menuHits: { title: string; action: string; covered: string | null }[] = []
async function menuAction(page: Page, title: string, action: string) {
  const button = card(page, title).getByRole("button", { name: action, exact: true })
  await expect(button).toBeVisible()
  menuHits.push({ title, action, covered: await coveredBy(button) })
  await button.focus()
  await page.keyboard.press("Enter")
}
async function coveredBy(button: import("@playwright/test").Locator) {
  return button.evaluate((node) => {
    const box = node.getBoundingClientRect()
    const hit = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2)
    return !hit || node.contains(hit) ? null : `${hit.tagName.toLowerCase()}.${[...hit.classList].join(".")} "${(hit.textContent ?? "").trim().slice(0, 40)}"`
  })
}
const order = (page: Page) => page.locator(".dash-grid > section.dash-card h3").allTextContents()
// Settled: no skeletons left and the refresh control idle.
async function dashboard(page: Page, expectGrid = true, reload = false) {
  if (reload) await page.reload()
  else await page.goto(`${baseURL}/dashboard`)
  await expect(page.getByRole("heading", { name: /^Good (morning|afternoon|evening)/ })).toBeVisible()
  await expect(page.getByRole("button", { name: "Refresh", exact: true })).toBeEnabled()
  if (expectGrid) await expect(page.locator(".dash-skeleton")).toHaveCount(0)
}

const SIZES = [[375, 812], [768, 1024], [1440, 1000]] as const
const ALL_WIDGETS = ["journey", "setup", "spend", "matter-cost", "draft-cost", "documents", "week", "matter-spend", "attention", "activity", "recent", "tokens", "trend", "actions", "library", "sample", "ideas"]
const STARTING = ["First steps", "Finish setting up", "Sample matter", "Your library", "Things to try"]
const journey = (page: Page) => card(page, "First steps").locator(".dash-journey > li")
const doneSteps = (page: Page) => card(page, "First steps").locator(".dash-journey > li.is-done strong").allTextContents()

try {
  for (const [width, height] of [...SIZES, [1180, 900] as const]) {
    const size = `${width}x${height}`
    if (width !== 1180) await scenario(`New install (${size})`, async () => {
      const state = fixture("new"); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        await expect(page.getByRole("heading", { name: "Good morning", exact: true })).toBeVisible()
        // No figures yet: the dashboard leads with first steps, not empty charts.
        expect(await order(page)).toEqual(STARTING)
        await expect(page.getByRole("group", { name: "Period", exact: true })).toHaveCount(0)
        await expect(journey(page).nth(0)).toHaveAttribute("aria-current", "step")
        await expect(journey(page).nth(0).getByRole("link", { name: "New matter", exact: true })).toHaveAttribute("href", "/")
        await expect(card(page, "First steps").locator(".dash-step-action")).toHaveCount(1)
        const setup = card(page, "Finish setting up")
        await expect(setup.getByText("4 essentials left", { exact: true })).toBeVisible()
        await expect(setup.getByRole("img", { name: "1 of 9 done", exact: true })).toBeVisible()
        await expect(setup.locator(".dash-setup-items > li")).toHaveCount(6)
        await expect(card(page, "Sample matter").getByRole("button", { name: "Add the sample matter", exact: true })).toBeVisible()
        await expect(card(page, "Your library").locator(".dash-library .is-zero")).toHaveCount(4)
        await expect(card(page, "Things to try").locator(".dash-actions > li")).toHaveCount(4)
        await capture(page, `new-install-overview-${size}`)
        expect(state.requests.filter((row) => row.path === "/api/session")).toEqual([])
        clean(state)
      } finally { await context.close() }
    })
    await scenario(`Populated practice (${size})`, async () => {
      const state = fixture("populated"); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        await expect(page.getByRole("heading", { name: "Good morning, Alex Example", exact: true })).toBeVisible()
        // Essentials done and documents drafted: figures only, no onboarding cards.
        await expect(card(page, "Finish setting up")).toHaveCount(0)
        await expect(card(page, "First steps")).toHaveCount(0)
        await expect(card(page, "Documents drafted").locator(".dash-stat-value")).toHaveText("4")
        await expect(card(page, "Needs you").locator(".dash-attention > li")).toHaveCount(2)
        await expect(card(page, "Needs you").getByText("Drafting run failed", { exact: true })).toBeVisible()
        await expect(card(page, "Needs you").getByText("2 questions waiting", { exact: true })).toBeVisible()
        await expect(card(page, "Spend by matter").locator(".dash-hbars li")).toHaveCount(6)
        await expect(card(page, "Activity").locator(".dash-chips")).toContainText("7 day streak now")
        // Each matter's sessions were read with its own folder.
        // (The dev build's StrictMode loads the dashboard twice.)
        expect([...new Set(state.sessionDirs)].sort()).toEqual(state.matters.map((row) => row.dir).sort())
        await capture(page, `populated-overview-${size}`)
        clean(state)
      } finally { await context.close() }
    })
    if (width === 1180) continue
    await scenario(`Error (${size})`, async () => {
      const state = fixture("error"); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page, false)
        const alert = page.locator(".dash-error")
        await expect(alert).toBeVisible()
        await expect(alert).toContainText("Could not load your matters")
        await expect(alert).toContainText("The document service is temporarily unavailable.")
        await expect(page.locator(".dash-grid")).toHaveCount(0)
        await capture(page, `error-overview-${size}`)
        const before = state.requests.filter((row) => row.path === "/matters").length
        await alert.getByRole("button", { name: "Retry", exact: true }).click()
        await expect.poll(() => state.requests.filter((row) => row.path === "/matters").length).toBeGreaterThan(before)
        await expect(alert).toBeVisible()
        clean(state)
      } finally { await context.close() }
    })
    await scenario(`Empty layout (${size})`, async () => {
      const state = fixture("populated", { widgets: { practice: [] }, period: 30 }); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page, false)
        await expect(page.getByText("Your dashboard is empty", { exact: true })).toBeVisible()
        await expect(page.locator(".dash-grid")).toHaveCount(0)
        await capture(page, `empty-layout-overview-${size}`)
        await page.getByRole("button", { name: "Restore defaults", exact: true }).click()
        await expect(page.locator(".dash-grid > section.dash-card")).toHaveCount(10)
        clean(state)
      } finally { await context.close() }
    })
    await scenario(`Every widget (${size})`, async () => {
      const state = fixture("populated", { widgets: { practice: ALL_WIDGETS }, period: 30 }); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        await expect(page.locator(".dash-grid > section.dash-card")).toHaveCount(ALL_WIDGETS.length)
        await expect(card(page, "Spend over time").locator(".dash-area svg")).toBeVisible()
        await capture(page, `populated-all-widgets-${size}`)
        clean(state)
      } finally { await context.close() }
    })
  }

  // A new lawyer's first days, step by step. The sample matter never counts
  // as their own progress.
  for (const [width, height] of [[1440, 1000], [375, 812]] as const) await scenario(`First matter with documents (${width}x${height})`, async () => {
    const state = fixture("new", undefined, early({ first: 3 })); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      expect(await order(page)).toEqual(STARTING)
      expect(await doneSteps(page)).toEqual(["Open a matter", "Add documents"])
      const current = journey(page).nth(2)
      await expect(current).toHaveAttribute("aria-current", "step")
      await expect(current.getByRole("link", { name: "Ask a question", exact: true })).toHaveAttribute("href", /^\/matter\/m-harbour\?view=chat&prompt=What/)
      await expect(card(page, "Things to try").getByText(`Key points in ${FIRST.title}`, { exact: true })).toBeVisible()
      await expect(card(page, "Things to try").getByRole("link", { name: /Review across documents/ })).toHaveAttribute("href", "/matter/m-harbour?view=review")
      await capture(page, `first-matter-${width}x${height}`)
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Sample explored, no matter of their own (1440x1000)", async () => {
    const state = fixture("new", undefined, early({ sample: true })); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      expect(await order(page)).toEqual(STARTING)
      expect(await doneSteps(page)).toEqual([])
      await expect(journey(page).nth(0)).toHaveAttribute("aria-current", "step")
      const sample = card(page, "Sample matter")
      await expect(sample.getByText("Pick up where you left off", { exact: true })).toBeVisible()
      await expect(sample.getByRole("link", { name: "Open the sample matter", exact: true })).toHaveAttribute("href", /^\/matter\/m-sample\?view=chat&prompt=/)
      await capture(page, "sample-explored-1440x1000")
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Add the sample matter (1440x1000)", async () => {
    const state = fixture("new"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      await card(page, "Sample matter").getByRole("button", { name: "Add the sample matter", exact: true }).click()
      // It opens with a first question seated in the composer.
      await expect(page).toHaveURL(/\/matter\/m-sample\?view=chat&prompt=What%20are%20the%20key%20points/)
      expect(state.samplePosts).toBe(1)
      expect(state.errors, "page errors").toEqual([])
    } finally { await context.close() }
  })

  for (const [width, height] of [[1440, 1000], [375, 812]] as const) await scenario(`First conversation, drafting now (${width}x${height})`, async () => {
    const state = fixture("new", undefined, early({ first: 3, sample: true, conversations: 2, drafting: true })); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      // Work has begun: figures come in, first steps stay until the first draft.
      // The drafting run is a conversation of its own, as in the sidebar.
      const titles = await order(page)
      expect(titles.slice(0, 3)).toEqual(["First steps", "Finish setting up", "Total spend"])
      expect(titles).not.toContain("Sample matter")
      await expect(page.getByRole("group", { name: "Period", exact: true })).toBeVisible()
      expect(await doneSteps(page)).toEqual(["Open a matter", "Add documents", "Ask a question"])
      await expect(journey(page).nth(2).locator(".dash-step-note")).toHaveText("Done: 3 conversations")
      await expect(journey(page).nth(3).getByRole("link", { name: "Drafting now", exact: true })).toHaveAttribute("href", "/matter/m-harbour?session=ses_first_run")
      await capture(page, `first-conversation-${width}x${height}`)
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Arranging first steps never keeps the figures away (1440x1000)", async () => {
    const state = fixture("new", undefined, early({ first: 3 })); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      await menuButton(page, "Things to try").click()
      await menuAction(page, "Things to try", "Hide widget")
      await dashboard(page, true, true)
      expect(await order(page)).toEqual(STARTING.filter((title) => title !== "Things to try"))
      // The first conversation arrives: practice opens on its own default.
      state.sessions.push(chat(FIRST, "ses_first_1", at(0, 9), "What are the key points in these materials?"))
      await dashboard(page, true, true)
      const practice = await order(page)
      expect(practice.slice(0, 3)).toEqual(["First steps", "Finish setting up", "Total spend"])
      // Reset brings back that default after the lawyer arranges practice too.
      await menuButton(page, "Total spend").click()
      await menuAction(page, "Total spend", "Hide widget")
      await expect(card(page, "Total spend")).toHaveCount(0)
      const opener = page.locator(".dash-toolbar").getByRole("button", { name: "Add widgets", exact: true })
      await opener.click()
      const dialog = page.getByRole("dialog", { name: "Add widgets", exact: true })
      await dialog.getByRole("button", { name: "Reset to default", exact: true }).click()
      await expect(dialog.getByRole("button", { name: "Reset to default", exact: true })).toBeDisabled()
      await page.keyboard.press("Escape")
      await expect(dialog).toHaveCount(0)
      expect(await order(page)).toEqual(practice)
      await dashboard(page, true, true)
      expect(await order(page)).toEqual(practice)
      clean(state)
    } finally { await context.close() }
  })

  // The fixture's last seven days (Fri 23 to Thu 29 October) cost
  // 4.00 + 0.45 + 0.30 + 11.90 + 5.80 + 9.30 + 1.75 = 33.50. London's clocks
  // went back on Sunday 25 October; UTC is the control with no clock change.
  for (const timezone of ["UTC", "Europe/London"]) await scenario(`Week total and bars across the clock change (${timezone}, 1440x1000)`, async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000, timezone)
    try {
      await dashboard(page)
      const week = card(page, "This week")
      const bars = await week.locator(".dash-week-track").evaluateAll((nodes) => nodes.map((node) => node.getAttribute("aria-label")))
      screens.push({ name: `week-bars-${timezone}`, hero: await week.locator(".dash-hero").textContent(), bars, total: await card(page, "Total spend").locator(".dash-stat-value").textContent() })
      if (timezone !== "UTC") await capture(page, "populated-week-clock-change-1440x1000", { fullPage: false })
      expect(bars).toEqual(["Fri: $4.00", "Sat: $0.45", "Sun: $0.30", "Mon: $11.90", "Tue: $5.80", "Wed: $9.30", "Today: $1.75"])
      await expect(week.locator(".dash-hero")).toHaveText("$33.50")
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Hide, add back, move and reload (1440x1000)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      const start = await order(page)
      // Hide via the widget's own menu; a toast points to Add widgets.
      for (const title of ["Spend by matter", "Where tokens go"]) {
        await menuButton(page, title).click()
        await menuAction(page, title, "Hide widget")
        await expect(card(page, title)).toHaveCount(0)
        await expect(page.locator(".toast").filter({ hasText: `${title} hidden` })).toContainText("Add widgets")
      }
      await capture(page, "populated-hide-toast-1440x1000")
      expect(await order(page)).toEqual(start.filter((title) => title !== "Spend by matter" && title !== "Where tokens go"))
      // Add widgets lists both; Add puts one back at the end.
      const opener = page.locator(".dash-toolbar").getByRole("button", { name: "Add widgets", exact: true })
      await opener.click()
      const dialog = page.getByRole("dialog", { name: "Add widgets", exact: true })
      await expect(dialog.getByRole("button", { name: "Add Spend by matter", exact: true })).toBeVisible()
      await expect(dialog.getByRole("button", { name: "Add Where tokens go", exact: true })).toBeVisible()
      await capture(page, "populated-add-dialog-1440x1000")
      await dialog.getByRole("button", { name: "Add Spend by matter", exact: true }).click()
      await expect(dialog.getByRole("button", { name: "Add Spend by matter", exact: true })).toHaveCount(0)
      expect((await order(page)).at(-1)).toBe("Spend by matter")
      // Escape closes the dialog and focus returns to its opener.
      await page.keyboard.press("Escape")
      await expect(dialog).toHaveCount(0)
      await expect(opener).toBeFocused()
      // Move earlier, then later, swaps with the neighbour.
      const before = await order(page)
      const third = before[2]!
      await menuButton(page, third).click()
      await capture(page, "populated-menu-open-1440x1000")
      const menu = await card(page, third).locator(".viewer-more-menu").boundingBox()
      expect(menu, "menu box").not.toBeNull()
      expect(menu!.x + menu!.width).toBeLessThanOrEqual(1440)
      await menuAction(page, third, "Move earlier")
      const moved = await order(page)
      expect(moved.slice(0, 3)).toEqual([before[0], before[2], before[1]])
      await menuButton(page, third).click()
      await menuAction(page, third, "Move later")
      expect(await order(page)).toEqual(before)
      await menuButton(page, third).click()
      await menuAction(page, third, "Move earlier")
      // The layout survives a reload.
      const saved = await order(page)
      await page.reload()
      await expect(page.locator(".dash-skeleton")).toHaveCount(0)
      await expect(page.locator(".dash-grid > section.dash-card")).toHaveCount(saved.length)
      expect(await order(page)).toEqual(saved)
      await expect(card(page, "Where tokens go")).toHaveCount(0)
      clean(state)
    } finally { await context.close() }
  })

  // Every widget's open menu must be on top of its own card's content.
  for (const [width, height] of [[1440, 1000], [375, 812]] as const) await scenario(`Open widget menus are on top (${width}x${height})`, async () => {
    const state = fixture("populated", { widgets: { practice: ALL_WIDGETS }, period: 30 }); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      const covered: string[] = []
      for (const title of await order(page)) {
        await menuButton(page, title).click()
        const hide = card(page, title).getByRole("button", { name: "Hide widget", exact: true })
        await expect(hide).toBeVisible()
        const hit = await coveredBy(hide)
        if (hit) {
          covered.push(`${title}: Hide widget covered by ${hit}`)
          if (covered.length === 1) await capture(page, `populated-menu-covered-${width}x${height}`, { fullPage: false })
        }
        await page.keyboard.press("Escape")
        await expect(card(page, title).locator(".dash-menu[open]")).toHaveCount(0)
      }
      screens.push({ name: `menu-stacking-${width}x${height}`, covered })
      expect(covered).toEqual([])
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Hide and Add widgets dialog (375x812)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 375, 812)
    try {
      await dashboard(page)
      for (const title of ["Spend by matter", "Where tokens go"]) {
        await menuButton(page, title).click()
        await menuAction(page, title, "Hide widget")
        await expect(card(page, title)).toHaveCount(0)
      }
      await expect(page.locator(".toast").last()).toContainText("Add widgets")
      await capture(page, "populated-hide-toast-375x812", { fullPage: false })
      const opener = page.locator(".dash-toolbar").getByRole("button", { name: "Add widgets", exact: true })
      await opener.click()
      const dialog = page.getByRole("dialog", { name: "Add widgets", exact: true })
      await expect(dialog.getByRole("button", { name: "Add Where tokens go", exact: true })).toBeVisible()
      await capture(page, "populated-add-dialog-375x812")
      await dialog.getByRole("button", { name: "Add Where tokens go", exact: true }).click()
      expect((await order(page)).at(-1)).toBe("Where tokens go")
      await page.keyboard.press("Escape")
      await expect(dialog).toHaveCount(0)
      await expect(opener).toBeFocused()
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Periods and tooltips (1440x1000)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      const period = page.getByRole("group", { name: "Period", exact: true })
      await period.getByRole("button", { name: "7 days", exact: true }).click()
      await expect(period.getByRole("button", { name: "7 days", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(card(page, "Total spend").locator(".dash-stat-value")).toHaveText("$33.50")
      await expect(card(page, "Total spend").getByText("The last 7 days", { exact: true })).toBeVisible()
      await capture(page, "populated-period-7-days-1440x1000")
      await period.getByRole("button", { name: "All time", exact: true }).click()
      await expect(card(page, "Total spend").getByText("All time", { exact: true })).toBeVisible()
      await expect(card(page, "Documents drafted").locator(".dash-stat-value")).toHaveText("5")
      await capture(page, "populated-period-all-time-1440x1000")
      // The period is saved with the layout.
      await page.reload()
      await expect(page.getByRole("group", { name: "Period", exact: true }).getByRole("button", { name: "All time", exact: true })).toHaveAttribute("aria-pressed", "true")
      await period.getByRole("button", { name: "30 days", exact: true }).click()
      await expect(page.locator(".dash-skeleton")).toHaveCount(0)
      // Week bar tooltip on the highest day.
      const monday = card(page, "This week").locator(".dash-week-track").nth(3)
      await monday.hover()
      const weekTip = card(page, "This week").locator(".dash-tip")
      await expect(weekTip).toHaveText("Mon: $11.90")
      const tipBox = await weekTip.boundingBox(), weekBox = await card(page, "This week").boundingBox()
      screens.push({ name: "week-tip-geometry", tip: tipBox, card: weekBox })
      await capture(page, "populated-tip-week-1440x1000")
      // Calendar cell tooltip on one of the busiest days.
      await page.mouse.move(5, 5)
      const cell = card(page, "Activity").locator("rect.dash-cell.level-4").last()
      await cell.scrollIntoViewIfNeeded()
      await cell.hover()
      const cellTip = card(page, "Activity").locator(".dash-tip")
      await expect(cellTip).toContainText(/conversations? on /)
      screens.push({ name: "calendar-tip-geometry", tip: await cellTip.boundingBox(), card: await card(page, "Activity").boundingBox() })
      await capture(page, "populated-tip-calendar-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "dashboard-results.json"), JSON.stringify({ now: new Date(NOW).toISOString(), results, screens, menuHits, fixtures: allStates.map(({ kind, errors, unexpected, requests }) => ({ kind, errors, unexpected, requests: requests.length, paths: [...new Set(requests.map((row) => `${row.method} ${row.path}`))] })) }, null, 2) + "\n")
  await browser.close()
}
expect(results.filter(result => result.status === "failed")).toEqual([])
console.log(`PASS ${results.length} scenarios; ${screens.length} screenshots.`)
