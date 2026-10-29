// Dashboard visual and interaction regression against an already-running web
// preview. Every API request is intercepted with synthetic data relative to a
// pinned clock; unknown requests are refused before any network I/O.
import { mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
type Page = import("@playwright/test").Page
type Locator = import("@playwright/test").Locator
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
type Tokens = { input: number; output: number; reasoning: number; cache: { read: number; write: number } }
type Session = { id: string; parentID?: string; projectID: string; title: string; cost: number; tokens: Tokens; time: { created: number; updated: number }; location: { directory: string }; metadata: Record<string, unknown> }
type Message = Record<string, unknown>

const MATTERS: Omit<Matter, "dir" | "created_at">[] = [
  { id: "m-alder", title: "Alder House · Studio lease renewal", reference: "AH-2026-014", client: "Alder House Studio", jurisdictions: ["EW"] },
  { id: "m-northbank", title: "Northbank Works · Supply agreement", reference: "NB-2026-008", client: "Northbank Works", jurisdictions: ["EW"] },
  { id: "m-fernlea", title: "Fernlea Bakery · Shareholders' agreement", reference: "FB-2026-003", client: "Fernlea Bakery Co.", jurisdictions: ["EW"] },
  { id: "m-quarry", title: "Quarry Lane Cycles · Franchise terms", reference: "QL-2026-021", client: "Quarry Lane Cycles", jurisdictions: ["SCT"] },
  { id: "m-marlowe", title: "Marlowe & Pike Design · Master services agreement and statement of work", reference: "MP-2026-011", client: "Marlowe & Pike Design", jurisdictions: ["EW"] },
  { id: "m-tidewater", title: "Tidewater Robotics · Mutual NDA set", reference: "TR-2026-030", client: "Tidewater Robotics", jurisdictions: ["EW"] },
  { id: "m-copperleaf", title: "Copperleaf Gardens · Employment contracts", jurisdictions: ["EW"] },
  { id: "m-sample", title: "Kestrel Books · Distribution agreement", client: "Kestrel Books", jurisdictions: ["EW"], sample: true },
]
const titleOf = (id: string) => MATTERS.find((row) => row.id === id)!.title

function tokensFor(cost: number, random: () => number): Tokens {
  const scale = 0.7 + random() * 0.6
  return {
    input: Math.round(cost * 160_000 * scale),
    output: Math.round(cost * 22_000 * scale),
    reasoning: Math.round(cost * 14_000 * scale),
    cache: { read: Math.round(cost * 520_000 * scale), write: Math.round(cost * 36_000 * scale) },
  }
}

const assistant = (id: string, created: number, cost: number): Message =>
  ({ id, type: "assistant", agent: "build", model: { id: "primary", providerID: "fixture" }, content: [], tokens: tokensFor(cost, rng(id.length)), cost, finish: "stop", time: { created, completed: created + 60_000 } })

// A populated practice: about a year of conversations, quiet at first, a
// holiday lull in August and a busy autumn; weekdays busier than weekends and
// Tuesday the busiest. The last fortnight is spelled out so the week card has
// clear high, middle and low days.
function practice(options: { unpriced?: boolean } = {}) {
  const random = rng(20261029)
  const sessions: (Session & { matter: string })[] = []
  const messages: Record<string, Message[]> = {}
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
    // Thursday 22 October has only the start of the overnight conversation below.
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
  // A conversation that ran overnight, from Thursday afternoon (the week
  // before) into Friday morning (this week): each reply's cost belongs on its
  // own day.
  const overnight = "ses_northbank_overnight"
  sessions.push({ matter: "m-northbank", id: overnight, projectID: "fixture-project", title: "Mark up the supplier's comments", cost: 1.2, tokens: tokensFor(1.2, random), time: { created: at(7, 15), updated: at(6, 9) }, location: { directory: "" }, metadata: {} })
  messages[overnight] = [
    { id: "msg_overnight_1", type: "user", text: "Please mark up the supplier's comments on the liability clause.", time: { created: at(7, 15) } },
    assistant("msg_overnight_2", at(7, 15, 1), 0.5),
    { id: "msg_overnight_3", type: "user", text: "Now the remaining comments, please.", time: { created: at(6, 8, 45) } },
    assistant("msg_overnight_4", at(6, 8, 50), 0.7),
  ]
  // A model with no price: it used tokens but reports no cost.
  if (options.unpriced) sessions.push({ matter: "m-copperleaf", id: "ses_unpriced", projectID: "fixture-project", title: "Compare the notice periods", cost: 0, tokens: tokensFor(0.6, random), time: { created: at(0, 9, 10), updated: at(0, 9, 30) }, location: { directory: "" }, metadata: {} })
  const firstSession = new Map<string, number>()
  for (const session of sessions) firstSession.set(session.matter, Math.min(firstSession.get(session.matter) ?? Infinity, session.time.created))
  const matters: Matter[] = MATTERS.map((row) => ({ ...row, dir: `/fixture/matters/${row.id.slice(2)}`, created_at: (firstSession.get(row.id) ?? at(30, 9)) - 2 * 3_600_000 }))
  for (const session of sessions) session.location.directory = matters.find((row) => row.id === session.matter)!.dir
  return { matters, sessions, messages, owners, stageIds }
}

const stage = (status: string, sessions: string[], started: number, finished: number | null, attempts = 1) => ({ status, attempts, session_ids: sessions, started_at: iso(started), finished_at: finished === null ? null : iso(finished) })

type Item = { id: string; stage: string; kind: string; detail: string; status: "open" | "resolved" | "disclosed" }
// A lean run's status in the shape isLeanDraftingStatus accepts.
function run(options: { status: "running" | "blocked" | "failed" | "completed"; owner: string; stages: string[]; started: number; updated: number; documents?: { title: string; kind?: "body" | "combined" }[]; questions?: string[]; stop?: string; inFlight?: boolean; delivery?: string; findings?: Item[]; changed?: string[] }) {
  const [drafting = [], review = [], finalize = []] = [options.stages.slice(0, 1), options.stages.slice(1, -1), options.stages.slice(-1)]
  const done = options.status === "completed"
  const documents = options.documents ?? []
  const findings = options.findings ?? []
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
    in_flight: options.inFlight ?? options.status === "running", stopped: false, pending_requests: 0, open_items: [],
    questions: (options.questions ?? []).map((question, index) => ({ id: `q${index + 1}`, stage: "review-legal", question, context: "The instructions do not settle this point.", status: "open", kind: "instruction" })),
    gate: done ? { status: "pass", delivery: options.delivery ?? (findings.length ? "review_draft_with_findings" : "review_draft"), blocking: [], disclosed: findings, items: findings } : null,
    documents: done ? { status: "completed", documents: [
      // A combined deliverable keeps its clean body row beside the one combined file.
      ...documents.map((row, index) => ({ kind: "body", document: `d${index + 1}`, title: row.title, path: `.dochaus/lean/out/d${index + 1}.docx`, sha256: "0".repeat(64), deliverable: `${row.title}.docx` })),
      ...documents.flatMap((row, index) => row.kind === "combined" ? [{ kind: "combined", document: `d${index + 1}`, title: row.title, path: `${row.title}.docx`, sha256: "3".repeat(64), deliverable: `${row.title}.docx` }] : []),
      ...documents.filter((row) => (row.kind ?? "body") === "body").map((row, index) => ({ kind: "redline", document: `d${index + 1}`, title: row.title, path: `.dochaus/lean/out/d${index + 1}-redline.docx`, sha256: "1".repeat(64), deliverable: `${row.title} - redline.docx` })),
      { kind: "drafting_notes", document: null, title: null, path: ".dochaus/lean/out/notes.docx", sha256: "2".repeat(64), deliverable: "Drafting notes.docx" },
    ] } : null,
    // A delivered draft the lawyer edited in the matter after the final check.
    changed_after_check: options.changed ? { since: iso(options.updated), documents: options.changed, missing: [] } : null,
    plan: { documents: documents.map((row, index) => ({ id: `d${index + 1}`, title: row.title, purpose: "Synthetic fixture document." })) },
    memo: null, combined_output: null, request: "Synthetic fixture request.",
  }
}

const finding = (id: string, detail: string): Item => ({ id, stage: "gate", kind: "checklist_unresolved", detail, status: "disclosed" })

function populated(options: { unpriced?: boolean } = {}) {
  const { matters, sessions, messages, owners, stageIds } = practice(options)
  const owned = (key: string) => ({ owner: owners[key]!, stages: stageIds[key]! })
  const created = (key: string) => sessions.find((row) => row.id === owners[key])!.time.created
  const drafting: Record<string, ReturnType<typeof run>> = {
    "m-alder": run({ status: "completed", ...owned("alder"), started: created("alder"), updated: created("alder") + 3_000_000, documents: [{ title: "Lease renewal" }, { title: "Side letter" }], changed: ["Lease renewal.docx"] }),
    "m-northbank": run({ status: "completed", ...owned("northbank"), started: created("northbank"), updated: created("northbank") + 2_400_000, documents: [{ title: "Supply agreement" }], delivery: "execution_ready" }),
    "m-fernlea": run({ status: "completed", ...owned("fernlea"), started: created("fernlea"), updated: created("fernlea") + 2_700_000, documents: [{ title: "Shareholders' agreement", kind: "combined" }], findings: [finding("f1", "The drag-along notice period is not stated."), finding("f2", "The deadlock procedure has no final step.")] }),
    "m-quarry": run({ status: "completed", ...owned("quarry"), started: created("quarry"), updated: created("quarry") + 2_100_000, documents: [{ title: "Franchise terms" }] }),
    "m-marlowe": run({ status: "blocked", ...owned("marlowe"), started: created("marlowe"), updated: created("marlowe") + 1_300_000, questions: ["Should the statement of work cap change requests?", "Which party owns the design files on termination?"] }),
    "m-tidewater": run({ status: "failed", ...owned("tidewater"), started: created("tidewater"), updated: created("tidewater") + 1_500_000, stop: "The review stage did not produce a valid output after three tries." }),
    // Still marked running, but nothing is working on it: it waits for Continue.
    "m-copperleaf": run({ status: "running", ...owned("copperleaf"), started: created("copperleaf"), updated: created("copperleaf") + 600_000, inFlight: false }),
  }
  const documents: Record<string, number> = { "m-alder": 9, "m-northbank": 6, "m-fernlea": 4, "m-quarry": 5, "m-marlowe": 7, "m-tidewater": 3, "m-copperleaf": 8, "m-sample": 3 }
  return { matters, sessions, messages, drafting: drafting as Record<string, unknown>, documents }
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
  return { matters, sessions, messages: {} as Record<string, Message[]>, drafting, documents: { [FIRST.id]: options.first ?? 0, [SAMPLE.id]: 3 } as Record<string, number> }
}

type Kind = "new" | "populated" | "error"
type Data = ReturnType<typeof populated>
type Preferences = typeof DEFAULT_DRAFTING
const PROFILE: Partial<Preferences> = { attorney: "Alex Example", firm: "Example Partners LLP" }
function fixture(kind: Kind, options: { layout?: unknown; data?: Data; preferences?: Partial<Preferences>; jurisdictions?: string[] } = {}) {
  const data = options.data ?? (kind === "populated" ? populated() : { matters: [], sessions: [], messages: {}, drafting: {}, documents: {} })
  return {
    kind, layout: options.layout, ...data, samplePosts: 0, failPreferences: false,
    aiPlan: { name: "Practice costs", widgets: [
      { type: "chart", title: "Spend by client", metric: "spend", groupBy: "client", form: "bars", period: "dashboard" },
      { type: "chart", title: "Conversations over time", metric: "conversations", groupBy: "week", form: "line", period: "quarter" },
    ], unsupported: ["Time per fee earner is not tracked."] } as unknown,
    aiFailure: false, aiDelay: 0, aiSession: "", aiCreated: [] as Record<string, unknown>[], aiPrompts: [] as Record<string, unknown>[], aiDeleted: [] as string[],
    // Folders whose session list the engine fails to return.
    failSessions: new Set<string>(),
    jurisdictions: options.jurisdictions ?? (kind === "populated" ? ["EW"] : undefined),
    preferences: kind === "populated"
      ? { ...DEFAULT_DRAFTING, ...PROFILE, practiceAreas: ["Commercial contracts", "Property"], houseStyle: "Arial 11pt, numbered clauses.", searchApiKey: "fixture-exa-key" }
      : { ...DEFAULT_DRAFTING, ...options.preferences },
    saved: [] as Record<string, unknown>[],
    unexpected: [] as string[], errors: [] as string[], sessionDirs: [] as string[], requests: [] as { method: string; path: string }[],
  }
}
type State = ReturnType<typeof fixture>

// A slice of the shipped packs, so a browser in Britain is offered its three.
const jurisdictions = [["AU", "Australia (Federal)"], ["CA-ON", "Ontario"], ["DE", "Germany"], ["EU", "European Union"], ["EW", "England & Wales"], ["IE", "Ireland"], ["NIR", "Northern Ireland"], ["SCT", "Scotland"], ["SG", "Singapore"], ["US", "United States (Federal)"], ["US-DE", "Delaware"], ["US-NY", "New York"]]
  .map(([code, name]) => ({ code, name, citationStyle: "fixture" }))
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
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId, locale: "en-GB", acceptDownloads: true })
  await context.clock.setFixedTime(new Date(NOW))
  await context.addInitScript(({ layout, jurisdiction }) => {
    localStorage.setItem("dh.launch-notice.v1", new Date().toDateString())
    localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1")
    // Seed the saved layout and jurisdictions once; a reload keeps what the page saved.
    if (sessionStorage.getItem("dashboard-fixture-seeded")) return
    if (jurisdiction) localStorage.setItem("dochaus.prefs", JSON.stringify({ defaultJurisdictions: jurisdiction }))
    if (layout) localStorage.setItem("dh.dashboard.v3", JSON.stringify(layout))
    sessionStorage.setItem("dashboard-fixture-seeded", "1")
  }, { layout: state.layout ?? null, jurisdiction: state.jurisdictions ?? null })
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
    // Ingest merges the top-level keys it is sent.
    if (method === "PUT" && pathname === "/preferences") {
      if (state.failPreferences) return reply({ error: "Could not write settings." }, 500)
      const patch = request.postDataJSON() as Record<string, unknown>
      state.saved.push(patch)
      state.preferences = { ...state.preferences, ...patch }
      return reply(state.preferences)
    }
    // Settings opens a provider candidate session on open, on every tab, and
    // may list its models before the scenario moves on.
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "POST") return reply({ sessionID: "synthetic-candidate" })
    if (pathname === "/api/dochaus/provider-candidate/models" && method === "POST") return reply({ data: [], customModels: false })
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "DELETE") return route.fulfill({ status: 204, headers })
    // The AI boundary is native OpenCode session/tool output; values still
    // come from the synthetic matters. Every other network call is refused.
    if (pathname === "/dashboard/directory") return reply({ dir: "/fixture/.dashboard" })
    if (pathname === "/api/session" && method === "POST") {
      const input = request.postDataJSON() as Record<string, unknown>
      state.aiCreated.push(input)
      state.aiSession = String(input.id)
      return reply({ data: { id: state.aiSession, projectID: "fixture-project", title: "doc.haus dashboard", agent: "dashboard", model: { providerID: "fixture", id: "fast" }, location: { directory: "/fixture/.dashboard" }, metadata: { dochaus_role: "dashboard" }, cost: 0, tokens: tokensFor(0, rng(1)), time: { created: NOW, updated: NOW } } })
    }
    if (pathname === `/api/session/${state.aiSession}/prompt` && method === "POST") {
      state.aiPrompts.push(request.postDataJSON() as Record<string, unknown>)
      if (state.aiDelay) await new Promise(resolve => setTimeout(resolve, state.aiDelay))
      return reply({ data: { id: "inbox_dashboard", type: "user", text: "Please update the dashboard.", time: { created: NOW } } })
    }
    if (pathname === `/api/experimental/session/${state.aiSession}/wait` && method === "POST") return route.fulfill({ status: 204, headers })
    if (pathname === `/api/session/${state.aiSession}/interrupt` && method === "POST") return reply({ interrupted: false })
    if (pathname === `/api/session/${state.aiSession}` && method === "DELETE") { state.aiDeleted.push(state.aiSession); return route.fulfill({ status: 204, headers }) }
    if (pathname === `/api/session/${state.aiSession}/message`) return reply({ data: [
      { id: "msg_dashboard_plan", type: "assistant", agent: "dashboard", model: { providerID: "fixture", id: "fast" }, time: { created: NOW, completed: NOW + 1 }, cost: 0, tokens: tokensFor(0, rng(1)), content: state.aiFailure ? [] : [{ type: "tool", id: "call_dashboard_plan", name: "dashboard-plan", time: { created: NOW, completed: NOW + 1 }, state: { status: "completed", input: {}, content: [{ type: "text", text: "Dashboard plan accepted." }], metadata: { plan: state.aiPlan } } }] },
      { id: "msg_dashboard_idle", type: "idle", outcome: state.aiFailure ? "failed" : "succeeded", time: { created: NOW + 2 } },
    ], cursor: { previous: null, next: null } })
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    const populatedState = state.kind === "populated"
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    // A new install finished onboarding with only the model step: details and
    // research were skipped, so the app opens normally and nothing else is set.
    if (pathname === "/setup") return reply(populatedState ? { version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" } : { version: 1, ready: true, completed: ["models"], skipped: ["details", "research"], next: "done" })
    if (pathname === "/preferences") return reply(state.preferences)
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
      if (state.failSessions.has(directory)) return reply({ error: "The engine could not read this folder." }, 503)
      const data = state.sessions.filter((row) => row.location.directory === directory).map(({ matter: _, ...row }) => row).sort((a, b) => b.time.updated - a.time.updated)
      return reply({ data, cursor: { previous: null, next: null } })
    }
    // A conversation's messages, read for one that ran over several days.
    const messageMatch = pathname.match(/^\/api\/session\/([^/]+)\/message$/)
    if (messageMatch && state.sessions.some((row) => row.id === messageMatch[1])) return reply({ data: state.messages[messageMatch[1]!] ?? [], cursor: { previous: null, next: null } })
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
    // What sticks out past the page's right edge, when something does.
    wide: document.documentElement.scrollWidth > document.documentElement.clientWidth + 1 ? Array.from(document.querySelectorAll<HTMLElement>("body *")).filter(node => node.getBoundingClientRect().right > document.documentElement.clientWidth + 1 && ![...node.children].some(child => child.getBoundingClientRect().right > document.documentElement.clientWidth + 1)).slice(0, 8).map(node => ({ node: `${node.tagName.toLowerCase()}.${[...node.classList].join(".")}`, right: Math.round(node.getBoundingClientRect().right), text: (node.textContent ?? "").trim().slice(0, 60) })) : [],
    // Text cut off inside a card: an element whose content is wider than its box.
    clipped: Array.from(document.querySelectorAll<HTMLElement>(".dash-card *")).filter(node => node.children.length === 0 && !node.closest(".sr-only") && node.textContent?.trim() && node.scrollWidth > node.clientWidth + 1 && getComputedStyle(node).overflow !== "visible").map(node => ({ text: node.textContent!.trim().slice(0, 80), card: node.closest(".dash-card")?.querySelector("h2")?.textContent })),
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
const card = (page: Page, title: string) => page.locator("section.dash-card").filter({ has: page.getByRole("heading", { name: title, exact: true, level: 2 }) })
const askButton = (page: Page) => page.getByRole("button", { name: "Ask AI", exact: true })
const menuHits: unknown[] = []
async function hideWidget(page: Page, title: string) {
  await card(page, title).hover()
  await page.getByRole("button", { name: `Remove ${title}`, exact: true }).click()
  await expect(card(page, title)).toHaveCount(0)
}
const order = (page: Page) => page.locator(".dash-grid > section.dash-card h2").allTextContents()
const greeting = (page: Page, text: string) => expect(page.locator(".page-bar").getByText(`${text}. What needs you, your drafts and what your matters cost.`, { exact: true })).toBeVisible()
const period = (page: Page) => page.getByRole("group", { name: "Period", exact: true })
// Settled: the page bar shown, no skeletons left and the refresh control idle.
async function dashboard(page: Page, expectGrid = true, reload = false) {
  if (reload) await page.reload()
  else await page.goto(`${baseURL}/dashboard`)
  await expect(page.getByRole("heading", { name: "Dashboard", exact: true, level: 1 })).toBeVisible()
  await expect(page.getByRole("button", { name: "Refresh", exact: true })).toBeEnabled()
  if (expectGrid) await expect(page.locator(".dash-skeleton")).toHaveCount(0)
}

const SIZES = [[320, 640], [375, 812], [768, 1024], [1440, 1000]] as const
const ALL_WIDGETS = ["journey", "setup", "attention", "drafts", "spend", "matter-cost", "draft-cost", "week", "recent", "matter-spend", "activity", "trend", "tokens", "sample", "ideas", "actions", "library"]
const PRACTICE_DEFAULTS = ["attention", "drafts", "setup", "spend", "matter-cost", "draft-cost", "week", "recent", "matter-spend", "activity"]
const STARTING = ["First steps", "Your practice"]
// The practice default once every essential detail is in: Your practice steps aside.
const PRACTICE = ["Needs you", "Recent drafts", "Total spend", "Average matter cost", "Average draft cost", "This week", "Recent matters", "Spend by matter", "Activity"]
// Every widget a practice can show: all but First steps and Your practice.
const EVERY = [...PRACTICE, "Spend over time", "Where tokens go", "Sample matter", "Things to try", "Quick actions", "Your library"]
const journey = (page: Page) => card(page, "First steps").locator(".dash-journey > li")
const doneSteps = (page: Page) => card(page, "First steps").locator(".dash-journey > li.is-done strong").allTextContents()
// A practice detail: an open tile named by its title, or a finished row named
// "<title>, done" (the visually hidden ", done" is computed with a space before it).
const detail = (setup: Locator, title: string, done = false) => setup.getByRole("region", { name: done ? new RegExp(`^${title} ?, done$`) : title, exact: !done })
const ring = (setup: Locator, done: number) => expect(setup.getByRole("img", { name: `${done} of 3 done`, exact: true })).toBeVisible()

// Expected figures straight from the fixture: own spend by calendar month
// (every fixture time falls on the same day in UTC and London), and each own
// matter's October spend.
const money = (value: number) => `$${value.toFixed(2)}`
const own = (state: State) => state.sessions.filter((row) => row.matter !== "m-sample")
const monthSpend = (state: State, month: number) => own(state).filter((row) => new Date(row.time.created).getUTCMonth() === month).reduce((sum, row) => sum + row.cost, 0)
function octoberByMatter(state: State) {
  const totals = new Map<string, number>()
  for (const row of own(state)) if (new Date(row.time.created).getUTCMonth() === 9 && row.cost > 0) totals.set(row.matter, (totals.get(row.matter) ?? 0) + row.cost)
  return [...totals].map(([id, cost]) => ({ id, title: titleOf(id), cost })).sort((a, b) => b.cost - a.cost)
}
// The overnight conversation splits $0.50 on Thursday 22 and $0.70 on Friday 23 October.
const WEEK = [
  "Fri: $4.70. Marlowe & Pike Design · Master services agreement and statement of work $2.40; Tidewater Robotics · Mutual NDA set $1.60; 1 other matter",
  "Sat: $0.45. Tidewater Robotics · Mutual NDA set $0.45",
  "Sun: $0.30. Copperleaf Gardens · Employment contracts $0.30",
  "Mon: $11.90. Alder House · Studio lease renewal $5.80; Marlowe & Pike Design · Master services agreement and statement of work $3.90; 1 other matter",
  "Tue: $5.80. Marlowe & Pike Design · Master services agreement and statement of work $3.10; Fernlea Bakery · Shareholders' agreement $2.70",
  "Wed: $9.30. Fernlea Bakery · Shareholders' agreement $6.40; Alder House · Studio lease renewal $2.90",
  "Today: $1.75. Copperleaf Gardens · Employment contracts $1.75",
]

try {
  for (const [width, height] of [...SIZES, [1180, 900] as const]) {
    const size = `${width}x${height}`
    if (width !== 1180) await scenario(`New install (${size})`, async () => {
      const state = fixture("new"); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        await greeting(page, "Good morning")
        // No figures yet: the dashboard leads with first steps and the profile, not empty charts.
        expect(await order(page)).toEqual(STARTING)
        await expect(period(page)).toHaveCount(0)
        await expect(journey(page).nth(0)).toHaveAttribute("aria-current", "step")
        await expect(journey(page).nth(0).getByRole("link", { name: "New matter", exact: true })).toHaveAttribute("href", "/?new=1")
        await expect(journey(page).nth(0).getByRole("button", { name: "or explore a sample matter first", exact: true })).toBeVisible()
        await expect(card(page, "First steps").locator(".dash-step-action")).toHaveCount(1)
        await expect(card(page, "First steps").getByText(/^Each draft arrives as a \.docx file with a memo/)).toBeVisible()
        const setup = card(page, "Your practice")
        await expect(setup.getByText("3 details to go", { exact: true })).toBeVisible()
        await ring(setup, 0)
        for (const title of ["Name and firm", "Jurisdictions", "Practice areas"]) await expect(detail(setup, title)).toBeVisible()
        await expect(setup.locator(".dash-profile-row")).toHaveCount(0)
        await expect(setup.locator(".dash-profile-extras li")).toHaveText(["Web research", "Word add-in", "House style notes", "A template"])
        // Settings items open a tab; the template goes to the library page.
        await expect(setup.getByRole("link", { name: "A template", exact: true })).toHaveAttribute("href", "/templates")
        await capture(page, `starting-new-user-${size}`)
        expect(state.requests.filter((row) => row.path === "/api/session")).toEqual([])
        clean(state)
      } finally { await context.close() }
    })
    if (width !== 1180) await scenario(`Starting with a partly filled profile (${size})`, async () => {
      const state = fixture("new", { data: early({ first: 3 }), preferences: PROFILE, jurisdictions: ["EW"] }); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        await greeting(page, "Good morning, Alex Example")
        expect(await order(page)).toEqual(STARTING)
        const setup = card(page, "Your practice")
        await expect(setup.getByText("1 detail to go", { exact: true })).toBeVisible()
        await ring(setup, 2)
        // Finished details fold into a line with a way to change them; the open one stays a tile.
        const signed = detail(setup, "Name and firm", true)
        await expect(signed).toContainText("Alex Example, Example Partners LLP")
        await expect(signed.getByRole("button", { name: "Edit name and firm", exact: true })).toBeVisible()
        const places = detail(setup, "Jurisdictions", true)
        await expect(places).toContainText("England & Wales")
        await expect(places.getByRole("button", { name: "Change jurisdictions", exact: true })).toBeVisible()
        const areas = detail(setup, "Practice areas")
        await expect(areas.locator(".dash-profile-areas button")).toHaveText(["Commercial", "Corporate", "Employment", "Litigation", "Real estate", "Private client"])
        await capture(page, `starting-partial-profile-${size}`)
        clean(state)
      } finally { await context.close() }
    })
    await scenario(`Populated practice (${size})`, async () => {
      const state = fixture("populated"); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        await greeting(page, "Good morning, Alex Example")
        // Essentials done and drafts finished: the work and the figures, no onboarding.
        expect(await order(page)).toEqual(PRACTICE)
        await expect(period(page).getByRole("button", { name: "This month", exact: true })).toHaveAttribute("aria-pressed", "true")
        const needs = card(page, "Needs you").locator(".dash-attention > li")
        await expect(needs.locator("strong")).toHaveText(["Drafting run failed", "2 questions waiting", "Drafting stalled, press Continue"])
        for (const [index, id] of ["m-tidewater", "m-marlowe", "m-copperleaf"].entries()) {
          const owner = (state.drafting[id] as { owner_session_id: string }).owner_session_id
          await expect(needs.nth(index).getByRole("link", { name: `Open ${titleOf(id)}`, exact: true })).toHaveAttribute("href", `/matter/${id}?session=${encodeURIComponent(owner)}`)
        }
        // Newest first, each with what it needs next.
        const drafts = card(page, "Recent drafts").locator(".dash-rows > li")
        await expect(drafts.locator("strong")).toHaveText(["Shareholders' agreement", "Lease renewal +1 more", "Supply agreement", "Franchise terms"])
        await expect(drafts.locator(".dash-pill")).toHaveText(["2 findings to resolve", "Changed since final check", "Ready for signature", "Review draft"])
        await expect(drafts.nth(1)).toContainText(titleOf("m-alder"))
        const week = card(page, "This week")
        await expect(week.locator(".dash-hero")).toHaveText("$34.20")
        await expect(week.getByText("Week before $11.80", { exact: true })).toBeVisible()
        const total = card(page, "Total spend")
        await expect(total.locator(".dash-stat-value")).toHaveText(money(monthSpend(state, 9)))
        await expect(total.getByText(`Last month ${money(monthSpend(state, 8))}`, { exact: true })).toBeVisible()
        const spend = octoberByMatter(state)
        const bars = card(page, "Spend by matter").locator(".dash-hbars li")
        await expect(bars.locator(".dash-hbar-name")).toHaveText(spend.slice(0, 6).map((row) => row.title))
        await expect(bars.locator(".dash-hbar-value")).toHaveText(spend.slice(0, 6).map((row) => money(row.cost)))
        const rest = spend.slice(6)
        await expect(card(page, "Spend by matter").locator(".dash-footnote")).toHaveText(`${rest.length} other matter${rest.length === 1 ? "" : "s"} · ${money(rest.reduce((sum, row) => sum + row.cost, 0))}`)
        const conversations = own(state).filter((row) => !row.parentID).length
        await expect(card(page, "Activity").getByText(`${conversations} conversations in the last year`, { exact: true })).toBeVisible()
        // Each matter's sessions were read with its own folder.
        // (The dev build's StrictMode loads the dashboard twice.)
        expect([...new Set(state.sessionDirs)].sort()).toEqual(state.matters.map((row) => row.dir).sort())
        await expect(page.locator(".dash-banner")).toHaveCount(0)
        await capture(page, `practice-dashboard-${size}`)
        if (width === 1440) await capture(page, "practice-needs-you-1440x1000", { fullPage: false })
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
        await expect(alert).toContainText("Your figures will show once the matter list loads.")
        // The service's own wording stays out of the lawyer's way.
        await expect(alert).not.toContainText("temporarily unavailable")
        await expect(page.locator(".dash-grid")).toHaveCount(0)
        await expect(period(page)).toHaveCount(0)
        await capture(page, `error-matters-unavailable-${size}`)
        const before = state.requests.filter((row) => row.path === "/matters").length
        await alert.getByRole("button", { name: "Retry", exact: true }).click()
        await expect.poll(() => state.requests.filter((row) => row.path === "/matters").length).toBeGreaterThan(before)
        await expect(alert).toBeVisible()
        clean(state)
      } finally { await context.close() }
    })
    await scenario(`Empty layout (${size})`, async () => {
      const state = fixture("populated", { layout: { stages: { practice: { hidden: PRACTICE_DEFAULTS, added: [] } }, period: "month" } }); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page, false)
        await expect(page.getByText("A dashboard for the way you work", { exact: true })).toBeVisible()
        await expect(page.locator(".dash-grid")).toHaveCount(0)
        await capture(page, `practice-empty-layout-${size}`)
        await page.getByRole("button", { name: "Ask AI to build it", exact: true }).click()
        await expect(page.getByRole("textbox", { name: "Tell AI what you want from this dashboard", exact: true })).toBeFocused()
        clean(state)
      } finally { await context.close() }
    })
    await scenario(`Every widget (${size})`, async () => {
      const state = fixture("populated", { layout: { stages: { practice: { hidden: [], added: ALL_WIDGETS } }, period: "month" } }); const { page, context } = await open(state, width, height)
      try {
        await dashboard(page)
        expect(await order(page)).toEqual(EVERY)
        await expect(card(page, "Spend over time").locator(".dash-area svg")).toBeVisible()
        await capture(page, `practice-every-widget-${size}`)
        clean(state)
      } finally { await context.close() }
    })
  }

  // A new lawyer's first days, step by step. The sample matter never counts
  // as their own progress.
  for (const [width, height] of [[1440, 1000], [375, 812]] as const) await scenario(`First matter with documents (${width}x${height})`, async () => {
    const state = fixture("new", { data: early({ first: 3 }) }); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      expect(await order(page)).toEqual(STARTING)
      expect(await doneSteps(page)).toEqual(["Open a matter", "Add documents"])
      await expect(journey(page).nth(1).getByText("Done: Documents added", { exact: true })).toBeVisible()
      const current = journey(page).nth(2)
      await expect(current).toHaveAttribute("aria-current", "step")
      await expect(current.getByRole("link", { name: "Ask a question", exact: true })).toHaveAttribute("href", /^\/matter\/m-harbour\?view=chat&prompt=What/)
      await expect(card(page, "First steps").locator(".dash-step-action")).toHaveCount(1)
      await capture(page, `starting-first-matter-${width}x${height}`)
      clean(state)
    } finally { await context.close() }
  })

  // Documents went into an older matter; the three newest are still empty.
  await scenario("Documents only in an older matter (1440x1000)", async () => {
    const newer = [["m-larkspur", "Larkspur Dental · Practice lease", 3], ["m-oakmoor", "Oakmoor Joinery · Terms of business", 2], ["m-brightwater", "Brightwater Tutors · Consultancy agreement", 1]] as const
    const data = early({ first: 4 })
    data.matters = [{ ...FIRST, created_at: at(6, 9) }, ...newer.map(([id, title, daysAgo]) => ({ id, title, jurisdictions: ["EW"], dir: `/fixture/matters/${id.slice(2)}`, created_at: at(daysAgo, 9) }))]
    const state = fixture("new", { data }); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      expect(await doneSteps(page)).toEqual(["Open a matter", "Add documents"])
      await expect(journey(page).nth(2)).toHaveAttribute("aria-current", "step")
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Sample explored, no matter of their own (1440x1000)", async () => {
    const state = fixture("new", { data: early({ sample: true }) }); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      expect(await order(page)).toEqual(STARTING)
      expect(await doneSteps(page)).toEqual([])
      const first = journey(page).nth(0)
      await expect(first).toHaveAttribute("aria-current", "step")
      await expect(first.getByRole("link", { name: "New matter", exact: true })).toHaveAttribute("href", "/?new=1")
      await expect(first.getByRole("link", { name: "or open the sample matter", exact: true })).toHaveAttribute("href", /^\/matter\/m-sample\?view=chat&prompt=What/)
      await capture(page, "starting-sample-explored-1440x1000")
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Add the sample matter from first steps (1440x1000)", async () => {
    const state = fixture("new"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      await journey(page).nth(0).getByRole("button", { name: "or explore a sample matter first", exact: true }).click()
      // It opens with a first question seated in the composer.
      await expect(page).toHaveURL(/\/matter\/m-sample\?view=chat&prompt=What%20are%20the%20key%20points/)
      expect(state.samplePosts).toBe(1)
      expect(state.errors, "page errors").toEqual([])
    } finally { await context.close() }
  })

  // Each detail fills in where it shows, and the card stays put until reload.
  for (const [width, height] of [[1440, 1000], [375, 812]] as const) await scenario(`Fill in the practice profile (${width}x${height})`, async () => {
    const state = fixture("new"); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      const setup = card(page, "Your practice")
      await ring(setup, 0)
      // A browser set to en-GB is offered the three UK jurisdictions first.
      const places = detail(setup, "Jurisdictions")
      expect(await places.locator(".dash-profile-flags button").allTextContents()).toEqual(["England & Wales", "Northern Ireland", "Scotland"])
      await places.getByRole("button", { name: "England & Wales", exact: true }).click()
      await expect(places.getByRole("button", { name: "England & Wales", exact: true })).toHaveAttribute("aria-pressed", "true")
      expect(await page.evaluate(() => JSON.parse(localStorage.getItem("dochaus.prefs") ?? "{}").defaultJurisdictions)).toEqual(["EW"])
      await ring(setup, 1)

      const areas = detail(setup, "Practice areas")
      await areas.getByRole("button", { name: "Commercial", exact: true }).click()
      await expect.poll(() => state.preferences.practiceAreas).toEqual(["Commercial"])
      await ring(setup, 2)
      await expect(setup.getByText("1 detail to go", { exact: true })).toBeVisible()
      if (width === 1440) {
        // A failed save says so and puts the chip back.
        state.failPreferences = true
        await areas.getByRole("button", { name: "Corporate", exact: true }).click()
        await expect(page.getByText("Could not save your practice areas. Try again.", { exact: true })).toBeVisible()
        await expect(areas.getByRole("button", { name: "Corporate", exact: true })).toHaveAttribute("aria-pressed", "false")
        state.failPreferences = false
        // The full picker carries the same flags.
        await places.getByRole("button", { name: /All 12 jurisdictions/ }).click()
        const picker = page.getByRole("group", { name: "Jurisdictions", exact: true })
        await expect(picker.locator(".jx-option .jx-flag svg")).toHaveCount(12)
        await capture(page, "starting-profile-picker-1440x1000", { fullPage: false })
        await page.keyboard.press("Escape")
        await expect(picker).toHaveCount(0)
      }

      // The Exa key lives in Settings; the extra opens the right tab.
      await setup.getByRole("button", { name: "Web research", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Research", exact: true })).toHaveAttribute("aria-selected", "true")
      await page.keyboard.press("Escape")

      // Next time, the finished details fold into lines with a way to change them.
      await dashboard(page, true, true)
      await expect(detail(setup, "Jurisdictions", true)).toContainText("England & Wales")
      await expect(detail(setup, "Jurisdictions", true).getByRole("button", { name: "Change jurisdictions", exact: true })).toBeVisible()
      await expect(detail(setup, "Practice areas", true)).toContainText("Commercial")
      await expect(detail(setup, "Practice areas", true).getByRole("button", { name: "Change practice areas", exact: true })).toBeVisible()
      await ring(setup, 2)

      // The last one, name and firm, folds the moment it is saved.
      const signed = detail(setup, "Name and firm")
      await signed.getByLabel("Your name", { exact: true }).fill("Alex Example")
      await signed.getByLabel("Firm", { exact: true }).fill("Example Partners LLP")
      await signed.getByRole("button", { name: "Save", exact: true }).click()
      await expect(detail(setup, "Name and firm", true)).toContainText("Alex Example, Example Partners LLP")
      expect(state.saved.at(-1)).toEqual({ attorney: "Alex Example", firm: "Example Partners LLP" })
      await greeting(page, "Good morning, Alex Example")
      await ring(setup, 3)
      await expect(setup.getByText("Your practice is set up", { exact: true })).toBeVisible()
      await setup.scrollIntoViewIfNeeded()
      await expect(detail(setup, "Name and firm", true).getByRole("button", { name: "Edit name and firm", exact: true })).toBeFocused()
      await capture(page, `starting-profile-filled-${width}x${height}`, { fullPage: width === 375 })

      // With every essential in, the card steps aside on the next visit.
      await dashboard(page, true, true)
      expect(await order(page)).toEqual(["First steps"])
      clean(state)
    } finally { await context.close() }
  })

  for (const [width, height] of [[1440, 1000], [375, 812]] as const) await scenario(`First conversation, drafting now (${width}x${height})`, async () => {
    const state = fixture("new", { data: early({ first: 3, sample: true, conversations: 2, drafting: true }) }); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      // Still starting out until a draft of their own finishes, so no figures yet.
      expect(await order(page)).toEqual(STARTING)
      await expect(period(page)).toHaveCount(0)
      // The drafting run is a conversation of its own, as in the sidebar.
      expect(await doneSteps(page)).toEqual(["Open a matter", "Add documents", "Ask a question"])
      await expect(journey(page).nth(2).locator(".dash-step-note")).toHaveText("Done: 3 conversations")
      await expect(journey(page).nth(3).getByRole("link", { name: "Drafting now", exact: true })).toHaveAttribute("href", "/matter/m-harbour?session=ses_first_run")
      await capture(page, `starting-drafting-now-${width}x${height}`)
      clean(state)
    } finally { await context.close() }
  })

  // The fixture's last seven days (Fri 23 to Thu 29 October). London's clocks
  // went back on Sunday 25 October; UTC is the control with no clock change.
  for (const timezone of ["UTC", "Europe/London"]) await scenario(`Week total and bars across the clock change (${timezone}, 1440x1000)`, async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000, timezone)
    try {
      await dashboard(page)
      const week = card(page, "This week")
      const bars = await week.locator(".dash-week-track").evaluateAll((nodes) => nodes.map((node) => node.getAttribute("aria-label")))
      screens.push({ name: `week-bars-${timezone}`, hero: await week.locator(".dash-hero").textContent(), bars })
      if (timezone !== "UTC") await capture(page, "practice-week-clock-change-1440x1000", { fullPage: false })
      expect(bars).toEqual(WEEK)
      await expect(week.locator(".dash-hero")).toHaveText("$34.20")
      await expect(week.getByText("Week before $11.80", { exact: true })).toBeVisible()
      // Only the conversation that crossed midnight needed its messages read.
      expect([...new Set(state.requests.filter((row) => /^\/api\/session\/[^/]+\/message$/.test(row.path)).map((row) => row.path))]).toEqual(["/api/session/ses_northbank_overnight/message"])
      clean(state)
    } finally { await context.close() }
  })

  for (const [width, height] of SIZES) await scenario(`AI update keeps pins and supports Undo (${width}x${height})`, async () => {
    const state = fixture("populated"); const { page, context } = await open(state, width, height)
    try {
      await dashboard(page)
      const before = await order(page)
      await card(page, "Needs you").hover()
      await page.getByRole("button", { name: "Pin Needs you", exact: true }).click()
      await expect(page.getByRole("button", { name: "Unpin Needs you", exact: true })).toHaveAttribute("aria-pressed", "true")
      await askButton(page).click()
      const input = page.getByRole("textbox", { name: "Tell AI what you want from this dashboard", exact: true })
      await expect(input).toBeFocused()
      await input.fill("Show spend by client and weekly conversations, and time per fee earner")
      await capture(page, `ai-prompt-open-${width}x${height}`, { fullPage: false })
      await page.getByRole("button", { name: "Update dashboard", exact: true }).click()
      await expect.poll(() => order(page)).toEqual(["Needs you", "Spend by client", "Conversations over time"])
      await expect(page.getByText("Time per fee earner is not tracked.", { exact: true })).toBeVisible()
      await expect(card(page, "Spend by client").locator(".dash-chart-total")).toHaveText(money(monthSpend(state, 9)))
      await expect(card(page, "Conversations over time").locator(".dash-area svg")).toBeVisible()
      expect(state.aiCreated).toHaveLength(1)
      expect(state.aiCreated[0]).toMatchObject({ agent: "dashboard", model: { providerID: "fixture", id: "fast" }, metadata: { dochaus_role: "dashboard" }, location: { directory: "/fixture/.dashboard" } })
      expect(state.aiPrompts).toHaveLength(1)
      expect(String(state.aiPrompts[0]?.text)).toContain('"pinned":true')
      await expect.poll(() => state.aiDeleted.length).toBe(1)
      await capture(page, `ai-updated-pinned-${width}x${height}`)
      await page.locator(".dash-snackbar").getByRole("button", { name: "Undo", exact: true }).click()
      expect(await order(page)).toEqual(before)
      await expect(page.getByRole("button", { name: "Unpin Needs you", exact: true })).toHaveAttribute("aria-pressed", "true")
      await dashboard(page, true, true)
      expect(await order(page)).toEqual(before)
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Named tabs, keyboard navigation, rename, delete and Undo (1440x1000)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      await page.getByRole("button", { name: "New dashboard", exact: true }).click()
      await expect(page.getByRole("tab", { name: "New dashboard", exact: true })).toHaveAttribute("aria-selected", "true")
      await page.getByRole("textbox", { name: "Tell AI what you want from this dashboard", exact: true }).fill("A dashboard about client spend")
      await page.getByRole("button", { name: "Update dashboard", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Practice costs", exact: true })).toHaveAttribute("aria-selected", "true")
      await capture(page, "ai-named-tabs-1440x1000")
      await page.getByRole("tab", { name: "Practice costs", exact: true }).focus()
      await page.keyboard.press("Home")
      await expect(page.getByRole("tab", { name: "Overview", exact: true })).toBeFocused()
      expect(await order(page)).toEqual(PRACTICE)
      await page.keyboard.press("End")
      await expect(page.getByRole("tab", { name: "Practice costs", exact: true })).toBeFocused()
      expect(await order(page)).toEqual(["Spend by client", "Conversations over time"])
      await page.locator('summary[aria-label="Options for Practice costs dashboard"]').click()
      await page.getByRole("button", { name: "Rename", exact: true }).click()
      await page.getByRole("textbox", { name: "Dashboard name", exact: true }).fill("Client work")
      await page.getByRole("button", { name: "Save", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Client work", exact: true })).toBeVisible()
      await dashboard(page, true, true)
      await expect(page.getByRole("tab", { name: "Client work", exact: true })).toHaveAttribute("aria-selected", "true")
      await page.locator('summary[aria-label="Options for Client work dashboard"]').click()
      await page.getByRole("button", { name: "Delete dashboard", exact: true }).click()
      await page.getByRole("group", { name: "Delete dashboard", exact: true }).getByRole("button", { name: "Delete dashboard", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Client work", exact: true })).toHaveCount(0)
      await page.locator(".dash-snackbar").getByRole("button", { name: "Undo", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Client work", exact: true })).toHaveAttribute("aria-selected", "true")
      clean(state)
    } finally { await context.close() }
  })

  for (const mode of ["failed", "invalid"] as const) await scenario(`AI ${mode} response leaves the layout and cleans its session`, async () => {
    const state = fixture("populated"); state.aiFailure = mode === "failed"
    if (mode === "invalid") state.aiPlan = { widgets: [{ type: "chart", title: "Undated drafts", metric: "drafts", groupBy: "day", form: "line" }] }
    const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      await askButton(page).click()
      await page.getByRole("textbox", { name: "Tell AI what you want from this dashboard", exact: true }).fill("Please update the dashboard")
      await page.getByRole("button", { name: "Update dashboard", exact: true }).click()
      await expect(page.locator(".dash-ai-error")).toBeVisible()
      expect(await order(page)).toEqual(PRACTICE)
      await expect.poll(() => state.aiDeleted.length).toBe(1)
      await capture(page, `ai-${mode}-1440x1000`, { fullPage: false })
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Remove and Undo on a phone (375x812)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 375, 812)
    try {
      await dashboard(page)
      await page.getByRole("button", { name: "Remove Needs you", exact: true }).click()
      await expect(card(page, "Needs you")).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Pin Recent drafts", exact: true })).toBeFocused()
      await capture(page, "ai-remove-undo-375x812", { fullPage: false })
      await page.locator(".dash-snackbar").getByRole("button", { name: "Undo", exact: true }).click()
      expect(await order(page)).toEqual(PRACTICE)
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Periods and tooltips (1440x1000)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      const total = card(page, "Total spend")
      await expect(period(page).getByRole("button", { name: "This month", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(total.getByText("This month", { exact: true })).toBeVisible()
      await period(page).getByRole("button", { name: "Last month", exact: true }).click()
      await expect(period(page).getByRole("button", { name: "Last month", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(total.getByText("Last month", { exact: true })).toBeVisible()
      await expect(total.locator(".dash-stat-value")).toHaveText(money(monthSpend(state, 8)))
      await expect(total.getByText(`Month before ${money(monthSpend(state, 7))}`, { exact: true })).toBeVisible()
      await expect(card(page, "Spend by matter").getByText("Last month", { exact: true })).toBeVisible()
      await capture(page, "practice-period-last-month-1440x1000")
      await period(page).getByRole("button", { name: "All time", exact: true }).click()
      await expect(total.getByText("All time", { exact: true })).toBeVisible()
      await expect(total.locator(".dash-stat-value")).toHaveText(money(own(state).reduce((sum, row) => sum + row.cost, 0)))
      // The period is saved with the layout.
      await dashboard(page, true, true)
      await expect(period(page).getByRole("button", { name: "All time", exact: true })).toHaveAttribute("aria-pressed", "true")
      await period(page).getByRole("button", { name: "This month", exact: true }).click()
      await expect(page.locator(".dash-skeleton")).toHaveCount(0)
      // Week bar tooltip on the busiest day names the matters behind it.
      const week = card(page, "This week")
      await week.scrollIntoViewIfNeeded()
      await week.locator(".dash-week-track").nth(3).hover()
      const weekTip = week.getByRole("tooltip")
      await expect(weekTip).toContainText("Mon: $11.90")
      await expect(weekTip).toContainText("Alder House · Studio lease renewal $5.80")
      await expect(weekTip).toContainText("1 other matter")
      screens.push({ name: "week-tip-geometry", tip: await weekTip.boundingBox(), card: await week.boundingBox() })
      await capture(page, "practice-week-tooltip-1440x1000", { fullPage: false })
      // A tooltip by the right-hand edge keeps its width and stays on the dashboard.
      const area = (await page.locator(".dash").boundingBox())!
      for (const index of [3, 6]) {
        await week.locator(".dash-week-track").nth(index).hover()
        const tip = (await weekTip.boundingBox())!
        expect(tip.x).toBeGreaterThanOrEqual(area.x - 0.5)
        expect(tip.x + tip.width).toBeLessThanOrEqual(area.x + area.width + 0.5)
        if (index === 3) expect(tip.width).toBeGreaterThan(200)
      }
      // Calendar cell tooltip on one of the busiest days.
      await page.mouse.move(5, 5)
      const cell = card(page, "Activity").locator("rect.dash-cell.level-4").last()
      await cell.scrollIntoViewIfNeeded()
      await cell.hover()
      const cellTip = card(page, "Activity").getByRole("tooltip")
      await expect(cellTip).toContainText(/conversations? on /)
      screens.push({ name: "calendar-tip-geometry", tip: await cellTip.boundingBox(), card: await card(page, "Activity").boundingBox() })
      await capture(page, "practice-calendar-tooltip-1440x1000", { fullPage: false })
      // With no widget left that follows the period, the choice goes too.
      await page.mouse.move(5, 5)
      await hideWidget(page, "Total spend")
      await expect(period(page)).toBeVisible()
      await hideWidget(page, "Spend by matter")
      await expect(period(page)).toHaveCount(0)
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Partial figures say so (1440x1000)", async () => {
    const state = fixture("populated", { data: populated({ unpriced: true }) })
    state.failSessions.add("/fixture/matters/quarry")
    const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      const banner = page.getByRole("status").filter({ hasText: "Some matters could not be read, so their spend and drafts are missing here." })
      await expect(banner).toBeVisible()
      // A model that reports no price makes the total an estimate.
      // The total leaves out the matter it could not read, and says why.
      const month = monthSpend(state, 9)
      const quarry = octoberByMatter(state).find((row) => row.id === "m-quarry")!.cost
      await expect(card(page, "Total spend").locator(".dash-stat-value")).toHaveText(money(month - quarry))
      await expect(card(page, "Total spend").getByText("Estimate: some AI use has no price on record, so the total may be low.", { exact: true })).toBeVisible()
      await capture(page, "practice-matters-unread-1440x1000", { fullPage: false })
      // Once the folder reads again, Retry clears the warning.
      state.failSessions.clear()
      await banner.getByRole("button", { name: "Retry", exact: true }).click()
      await expect(banner).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Refresh", exact: true })).toBeEnabled()
      await expect(card(page, "Total spend").locator(".dash-stat-value")).toHaveText(money(month))
      clean(state)
    } finally { await context.close() }
  })

  await scenario("Export spend by matter as CSV (1440x1000)", async () => {
    const state = fixture("populated"); const { page, context } = await open(state, 1440, 1000)
    try {
      await dashboard(page)
      const [download] = await Promise.all([
        page.waitForEvent("download"),
        card(page, "Spend by matter").getByRole("button", { name: "Export CSV", exact: true }).click(),
      ])
      expect(download.suggestedFilename()).toBe("spend-month-2026-10-29.csv")
      const csv = await readFile(await download.path(), "utf8")
      // A byte-order mark, so Excel reads "&" and "·" in matter titles as UTF-8.
      expect(csv.startsWith("\uFEFF"), "CSV starts with a UTF-8 BOM").toBe(true)
      const lines = csv.slice(1).split("\r\n").filter(Boolean)
      expect(lines[0]).toBe("Client,Matter,Reference,Conversations,Spend (USD)")
      // One line per matter that spent this month, sample left out, each with its spend.
      const spend = octoberByMatter(state)
      expect(lines.length - 1).toBe(spend.length)
      for (const row of spend) {
        const matter = MATTERS.find((entry) => entry.id === row.id)!
        expect(lines.some((line) => line.startsWith(`${matter.client ?? ""},${matter.title},${matter.reference ?? ""},`) && line.endsWith(`,${row.cost.toFixed(2)}`)), `${matter.title} in the CSV`).toBe(true)
      }
      expect(csv).not.toContain(titleOf("m-sample"))
      clean(state)
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "dashboard-results.json"), JSON.stringify({ now: new Date(NOW).toISOString(), results, screens, menuHits, fixtures: allStates.map(({ kind, errors, unexpected, requests }) => ({ kind, errors, unexpected, requests: requests.length, paths: [...new Set(requests.map((row) => `${row.method} ${row.path}`))] })) }, null, 2) + "\n")
  await browser.close()
}
expect(results.filter(result => result.status === "failed")).toEqual([])
console.log(`PASS ${results.length} scenarios; ${screens.length} screenshots.`)
