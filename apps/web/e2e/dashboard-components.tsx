// Isolated, synthetic component gallery. No product APIs are imported or called.
import { useEffect, useState, type ComponentProps } from "react"
import { createRoot } from "react-dom/client"
import { BrowserRouter } from "react-router-dom"
import "@fontsource/space-grotesk/500.css"
import "@fontsource/space-grotesk/600.css"
import "@fontsource/space-grotesk/700.css"
import "../src/styles.css"
import "../src/styles/primitives.css"
import "../src/styles/design-system.css"
import "../src/styles/dashboard.css"
import "../src/styles/dashboard-chart.css"
import "../src/styles/contrast.css"
import logo from "../src/assets/dochaus-logo.svg"
import { icons, type IconName } from "../src/components/dashboard/icons"
import { AdaptiveGrid } from "../src/components/dashboard/visual/DashboardChrome"
import { SetupChecklist, DemoDataCard, type SetupChecklistConfig, type DemoDataCardConfig } from "../src/components/dashboard/visual/SetupCards"
import { LiveComponent, useDashboardSource } from "../src/components/dashboard/visual/LiveComponents"
import { createDashboardSource, type JsonValue } from "../src/components/dashboard/visual/source"

type Definition = ComponentProps<typeof LiveComponent>["definition"]
type Source = ReturnType<typeof createDashboardSource>
type DemoAPI = {
  publish: (id: string, data: JsonValue) => void
  snapshot: (id: string) => ReturnType<Source["getSnapshot"]> | undefined
  fail: (id: string, message?: string) => void
  replaceSource: (id: string, data?: JsonValue) => void
  loading: (id: string) => void
  select: (tab: string) => void
  openComposer: () => void
  definitions: () => Definition[]
  setDefinition: (id: string, definition: Definition) => void
  setDefinitions: (definitions: Definition[]) => void
  actions: () => string[]
}
declare global { interface Window { __dashboardComponentDemo: DemoAPI } }

const DAY = 86_400_000
const TODAY = new Date(2026, 9, 29).getTime()
const clients = ["Alder House Studio", "Northbank Works", "Fernlea Bakery Co.", "Marlowe & Pike Design", "No client recorded", "Tidewater Robotics", "Quarry Lane Cycles"]
const spend = [15.28, 11.29, 10.53, 9.4, 7.67, 5.05, 3]
const attentionRows = [
  { id: "failed", title: "Drafting run failed", detail: "Tidewater Robotics · Mutual NDA set", icon: "error", tone: "error", action: { label: "Open" } },
  { id: "questions", title: "2 questions waiting", detail: "Marlowe & Pike Design · Master services agreement", icon: "question", tone: "info", action: { label: "Answer" } },
  { id: "stalled", title: "Drafting stalled, press Continue", detail: "Copperleaf Gardens · Employment contracts", icon: "pause", tone: "warning", action: { label: "Resume" } },
]
const matterRows = [
  { id: "copperleaf", title: "Copperleaf Gardens · Employment contracts", detail: "No client recorded", href: "/component-preview/copperleaf", updated: "1h ago", spend: "$35.94", status: { icon: "draft", label: "Drafting", tone: "neutral" } },
  { id: "alder", title: "Alder House · Studio lease renewal", detail: "Alder House Studio · AH-2026-014", href: "/component-preview/alder", updated: "20h ago", spend: "$57.26", status: { icon: "check", label: "Drafted", tone: "good" } },
  { id: "fernlea", title: "Fernlea Bakery · Shareholders’ agreement", detail: "Fernlea Bakery Co. · FB-2026-003", href: "/component-preview/fernlea", updated: "28 Oct 2026", spend: "$72.50", status: { icon: "check", label: "Drafted", tone: "good" } },
  { id: "marlowe", title: "Marlowe & Pike Design · Master services agreement", detail: "Marlowe & Pike Design · MP-2026-011", href: "/component-preview/marlowe", updated: "27 Oct 2026", spend: "$30.66", status: { icon: "warning", label: "Needs you", tone: "warning" } },
  { id: "tidewater", title: "Tidewater Robotics · Mutual NDA set", detail: "Tidewater Robotics · TR-2026-030", href: "/component-preview/tidewater", updated: "26 Oct 2026", spend: "$32.99", status: { icon: "error", label: "Failed", tone: "error" } },
]
const draftRows = [
  { id: "shareholders", title: "Shareholders’ agreement", detail: "Fernlea Bakery · 28 Oct 2026", href: "/component-preview/shareholders", updated: "", spend: "", status: { icon: "warning", label: "2 findings to resolve", tone: "warning" } },
  { id: "lease", title: "Lease renewal +1 more", detail: "Alder House · 26 Oct 2026", href: "/component-preview/lease", updated: "", spend: "", status: { icon: "warning", label: "Changed since final check", tone: "warning" } },
  { id: "supply", title: "Supply agreement", detail: "Northbank Works · 20 Oct 2026", href: "/component-preview/supply", updated: "", spend: "", status: { icon: "check", label: "Ready for signature", tone: "good" } },
  { id: "franchise", title: "Franchise terms", detail: "Quarry Lane Cycles · 4 Sept 2026", href: "/component-preview/franchise", updated: "", spend: "", status: { icon: "draft", label: "Review draft", tone: "neutral" } },
]

const initialPractice: JsonValue = {
  attention: attentionRows, matters: matterRows, drafts: draftRows,
  spend: { value: "$62.22", caption: "This month · 36 conversations", tag: "+$12.03 vs $50.19", visual: { type: "sparkline", values: [6.1, 1.4, 3.8, 2.7, 6.4, 1.9, 4.6, 5.2, 15.3, 14.82], label: "Spend through this month" } },
  matterCost: { value: "$43.86", caption: "All time · across 7 matters", tag: "$30–$72 range", visual: { type: "spread", values: [72.5, 57.26, 35.94, 32.99, 30.66, 39.67, 38], average: 43.86, label: "All-time spend per matter", format: "money" } },
  draftCost: { value: "$4.00", caption: "All time · 5 documents from 4 runs", tag: "4 runs · 5 docs", visual: { type: "spread", values: [6, 3, 3.5, 4.5, 3], average: 4, label: "Cost per finished draft", format: "money" } },
  week: { value: "$34.20", caption: "Last seven days · $11.80 prior", tag: "+$22.40", visual: { type: "week", days: [3.4, .45, .3, 10.8, 6.5, 10.55, 2.2].map((cost, index) => ({ day: TODAY - (6 - index) * DAY, cost, label: ["Fri", "Sat", "Sun", "Mon", "Tue", "Wed", "Thu"][index]!, today: index === 6 })) } },
  library: [
    { id: "templates", icon: "template", label: "Templates", count: 12, detail: "Your firm’s precedents", emptyLabel: "Add a template", href: "/component-preview/templates" },
    { id: "playbooks", icon: "playbook", label: "Playbooks", count: 4, detail: "Shared review standards", emptyLabel: "Add a playbook", href: "/component-preview/playbooks" },
    { id: "clients", icon: "client", label: "Clients", count: 6, detail: "People and organisations", emptyLabel: "Add a client", href: "/component-preview/clients" },
    { id: "matters", icon: "matter", label: "Matters", count: 7, detail: "Your active practice", emptyLabel: "Open a matter", href: "/component-preview/matters" },
  ],
}
const activity = Array.from({ length: 365 }, (_, index) => ({ key: `day-${index}`, label: `Day ${index + 1}`, time: TODAY - (364 - index) * DAY, value: index % 7 > 4 ? 0 : (index * 7 + Math.floor(index / 10)) % 5 }))
const initialCharts: JsonValue = {
  clientSpend: clients.map((label, index) => ({ key: label, label, value: spend[index]! })),
  matterSpend: ["Alder House · Studio lease renewal", "Northbank Works · Supply agreement", "Fernlea Bakery · Shareholders’ agreement", "Marlowe & Pike Design · Master services agreement", "Copperleaf Gardens · Employment contracts", "Tidewater Robotics · Mutual NDA set", "Quarry Lane Cycles · Franchise terms"].map((label, index) => ({ key: `matter-${index}`, label, value: spend[index]!, detail: clients[index]! })),
  conversations: [0, 7, 0, 1, 0, 11, 12, 6, 12, 16, 4, 9, 11].map((value, index) => ({ key: `week-${index}`, label: new Date(TODAY - (12 - index) * 7 * DAY).toLocaleDateString("en-GB", { day: "numeric", month: "short" }), time: TODAY - (12 - index) * 7 * DAY, value })),
  activity,
  dailySpend: [1.8, 0.6, 0, 0, 9.8, 2.1, 4.3, 3.8, 1.5, 0, 0, 6.2, 8.1, 3.6].map((value, index) => ({ key: `spend-${index}`, label: `${16 + index} Oct`, time: TODAY - (13 - index) * DAY, value })),
  tokens: [{ key: "input", label: "Input", value: 9520000 }, { key: "cached", label: "Cached input", value: 30464000 }, { key: "output", label: "Output", value: 5236000 }, { key: "reasoning", label: "Reasoning", value: 2380000 }],
}
const setupConfig: SetupChecklistConfig = {
  title: "Finish setting up", description: "Connect your tools and defaults so every draft and search is ready", items: [
    { id: "models", icon: "model", title: "AI provider & models connected", detail: "OpenCode Go · Primary & fast models configured", done: true, action: { id: "models", label: "Change" } },
    { id: "identity", icon: "person", title: "Practice identity & default law", detail: "Stamps your comments and preselects your jurisdiction on new matters", done: false, editable: true },
    { id: "research", icon: "search", title: "Connect Exa web research", detail: "Add your Exa API key to enable live legal and company web search", done: false, editable: true },
    { id: "word", icon: "word", title: "Install Microsoft Word add-in", detail: "Review clauses, run playbooks and apply tracked changes inside Word", done: false, action: { id: "word", label: "Set up Word", variant: "outline" } },
    { id: "style", icon: "template", title: "House style & firm templates", detail: "Add drafting conventions and upload your precedent .docx agreements", done: false, action: { id: "style", label: "Configure", variant: "outline" } },
  ],
}
const demoConfig: DemoDataCardConfig = { title: "Explore with demo data", description: "See a populated workspace in one click", features: [{ value: "5", label: "Sample matters & drafts" }, { value: "3", label: "Precedent templates" }, { value: "2", label: "Review playbooks" }, { value: "4", label: "Client records & spend" }], action: { id: "populate", label: "Populate with demo data", icon: "plus" }, note: "Fictional data only · You can clear demo data at any time" }

const frame = (title: string, icon: IconName, size: "tile" | "wide" | "full" = "wide", tone = "default", subtitle?: string) => ({ title, icon, size, tone, ...(subtitle ? { subtitle } : {}) })
const binding = (source: string, path: string) => ({ source, path })
const stat = (id: string, title: string, path: string, icon: IconName, tone: string) => ({ id, kind: "stat", frame: frame(title, icon, "tile", tone), bindings: { value: binding("practice", `${path}/value`), caption: binding("practice", `${path}/caption`), tag: binding("practice", `${path}/tag`), visual: binding("practice", `${path}/visual`) } })
const allDefinitions = [
  { id: "attention", kind: "attention", frame: frame("Needs you", "attention", "wide", "bold", "3 drafting runs waiting on your input"), bindings: { rows: binding("practice", "/attention") } },
  { id: "drafts", kind: "matters", frame: frame("Recent drafts", "draft", "wide", "default", "Latest first, with what each needs"), bindings: { rows: binding("practice", "/drafts") } },
  stat("spend", "Total spend", "/spend", "spend", "tint-1"),
  stat("matter-cost", "Average matter cost", "/matterCost", "matter", "tint-2"),
  stat("draft-cost", "Average draft cost", "/draftCost", "draft", "tint-3"),
  stat("week", "This week", "/week", "week", "ink"),
  { id: "recent", kind: "matters", frame: frame("Recent matters", "recent", "wide", "default", "Latest first, with all-time spend"), bindings: { rows: binding("practice", "/matters") } },
  { id: "matter-spend", kind: "chart", frame: frame("Spend by matter", "bars", "wide", "default", "This month · estimated AI spend"), props: { form: "bars", label: "Spend by matter", format: "money", showShares: true }, bindings: { points: binding("charts", "/matterSpend") } },
  { id: "activity", kind: "chart", frame: frame("Activity", "activity", "full", "default", "Conversations started each day over the last year"), props: { form: "calendar", label: "Daily conversations", basis: "conversations in the last year", format: "count" }, bindings: { points: binding("charts", "/activity") } },
  { id: "trend", kind: "chart", frame: frame("Spend over time", "trend", "wide", "default", "Last 14 days · estimated spend"), props: { form: "line", label: "Spend over time", format: "money", basis: "estimated AI spend" }, bindings: { points: binding("charts", "/dailySpend") } },
  { id: "tokens", kind: "tokens", frame: frame("Where tokens go", "tokens", "wide", "default", "This month · input, cache, output and reasoning"), props: { period: "This month" }, bindings: { segments: binding("charts", "/tokens") } },
  { id: "sample", kind: "sample", frame: frame("Sample matter", "sample", "wide", "default", "A fictional matter, ready to explore"), props: { items: [{ id: "engagement", icon: "draft", title: "Engagement letter", detail: "A precedent to explore" }, { id: "review", icon: "review", title: "Legal review", detail: "Findings and cited answers" }, { id: "sources", icon: "documents", title: "Matter documents", detail: "The material behind the work" }] } },
  { id: "ideas", kind: "actions", frame: frame("Things to try", "idea", "wide", "default", "A few ways to start"), props: { items: [{ id: "ask", icon: "documents", title: "Ask about documents", detail: "Find key points, deadlines and gaps" }, { id: "draft", icon: "template", title: "Draft from a template", detail: "Start from your firm’s own precedent" }, { id: "review", icon: "review", title: "Review across documents", detail: "Ask the same questions of every file" }, { id: "workflow", icon: "workflow", title: "Run a workflow", detail: "A repeatable routine for a matter" }] } },
  { id: "actions", kind: "actions", frame: frame("Quick actions", "actions", "wide", "default", "Start something new"), props: { items: [{ id: "new", icon: "matter", title: "New matter", detail: "Open a folder for an engagement" }, { id: "template", icon: "template", title: "Add a template", detail: "Draft from your firm’s precedent" }, { id: "client", icon: "client", title: "Add a client", detail: "Keep parties and details on file" }, { id: "sample", icon: "sample", title: "Try the sample", detail: "Explore a matter with documents" }] } },
  { id: "library", kind: "library", frame: frame("Your library", "library", "wide", "default", "Templates, playbooks, clients and matters"), bindings: { cards: binding("practice", "/library") } },
] as Definition[]
const customDefinitions = [
  allDefinitions[0]!,
  { id: "client-spend", kind: "chart", frame: frame("Spend by client", "spend", "wide", "default", "This month · grouped by client record"), props: { form: "bars", label: "Spend by client", format: "money", basis: "estimated spend (USD)", showShares: true, badge: "7 clients" }, bindings: { points: binding("charts", "/clientSpend") } },
  { id: "conversations", kind: "chart", frame: frame("Conversations over time", "trend", "wide", "default", "Weekly trend · Last 3 months"), props: { form: "line", label: "Weekly conversations", format: "count", basis: "conversations started", badge: "Peak: 16 / wk" }, bindings: { points: binding("charts", "/conversations") } },
] as Definition[]

function Gallery() {
  const [sources, setSources] = useState<Record<string, Source>>(() => ({ practice: createDashboardSource(initialPractice), charts: createDashboardSource(initialCharts), pending: createDashboardSource(), ui: createDashboardSource({ activeId: new URLSearchParams(location.search).get("view") ?? "overview", open: false, request: "", working: false, tabs: [{ id: "overview", label: "Overview" }, { id: "catalog", label: "All widgets" }, { id: "costs", label: "Practice costs" }, { id: "setup", label: "Setup", badge: { value: 1, total: 5 } }] }) }))
  const [definitions, setDefinitions] = useState<Definition[]>(allDefinitions)
  const [custom, setCustom] = useState<Definition[]>(customDefinitions)
  const uiSnapshot = useDashboardSource(sources.ui!)
  const ui = (uiSnapshot.data ?? {}) as Record<string, JsonValue>
  const tab = typeof ui.activeId === "string" ? ui.activeId : "overview"
  const updateUI = (next: Record<string, JsonValue>) => sources.ui!.publish({ ...(sources.ui!.getSnapshot().data as Record<string, JsonValue> ?? {}), ...next })
  const setTab = (activeId: string) => updateUI({ activeId })
  const setComposer = (open: boolean) => updateUI({ open })
  const [name, setName] = useState("Alex Example"), [firm, setFirm] = useState("Example Partners LLP")
  const [setup, setSetup] = useState(setupConfig)
  const [actions, setActions] = useState<string[]>([])
  const [notice, setNotice] = useState("")
  const record = (id: string) => { setActions((list) => [...list, id]); setNotice(`Preview action: ${id}`) }
  useEffect(() => {
    window.__dashboardComponentDemo = {
      publish: (id, data) => { if (!sources[id]) throw new Error(`Unknown source: ${id}`); sources[id].publish(data) },
      snapshot: (id) => sources[id]?.getSnapshot(),
      fail: (id, message) => sources[id]?.fail(message),
      replaceSource: (id, data) => setSources((value) => ({ ...value, [id]: createDashboardSource(data) })),
      loading: (id) => setSources((value) => ({ ...value, [id]: createDashboardSource() })),
      select: setTab, openComposer: () => setComposer(true), definitions: () => definitions,
      setDefinition: (id, definition) => setDefinitions((value) => value.map((item) => item.id === id ? definition : item)),
      setDefinitions, actions: () => actions,
    }
  }, [sources, definitions, actions])
  const selected = tab === "costs" ? custom : tab === "catalog" ? definitions : definitions.filter((item) => ["attention", "drafts", "spend", "matter-cost", "draft-cost", "week", "recent", "matter-spend", "activity"].includes(item.id))
  const renderItem = (id: string) => <LiveComponent definition={selected.find((item) => item.id === id)!} sources={sources} onAction={record} />
  const finishItem = (id: string) => {
    const items = setup.items.map((item) => item.id === id ? { ...item, done: true } : item)
    setSetup({ ...setup, items })
    const tabs = (ui.tabs as { id: string; [key: string]: JsonValue }[]).map((item) => item.id === "setup" ? { ...item, badge: { value: items.filter((item) => item.done).length, total: 5 } } : item)
    updateUI({ tabs }); setNotice("Saved for this component preview")
  }
  const chromeAction = (action: string, _id: string, payload?: JsonValue) => {
    const value = payload && typeof payload === "object" && !Array.isArray(payload) ? payload : {}
    if (action === "select-tab" && typeof value.id === "string") setTab(value.id)
    if (action === "new-dashboard" || action === "refine-view") setComposer(true)
    if (action === "change-request" && typeof value.value === "string") updateUI({ request: value.value })
    if (action === "cancel-dashboard") updateUI({ open: false, working: false })
    if (action === "stop-dashboard") updateUI({ working: false })
    if (action === "generate-dashboard") { updateUI({ working: true }); window.setTimeout(() => { setCustom(customDefinitions); updateUI({ activeId: "costs", open: false, working: false }); setNotice("Dashboard preview generated from synthetic data") }, 350) }
  }
  return <div className="component-gallery">
    <style>{galleryCSS}</style>
    <aside className="component-rail" aria-label="App navigation"><div className="component-brand"><img src={logo} alt="" /><strong>Doc.Haus</strong></div><nav>{([ ["activity", "Dashboard"], ["matter", "Matters"], ["person", "Clients"], ["template", "Templates"], ["playbook", "Playbooks"], ["workflow", "Workflows"], ["library", "Skills"], ["model", "Agents"] ] as [IconName, string][]).map(([icon, label], index) => <a key={label} href={`#${label.toLowerCase()}`} className={index === 0 ? "active" : ""} aria-current={index === 0 ? "page" : undefined}>{icons[icon]}<span>{label}</span></a>)}</nav><span className="component-rail-note">Component preview<br />Fictional practice data</span></aside>
    <main className="component-main"><header className="component-head"><h1>Dashboard</h1><p>{tab === "setup" ? "Finish configuring your workspace or load demo data to explore." : `Good morning, ${name}. What needs you, your drafts and what your matters cost.`}</p></header><div className="dash component-content">
      <LiveComponent definition={{ id: "tabs", kind: "tabs", bindings: { items: binding("ui", "/tabs"), activeId: binding("ui", "/activeId") } }} sources={sources} onAction={chromeAction} />
      {tab === "costs" && <LiveComponent definition={{ id: "about", kind: "about", props: { config: { text: "Time per fee earner is not tracked in matter data — showing client spend and weekly conversations instead." } } }} sources={sources} onAction={chromeAction} />}
      {tab === "setup" ? <div className="dash-setup-grid"><SetupChecklist config={setup} onAction={record} renderContent={(id) => id === "identity" ? <div className="dash-setup-inline"><div className="dash-setup-input-row"><label><span className="sr-only">Your name</span><input value={name} onChange={(event) => setName(event.target.value)} /></label><label><span className="sr-only">Firm name</span><input value={firm} onChange={(event) => setFirm(event.target.value)} /></label><button type="button" onClick={() => finishItem("identity")}>Save</button></div><div className="dash-setup-chips">{["England & Wales", "Scotland", "Northern Ireland", "Commercial", "Corporate", "Employment"].map((label, index) => <button type="button" key={label} aria-pressed={index === 0 || index === 3} onClick={() => record(label)}>{label}{(index === 0 || index === 3) && icons.check}</button>)}</div></div> : id === "research" ? <form className="dash-setup-exa" onSubmit={(event) => { event.preventDefault(); finishItem("research") }}><label><span className="sr-only">Exa API key</span><input type="password" autoComplete="off" placeholder="Paste Exa API key (exa_…)" required /></label><button type="submit">Save key</button></form> : null} /><DemoDataCard config={demoConfig} onAction={(id) => { record(id); setTab("catalog"); setNotice("Synthetic component data is already populated") }} /></div>
        : <AdaptiveGrid items={selected.map((item) => ({ id: item.id, size: item.frame?.size ?? "wide" }))} renderItem={renderItem} mode={tab === "costs" ? "adaptive" : "catalog"} label={tab === "catalog" ? "Complete 15-widget catalog" : "Dashboard widgets"} />}
      <p className="component-notice" role="status">{notice}</p>
      <LiveComponent definition={{ id: "composer", kind: "composer", props: { config: { description: "Describe what you want to see. Your AI arranges the widgets and time periods; your app works out every figure locally.", suggestions: ["Last month spend by client", "Weekly activity & token usage", "Focus on drafts that need me", "All-time matter cost comparison"], footnote: "Creates a new tab · Synthetic component preview" } }, bindings: { open: binding("ui", "/open"), request: binding("ui", "/request"), working: binding("ui", "/working") } }} sources={sources} onAction={chromeAction} />
    </div></main>
  </div>
}

const galleryCSS = `
html,body,#root{margin:0;min-height:100%;height:auto;overflow:auto}body{background:var(--panel);color:var(--text)}
.component-gallery{min-height:100vh;display:grid;grid-template-columns:232px minmax(0,1fr)}.component-rail{position:sticky;top:0;height:100vh;box-sizing:border-box;display:flex;flex-direction:column;padding:24px 14px;background:var(--md-sys-color-surface-container-low);border-right:1px solid var(--border)}.component-brand{display:flex;gap:10px;align-items:center;padding:0 10px 30px}.component-brand img{width:28px;height:28px}.component-brand strong{font-family:var(--font-display);font-size:23px;letter-spacing:-1px}.component-rail nav{display:grid;gap:6px}.component-rail nav a{display:flex;align-items:center;gap:12px;min-height:44px;padding:0 14px;border-radius:var(--radius-pill);color:var(--text-secondary);text-decoration:none;font-size:13px}.component-rail nav a.active{color:var(--accent);font-weight:650;background:var(--md-sys-color-primary-container)}.component-rail nav a:nth-child(4){margin-top:28px}.component-rail-note{margin-top:auto;padding:16px 12px 0;border-top:1px solid var(--border);font-size:11px;line-height:1.6;color:var(--text-secondary)}.component-main{min-width:0}.component-head{padding:24px 40px 20px;border-bottom:1px solid var(--border)}.component-head h1{margin:0;font-family:var(--font-display);font-size:23px;font-weight:600;letter-spacing:-.7px}.component-head p{margin:5px 0 0;font-size:13px;color:var(--text-secondary);line-height:1.5}.component-content{padding:24px 40px 30px}.component-notice{min-height:1px;margin:0;color:var(--text-secondary);font-size:12px}.component-content .dash-card{box-sizing:border-box}.component-content .dash-adaptive-grid .size-tile .dash-card{min-height:210px}.component-content .dash-stat-chart:empty{display:none}.component-content .dash-stat{min-height:115px}.component-content .dash-card:has(.dash-sample-preview) .dash-card-body{flex:none}.component-content .dash-sample-preview{min-height:145px}.component-content .dash-chart-widget.form-calendar .dash-chart-summary{font-size:12px}.component-content .dash-chart-widget.form-calendar .dash-chart-total{font-size:19px;letter-spacing:-.4px}.component-content .dash-chart-widget.form-calendar{gap:10px}
@media(max-width:1050px){.component-gallery{grid-template-columns:190px minmax(0,1fr)}.component-head{padding:22px 24px 18px}.component-content{padding:22px 24px}}@media(max-width:760px){.component-gallery{display:block}.component-rail{height:auto;position:relative;padding:12px 16px;border-right:0;border-bottom:1px solid var(--border)}.component-brand{padding:0}.component-brand img{width:24px;height:24px}.component-brand strong{font-size:20px}.component-rail nav,.component-rail-note{display:none}.component-head{padding:18px 20px}.component-content{padding:18px 20px}.component-head h1{font-size:22px}}@media(max-width:360px){.component-head{padding:16px}.component-content{padding:16px}.component-content .dash-card{padding:16px}.component-content .dash-card-heading h2{font-size:14px}}
`

createRoot(document.getElementById("root")!).render(<BrowserRouter><Gallery /></BrowserRouter>)
