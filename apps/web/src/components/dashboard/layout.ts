import { chartKey, checkPlan, MAX_WIDGETS, type Plan, type PlanWidget } from "../../../../../dochaus/lib/dashboard-plan"
import { readBrowserJSON, writeBrowserValue } from "../../browser-storage"
import { PERIODS, type Period } from "./data"

// tile: one column; wide: two columns; full: the whole row. Rows are as tall
// as their tallest card.
export type WidgetSize = "tile" | "wide" | "full"

export const WIDGETS = {
  journey: { title: "First steps", description: "From a new matter to your first draft, one step at a time.", size: "full" },
  setup: { title: "Your practice", description: "Your name, firm, jurisdictions and practice areas, filled in right here.", size: "full" },
  attention: { title: "Needs you", description: "Drafting runs waiting on an answer, a decision or a fix.", size: "wide" },
  drafts: { title: "Recent drafts", description: "Finished drafts and what each one needs next.", size: "wide" },
  spend: { title: "Total spend", description: "Estimated AI spend across your matters in the period.", size: "tile" },
  "matter-cost": { title: "Average matter cost", description: "Spend per matter, across every matter that has used AI.", size: "tile" },
  "draft-cost": { title: "Average draft cost", description: "Spend per drafted document, including review.", size: "tile" },
  week: { title: "This week", description: "Spend on each of the last seven days.", size: "tile" },
  recent: { title: "Recent matters", description: "The matters you worked on last.", size: "wide" },
  "matter-spend": { title: "Spend by matter", description: "Which matters and clients used the most in the period, with a CSV export.", size: "wide" },
  activity: { title: "Activity", description: "Conversations started each day over the last year.", size: "full" },
  trend: { title: "Spend over time", description: "Daily spend across the period.", size: "wide" },
  tokens: { title: "Where tokens go", description: "Input, cached input, output and reasoning.", size: "wide" },
  sample: { title: "Sample matter", description: "A fictional matter to explore before your own.", size: "tile" },
  ideas: { title: "Things to try", description: "Ask, draft, review across documents or run a workflow.", size: "wide" },
  actions: { title: "Quick actions", description: "Start a matter, add a template, client or playbook.", size: "wide" },
  library: { title: "Your library", description: "Templates, playbooks, clients and matters.", size: "wide" },
} as const satisfies Record<string, { title: string; description: string; size: WidgetSize }>

export type WidgetId = keyof typeof WIDGETS
const ALL = Object.keys(WIDGETS) as WidgetId[]

// Starting out, or in practice once a draft of the lawyer's own has finished.
export type Stage = "starting" | "practice"
const STAGES: Stage[] = ["starting", "practice"]

// The default dashboards. Starting out, the figures would all be zero, so
// first steps lead. Setup has its own dashboard tab. Once a draft has finished, the work and the
// figures take over. Explicit saved layouts can still include the practice widget.
export const DEFAULT_WIDGETS: Record<Stage, WidgetId[]> = {
  starting: ["journey"],
  practice: ["attention", "drafts", "spend", "matter-cost", "draft-cost", "week", "recent", "matter-spend", "activity"],
}

// Widgets whose figures follow the chosen period.
export const PERIODIC = new Set<WidgetId>(["spend", "matter-spend", "tokens", "trend"])

// The lawyer's changes to one stage's default: the order they set, the
// defaults they hid and the extras they added.
export type Arrangement = { order?: WidgetId[]; hidden: WidgetId[]; added: WidgetId[] }

export type Layout = {
  // Practice keeps its own arrangement, so arranging the first-steps cards
  // never keeps the practice figures away.
  stages: Partial<Record<Stage, Arrangement>>
  period: Period
  // The stage and widgets last shown, so placeholders take the same places on reload.
  stage?: Stage
  shown?: WidgetId[]
}

export const DEFAULT_LAYOUT: Layout = { stages: {}, period: "month" }

// Which widgets a stage can show: First steps only while starting out, and
// setup only while an essential is missing.
export function available(stage: Stage, setupDone: boolean) {
  return ALL.filter((id) => !(id === "journey" && stage !== "starting") && !(id === "setup" && setupDone))
}

// The stage's defaults first, then everything else it can show.
function canonical(stage: Stage) {
  return [...DEFAULT_WIDGETS[stage], ...ALL.filter((id) => !DEFAULT_WIDGETS[stage].includes(id))]
}

export function visibleWidgets(stage: Stage, setupDone: boolean, arrangement?: Arrangement) {
  const can = available(stage, setupDone)
  const shown = [...DEFAULT_WIDGETS[stage], ...(arrangement?.added ?? [])].filter((id) => can.includes(id) && !arrangement?.hidden.includes(id))
  // The lawyer's order where they set one; anything else takes its default
  // place after the nearest widget before it.
  const order = canonical(stage)
  const result = (arrangement?.order ?? []).filter((id) => shown.includes(id))
  for (const id of order) {
    if (!shown.includes(id) || result.includes(id)) continue
    const before = order.slice(0, order.indexOf(id)).reverse().find((other) => result.includes(other))
    result.splice(before ? result.indexOf(before) + 1 : 0, 0, id)
  }
  return result
}

// Per browser, like the sidebar's collapsed state.
const KEY = "dh.dashboard.v3"

const ids = (value: unknown) =>
  Array.isArray(value) ? [...new Set(value.filter((id): id is WidgetId => typeof id === "string" && Object.hasOwn(WIDGETS, id)))] : undefined

function arrangement(value: unknown): Arrangement | undefined {
  if (typeof value !== "object" || value === null) return undefined
  const saved = value as Record<string, unknown>
  return { order: ids(saved.order), hidden: ids(saved.hidden) ?? [], added: ids(saved.added) ?? [] }
}

export function loadLayout(): Layout {
  const saved = readBrowserJSON(KEY)
  if (typeof saved !== "object" || saved === null || Array.isArray(saved)) return DEFAULT_LAYOUT
  const value = saved as Record<string, unknown>
  const period = PERIODS.includes(value.period as Period) ? (value.period as Period) : DEFAULT_LAYOUT.period
  const arranged = typeof value.stages === "object" && value.stages !== null ? (value.stages as Record<string, unknown>) : {}
  const stages = Object.fromEntries(STAGES.flatMap((stage) => { const entry = arrangement(arranged[stage]); return entry ? [[stage, entry]] : [] }))
  const stage = STAGES.includes(value.stage as Stage) ? (value.stage as Stage) : undefined
  return { stages, period, stage, shown: ids(value.shown) }
}

export function saveLayout(layout: Layout) {
  writeBrowserValue(KEY, JSON.stringify(layout))
}


export type WidgetInstance = { key: string; widget: PlanWidget; pinned: boolean }
export type DashboardTab = {
  id: string
  name: string
  period: Period
  // Overview keeps a separate arrangement for each first-run stage until AI
  // replaces it with an explicit plan.
  stages: Partial<Record<Stage, WidgetInstance[]>>
  widgets?: WidgetInstance[]
  // Keep the AI's limitations with the view they describe, including on reload.
  notes?: string[]
  legacy?: Partial<Record<Stage, Arrangement>>
}
export type Dashboards = { active: string; dashboards: DashboardTab[] }
export const DASHBOARDS_KEY = "dh.dashboard.v4"

export function instance(widget: PlanWidget, pinned = false): WidgetInstance {
  return { key: widget.type === "chart" ? chartKey(widget) : widget.type, widget, pinned }
}

export function tabWidgets(tab: DashboardTab, stage: Stage, setupDone: boolean): WidgetInstance[] {
  const widgets = tab.widgets ?? tab.stages[stage]
  if (widgets) return widgets.filter(({ widget }) => widget.type === "chart" || available(stage, setupDone).includes(widget.type))
  return visibleWidgets(stage, setupDone, tab.legacy?.[stage]).map((type) => instance({ type }))
}

// The AI controls unpinned widgets. Pins keep their original chart settings,
// identity and position even if the proposed plan omits or changes them.
export function applyPlan(current: WidgetInstance[], plan: Plan): WidgetInstance[] {
  const pins = current.map((widget, index) => ({ widget, index })).filter(({ widget }) => widget.pinned)
  const pinned = new Set(pins.map(({ widget }) => widget.key))
  const result = plan.widgets.map((widget) => instance(widget)).filter((widget) => !pinned.has(widget.key)).slice(0, Math.max(0, MAX_WIDGETS - pins.length))
  for (const { widget, index } of pins) result.splice(Math.min(index, result.length), 0, widget)
  return result
}

function savedInstances(value: unknown): WidgetInstance[] | undefined {
  if (!Array.isArray(value)) return undefined
  if (!value.length) return []
  const result: WidgetInstance[] = []
  const seen = new Set<string>()
  // Validate one persisted widget at a time: the older catalog allowed all 17
  // templates, while new model plans are capped at 16. Keep those saved views.
  for (const entry of value.slice(0, 32)) {
    if (typeof entry !== "object" || entry === null || Array.isArray(entry)) continue
    const saved = entry as Record<string, unknown>
    const checked = checkPlan({ widgets: [saved.widget] })
    if (!checked.plan) continue
    const next = instance(checked.plan.widgets[0]!, saved.pinned === true)
    if (seen.has(next.key)) continue
    seen.add(next.key)
    result.push(next)
  }
  return result.length || value.length === 0 ? result : undefined
}

export function readDashboards(saved: unknown, legacy: Layout = DEFAULT_LAYOUT): Dashboards {
  if (typeof saved === "object" && saved !== null && !Array.isArray(saved)) {
    const value = saved as Record<string, unknown>
    if (Array.isArray(value.dashboards)) {
      const seen = new Set<string>()
      const dashboards = value.dashboards.flatMap((entry): DashboardTab[] => {
        if (typeof entry !== "object" || entry === null || Array.isArray(entry)) return []
        const tab = entry as Record<string, unknown>
        if (typeof tab.id !== "string" || !tab.id || tab.id === "__setup" || seen.has(tab.id) || typeof tab.name !== "string" || !tab.name.trim()) return []
        seen.add(tab.id)
        const stages = typeof tab.stages === "object" && tab.stages !== null ? tab.stages as Record<string, unknown> : {}
        const legacyStages = typeof tab.legacy === "object" && tab.legacy !== null ? tab.legacy as Record<string, unknown> : {}
        const notes = Array.isArray(tab.notes) ? tab.notes.filter((note): note is string => typeof note === "string" && !!note.trim()).slice(0, 5).map((note) => note.trim().slice(0, 160)) : undefined
        return [{ id: tab.id, name: tab.name.trim().slice(0, 30), period: PERIODS.includes(tab.period as Period) ? tab.period as Period : "month", widgets: savedInstances(tab.widgets), ...(notes ? { notes } : {}),
          stages: Object.fromEntries(STAGES.flatMap((stage) => { const widgets = savedInstances(stages[stage]); return widgets ? [[stage, widgets]] : [] })),
          legacy: Object.fromEntries(STAGES.flatMap((stage) => { const entry = arrangement(legacyStages[stage]); return entry ? [[stage, entry]] : [] })),
        }]
      }).slice(0, 12)
      if (!value.dashboards.length && value.active === "") return { active: "", dashboards: [] }
      if (dashboards.length) return { active: dashboards.some((tab) => tab.id === value.active) ? value.active as string : dashboards[0]!.id, dashboards }
    }
  }
  return { active: "overview", dashboards: [{ id: "overview", name: "Overview", period: legacy.period, stages: {}, legacy: legacy.stages }] }
}

export function loadDashboards(): Dashboards {
  return readDashboards(readBrowserJSON(DASHBOARDS_KEY), loadLayout())
}

export function saveDashboards(dashboards: Dashboards) {
  return writeBrowserValue(DASHBOARDS_KEY, JSON.stringify(dashboards))
}
