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
// first steps and setup lead. Once a draft has finished, the work and the
// figures take over. Setup joins either while an essential is missing.
export const DEFAULT_WIDGETS: Record<Stage, WidgetId[]> = {
  starting: ["journey", "setup"],
  practice: ["attention", "drafts", "setup", "spend", "matter-cost", "draft-cost", "week", "recent", "matter-spend", "activity"],
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
