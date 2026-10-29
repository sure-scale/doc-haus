import { readBrowserJSON, writeBrowserValue } from "../../browser-storage"
import type { Period } from "./data"

// tile: one column; tall: one column, two rows; wide: two columns, two rows;
// full: the whole row, as tall as its content.
export type WidgetSize = "tile" | "tall" | "wide" | "full"

export const WIDGETS = {
  setup: { title: "Your practice", description: "Your name, firm, jurisdictions, practice areas and web research, filled in right here.", size: "full" },
  spend: { title: "Total spend", description: "Model spend across every matter in the period.", size: "tile" },
  "matter-cost": { title: "Average matter cost", description: "Spend per matter that had work in the period.", size: "tile" },
  "draft-cost": { title: "Average draft cost", description: "Spend per drafted document, including review.", size: "tile" },
  documents: { title: "Documents drafted", description: "Documents delivered by finished drafting runs.", size: "tile" },
  week: { title: "This week", description: "Spend on each of the last seven days.", size: "tall" },
  "matter-spend": { title: "Spend by matter", description: "Which matters used the most in the period.", size: "wide" },
  attention: { title: "Needs you", description: "Drafting runs waiting on an answer or a decision.", size: "tall" },
  activity: { title: "Activity", description: "Conversations started each day over the last year.", size: "full" },
  recent: { title: "Recent matters", description: "The matters you worked on last.", size: "wide" },
  tokens: { title: "Where tokens go", description: "Input, cached input, output and reasoning.", size: "wide" },
  trend: { title: "Spend over time", description: "Daily spend across the period.", size: "wide" },
  actions: { title: "Quick actions", description: "Start a matter, add a template, client or playbook.", size: "wide" },
  library: { title: "Your library", description: "Templates, playbooks, clients and matters.", size: "tall" },
  journey: { title: "First steps", description: "From a new matter to your first draft, one step at a time.", size: "full" },
  sample: { title: "Sample matter", description: "A fictional matter to explore before your own.", size: "tall" },
  ideas: { title: "Things to try", description: "Ask, draft, review across documents or run a workflow.", size: "wide" },
} as const satisfies Record<string, { title: string; description: string; size: WidgetSize }>

export type WidgetId = keyof typeof WIDGETS

// The default dashboards. While the lawyer starts out, figures would all be
// zero, so first steps and the sample lead; once they have talked to the
// assistant in a matter of their own, the practice figures take over. Setup
// and first steps join either while unfinished.
export const STARTING_WIDGETS: WidgetId[] = ["sample", "library", "ideas"]
export const PRACTICE_WIDGETS: WidgetId[] = ["spend", "matter-cost", "draft-cost", "documents", "week", "matter-spend", "attention", "activity", "recent", "tokens"]

// Widgets whose figures follow the chosen period.
export const PERIODIC = new Set<WidgetId>(["spend", "matter-cost", "draft-cost", "documents", "matter-spend", "tokens", "trend"])

// Starting out, or in practice once the lawyer has talked to the assistant
// in a matter of their own.
export type Stage = "starting" | "practice"
const STAGES: Stage[] = ["starting", "practice"]

export type Layout = {
  // The lawyer's own arrangement for each stage, from their first move, hide
  // or add; until then the dashboard follows the default for where they are.
  // Practice starts from its own default, so arranging the first-steps cards
  // never keeps the practice figures away.
  widgets?: Partial<Record<Stage, WidgetId[]>>
  period: Period
  // The stage and widgets last shown, so placeholders take the same places on reload.
  stage?: Stage
  shown?: WidgetId[]
}

export const DEFAULT_LAYOUT: Layout = { period: 30 }

// Per browser, like the sidebar's collapsed state: which widgets show, in
// what order. A widget not in the list is hidden and can be added back.
const KEY = "dh.dashboard.v2"
const PERIODS: Period[] = [7, 30, 90, 0]

const ids = (value: unknown) =>
  Array.isArray(value) ? [...new Set(value.filter((id): id is WidgetId => typeof id === "string" && Object.hasOwn(WIDGETS, id)))] : undefined

export function loadLayout(): Layout {
  const saved = readBrowserJSON(KEY)
  if (typeof saved !== "object" || saved === null || Array.isArray(saved)) return DEFAULT_LAYOUT
  const period = "period" in saved && PERIODS.includes(saved.period as Period) ? (saved.period as Period) : DEFAULT_LAYOUT.period
  const arranged = "widgets" in saved && typeof saved.widgets === "object" && saved.widgets !== null ? (saved.widgets as Record<string, unknown>) : {}
  const widgets = Object.fromEntries(STAGES.flatMap((stage) => { const list = ids(arranged[stage]); return list ? [[stage, list]] : [] }))
  const stage = "stage" in saved && STAGES.includes(saved.stage as Stage) ? (saved.stage as Stage) : undefined
  return { widgets, period, stage, shown: ids("shown" in saved && saved.shown) }
}

export function saveLayout(layout: Layout) {
  writeBrowserValue(KEY, JSON.stringify(layout))
}
