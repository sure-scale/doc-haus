import { readBrowserJSON, writeBrowserValue } from "../../browser-storage"
import type { Period } from "./data"

// tile: one column; tall: one column, two rows; wide: two columns, two rows;
// full: the whole row, as tall as its content.
export type WidgetSize = "tile" | "tall" | "wide" | "full"

export const WIDGETS = {
  setup: { title: "Finish setting up", description: "What is left to set up, with a shortcut to each setting.", size: "full" },
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
  library: { title: "Your library", description: "Templates, playbooks, clients and matters.", size: "wide" },
} as const satisfies Record<string, { title: string; description: string; size: WidgetSize }>

export type WidgetId = keyof typeof WIDGETS

export const DEFAULT_WIDGETS: WidgetId[] = ["setup", "spend", "matter-cost", "draft-cost", "documents", "week", "matter-spend", "attention", "activity", "recent", "tokens"]

export type Layout = { widgets: WidgetId[]; period: Period }

export const DEFAULT_LAYOUT: Layout = { widgets: DEFAULT_WIDGETS, period: 30 }

// Per browser, like the sidebar's collapsed state: which widgets show, in
// what order. A widget not in the list is hidden and can be added back.
const KEY = "dh.dashboard.v1"
const PERIODS: Period[] = [7, 30, 90, 0]

export function loadLayout(): Layout {
  const saved = readBrowserJSON(KEY)
  if (typeof saved !== "object" || saved === null || Array.isArray(saved)) return DEFAULT_LAYOUT
  const widgets = "widgets" in saved && Array.isArray(saved.widgets)
    ? [...new Set(saved.widgets.filter((id): id is WidgetId => typeof id === "string" && Object.hasOwn(WIDGETS, id)))]
    : DEFAULT_WIDGETS
  const period = "period" in saved && PERIODS.includes(saved.period as Period) ? (saved.period as Period) : DEFAULT_LAYOUT.period
  return { widgets, period }
}

export function saveLayout(layout: Layout) {
  writeBrowserValue(KEY, JSON.stringify(layout))
}
