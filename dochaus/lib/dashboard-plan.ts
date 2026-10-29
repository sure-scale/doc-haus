// The widgets an AI-built dashboard can hold and the check its plan passes.
// Shared by the dashboard-plan tool, which accepts a plan or returns these
// errors for the dashboard agent to fix, and by the web, which re-checks the
// plan it reads back and computes every figure itself. A plan names widgets
// and chart settings only; it never carries a number.

// The ready-made widgets, as the web's dashboard catalog shows them.
export const TEMPLATES = {
  journey: { title: "First steps", shows: "The steps from a new matter to a first draft. Only shown before the first draft finishes." },
  setup: { title: "Your practice", shows: "A form for the lawyer's name, firm, jurisdictions and practice areas. Only shown while one is missing." },
  attention: { title: "Needs you", shows: "Drafting runs waiting on an answer, a decision or a fix." },
  drafts: { title: "Recent drafts", shows: "Finished drafts and what each one needs next." },
  spend: { title: "Total spend", shows: "Estimated AI spend across the lawyer's matters in the period, with the change on the period before." },
  "matter-cost": { title: "Average matter cost", shows: "Average spend per matter that has used AI." },
  "draft-cost": { title: "Average draft cost", shows: "Average spend per drafted document, including review." },
  week: { title: "This week", shows: "Spend on each of the last seven days." },
  recent: { title: "Recent matters", shows: "The matters worked on last." },
  "matter-spend": { title: "Spend by matter", shows: "The matters and clients that spent the most in the period, with a CSV export." },
  activity: { title: "Activity", shows: "Conversations started each day over the last year." },
  trend: { title: "Spend over time", shows: "Daily spend across the period." },
  tokens: { title: "Where tokens go", shows: "Input, cached input, output and reasoning tokens in the period." },
  sample: { title: "Sample matter", shows: "A link to the fictional sample matter." },
  ideas: { title: "Things to try", shows: "Suggestions: ask, draft, review across documents or run a workflow." },
  actions: { title: "Quick actions", shows: "Buttons to start a matter or add a template, client or playbook." },
  library: { title: "Your library", shows: "Counts of templates, playbooks, clients and matters, with links." },
} as const

export type TemplateId = keyof typeof TEMPLATES

// The chart builder. Every chart counts one metric, grouped one way, over one
// period, drawn in one form.
export const METRICS = {
  spend: "Estimated AI spend, in US dollars, when it was spent.",
  conversations: "Conversations the lawyer started, counted on the day each started.",
  matters: "Matters, counted on the day each was opened.",
  drafts: "Finished drafted documents in the latest run per matter. Has no completion dates: use the all-time period and group by none, matter or client.",
  documents: "Ingested source documents in each matter. Has no reliable dates: use the all-time period and group by none, matter or client.",
  tokens: "Total input, cached input, output and reasoning tokens across sessions, counted on the day each session started.",
} as const

export const GROUPINGS = {
  none: "One total.",
  day: "Each day.",
  week: "Each week.",
  month: "Each month.",
  matter: "Each matter, largest first.",
  client: "Each client, largest first. Matters with no client form one group.",
} as const

export const FORMS = {
  stat: "One large figure. Needs grouping none.",
  bars: "Bars, one per group.",
  line: "A line over time. Needs grouping day, week or month.",
  list: "A ranked list. Needs grouping matter or client.",
  calendar: "A calendar with one cell per day. Needs grouping day.",
} as const

// "dashboard" follows the period picked above the dashboard.
export const CHART_PERIODS = ["dashboard", "month", "last-month", "quarter", "all"] as const

export type Metric = keyof typeof METRICS
export type Grouping = keyof typeof GROUPINGS
export type Form = keyof typeof FORMS
export type ChartPeriod = (typeof CHART_PERIODS)[number]

export type Chart = {
  type: "chart"
  source?: { kind: "local-query"; query: "practice-chart"; version: 1 }
  title: string
  metric: Metric
  groupBy: Grouping
  form: Form
  period: ChartPeriod
  // Narrows the chart to one client's matters or one matter, by exact name.
  client?: string
  matter?: string
  // How many groups a matter or client chart shows.
  limit?: number
}

export type PlanWidget = { type: TemplateId } | Chart

export type Plan = {
  // A short tab name for a dashboard that has none yet.
  name?: string
  widgets: PlanWidget[]
  // Parts of the request the dashboard cannot show, in plain words for the lawyer.
  unsupported: string[]
}

export const MAX_WIDGETS = 16
export const MAX_TITLE = 60
export const MAX_LIMIT = 12
export const MAX_NAME = 30
const MAX_NOTES = 5
const MAX_NOTE = 160

const TIME: Grouping[] = ["day", "week", "month"]
const RANKED: Grouping[] = ["matter", "client"]

const has = <T extends object>(table: T, key: unknown): key is keyof T => typeof key === "string" && Object.hasOwn(table, key)
const quoted = (values: readonly string[]) => values.map((value) => `"${value}"`).join(", ")

// Why one chart cannot be drawn, or nothing when it can.
function chartErrors(chart: Record<string, unknown>, at: string) {
  const errors: string[] = []
  if (chart.source !== undefined) {
    const source = chart.source
    if (!source || typeof source !== "object" || Array.isArray(source) || Object.keys(source).length !== 3 || (source as Record<string, unknown>).kind !== "local-query" || (source as Record<string, unknown>).query !== "practice-chart" || (source as Record<string, unknown>).version !== 1)
      errors.push(`${at}: source must be the supported local practice-chart query, version 1.`)
  }
  if (typeof chart.title !== "string" || !chart.title.trim()) errors.push(`${at}: give the chart a title.`)
  else if (chart.title.trim().length > MAX_TITLE) errors.push(`${at}: shorten the title to ${MAX_TITLE} characters or fewer.`)
  if (!has(METRICS, chart.metric)) errors.push(`${at}: metric must be one of ${quoted(Object.keys(METRICS))}.`)
  if (!has(GROUPINGS, chart.groupBy)) errors.push(`${at}: groupBy must be one of ${quoted(Object.keys(GROUPINGS))}.`)
  if (!has(FORMS, chart.form)) errors.push(`${at}: form must be one of ${quoted(Object.keys(FORMS))}.`)
  if (chart.period !== undefined && !CHART_PERIODS.includes(chart.period as ChartPeriod))
    errors.push(`${at}: period must be one of ${quoted(CHART_PERIODS)}.`)
  for (const key of ["client", "matter"] as const)
    if (chart[key] !== undefined && (typeof chart[key] !== "string" || !chart[key].trim())) errors.push(`${at}: ${key} must be a name, or left out.`)
  if (chart.client !== undefined && chart.matter !== undefined) errors.push(`${at}: narrow to a client or a matter, not both.`)
  if (errors.length) return errors
  const { metric, groupBy, form } = chart as Pick<Chart, "metric" | "groupBy" | "form">
  if (form === "stat" && groupBy !== "none") errors.push(`${at}: a stat shows one total, so groupBy must be "none".`)
  if (form !== "stat" && groupBy === "none") errors.push(`${at}: groupBy "none" is one total, so form must be "stat".`)
  if (form === "line" && !TIME.includes(groupBy)) errors.push(`${at}: a line runs over time, so groupBy must be "day", "week" or "month".`)
  if (form === "list" && !RANKED.includes(groupBy)) errors.push(`${at}: a list ranks matters or clients, so groupBy must be "matter" or "client".`)
  if (form === "calendar" && groupBy !== "day") errors.push(`${at}: a calendar has one cell per day, so groupBy must be "day".`)
  if (metric === "drafts" && chart.period !== "all") errors.push(`${at}: drafts have no completion dates, so period must be "all".`)
  if (metric === "drafts" && TIME.includes(groupBy)) errors.push(`${at}: drafts have no dates, so group them by "matter" or "client", or use "none".`)
  if (metric === "documents" && chart.period !== "all") errors.push(`${at}: source documents have no reliable dates, so period must be "all".`)
  if (metric === "documents" && TIME.includes(groupBy)) errors.push(`${at}: source documents have no reliable dates, so group them by "matter" or "client", or use "none".`)
  if (groupBy === "matter" && chart.matter !== undefined) errors.push(`${at}: a chart of one matter cannot group by matter.`)
  if (chart.limit !== undefined) {
    if (!RANKED.includes(groupBy)) errors.push(`${at}: limit applies only when grouping by matter or client.`)
    else if (!Number.isInteger(chart.limit) || (chart.limit as number) < 1 || (chart.limit as number) > MAX_LIMIT)
      errors.push(`${at}: limit must be a whole number from 1 to ${MAX_LIMIT}.`)
  }
  return errors
}

// The same chart twice, whatever its title.
export function chartKey(chart: Chart) {
  return ["chart", chart.metric, chart.groupBy, chart.form, chart.period, chart.client ?? "", chart.matter ?? "", chart.limit ?? ""].join("|")
}

// Accepts a plan as the model wrote it, or lists every reason it cannot be
// shown, each naming the widget it is about.
export function checkPlan(input: unknown): { plan: Plan; errors?: undefined } | { plan?: undefined; errors: string[] } {
  if (typeof input !== "object" || input === null || Array.isArray(input)) return { errors: ["The plan must be an object with a widgets list."] }
  const value = input as Record<string, unknown>
  if (!Array.isArray(value.widgets) || value.widgets.length === 0) return { errors: ["widgets must list at least one widget."] }
  const errors: string[] = []
  if (value.widgets.length > MAX_WIDGETS) errors.push(`Keep the dashboard to ${MAX_WIDGETS} widgets or fewer.`)
  if (value.name !== undefined && (typeof value.name !== "string" || !value.name.trim() || value.name.trim().length > MAX_NAME))
    errors.push(`name must be a short tab name of ${MAX_NAME} characters or fewer, or left out.`)
  const notes = value.unsupported ?? []
  if (!Array.isArray(notes) || notes.some((note) => typeof note !== "string" || !note.trim() || note.length > MAX_NOTE) || notes.length > MAX_NOTES)
    errors.push(`unsupported must list up to ${MAX_NOTES} short sentences of ${MAX_NOTE} characters or fewer.`)
  const widgets: PlanWidget[] = []
  const seen = new Map<string, number>()
  value.widgets.forEach((entry, index) => {
    const at = `widgets[${index}]`
    if (typeof entry !== "object" || entry === null || Array.isArray(entry)) return errors.push(`${at}: each widget must be an object with a type.`)
    const widget = entry as Record<string, unknown>
    if (widget.type !== "chart" && !has(TEMPLATES, widget.type))
      return errors.push(`${at}: type must be "chart" or one of ${quoted(Object.keys(TEMPLATES))}.`)
    let next: PlanWidget
    if (widget.type === "chart") {
      const problems = chartErrors(widget, at)
      if (problems.length) return errors.push(...problems)
      const chart = widget as unknown as Chart
      next = {
        type: "chart",
        ...(chart.source ? { source: { kind: "local-query" as const, query: "practice-chart" as const, version: 1 as const } } : {}),
        title: chart.title.trim(),
        metric: chart.metric,
        groupBy: chart.groupBy,
        form: chart.form,
        period: chart.period ?? "dashboard",
        ...(chart.client !== undefined ? { client: chart.client.trim() } : {}),
        ...(chart.matter !== undefined ? { matter: chart.matter.trim() } : {}),
        ...(chart.limit !== undefined ? { limit: chart.limit } : {}),
      }
    } else next = { type: widget.type as TemplateId }
    const key = next.type === "chart" ? chartKey(next) : next.type
    const first = seen.get(key)
    if (first !== undefined) return errors.push(`${at}: repeats widgets[${first}]; show each widget once.`)
    seen.set(key, index)
    widgets.push(next)
  })
  if (errors.length) return { errors }
  const name = typeof value.name === "string" ? value.name.trim() : undefined
  return { plan: { ...(name ? { name } : {}), widgets, unsupported: (notes as string[]).map((note) => note.trim()) } }
}

// The catalog in words, for the dashboard agent's request.
export function describeCatalog() {
  const lines = (table: Record<string, string>) => Object.entries(table).map(([key, text]) => `- ${key}: ${text}`)
  return [
    "Ready-made widgets (type: what it shows):",
    ...Object.entries(TEMPLATES).map(([id, template]) => `- ${id}: ${template.title}. ${template.shows}`),
    "",
    `Charts (type "chart") map to source {"kind":"local-query","query":"practice-chart","version":1}; take a title of up to ${MAX_TITLE} characters, a metric, a groupBy, a form and a period, and may narrow to one client or one matter by exact name.`,
    "Metrics:",
    ...lines(METRICS),
    "Groupings (groupBy):",
    ...lines(GROUPINGS),
    "Forms:",
    ...lines(FORMS),
    `Periods: ${quoted(CHART_PERIODS)}. "dashboard" follows the period picked above the dashboard; the others are this month, last month, the last three months and all time.`,
    `A matter or client chart may set limit, from 1 to ${MAX_LIMIT} groups.`,
    `A dashboard holds 1 to ${MAX_WIDGETS} widgets, each once.`,
  ].join("\n")
}
