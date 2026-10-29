import { expect, test } from "bun:test"
import { isComponentData, isLocalDashboardHref, validateComponentData, type ComponentDataMap, type ComponentKind } from "../src/components/dashboard/visual/schema"

const cases: ComponentDataMap = {
  stat: { value: "−$12.00", caption: "Net change", tag: "vs prior period" },
  attention: { rows: [{ id: "question", title: "A question is waiting", detail: "Example matter", icon: "question", tone: "info", action: { label: "Answer", href: "/matter/example?session=one" } }] },
  matters: { rows: [{ id: "matter", title: "Example matter", detail: "Example client", href: "/matter/example", updated: "Today", spend: "$3.00", status: { icon: "check", label: "Drafted", tone: "good" } }] },
  library: { cards: [{ id: "templates", icon: "template", label: "Templates", count: 3, detail: "Firm precedents", emptyLabel: "Add your first", href: "/templates" }] },
  sample: { items: [{ id: "sources", icon: "documents", title: "Source files", detail: "Engagement letter" }] },
  actions: { items: [{ id: "new", icon: "matter", title: "New matter", detail: "Open a matter", href: "/?new=1" }, { id: "sample", icon: "sample", title: "Try a sample", detail: "Explore fictional material" }] },
  tokens: { segments: [{ key: "input", label: "Input", value: 100 }, { key: "cached", label: "Cached input", value: 200 }], total: 300, period: "This month" },
  chart: { form: "bars", label: "Spend", format: "money", currency: "GBP", showShares: true, total: 10, points: [{ key: "a", label: "Client A", value: 3 }] },
  setup: { config: { title: "Setup", description: "Connect your tools", items: [{ id: "ai", title: "AI", detail: "Connected", icon: "check", done: true, action: { id: "ai", label: "Change", variant: "text" } }] } },
  demo: { config: { title: "Demo", description: "Explore a sample", note: "Fictional data", features: [{ label: "Sample matter", value: "1" }], action: { id: "populate", label: "Populate", icon: "plus" } }, busy: false, disabled: false },
  tabs: { items: [{ id: "overview", label: "Overview", badge: { value: 3 } }, { id: "setup", label: "Setup", badge: { value: 1, total: 5 } }], activeId: "overview" },
  about: { config: { title: "About this view", text: "Figures come from your practice data", actionLabel: "Refine view" } },
  composer: { open: true, request: "Show spend by client", config: { title: "New dashboard", suggestions: ["Spend by client", "Weekly overview"] } },
}

test("each component accepts its JSON roundtrip and returns a typed view", () => {
  for (const kind of Object.keys(cases) as ComponentKind[]) {
    const value: unknown = JSON.parse(JSON.stringify(cases[kind]))
    expect(validateComponentData(kind, value).data).toEqual(cases[kind])
    expect(isComponentData(kind, value)).toBe(true)
  }
  const value: unknown = cases.chart
  if (isComponentData("chart", value)) expect(value.points[0]?.value).toBe(3)
})

test("missing structure, executable values and unknown icons are refused", () => {
  expect(validateComponentData("unknown", {}).error).toContain("Unknown")
  for (const kind of Object.keys(cases) as ComponentKind[]) expect(validateComponentData(kind, {}).error).toBeDefined()
  expect(isComponentData("stat", { value: "$1", caption: "Spend", run: () => 1 })).toBe(false)
  expect(isComponentData("sample", { items: [{ ...cases.sample.items[0], icon: "<svg>" }] })).toBe(false)
  expect(isComponentData("attention", { rows: [{ ...cases.attention.rows[0], action: {} }] })).toBe(false)
})

test("links accept local root paths and reject remote, protocol, slash and control tricks", () => {
  expect(isLocalDashboardHref("/matter/one?view=chat#latest")).toBe(true)
  for (const href of ["javascript:alert(1)", "https://example.com", "//example.com", "/\\example.com", "/\n/example.com", " /matters", "matter/one"]) {
    expect(isLocalDashboardHref(href)).toBe(false)
    expect(isComponentData("attention", { rows: [{ ...cases.attention.rows[0], action: { label: "Open", href } }] })).toBe(false)
    expect(isComponentData("matters", { rows: [{ ...cases.matters.rows[0], href }] })).toBe(false)
    expect(isComponentData("library", { cards: [{ ...cases.library.cards[0], href }] })).toBe(false)
  }
})

test("row IDs are unique and library counts are finite nonnegative integers or unknown", () => {
  expect(isComponentData("attention", { rows: [...cases.attention.rows, ...cases.attention.rows] })).toBe(false)
  expect(isComponentData("setup", { config: { ...cases.setup.config, items: [...cases.setup.config.items, ...cases.setup.config.items] } })).toBe(false)
  for (const count of [-1, 1.5, NaN, Infinity, undefined]) expect(isComponentData("library", { cards: [{ ...cases.library.cards[0], count }] })).toBe(false)
  expect(isComponentData("library", { cards: [{ ...cases.library.cards[0], count: null }] })).toBe(true)
  expect(isComponentData("sample", { items: Array.from({ length: 201 }, (_, id) => ({ ...cases.sample.items[0], id: String(id) })) })).toBe(false)
})

test("series reject negative, nonfinite and overflowing values but signed stats work", () => {
  for (const form of ["line", "bars", "list", "calendar"]) expect(isComponentData("chart", { form, label: "Bad", points: [{ label: "A", value: -1, time: 0 }] })).toBe(false)
  for (const value of [NaN, Infinity, -Infinity]) expect(isComponentData("chart", { form: "stat", label: "Bad", points: [{ label: "A", value }] })).toBe(false)
  expect(isComponentData("chart", { form: "stat", label: "Change", total: -2, points: [{ label: "A", value: -2 }] })).toBe(true)
  expect(isComponentData("chart", { form: "bars", label: "Huge", points: [{ label: "A", value: 1e308 }, { label: "B", value: 1e308 }] })).toBe(false)
  expect(isComponentData("chart", { form: "bars", label: "Huge", points: Array.from({ length: 5001 }, () => ({ label: "A", value: 0 })) })).toBe(false)
})

test("chart identities, dates, currency and percentage denominators stay meaningful", () => {
  expect(isComponentData("chart", { ...cases.chart, points: [{ key: "a", label: "A", value: 1 }, { key: "a", label: "B", value: 2 }] })).toBe(false)
  for (const time of [NaN, 9e15, "2026-10-29"]) expect(isComponentData("chart", { form: "line", label: "Date", points: [{ label: "A", value: 1, time }] })).toBe(false)
  for (const currency of ["gbp", "US", "ZZZ", "invalid"]) expect(isComponentData("chart", { ...cases.chart, currency })).toBe(false)
  expect(isComponentData("chart", { ...cases.chart, total: 2 })).toBe(false)
  expect(isComponentData("chart", { ...cases.chart, total: 0 })).toBe(false)
  expect(isComponentData("chart", { ...cases.chart, total: 100 })).toBe(true)
  expect(isComponentData("chart", { ...cases.chart, total: 0.3, points: [{ label: "A", value: 0.1 }, { label: "B", value: 0.2 }] })).toBe(true)
})

test("calendars require one value per local day within ten years", () => {
  const morning = new Date(2026, 9, 29, 9).getTime()
  const evening = new Date(2026, 9, 29, 20).getTime()
  const calendar = { form: "calendar", label: "Activity", points: [{ label: "Today", value: 1, time: morning }] }
  expect(isComponentData("chart", calendar)).toBe(true)
  expect(isComponentData("chart", { ...calendar, points: [{ label: "A", value: 1 }] })).toBe(false)
  expect(isComponentData("chart", { ...calendar, points: [...calendar.points, { label: "B", value: 2, time: evening }] })).toBe(false)
  expect(isComponentData("chart", { ...calendar, points: [...calendar.points, { label: "B", value: 2, time: new Date(2040, 1, 1).getTime() }] })).toBe(false)
})


test("tab counts are finite nonnegative integers with honest progress totals", () => {
  for (const value of [-1, 1.5, NaN, Infinity]) expect(isComponentData("tabs", { ...cases.tabs, items: [{ id: "overview", label: "Overview", badge: { value } }] })).toBe(false)
  for (const total of [0, -1, 1.5, Infinity, 2]) expect(isComponentData("tabs", { ...cases.tabs, items: [{ id: "overview", label: "Overview", badge: { value: 3, total } }] })).toBe(false)
  expect(isComponentData("tabs", { ...cases.tabs, activeId: "missing" })).toBe(false)
  expect(isComponentData("tabs", { ...cases.tabs, items: [...cases.tabs.items, ...cases.tabs.items] })).toBe(false)
  expect(isComponentData("tabs", { items: [], activeId: "" })).toBe(true)
})

test("composer and about content remain plain structured data", () => {
  expect(isComponentData("composer", { open: true, request: "" })).toBe(true)
  expect(isComponentData("composer", { open: "yes", request: "A chart" })).toBe(false)
  expect(isComponentData("composer", { ...cases.composer, config: { suggestions: ["Repeated", "Repeated"] } })).toBe(false)
  expect(isComponentData("composer", { ...cases.composer, config: { suggestions: [42] } })).toBe(false)
  expect(isComponentData("about", { config: { text: "Hello", actionLabel: { code: "run" } } })).toBe(false)
})


test("JSON stat visuals use finite nonnegative series, valid averages and unique dated days", () => {
  for (const visual of [
    { type: "sparkline", values: [0, 2, 1], label: "Spend" },
    { type: "spread", values: [1, 2], average: 1.5, label: "Per matter", format: "money" },
    { type: "week", days: [{ day: new Date(2026, 9, 29).getTime(), cost: 2, label: "Thu", today: true }] },
  ]) expect(isComponentData("stat", JSON.parse(JSON.stringify({ ...cases.stat, visual })))).toBe(true)
  for (const visual of [
    { type: "sparkline", values: [-1], label: "Spend" },
    { type: "sparkline", values: [Infinity], label: "Spend" },
    { type: "spread", values: [1], average: -1, label: "Per matter" },
    { type: "spread", values: [1], average: NaN, label: "Per matter" },
    { type: "spread", values: [1], average: 1, label: "Per matter", format: "invalid" },
    { type: "week", days: [{ day: "2026-10-29", cost: 2, label: "Thu", today: true }] },
    { type: "week", days: [{ day: 0, cost: -2, label: "Thu", today: true }] },
    { type: "week", days: [{ day: 0, cost: 2, label: "Thu", today: true }, { day: 0, cost: 3, label: "Fri", today: false }] },
    { type: "week", days: Array.from({ length: 8 }, (_, index) => ({ day: index * 86400000, cost: 2, label: String(index), today: false })) },
  ]) expect(isComponentData("stat", { ...cases.stat, visual })).toBe(false)
})


test("action tiles validate local links, unique IDs and structured labels", () => {
  expect(isComponentData("actions", JSON.parse(JSON.stringify(cases.actions)))).toBe(true)
  for (const href of ["//example.com", "/\\example.com", "javascript:alert(1)"]) expect(isComponentData("actions", { items: [{ ...cases.actions.items[0], href }] })).toBe(false)
  expect(isComponentData("actions", { items: [...cases.actions.items, ...cases.actions.items] })).toBe(false)
  expect(isComponentData("actions", { items: [{ ...cases.actions.items[0], icon: "missing" }] })).toBe(false)
})

test("count charts reject fractional values and totals while money retains decimals", () => {
  for (const format of [undefined, "count"]) {
    const base = { form: "bars", label: "Conversations", points: [{ label: "A", value: 1 }], ...(format ? { format } : {}) }
    expect(isComponentData("chart", { ...base, total: 1.5 })).toBe(false)
    expect(isComponentData("chart", { ...base, points: [{ label: "A", value: 1.5 }] })).toBe(false)
  }
  expect(isComponentData("chart", { form: "stat", label: "Change", format: "count", points: [{ label: "Delta", value: -2 }], total: -2 })).toBe(true)
  expect(isComponentData("chart", { form: "bars", label: "Spend", format: "money", points: [{ label: "A", value: 1.5 }], total: 1.5 })).toBe(true)
})

test("token totals and shares use the same unique nonnegative integer segments", () => {
  expect(isComponentData("tokens", { segments: cases.tokens.segments })).toBe(true)
  expect(isComponentData("tokens", { ...cases.tokens, total: 301 })).toBe(false)
  expect(isComponentData("tokens", { ...cases.tokens, total: 299 })).toBe(false)
  expect(isComponentData("tokens", { segments: [...cases.tokens.segments, ...cases.tokens.segments] })).toBe(false)
  for (const value of [-1, 1.5, NaN, Infinity]) expect(isComponentData("tokens", { segments: [{ key: "input", label: "Input", value }] })).toBe(false)
  expect(isComponentData("tokens", { segments: Array.from({ length: 51 }, (_, index) => ({ key: String(index), label: "Input", value: 1 })) })).toBe(false)
})
