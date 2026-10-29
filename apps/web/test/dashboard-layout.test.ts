import { expect, test } from "bun:test"
import type { Chart, Plan } from "../../../dochaus/lib/dashboard-plan"
import { applyPlan, instance, readDashboards, tabWidgets, WIDGETS, type WidgetId, type Layout } from "../src/components/dashboard/layout"

const chart: Chart = { type: "chart", title: "Spend by client", metric: "spend", groupBy: "client", form: "bars", period: "dashboard" }
const plan = (widgets: Plan["widgets"]): Plan => ({ widgets, unsupported: [] })

test("an AI replacement preserves pinned widget settings and positions", () => {
  const current = [instance({ type: "spend" }), instance(chart, true), instance({ type: "drafts" }), instance({ type: "week" }, true)]
  const result = applyPlan(current, plan([{ type: "activity" }, { ...chart, title: "A new title" }, { type: "trend" }, { type: "recent" }]))
  expect(result.map((item) => item.widget.type)).toEqual(["activity", "chart", "trend", "week", "recent"])
  expect(result[1]).toEqual(current[1])
  expect(result[3]).toEqual(current[3])
})

test("pins omitted by the model remain when the replacement is shorter", () => {
  const current = [instance({ type: "spend" }), instance(chart, true)]
  const result = applyPlan(current, plan([{ type: "attention" }]))
  expect(result).toEqual([instance({ type: "attention" }), current[1]!])
})

test("the widget limit counts retained pins as well as new widgets", () => {
  const pins = [instance({ type: "setup" }, true), instance({ type: "journey" }, true)]
  const templates = ["spend", "matter-cost", "draft-cost", "week", "recent", "matter-spend", "activity", "trend", "tokens", "sample", "ideas", "actions", "library", "attention", "drafts"] as const
  const result = applyPlan(pins, plan(templates.map((type) => ({ type }))))
  expect(result).toHaveLength(16)
  expect(result.slice(0, 2)).toEqual(pins)
})

test("v3 migration retains both stages and a fully empty arrangement", () => {
  const legacy: Layout = { period: "quarter", stages: { starting: { hidden: ["journey", "setup"], added: [] }, practice: { order: ["week", "spend"], hidden: ["attention", "drafts", "setup", "matter-cost", "draft-cost", "recent", "matter-spend", "activity"], added: [] } } }
  const saved = readDashboards(null, legacy)
  const overview = saved.dashboards[0]!
  expect(saved.active).toBe("overview")
  expect(overview.period).toBe("quarter")
  expect(tabWidgets(overview, "starting", false)).toEqual([])
  expect(tabWidgets(overview, "practice", false).map((item) => item.widget.type)).toEqual(["week", "spend"])
})

test("a stored chart is checked again and retains its pin after reload", () => {
  const current = { active: "custom", dashboards: [{ id: "custom", name: "My work", period: "all", widgets: [instance(chart, true)], stages: {} }] }
  const saved = readDashboards(JSON.parse(JSON.stringify(current)))
  expect(saved).toMatchObject(current)
})

test("broken browser state falls back to a usable dashboard", () => {
  const saved = readDashboards({ active: "missing", dashboards: [{ id: "a", name: "Work", period: "yesterday", widgets: [{ widget: { type: "chart", title: "Wrong", metric: "hours" } }], stages: {} }, { id: "a", name: "Duplicate" }] })
  expect(saved.active).toBe("a")
  expect(saved.dashboards).toHaveLength(1)
  expect(saved.dashboards[0]!.period).toBe("month")
  expect(tabWidgets(saved.dashboards[0]!, "starting", false).map((item) => item.widget.type)).toEqual(["journey", "setup"])
})


test("all 17 legacy templates survive a pin and reload; pins do not expand an AI plan", () => {
  const widgets = (Object.keys(WIDGETS) as WidgetId[]).map((type) => instance({ type }, true))
  const saved = readDashboards({ active: "overview", dashboards: [{ id: "overview", name: "Overview", period: "month", stages: { starting: widgets } }] })
  expect(saved.dashboards[0]!.stages.starting).toHaveLength(17)
  expect(applyPlan(widgets, plan([{ type: "spend" }]))).toEqual(widgets)
})
