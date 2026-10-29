import { expect, test } from "bun:test"
import { checkPlan } from "../lib/dashboard-plan"
import dashboardPlan from "../tool/dashboard-plan"

const chart = { type: "chart", title: "Spend", metric: "spend", groupBy: "day", form: "line", period: "month" }

test("the tool returns only the checked spec, never model-supplied figures", async () => {
  const result = await dashboardPlan.execute({ widgets: [{ ...chart, title: " Spend ", values: [999], total: 999 }], unsupported: [] }, {} as never)
  expect(result.metadata.plan).toEqual({ widgets: [chart], unsupported: [] })
  expect(JSON.stringify(result.metadata.plan)).not.toContain("999")
})

test("invalid chart geometry returns actionable errors through the tool", async () => {
  await expect(dashboardPlan.execute({ widgets: [{ ...chart, groupBy: "client" }] }, {} as never)).rejects.toThrow('a line runs over time')
  expect(checkPlan({ widgets: [{ ...chart, form: "stat" }] }).errors?.join(" ")).toContain('groupBy must be "none"')
  expect(checkPlan({ widgets: [{ ...chart, form: "calendar", groupBy: "month" }] }).errors?.join(" ")).toContain('a calendar has one cell per day')
})

test("draft totals cannot claim dates and charts cannot silently repeat", () => {
  const draft = { ...chart, metric: "drafts", groupBy: "client", form: "bars" }
  expect(checkPlan({ widgets: [draft] }).errors?.join(" ")).toContain('period must be "all"')
  expect(checkPlan({ widgets: [{ ...draft, period: "all" }] }).plan?.widgets).toHaveLength(1)
  expect(checkPlan({ widgets: [chart, { ...chart, title: "Other title" }] }).errors?.join(" ")).toContain("repeats widgets[0]")
  expect(checkPlan({ widgets: [{ type: "spend" }, { type: "spend" }] }).errors?.join(" ")).toContain("repeats widgets[0]")
})

test("unsupported and conflicting filters do not enter an accepted plan", () => {
  expect(checkPlan({ widgets: [{ ...chart, client: "Example LLP", matter: "Example matter" }] }).errors?.join(" ")).toContain("not both")
  expect(checkPlan({ widgets: [{ ...chart, limit: 5 }] }).errors?.join(" ")).toContain("limit applies only")
  expect(checkPlan({ widgets: [{ type: "revenue" }] }).errors?.join(" ")).toContain("type must be")
  expect(checkPlan({ widgets: [], unsupported: ["No data"] }).errors?.join(" ")).toContain("at least one")
})


test("chart sources preserve old plans and reject unapproved mappings", () => {
  const old = checkPlan({ widgets: [chart] })
  expect(old.plan?.widgets[0]).toEqual(chart)
  const source = { kind: "local-query", query: "practice-chart", version: 1 }
  expect(checkPlan({ widgets: [{ ...chart, source }] }).plan?.widgets[0]).toEqual({ ...chart, source })
  for (const source of [{ kind: "local-query", query: "other", version: 1 }, { kind: "local-query", query: "practice-chart", version: 2 }, { kind: "local-query", query: "practice-chart", version: 1, url: "/other" }]) {
    expect(checkPlan({ widgets: [{ ...chart, source }] }).errors).toBeDefined()
  }
})

test("source documents have no date groupings; token charts use session-start periods", () => {
  expect(checkPlan({ widgets: [{ ...chart, metric: "documents", period: "all", groupBy: "client", form: "bars" }] }).plan).toBeDefined()
  expect(checkPlan({ widgets: [{ ...chart, metric: "documents", period: "month" }] }).errors?.join(" ")).toContain("no reliable dates")
  expect(checkPlan({ widgets: [{ ...chart, metric: "tokens" }] }).plan).toBeDefined()
})
