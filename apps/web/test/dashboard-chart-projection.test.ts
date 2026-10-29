import { expect, test } from "bun:test"
import type { Chart } from "../../../dochaus/lib/dashboard-plan"
import type { DashboardData, MatterActivity } from "../src/components/dashboard/data"
import { chartProjection } from "../src/components/dashboard/chart-widget"
import { validateComponentData } from "../src/components/dashboard/visual/schema"

const chart: Chart = { type: "chart", title: "Clients", metric: "matters", groupBy: "client", form: "bars", period: "all" }
test("client labels without usable IDs produce valid chart component data", () => {
  const rows = [undefined, "A".repeat(160)].map((client, index) => ({ matter: { id: `matter-${index}`, title: "Matter", client, dir: `/synthetic/matter-${index}`, created_at: Date.now() }, sessions: [], costs: [], unread: false, unpriced: 0 } as MatterActivity))
  const projection = chartProjection(chart, { matters: rows } as DashboardData, "all")
  const checked = validateComponentData("chart", { form: "bars", label: "Clients", format: "count", showShares: true, ...projection })
  expect(checked.error).toBeUndefined()
  expect(checked.data?.points.map((point) => point.label)).toEqual(["A".repeat(160), "No client recorded"])
})
