import { tool } from "./legacy-definition"
import { checkPlan } from "../lib/dashboard-plan"

// The model can inspect only the app's bounded inventory and checked queries.
export default tool({
  description: "Inspect actual dashboard dataset availability or preview proposed widgets before submitting a plan. Use only the period and timeZone supplied in the request.",
  args: {
    action: tool.schema.enum(["inventory", "preview"]),
    period: tool.schema.enum(["month", "last-month", "quarter", "all"]),
    timeZone: tool.schema.string().min(1).max(100),
    widgets: tool.schema.array(tool.schema.object({}).passthrough()).optional(),
  },
  async execute(input, context) {
    const base = process.env.INGEST_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_INGEST_PORT ?? 4501}`
    let widgets: unknown[] | undefined
    if (input.action === "preview") {
      const checked = checkPlan({ widgets: input.widgets })
      if (checked.errors) throw new Error(checked.errors.join("\n"))
      widgets = checked.plan.widgets
    }
    const response = await fetch(`${base}/dashboard/discovery`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: input.action, period: input.period, timeZone: input.timeZone, ...(widgets ? { widgets } : {}) }), signal: context.abort,
    })
    if (!response.ok) throw new Error("Could not inspect dashboard data. Check your data sources and try again.")
    const data: unknown = await response.json()
    return { title: input.action === "inventory" ? "Dashboard data inventory" : "Dashboard query preview", output: JSON.stringify(data), metadata: { action: input.action, period: input.period, timeZone: input.timeZone, ...(widgets ? { widgets } : {}) } }
  },
})
