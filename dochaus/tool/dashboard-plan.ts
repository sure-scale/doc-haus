import { tool } from "./legacy-definition"
import { checkPlan } from "../lib/dashboard-plan"

// Native tool completion carries the checked plan in metadata.plan. The web
// reads that field rather than trusting prose or the model's unvalidated input.
export default tool({
  description: "Submit a dashboard plan of ready-made widgets and chart settings. The app computes all figures. Returns specific errors for invalid settings; fix those before submitting again.",
  args: {
    widgets: tool.schema.array(tool.schema.object({}).passthrough()).describe("The complete ordered widget list. Each item has type (a ready-made widget id or chart); charts also need title, metric, groupBy, form, period and the supported local-query source mapping."),
    name: tool.schema.string().optional().describe("Optional short tab name."),
    unsupported: tool.schema.array(tool.schema.string()).optional().describe("Short plain sentences explaining requested data or charts that are unavailable."),
  },
  async execute(input) {
    const result = checkPlan(input)
    if (result.errors) throw new Error(result.errors.join("\n"))
    return { title: "Dashboard plan", output: `Dashboard plan accepted (${result.plan.widgets.length} widgets).`, metadata: { plan: result.plan } }
  },
})
