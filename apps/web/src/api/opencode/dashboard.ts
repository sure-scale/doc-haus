import type { SessionMessageInfo } from "@opencode/client"
import { chartKey, checkPlan, describeCatalog, type Plan, type PlanWidget } from "../../../../../dochaus/lib/dashboard-plan"
import { dashboardDirectory } from "../ingest/dashboard"
import { location, matterClient } from "./client"
import { getConfig, modelSpec } from "./config"
import { INGEST_URL } from "../../config"
import type { Period } from "../../components/dashboard/data"

export type DashboardRequest = {
  request: string
  current: Array<{ widget: PlanWidget; pinned: boolean }>
  clients: string[]
  matters: string[]
  period?: Period
  signal?: AbortSignal
}

// This is the protocol shared with dochaus/tool/dashboard-plan.ts: only a
// completed dashboard-plan tool's metadata.plan is a result. Prose and raw tool
// input cannot change a dashboard. Every figure is still computed in the web.
export function planFromMessages(messages: readonly SessionMessageInfo[]): Plan {
  for (const message of [...messages].reverse()) {
    if (message.type !== "assistant") continue
    for (const part of [...message.content].reverse()) {
      if (part.type !== "tool" || part.name !== "dashboard-plan" || part.state.status !== "completed") continue
      const result = checkPlan(part.state.metadata?.plan)
      if (result.errors) throw new Error(`The AI returned a dashboard that could not be shown. ${result.errors.join(" ")}`)
      return result.plan
    }
  }
  throw new Error("The AI did not finish a dashboard plan. Try again.")
}

export function dashboardPrompt(input: Omit<DashboardRequest, "signal">): string {
  return [
    "Build one dashboard for this request. Select widget settings and supported source mappings only; do not produce figures.",
    "Request (JSON string):", JSON.stringify(input.request.trim()),
    "Widget catalog:", describeCatalog(),
    "Current widgets with pinned flags (JSON data):", JSON.stringify(input.current),
    "Available client names (JSON data):", JSON.stringify([...new Set(input.clients)]),
    "Available matter names (JSON data):", JSON.stringify([...new Set(input.matters)]),
    `Data context: period ${input.period ?? "month"}; timeZone ${Intl.DateTimeFormat().resolvedOptions().timeZone}.`,
    "Use dashboard-data inventory first, then preview all proposed widgets with dashboard-data using that period and timeZone. Use task with subagent_type dashboard-chart for each new chart, then submit the complete plan with dashboard-plan.",
    "Pinned widgets must remain unchanged. Describe unsupported requests plainly; never substitute another metric without saying so.",
  ].join("\n\n")
}

// One bounded native agent job, on the lawyer's configured fast model. Running
// outside matter directories keeps it out of conversations and matter spend.
// Native remove recursively stops and deletes child sessions as well.
export async function askDashboard(input: DashboardRequest): Promise<Plan> {
  if (!input.request.trim()) throw new Error("Tell AI what you want from this dashboard.")
  const timeout = AbortSignal.timeout(180_000)
  const controller = new AbortController()
  const signal = AbortSignal.any([timeout, controller.signal, ...(input.signal ? [input.signal] : [])])
  signal.throwIfAborted()
  const [dir, config] = await Promise.all([dashboardDirectory(signal), getConfig()])
  signal.throwIfAborted()
  const model = modelSpec({ model: config.small_model ?? config.model })
  if (!model) throw new Error("Choose a fast model in Settings before asking AI for a dashboard.")
  const client = matterClient(dir)
  // Knowing the id before admission also lets cleanup remove a session when
  // the create response was lost after the engine accepted it.
  const sessionID = `ses_${crypto.randomUUID().replaceAll("-", "")}`
  try {
    await client.session.create({
      id: sessionID,
      title: "doc.haus dashboard",
      agent: "dashboard",
      model: { providerID: model.providerID, id: model.modelID },
      location: location(client),
      metadata: { dochaus_role: "dashboard" },
    }, { signal })
    await client.session.prompt({ sessionID, text: dashboardPrompt(input), delivery: "queue" }, { signal })
    await client.session.wait({ sessionID }, { signal })
    const messages: SessionMessageInfo[] = []
    let cursor: string | undefined
    do {
      const page = await client.message.list({ sessionID, limit: 200, ...(cursor ? { cursor } : { order: "asc" }) }, { signal })
      messages.push(...page.data)
      cursor = page.cursor.next ?? undefined
    } while (cursor)
    const terminal = messages.at(-1)
    if (terminal?.type !== "idle") throw new Error("The dashboard request did not finish. Try again.")
    if (terminal.outcome !== "succeeded") {
      const failed = messages.findLast((message) => message.type === "assistant" && message.error)
      throw new Error(failed?.type === "assistant" && failed.error ? failed.error.message :
        terminal.outcome === "interrupted" ? "The dashboard request was stopped." : "The dashboard request failed. Try again.")
    }
    const plan = planFromMessages(messages)
    const period = input.period ?? "month"
    const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone
    const newWidgets = plan.widgets.filter((widget) => !input.current.some((entry) => widgetIdentity(entry.widget) === widgetIdentity(widget)))
    const parts = messages.flatMap((message) => message.type === "assistant" ? message.content : [])
    const submission = parts.findLastIndex((part) => part.type === "tool" && part.name === "dashboard-plan" && part.state.status === "completed")
    const discoveries = parts.slice(0, submission).filter((part) => part.type === "tool" && part.name === "dashboard-data" && part.state.status === "completed").map((part) => part.type === "tool" && part.state.status === "completed" ? part.state.metadata : undefined).filter((metadata) => metadata?.period === period && metadata.timeZone === timeZone)
    const inventory = discoveries.findIndex((metadata) => metadata?.action === "inventory")
    if (inventory < 0) throw new Error("The dashboard could not inspect your data. Try again.")
    const previewed = new Set(discoveries.slice(inventory + 1).filter((metadata) => metadata?.action === "preview").flatMap((metadata) => {
      const result = checkPlan({ widgets: metadata?.widgets })
      return result.plan?.widgets.map(widgetIdentity) ?? []
    }))
    if (plan.widgets.some((widget) => !previewed.has(widgetIdentity(widget)))) throw new Error("The dashboard could not check its proposed views against your data. Try again.")
    await validateDashboardData(plan.widgets, newWidgets, period, timeZone, signal)
    for (const widget of plan.widgets) {
      if (widget.type !== "chart") continue
      if (widget.client && !input.clients.some((client) => client.trim() === widget.client)) throw new Error(`The AI chose an unknown client: ${widget.client}. Try again with a client shown in your matters.`)
      if (widget.matter && !input.matters.includes(widget.matter)) throw new Error(`The AI chose an unknown matter: ${widget.matter}. Try again with a matter shown in your list.`)
    }
    return plan
  } catch (error) {
    if (timeout.aborted) throw new Error("The dashboard request took too long. Try a smaller request.")
    if (input.signal?.aborted) throw new DOMException("The dashboard request was stopped.", "AbortError")
    throw error
  } finally {
    controller.abort()
    await client.session.interrupt({ sessionID }, { signal: AbortSignal.timeout(10_000) }).catch(() => {})
    await client.session.remove({ sessionID }, { signal: AbortSignal.timeout(10_000) }).catch(() => {})
  }
}

function widgetIdentity(widget: PlanWidget) { return widget.type === "chart" ? chartKey(widget) : widget.type }

export async function validateDashboardData(widgets: PlanWidget[], newWidgets: PlanWidget[], period: Period, timeZone: string, signal?: AbortSignal) {
  const response = await fetch(`${INGEST_URL}/dashboard/discovery`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ action: "validate", widgets, period, timeZone }), signal })
  if (!response.ok) throw new Error("Could not check dashboard data. Check your data sources and try again.")
  const value: unknown = await response.json()
  if (!value || typeof value !== "object" || !("results" in value) || !Array.isArray(value.results) || value.results.length !== widgets.length) throw new Error("Could not check dashboard data. Try again.")
  let hasData = false
  for (let index = 0; index < widgets.length; index++) {
    const result: unknown = value.results[index]
    if (!result || typeof result !== "object" || !("index" in result) || result.index !== index || !("usable" in result) || typeof result.usable !== "boolean" || !("dataBacked" in result) || typeof result.dataBacked !== "boolean") throw new Error("Could not check dashboard data. Try again.")
    if (result.usable && result.dataBacked) hasData = true
    if (!result.usable && newWidgets.some((widget) => widgetIdentity(widget) === widgetIdentity(widgets[index]!))) {
      const label = widgets[index]!.type === "chart" ? (widgets[index] as Extract<PlanWidget, { type: "chart" }>).title : widgets[index]!.type
      const reason = "reason" in result && typeof result.reason === "string" && result.reason.length <= 240 ? ` ${result.reason}` : ""
      throw new Error(`No usable data is available for ${label}.${reason} Add data or ask for a different view.`)
    }
  }
  if (!hasData) throw new Error("No usable data is available for this dashboard. Add data or ask for a different view.")
}
