import { expect, test } from "bun:test"
import type { SessionMessageInfo } from "@opencode/client"
import type { Plan } from "../../../dochaus/lib/dashboard-plan"
import { askDashboard, dashboardPrompt, planFromMessages, type DashboardRequest } from "../src/api/opencode/dashboard"

const plan: Plan = { widgets: [{ type: "spend" }], unsupported: [] }
const input: DashboardRequest = { request: "Show AI spend", current: [{ widget: { type: "matter-cost" }, pinned: true }], clients: ["Example LLP"], matters: ["Example matter"] }
const tool = (value: unknown, status = "completed") => ({ type: "tool", name: "dashboard-plan", state: { status, metadata: { plan: value } } })
const assistant = (...content: unknown[]) => ({ type: "assistant", content }) as SessionMessageInfo

test("reads the latest checked completed plan and never prose or raw tool input", () => {
  expect(planFromMessages([assistant(tool(plan)), assistant({ type: "text", text: JSON.stringify({ widgets: [{ type: "tokens" }] }) })])).toEqual(plan)
  expect(() => planFromMessages([assistant(tool(plan, "running"))])).toThrow("did not finish")
  const invalid = assistant(tool({ widgets: [{ type: "revenue" }] }))
  expect(() => planFromMessages([invalid])).toThrow("could not be shown")
  expect(planFromMessages([assistant(tool(plan)), assistant(tool({ widgets: [{ type: "tokens" }], unsupported: [] }))]).widgets).toEqual([{ type: "tokens" }])
})

test("the prompt carries settings, exact names and pins, with no practice figures", () => {
  const prompt = dashboardPrompt(input)
  expect(prompt).toContain('"pinned":true')
  expect(prompt).toContain('"Example LLP"')
  expect(prompt).toContain("dashboard-chart")
  expect(prompt).toContain("catalog")
})

async function nativeJob(outcome: "succeeded" | "failed", chosen = plan, options: { usable?: boolean; preview?: boolean; request?: DashboardRequest } = {}) {
  const original = globalThis.fetch
  const requests: Request[] = []
  const period = "month"
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone
  const discovery = (action: string) => ({ type: "tool", name: "dashboard-data", state: { status: "completed", metadata: { action, period, timeZone, ...(action === "preview" ? { widgets: chosen.widgets } : {}) } } })
  let created: Record<string, unknown> | undefined
  globalThis.fetch = (async (source: string | URL | Request, init?: RequestInit) => {
    const request = source instanceof Request ? source : new Request(source, init)
    requests.push(request.clone())
    const pathname = new URL(request.url).pathname
    if (pathname === "/engine/config") return Response.json({ model: "provider/primary", small_model: "provider/fast" })
    if (pathname === "/dashboard/discovery") return Response.json({ results: chosen.widgets.map((_, index) => ({ index, status: options.usable === false ? "empty" : "available", usable: options.usable !== false, dataBacked: options.usable !== false })) })
    if (pathname === "/dashboard/directory") return Response.json({ dir: "/synthetic-workspace/.dashboard" })
    if (pathname === "/api/session" && request.method === "POST") {
      created = await request.json() as Record<string, unknown>
      return Response.json({ data: { id: created.id } })
    }
    if (pathname.endsWith("/message")) return Response.json({ data: [assistant(discovery("inventory"), ...(options.preview === false ? [] : [discovery("preview")]), tool(chosen)), { type: "idle", outcome }], cursor: {} })
    if (pathname.endsWith("/interrupt")) return Response.json({ interrupted: true })
    if (pathname.endsWith("/prompt")) return Response.json({ data: { id: "msg_fixture" } })
    if (pathname.endsWith("/wait") || request.method === "DELETE") return new Response(null, { status: 204 })
    throw new Error(`Unexpected fixture request: ${request.method} ${pathname}`)
  }) as typeof fetch
  try {
    let result: unknown
    let error: unknown
    try { result = await askDashboard(options.request ?? input) } catch (caught) { error = caught }
    return { result, error, created, requests }
  } finally { globalThis.fetch = original }
}

test("uses the selected fast model, isolated directory and native cleanup", async () => {
  const { result, error, created, requests } = await nativeJob("succeeded")
  expect(error).toBeUndefined()
  expect(result).toEqual(plan)
  expect(created).toMatchObject({ agent: "dashboard", model: { providerID: "provider", id: "fast" }, metadata: { dochaus_role: "dashboard" }, location: { directory: "/synthetic-workspace/.dashboard" } })
  const native = requests.filter((request) => new URL(request.url).pathname.startsWith("/api/"))
  expect(native.every((request) => request.headers.get("x-opencode-directory") === "/synthetic-workspace/.dashboard")).toBe(true)
  expect(requests.some((request) => request.method === "DELETE" && request.url.includes(String(created?.id)))).toBe(true)
})

test("a failed native turn never applies an earlier accepted plan and still cleans up", async () => {
  const { result, error, requests } = await nativeJob("failed")
  expect(result).toBeUndefined()
  expect(String(error)).toContain("failed")
  expect(requests.some((request) => request.method === "DELETE")).toBe(true)
})


test("empty data and a missing native preview cannot create a dashboard", async () => {
  const empty = await nativeJob("succeeded", plan, { usable: false })
  expect(String(empty.error)).toContain("No usable data")
  expect(empty.result).toBeUndefined()
  expect(empty.requests.some((request) => request.method === "DELETE")).toBe(true)
  const unpreviewed = await nativeJob("succeeded", plan, { preview: false })
  expect(String(unpreviewed.error)).toContain("could not check its proposed views")
  expect(unpreviewed.result).toBeUndefined()
})

test("validation preserves old empty widgets, permits reorder, and requires data-backed content", async () => {
  const { validateDashboardData } = await import("../src/api/opencode/dashboard")
  const original = globalThis.fetch
  const existing: Plan["widgets"] = [{ type: "drafts" }, { type: "spend" }]
  let results = [
    { index: 0, status: "empty", usable: false, dataBacked: false },
    { index: 1, status: "available", usable: true, dataBacked: true },
  ]
  globalThis.fetch = (async () => Response.json({ results })) as unknown as typeof fetch
  try {
    await expect(validateDashboardData(existing, [], "month", "UTC")).resolves.toBeUndefined()
    await expect(validateDashboardData(existing, [existing[0]!], "month", "UTC")).rejects.toThrow("No usable data")
    results = results.map((result) => ({ ...result, usable: true, dataBacked: false }))
    await expect(validateDashboardData([{ type: "actions" }, { type: "ideas" }], [], "month", "UTC")).rejects.toThrow("No usable data")
  } finally { globalThis.fetch = original }
})


test("accepts a discovered client whose saved name has surrounding whitespace", async () => {
  const chosen: Plan = { widgets: [{ type: "chart", title: "Client matters", metric: "matters", groupBy: "none", form: "stat", period: "all", client: "Example LLP" }], unsupported: [] }
  const { error, result } = await nativeJob("succeeded", chosen, { request: { ...input, clients: [" Example LLP "] } })
  expect(error).toBeUndefined()
  expect(result).toEqual(chosen)
})
