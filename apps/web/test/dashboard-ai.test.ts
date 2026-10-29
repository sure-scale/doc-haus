import { expect, test } from "bun:test"
import type { SessionMessageInfo } from "@opencode/client"
import type { Plan } from "../../../dochaus/lib/dashboard-plan"
import { askDashboard, dashboardPrompt, planFromMessages, type DashboardRequest } from "../src/api/opencode/dashboard"

const plan: Plan = { widgets: [{ type: "spend" }], unsupported: [] }
const input: DashboardRequest = { request: "Show AI spend", current: [{ widget: { type: "spend" }, pinned: true }], clients: ["Example LLP"], matters: ["Example matter"] }
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

async function nativeJob(outcome: "succeeded" | "failed", chosen = plan) {
  const original = globalThis.fetch
  const requests: Request[] = []
  let created: Record<string, unknown> | undefined
  globalThis.fetch = (async (source: string | URL | Request, init?: RequestInit) => {
    const request = source instanceof Request ? source : new Request(source, init)
    requests.push(request.clone())
    const pathname = new URL(request.url).pathname
    if (pathname === "/engine/config") return Response.json({ model: "provider/primary", small_model: "provider/fast" })
    if (pathname === "/dashboard/directory") return Response.json({ dir: "/synthetic-workspace/.dashboard" })
    if (pathname === "/api/session" && request.method === "POST") {
      created = await request.json() as Record<string, unknown>
      return Response.json({ data: { id: created.id } })
    }
    if (pathname.endsWith("/message")) return Response.json({ data: [assistant(tool(chosen)), { type: "idle", outcome }], cursor: {} })
    if (pathname.endsWith("/interrupt")) return Response.json({ interrupted: true })
    if (pathname.endsWith("/prompt")) return Response.json({ data: { id: "msg_fixture" } })
    if (pathname.endsWith("/wait") || request.method === "DELETE") return new Response(null, { status: 204 })
    throw new Error(`Unexpected fixture request: ${request.method} ${pathname}`)
  }) as typeof fetch
  try {
    let result: unknown
    let error: unknown
    try { result = await askDashboard(input) } catch (caught) { error = caught }
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
