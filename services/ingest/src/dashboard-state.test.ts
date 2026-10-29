import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { existsSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { chartKey, TEMPLATES, type Chart } from "../../../dochaus/lib/dashboard-plan"
import { dashboardRoutes } from "./dashboard"
import { DASHBOARD_STATE_MAX_BYTES } from "./dashboard-state"

const previous = process.env.WORKSPACE_ROOT
const directories: string[] = []
afterEach(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
})
function workspace() {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-dashboard-state-"))
  directories.push(directory)
  process.env.WORKSPACE_ROOT = directory
  return directory
}
function state() {
  const chart: Chart = { type: "chart", title: "Conversations by client", metric: "conversations", groupBy: "client", form: "bars", period: "all" }
  const templates = Object.keys(TEMPLATES).map(type => ({ key: type, widget: { type }, pinned: false }))
  return { version: 1, workspace: { active: "custom", dashboards: [
    { id: "overview", name: "Overview", period: "month", stages: { starting: templates }, legacy: { practice: { hidden: ["setup"], added: [], order: ["attention", "drafts"] } } },
    { id: "custom", name: "Client activity", period: "all", stages: {}, widgets: [{ key: chartKey(chart), widget: chart, pinned: true }], notes: ["Revenue is not tracked."] },
  ] }, setupHidden: false, view: "setup" }
}
function put(app: ReturnType<typeof dashboardRoutes>, value: unknown, revision: number) {
  return app.request("/dashboard/tabs", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ state: value, revision }) })
}

test("saved dashboard layouts, visibility and selected view survive reopening the workspace database", async () => {
  const directory = workspace(), app = dashboardRoutes()
  const empty = await app.request("/dashboard/tabs")
  expect(empty.status).toBe(200)
  expect(empty.headers.get("cache-control")).toBe("no-store")
  expect(await empty.json()).toEqual({ state: null, revision: 0 })
  const saved = state()
  const response = await put(app, saved, 0)
  expect(response.status).toBe(200)
  expect(await response.json()).toEqual({ state: saved, revision: 1 })
  // Each route call closes its connection. A separately built router reads the
  // saved SQLite data, including the older 17-widget catalog and pinned chart.
  expect(await (await dashboardRoutes().request("/dashboard/tabs")).json()).toEqual({ state: saved, revision: 1 })
  const file = path.join(directory, ".dashboard", "dashboard.db")
  expect(existsSync(file)).toBe(true)
  const db = new Database(file, { readonly: true })
  try { expect((db.query("SELECT revision FROM dashboard_tabs WHERE id = 1").get() as { revision: number }).revision).toBe(1) }
  finally { db.close() }
})

test("a new browser cannot reseed existing tabs and a stale browser cannot overwrite newer edits", async () => {
  workspace(); const app = dashboardRoutes(), original = state()
  const first = await put(app, original, 0)
  expect(first.status).toBe(200)
  const reseed = await put(dashboardRoutes(), { ...original, setupHidden: true, view: "dashboard" }, 0)
  expect(reseed.status).toBe(409)
  expect(await reseed.json()).toMatchObject({ state: original, revision: 1 })
  const newer = { ...original, setupHidden: true, view: "dashboard" }
  expect((await put(app, newer, 1)).status).toBe(200)
  const stale = await put(app, original, 1)
  expect(stale.status).toBe(409)
  expect(await stale.json()).toMatchObject({ state: newer, revision: 2 })
  expect(await (await app.request("/dashboard/tabs")).json()).toEqual({ state: newer, revision: 2 })
})

test("malformed, oversized and unsupported widget state is rejected without replacing saved tabs", async () => {
  workspace(); const app = dashboardRoutes(), valid = state()
  expect((await put(app, valid, 0)).status).toBe(200)
  const invalidWidget = state()
  invalidWidget.workspace.dashboards[1]!.widgets = [{ key: "chart", widget: { type: "chart", title: "Untracked hours", metric: "hours", groupBy: "client", form: "bars", period: "all" } as unknown as Chart, pinned: false }]
  const invalid = [
    { ...valid, version: 2 },
    { ...valid, view: "unknown" },
    { ...valid, setupHidden: "yes" },
    { ...valid, setupHidden: true, view: "setup" },
    { ...valid, workspace: { ...valid.workspace, active: "missing" } },
    { ...valid, workspace: { ...valid.workspace, dashboards: [valid.workspace.dashboards[0], valid.workspace.dashboards[0]] } },
    { ...valid, workspace: { active: "__setup", dashboards: [{ ...valid.workspace.dashboards[1], id: "__setup" }] } },
    invalidWidget,
    JSON.parse('{"version":1,"workspace":{"active":"","dashboards":[]},"setupHidden":true,"__proto__":{"polluted":true}}'),
  ]
  for (const candidate of invalid) expect((await put(app, candidate, 1)).status).toBe(400)
  expect((await put(app, valid, -1)).status).toBe(400)
  expect((await app.request("/dashboard/tabs", { method: "PUT", body: "{" })).status).toBe(400)
  expect((await app.request("/dashboard/tabs", { method: "PUT", body: " ".repeat(DASHBOARD_STATE_MAX_BYTES + 1) })).status).toBe(413)
  expect(await (await app.request("/dashboard/tabs")).json()).toEqual({ state: valid, revision: 1 })
})

test("closed empty tabs persist and workspace roots keep separate dashboard databases", async () => {
  const first = workspace(), app = dashboardRoutes()
  const closed = { version: 1, workspace: { active: "", dashboards: [] }, setupHidden: true, view: "dashboard" }
  expect((await put(app, closed, 0)).status).toBe(200)
  const second = workspace()
  expect(second).not.toBe(first)
  expect(await (await app.request("/dashboard/tabs")).json()).toEqual({ state: null, revision: 0 })
  process.env.WORKSPACE_ROOT = first
  expect(await (await app.request("/dashboard/tabs")).json()).toEqual({ state: closed, revision: 1 })
})

test("dashboard descriptions persist with their view and reject malformed edits without replacing it", async () => {
  workspace(); const app = dashboardRoutes(), original = state()
  const saved = { ...original, workspace: { ...original.workspace, dashboards: original.workspace.dashboards.map(tab => ({ ...tab, description: "Spend by client and conversations over time" })) } }
  expect((await put(app, saved, 0)).status).toBe(200)
  expect(await (await dashboardRoutes().request("/dashboard/tabs")).json()).toEqual({ state: saved, revision: 1 })
  for (const description of ["", " ", "x".repeat(241), "A line\nAnother line", 123, { source: "generated-code" }]) {
    const invalid = { ...saved, workspace: { ...saved.workspace, dashboards: saved.workspace.dashboards.map(tab => ({ ...tab, description })) } }
    expect((await put(app, invalid, 1)).status).toBe(400)
  }
  expect(await (await app.request("/dashboard/tabs")).json()).toEqual({ state: saved, revision: 1 })
})
