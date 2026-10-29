import { Database } from "bun:sqlite"
import { chmodSync, mkdirSync } from "node:fs"
import path from "node:path"
import { chartKey, checkPlan, TEMPLATES, type PlanWidget } from "../../../dochaus/lib/dashboard-plan"
import { workspaceRoot } from "./matter"

const PERIODS = ["month", "last-month", "quarter", "all"]
const STAGES = ["starting", "practice"]
export const DASHBOARD_STATE_MAX_BYTES = 256 * 1024

type Widget = { key: string; widget: PlanWidget; pinned: boolean }
type Arrangement = { order?: string[]; hidden: string[]; added: string[] }
type Tab = { id: string; name: string; period: string; stages: Record<string, Widget[]>; widgets?: Widget[]; notes?: string[]; legacy?: Record<string, Arrangement> }
export type DashboardState = {
  version: 1
  workspace: { active: string; dashboards: Tab[] }
  setupHidden: boolean
  view?: "setup" | "dashboard"
}
export type DashboardSnapshot = { state: DashboardState | null; revision: number }
export class DashboardStateError extends Error {}

function object(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
function fields(value: Record<string, unknown>, allowed: string[]) {
  if (Object.keys(value).some(key => !allowed.includes(key))) throw new DashboardStateError("Dashboard settings contain an unsupported field.")
}
function json(value: unknown, depth = 0): boolean {
  if (depth > 16) return false
  if (value === null || typeof value === "boolean") return true
  if (typeof value === "number") return Number.isFinite(value)
  if (typeof value === "string") return value.length <= 4096
  if (Array.isArray(value)) return value.every(item => json(item, depth + 1))
  return object(value) && Object.entries(value).every(([key, item]) => !["__proto__", "prototype", "constructor"].includes(key) && json(item, depth + 1))
}
function text(value: unknown, max: number, message: string) {
  if (typeof value !== "string" || !value.trim() || value.length > max || /\p{Cc}/u.test(value)) throw new DashboardStateError(message)
  return value
}
function widgetList(value: unknown): Widget[] {
  if (!Array.isArray(value) || value.length > 32) throw new DashboardStateError("A saved dashboard can contain at most 32 widgets.")
  const keys = new Set<string>()
  return value.map(item => {
    if (!object(item)) throw new DashboardStateError("Choose valid dashboard widgets.")
    fields(item, ["key", "widget", "pinned"])
    if (typeof item.pinned !== "boolean" || !object(item.widget)) throw new DashboardStateError("Choose valid dashboard widgets.")
    fields(item.widget, item.widget.type === "chart" ? ["type", "title", "metric", "groupBy", "form", "period", "client", "matter", "limit"] : ["type"])
    const checked = checkPlan({ widgets: [item.widget] })
    if (!checked.plan) throw new DashboardStateError("One of the saved dashboard widgets cannot be displayed.")
    const widget = checked.plan.widgets[0]!
    const key = widget.type === "chart" ? chartKey(widget) : widget.type
    if (item.key !== key || keys.has(key)) throw new DashboardStateError("Saved dashboard widgets need distinct, valid keys.")
    keys.add(key)
    return { key, widget, pinned: item.pinned }
  })
}
function templateList(value: unknown): string[] {
  if (!Array.isArray(value) || value.length > 32 || value.some(id => typeof id !== "string" || !Object.hasOwn(TEMPLATES, id)) || new Set(value).size !== value.length)
    throw new DashboardStateError("Choose valid saved dashboard templates.")
  return value as string[]
}

// The server accepts only presentation state, never executable sources or model
// instructions. Widgets use the same chart contract as dashboard-plan.
export function parseDashboardState(value: unknown): DashboardState {
  if (!object(value) || !json(value)) throw new DashboardStateError("Choose valid dashboard settings.")
  fields(value, ["version", "workspace", "setupHidden", "view"])
  if (value.version !== 1 || typeof value.setupHidden !== "boolean" || !object(value.workspace) || value.view !== undefined && value.view !== "setup" && value.view !== "dashboard")
    throw new DashboardStateError("Choose valid dashboard settings.")
  if (value.setupHidden && value.view === "setup") throw new DashboardStateError("Show the Setup tab before selecting it.")
  fields(value.workspace, ["active", "dashboards"])
  const list = value.workspace.dashboards
  if (!Array.isArray(list) || list.length > 12) throw new DashboardStateError("Save no more than 12 dashboard tabs.")
  const ids = new Set<string>()
  const dashboards = list.map((item): Tab => {
    if (!object(item)) throw new DashboardStateError("Choose valid dashboard tabs.")
    fields(item, ["id", "name", "period", "stages", "widgets", "notes", "legacy"])
    const id = text(item.id, 128, "Each dashboard tab needs a short identifier.")
    if (id === "__setup") throw new DashboardStateError("The Setup tab identifier is reserved.")
    const name = text(item.name, 30, "Each dashboard tab needs a name of at most 30 characters.").trim()
    if (ids.has(id) || !PERIODS.includes(item.period as string) || !object(item.stages)) throw new DashboardStateError("Choose distinct dashboard tabs and a valid period.")
    ids.add(id)
    fields(item.stages, STAGES)
    const stages = Object.fromEntries(Object.entries(item.stages).map(([stage, widgets]) => [stage, widgetList(widgets)]))
    const widgets = item.widgets === undefined ? undefined : widgetList(item.widgets)
    let notes: string[] | undefined
    if (item.notes !== undefined) {
      if (!Array.isArray(item.notes) || item.notes.length > 5) throw new DashboardStateError("Save at most five short dashboard notes.")
      notes = item.notes.map(note => text(note, 160, "Dashboard notes must be short sentences.").trim())
    }
    let legacy: Record<string, Arrangement> | undefined
    if (item.legacy !== undefined) {
      if (!object(item.legacy)) throw new DashboardStateError("Choose valid saved dashboard arrangements.")
      fields(item.legacy, STAGES)
      legacy = Object.fromEntries(Object.entries(item.legacy).map(([stage, arrangement]) => {
        if (!object(arrangement)) throw new DashboardStateError("Choose valid saved dashboard arrangements.")
        fields(arrangement, ["order", "hidden", "added"])
        return [stage, { ...(arrangement.order === undefined ? {} : { order: templateList(arrangement.order) }), hidden: templateList(arrangement.hidden), added: templateList(arrangement.added) }]
      }))
    }
    return { id, name, period: item.period as string, stages, ...(widgets === undefined ? {} : { widgets }), ...(notes === undefined ? {} : { notes }), ...(legacy === undefined ? {} : { legacy }) }
  })
  const active = value.workspace.active
  if (typeof active !== "string" || (dashboards.length ? !ids.has(active) : active !== "")) throw new DashboardStateError("Select a saved dashboard tab.")
  return { version: 1, workspace: { active, dashboards }, setupHidden: value.setupHidden, ...(value.view === undefined ? {} : { view: value.view }) }
}

function open() {
  const directory = path.join(workspaceRoot(), ".dashboard")
  mkdirSync(directory, { recursive: true })
  const file = path.join(directory, "dashboard.db")
  const db = new Database(file)
  try {
    chmodSync(file, 0o600)
    db.run("PRAGMA busy_timeout=5000")
    db.run("CREATE TABLE IF NOT EXISTS dashboard_tabs (id INTEGER PRIMARY KEY CHECK (id = 1), state TEXT NOT NULL, revision INTEGER NOT NULL CHECK (revision > 0))")
    return db
  } catch (error) { db.close(); throw error }
}
type Row = { state: string; revision: number }
function snapshot(db: Database): DashboardSnapshot {
  const row = db.query<Row, []>("SELECT state, revision FROM dashboard_tabs WHERE id = 1").get()
  if (!row) return { state: null, revision: 0 }
  if (!Number.isSafeInteger(row.revision) || row.revision < 1) throw new Error("Invalid dashboard revision")
  return { state: parseDashboardState(JSON.parse(row.state)), revision: row.revision }
}
export function readDashboardState(): DashboardSnapshot {
  const db = open()
  try { return snapshot(db) } finally { db.close() }
}
export function writeDashboardState(state: DashboardState, revision: number): { saved: boolean; snapshot: DashboardSnapshot } {
  const db = open()
  try {
    const serialized = JSON.stringify(state)
    // One SQL statement chooses the winning seed or checks the saved revision.
    // A new browser cannot overwrite another browser's existing settings.
    const row = revision === 0
      ? db.query<Row, [string]>("INSERT INTO dashboard_tabs (id, state, revision) VALUES (1, ?, 1) ON CONFLICT(id) DO NOTHING RETURNING state, revision").get(serialized)
      : db.query<Row, [string, number]>("UPDATE dashboard_tabs SET state = ?, revision = revision + 1 WHERE id = 1 AND revision = ? RETURNING state, revision").get(serialized, revision)
    return { saved: !!row, snapshot: row ? { state: parseDashboardState(JSON.parse(row.state)), revision: row.revision } : snapshot(db) }
  } finally { db.close() }
}
