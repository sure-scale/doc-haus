import { INGEST_URL } from "../../config"
import { readDashboards, type Dashboards } from "../../components/dashboard/layout"

export type DashboardTabState = { version: 1; workspace: Dashboards; setupHidden: boolean; view?: "setup" | "dashboard" }
export type DashboardTabSnapshot = { state: DashboardTabState | null; revision: number }
export class DashboardTabConflict extends Error {}

function snapshot(value: unknown): DashboardTabSnapshot {
  if (!value || typeof value !== "object") throw new Error("Could not read saved dashboards.")
  const result = value as DashboardTabSnapshot
  if (!Number.isSafeInteger(result.revision) || result.revision < 0) throw new Error("Could not read saved dashboards.")
  if (result.state === null && result.revision === 0) return result
  const state = result.state
  if (!state || state.version !== 1 || typeof state.setupHidden !== "boolean" || !state.workspace || !Array.isArray(state.workspace.dashboards) || ![undefined, "setup", "dashboard"].includes(state.view)) throw new Error("Could not read saved dashboards.")
  const workspace = readDashboards(state.workspace)
  if (workspace.dashboards.length !== state.workspace.dashboards.length || workspace.active !== state.workspace.active) throw new Error("Could not read saved dashboards.")
  return { revision: result.revision, state: { version: 1, workspace, setupHidden: state.setupHidden, ...(state.view ? { view: state.view } : {}) } }
}

export async function getDashboardTabs(): Promise<DashboardTabSnapshot> {
  const response = await fetch(`${INGEST_URL}/dashboard/tabs`, { signal: AbortSignal.timeout(12000), cache: "no-store" })
  if (!response.ok) throw new Error("Could not load saved dashboards.")
  return snapshot(await response.json())
}

export async function putDashboardTabs(state: DashboardTabState, revision: number): Promise<DashboardTabSnapshot> {
  const response = await fetch(`${INGEST_URL}/dashboard/tabs`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ state, revision }), signal: AbortSignal.timeout(12000) })
  if (response.status === 409) throw new DashboardTabConflict("These dashboards changed in another window.")
  if (!response.ok) throw new Error("Could not save your dashboard changes.")
  return snapshot(await response.json())
}
