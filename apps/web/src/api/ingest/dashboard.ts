import { INGEST_URL } from "../../config"

export async function dashboardDirectory(signal?: AbortSignal): Promise<string> {
  const res = await fetch(`${INGEST_URL}/dashboard/directory`, { signal })
  if (!res.ok) throw new Error("Could not prepare the dashboard. Try again.")
  const value: unknown = await res.json()
  if (!value || typeof value !== "object" || !("dir" in value) || typeof value.dir !== "string" || !value.dir)
    throw new Error("Could not prepare the dashboard directory.")
  return value.dir
}

// Only app-constructed URLs and the checked chart contract reach this endpoint.
// The backend coalesces the underlying practice read across different charts.
import type { Chart } from "../../../../../dochaus/lib/dashboard-plan"
import type { Period } from "../../components/dashboard/data"
import { isJson, type DashboardSource, type SourceSnapshot } from "../../components/dashboard/visual/source"

export function chartSourceURL(chart: Chart, period: Period, timeZone: string) {
  const source = chart.source ?? { kind: "local-query", query: "practice-chart", version: 1 }
  const query = new URLSearchParams({ chart: JSON.stringify({ ...chart, source }), period, timeZone })
  return `${INGEST_URL}/dashboard/sources/chart?${query}`
}

export function createChartEndpointSource(url: string, intervalMs = 30_000): DashboardSource {
  let snapshot: SourceSnapshot = Object.freeze({ status: "loading" })
  const listeners = new Set<() => void>()
  let request: AbortController | undefined
  let timer: ReturnType<typeof setTimeout> | undefined
  let again = false
  const visible = () => document.visibilityState === "visible"
  const emit = (next: SourceSnapshot) => { snapshot = Object.freeze(next); listeners.forEach((listener) => listener()) }
  const refresh = async () => {
    if (!listeners.size || !visible()) return
    if (request) { again = true; return }
    clearTimeout(timer)
    const controller = new AbortController()
    request = controller
    try {
      const response = await fetch(url, { signal: AbortSignal.any([controller.signal, AbortSignal.timeout(15000)]), redirect: "error", cache: "no-store", headers: { Accept: "application/json" } })
      if (!response.ok) throw new Error("Source unavailable")
      const result: unknown = await response.json()
      if (!result || typeof result !== "object" || Array.isArray(result) || !("data" in result) || !isJson(result.data) || !("updatedAt" in result) || typeof result.updatedAt !== "number" || !Number.isFinite(result.updatedAt)) throw new Error("Invalid source response")
      if (!controller.signal.aborted && !again) emit({ status: "live", data: result.data, updatedAt: result.updatedAt })
    } catch {
      if (!controller.signal.aborted && !again) emit({ ...snapshot, status: "error", error: "Could not update this data source." })
    } finally {
      if (request === controller) request = undefined
      if (listeners.size && !controller.signal.aborted && visible()) {
        const delay = again ? 0 : intervalMs
        again = false
        timer = setTimeout(() => { void refresh() }, delay)
      }
    }
  }
  const onVisible = () => { if (visible()) void refresh() }
  return {
    getSnapshot: () => snapshot,
    subscribe(listener) {
      listeners.add(listener)
      if (listeners.size === 1) {
        document.addEventListener("visibilitychange", onVisible)
        window.addEventListener("focus", onVisible)
        void refresh()
      }
      return () => {
        listeners.delete(listener)
        if (!listeners.size) {
          clearTimeout(timer)
          request?.abort()
          request = undefined
          again = false
          document.removeEventListener("visibilitychange", onVisible)
          window.removeEventListener("focus", onVisible)
        }
      }
    },
    refresh: () => { void refresh() },
  }
}
