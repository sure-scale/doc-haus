// Live values are supplied by the application. JSON chooses bindings; it never
// supplies executable code. One source can feed several independently rendered cards.
export type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue }
export type SourceSnapshot = Readonly<{ status: "loading" | "live" | "error"; data?: JsonValue; updatedAt?: number; error?: string }>
export type DashboardSource = {
  getSnapshot: () => SourceSnapshot
  subscribe: (listener: () => void) => () => void
  refresh?: () => void
}
export type DashboardSources = Readonly<Record<string, DashboardSource>>

export function isJson(value: unknown, depth = 0): value is JsonValue {
  if (depth > 30) return false
  if (value === null || typeof value === "boolean" || typeof value === "string") return true
  if (typeof value === "number") return Number.isFinite(value)
  if (Array.isArray(value)) return value.every((entry) => isJson(entry, depth + 1))
  if (typeof value !== "object" || !value || Object.getPrototypeOf(value) !== Object.prototype) return false
  return Object.entries(value).every(([key, entry]) => !["__proto__", "constructor", "prototype"].includes(key) && isJson(entry, depth + 1))
}

function freeze(value: JsonValue): JsonValue {
  if (value && typeof value === "object") { Object.values(value).forEach(freeze); Object.freeze(value) }
  return value
}

export function createDashboardSource(initial?: JsonValue) {
  const listeners = new Set<() => void>()
  let snapshot: SourceSnapshot = { status: "loading" }
  const emit = (next: SourceSnapshot) => { snapshot = Object.freeze(next); listeners.forEach((listener) => listener()) }
  const source = {
    getSnapshot: () => snapshot,
    subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener) } },
    publish(data: JsonValue) {
      if (!isJson(data)) throw new Error("Dashboard source must contain finite JSON values.")
      emit({ status: "live", data: freeze(structuredClone(data)), updatedAt: Date.now() })
    },
    fail(message = "Could not update this data source.") { emit({ ...snapshot, status: "error", error: message }) },
  }
  if (initial !== undefined) source.publish(initial)
  return source
}

// A generated endpoint can use this adapter. It fetches while subscribed, never
// overlaps requests, refreshes on demand and aborts when the last card unmounts.
// Push transports (WebSocket/SSE/store events) publish directly to the source above.
export function createPollingDashboardSource({ url, intervalMs = 5000 }: { url: string; intervalMs?: number }): DashboardSource {
  if (!/^\/(?!\/)[^\u0000-\u0020\\]*$/.test(url)) throw new Error("Use a same-origin endpoint path for dashboard data.")
  if (!Number.isFinite(intervalMs) || intervalMs < 1000) throw new Error("Dashboard polling interval must be at least 1000 ms.")
  const state = createDashboardSource()
  let users = 0
  let request: AbortController | undefined
  let timer: ReturnType<typeof setTimeout> | undefined
  let again = false
  async function refresh() {
    if (!users) return
    if (request) { again = true; return }
    clearTimeout(timer)
    const controller = new AbortController()
    request = controller
    try {
      const response = await fetch(url, { signal: AbortSignal.any([controller.signal, AbortSignal.timeout(15000)]), redirect: "error", headers: { Accept: "application/json" }, cache: "no-store" })
      if (!response.ok) throw new Error("Could not update this data source.")
      const data: unknown = await response.json()
      if (!isJson(data)) throw new Error("This data source returned an invalid JSON value.")
      if (!controller.signal.aborted) state.publish(data)
    } catch (error) {
      if (!controller.signal.aborted) state.fail(error instanceof Error && error.message === "This data source returned an invalid JSON value." ? error.message : "Could not update this data source.")
    } finally {
      if (request === controller) request = undefined
      if (users && !controller.signal.aborted) { const delay = again ? 0 : intervalMs; again = false; timer = setTimeout(() => { void refresh() }, delay) }
    }
  }
  return {
    getSnapshot: state.getSnapshot,
    subscribe(listener) {
      const unsubscribe = state.subscribe(listener)
      if (++users === 1) void refresh()
      return () => {
        unsubscribe()
        if (--users === 0) { clearTimeout(timer); request?.abort(); request = undefined; again = false }
      }
    },
    refresh: () => { void refresh() },
  }
}

export type DataBinding = { source: string; path?: string }
// RFC 6901 JSON Pointer. An empty path selects the complete source object.
export function readBinding(data: JsonValue | undefined, path = ""): JsonValue | undefined {
  if (!path) return data
  if (!path.startsWith("/") || /~(?![01])/.test(path)) return undefined
  let value = data
  for (const encoded of path.slice(1).split("/")) {
    const key = encoded.replace(/~1/g, "/").replace(/~0/g, "~")
    if (["__proto__", "constructor", "prototype"].includes(key) || !value || typeof value !== "object" || !Object.hasOwn(value, key)) return undefined
    value = (value as Record<string, JsonValue>)[key]
  }
  return value
}
