// Shared display formatters. Dates read the same everywhere ("11 Oct 2026"),
// independent of the browser locale's numeric order, so a matter list and a
// document row never disagree about what 03/04 means.

const DATE = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", year: "numeric" })
const PRECISE_TIME = new Intl.DateTimeFormat("en-GB", { hour: "2-digit", minute: "2-digit", second: "2-digit" })
const TIME = new Intl.DateTimeFormat("en-GB", { hour: "2-digit", minute: "2-digit" })

// Accepts epoch milliseconds, epoch seconds (ingest stores some as seconds) or
// an ISO string. Returns "" for a missing or unparseable value.
export function formatDate(value: number | string | null | undefined) {
  const date = toDate(value)
  return date ? DATE.format(date) : ""
}

export function formatDateTime(value: number | string | null | undefined) {
  const date = toDate(value)
  return date ? `${DATE.format(date)}, ${TIME.format(date)}` : ""
}

// Timeline entries show seconds because several files can be saved in one minute.
export function formatTime(value: number | string | null | undefined) {
  const date = toDate(value)
  return date ? PRECISE_TIME.format(date) : ""
}

// A duration as "45s", "4m 12s" or "1h 05m". Negative or missing → "".
export function formatElapsed(ms: number | null | undefined) {
  if (ms === null || ms === undefined || !Number.isFinite(ms) || ms < 0) return ""
  const seconds = Math.floor(ms / 1000)
  if (seconds < 60) return `${seconds}s`
  const minutes = Math.floor(seconds / 60)
  if (minutes < 60) return `${minutes}m ${String(seconds % 60).padStart(2, "0")}s`
  return `${Math.floor(minutes / 60)}h ${String(minutes % 60).padStart(2, "0")}m`
}

// "just now", "12s ago", "4m ago", "2h ago", then a date.
export function formatAgo(from: number | string | null | undefined, now = Date.now()) {
  const date = toDate(from)
  if (!date) return ""
  const seconds = Math.max(0, Math.floor((now - date.getTime()) / 1000))
  if (seconds < 5) return "just now"
  if (seconds < 60) return `${seconds}s ago`
  if (seconds < 3600) return `${Math.floor(seconds / 60)}m ago`
  if (seconds < 86_400) return `${Math.floor(seconds / 3600)}h ago`
  return formatDate(date.getTime())
}

// Dollar cost to two decimals. A cost the rate card cannot price is "Not priced"
// rather than a misleading $0.00; a partly priced total is a lower bound.
export function formatCost(value: number, unpriced = false) {
  if (unpriced && value <= 0) return "Not priced"
  const amount = value > 0 && value < 0.01 ? "< $0.01" : `$${value.toFixed(2)}`
  return unpriced ? `≥ ${amount.replace("< ", "")}` : amount
}

function toDate(value: number | string | null | undefined) {
  if (value === null || value === undefined || value === "") return undefined
  const date = typeof value === "number" ? new Date(value < 1e11 ? value * 1000 : value) : new Date(value)
  return Number.isNaN(date.getTime()) ? undefined : date
}
