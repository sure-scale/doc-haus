import { icons } from "../icons"
import type { ChartCardData } from "./ChartCards"
import type { TokenBreakdownData } from "./TokenBreakdown"
import type { AIComposerConfig, AboutViewConfig, DashboardTabItem } from "./DashboardChrome"
import type { DemoDataCardConfig, SetupChecklistConfig } from "./SetupCards"
import type { ActionTileData, AttentionRowData, LibraryCardData, MatterRowData, SamplePreviewData, StatCardData } from "./WidgetCards"
import { isJson } from "./source"

export type ComponentDataMap = {
  stat: StatCardData
  attention: { rows: AttentionRowData[] }
  matters: { rows: MatterRowData[] }
  library: { cards: LibraryCardData[] }
  sample: { items: SamplePreviewData[] }
  chart: ChartCardData
  actions: { items: ActionTileData[] }
  tokens: TokenBreakdownData
  setup: { config: SetupChecklistConfig }
  demo: { config: DemoDataCardConfig; busy?: boolean; disabled?: boolean; error?: string }
  tabs: { items: DashboardTabItem[]; activeId: string; newDisabled?: boolean; label?: string }
  about: { config: AboutViewConfig }
  composer: { open: boolean; request: string; config?: AIComposerConfig; working?: boolean; error?: string; disabled?: boolean }
}
export type ComponentKind = keyof ComponentDataMap
export type ComponentValidation<T> = { data: T; error?: undefined } | { data?: undefined; error: string }

const KINDS = new Set<string>(["stat", "attention", "matters", "library", "sample", "chart", "setup", "demo", "tabs", "about", "composer", "actions", "tokens"])
const currencies = new Set<string>(Intl.supportedValuesOf("currency"))
type RecordValue = Record<string, unknown>
const object = (value: unknown): value is RecordValue => typeof value === "object" && value !== null && !Array.isArray(value)
const text = (value: unknown) => typeof value === "string" && value.length <= 10_000
const id = (value: unknown) => typeof value === "string" && value.trim().length > 0 && value.length <= 128
const optional = (row: RecordValue, key: string, check: (value: unknown) => boolean) => row[key] === undefined || check(row[key])
const number = (value: unknown): value is number => typeof value === "number" && Number.isFinite(value)
const bool = (value: unknown) => typeof value === "boolean"
const icon = (value: unknown) => typeof value === "string" && Object.hasOwn(icons, value)
const member = (values: string[]) => (value: unknown) => typeof value === "string" && values.includes(value)
// Browser URL parsing strips control characters and treats backslashes as
// separators. Refuse them before a JSON-fed link can become a remote URL.
export const isLocalDashboardHref = (value: unknown): value is string => typeof value === "string" && /^\/(?!\/)[^\u0000-\u0020\\]*$/.test(value)

function list(value: unknown, check: (row: RecordValue) => boolean, key = "id", maximum = 200): boolean {
  if (!Array.isArray(value) || value.length > maximum) return false
  const seen = new Set<string>()
  return value.every((row) => {
    if (!object(row) || !id(row[key]) || seen.has(row[key] as string) || !check(row)) return false
    seen.add(row[key] as string)
    return true
  })
}
const strings = (value: unknown) => Array.isArray(value) && value.length <= 200 && value.every(text)
const date = (value: unknown): value is number => number(value) && Number.isFinite(new Date(value).getTime())
const localDay = (time: number) => { const day = new Date(time); day.setHours(0, 0, 0, 0); return day.getTime() }

function validStatVisual(value: unknown): boolean {
  if (!object(value)) return false
  if (value.type === "week") {
    if (!Array.isArray(value.days) || value.days.length > 7) return false
    const dates = new Set<number>()
    let today = false
    return value.days.every((day) => {
      if (!object(day) || !date(day.day) || !number(day.cost) || day.cost < 0 || !text(day.label) || !bool(day.today) || !optional(day, "detail", text)) return false
      const key = localDay(day.day)
      if (dates.has(key) || (day.today && today)) return false
      dates.add(key)
      today ||= day.today as boolean
      return true
    })
  }
  return member(["sparkline", "spread"])(value.type) && text(value.label) && Array.isArray(value.values) && value.values.length <= 5000 && value.values.every((entry) => number(entry) && entry >= 0) && (value.type !== "spread" || (number(value.average) && value.average >= 0 && optional(value, "format", member(["money", "count", "percent"]))))
}

function validChart(row: RecordValue): boolean {
  if (!member(["stat", "line", "bars", "list", "calendar"])(row.form) || !text(row.label) || !Array.isArray(row.points) || row.points.length > 5000) return false
  if (!optional(row, "format", member(["money", "count", "percent"])) || !optional(row, "currency", (value) => typeof value === "string" && currencies.has(value))) return false
  if (!["basis", "period", "badge", "empty"].every((key) => optional(row, key, text)) || !["showShares", "unavailable", "integer"].every((key) => optional(row, key, bool)) || !optional(row, "notes", strings)) return false
  const countFormat = row.format === undefined || row.format === "count"
  if (!optional(row, "total", (value) => number(value) && (row.form === "stat" || value >= 0) && (!countFormat || Number.isInteger(value)))) return false
  const keys = new Set<string>()
  const days = new Set<number>()
  let sum = 0
  for (const point of row.points) {
    if (!object(point) || !text(point.label) || !number(point.value) || (row.form !== "stat" && point.value < 0) || (countFormat && !Number.isInteger(point.value))) return false
    if (!optional(point, "key", id) || !optional(point, "detail", text) || !optional(point, "time", date)) return false
    if (typeof point.key === "string") { if (keys.has(point.key)) return false; keys.add(point.key) }
    if (row.form === "calendar") {
      if (!date(point.time)) return false
      const day = localDay(point.time)
      if (days.has(day)) return false
      days.add(day)
    }
    sum += point.value
  }
  if (!Number.isFinite(sum)) return false
  if (row.showShares && typeof row.total === "number" && row.total < sum - Math.max(1, Math.abs(sum)) * 1e-9) return false
  if (days.size) {
    const first = new Date(Math.min(...days))
    first.setDate(first.getDate() - (first.getDay() + 6) % 7)
    const last = Math.max(...days)
    // Count calendar days rather than elapsed hours across daylight saving.
    const limit = new Date(first); limit.setDate(limit.getDate() + 3660)
    if (last >= limit.getTime()) return false
  }
  return true
}

function valid(kind: ComponentKind, row: RecordValue): boolean {
  switch (kind) {
    case "stat": return text(row.value) && text(row.caption) && ["tag", "tagTitle", "note"].every((key) => optional(row, key, text)) && optional(row, "visual", validStatVisual)
    case "attention": return list(row.rows, (item) => text(item.title) && text(item.detail) && icon(item.icon) && member(["error", "info", "warning", "neutral"])(item.tone) && object(item.action) && text(item.action.label) && optional(item.action, "href", isLocalDashboardHref))
    case "matters": return list(row.rows, (item) => ["title", "detail", "updated", "spend"].every((key) => text(item[key])) && isLocalDashboardHref(item.href) && optional(item, "status", (status) => object(status) && icon(status.icon) && text(status.label) && member(["neutral", "warning", "error", "good"])(status.tone)))
    case "library": return list(row.cards, (item) => icon(item.icon) && ["label", "detail", "emptyLabel"].every((key) => text(item[key])) && (item.count === null || (number(item.count) && Number.isInteger(item.count) && item.count >= 0)) && isLocalDashboardHref(item.href))
    case "sample": return list(row.items, (item) => icon(item.icon) && text(item.title) && text(item.detail))
    case "chart": return validChart(row)
    case "actions": return list(row.items, (item) => icon(item.icon) && text(item.title) && text(item.detail) && optional(item, "href", isLocalDashboardHref))
    case "tokens": {
      if (!list(row.segments, (segment) => text(segment.label) && number(segment.value) && Number.isInteger(segment.value) && segment.value >= 0, "key", 50) || !optional(row, "period", text) || !optional(row, "notes", strings)) return false
      const total = (row.segments as { value: number }[]).reduce((sum, segment) => sum + segment.value, 0)
      return Number.isFinite(total) && optional(row, "total", (value) => number(value) && Number.isInteger(value) && value === total)
    }
    case "setup": return object(row.config) && text(row.config.title) && text(row.config.description) && list(row.config.items, (item) => icon(item.icon) && text(item.title) && text(item.detail) && bool(item.done) && optional(item, "editable", bool) && optional(item, "action", (action) => object(action) && id(action.id) && text(action.label) && optional(action, "variant", member(["text", "outline"]))))
    case "tabs": return text(row.activeId) && optional(row, "label", text) && optional(row, "newDisabled", bool) && list(row.items, (item) => text(item.label) && optional(item, "disabled", bool) && optional(item, "badge", (badge) => object(badge) && number(badge.value) && Number.isInteger(badge.value) && badge.value >= 0 && optional(badge, "label", text) && optional(badge, "total", (total) => number(total) && Number.isInteger(total) && total > 0 && (badge.value as number) <= total))) && ((row.items as RecordValue[]).length === 0 || (row.items as RecordValue[]).some((item) => item.id === row.activeId))
    case "about": return object(row.config) && text(row.config.text) && optional(row.config, "title", text) && optional(row.config, "actionLabel", text)
    case "composer": return bool(row.open) && text(row.request) && ["working", "disabled"].every((key) => optional(row, key, bool)) && optional(row, "error", text) && optional(row, "config", (config) => object(config) && ["title", "description", "promptLabel", "placeholder", "submitLabel", "footnote"].every((key) => optional(config, key, text)) && optional(config, "suggestions", (suggestions) => strings(suggestions) && new Set(suggestions as string[]).size === (suggestions as string[]).length))
    case "demo": {
      const config = row.config
      return ["busy", "disabled"].every((key) => optional(row, key, bool)) && optional(row, "error", text) && object(config) && ["title", "description", "note"].every((key) => text(config[key])) && list(config.features, (feature) => text(feature.label) && optional(feature, "value", text) && optional(feature, "icon", icon), "label") && object(config.action) && id(config.action.id) && text(config.action.label) && optional(config.action, "icon", icon)
    }
  }
}

export function validateComponentData<K extends ComponentKind>(kind: K, value: unknown): ComponentValidation<ComponentDataMap[K]>
export function validateComponentData(kind: string, value: unknown): ComponentValidation<ComponentDataMap[ComponentKind]>
export function validateComponentData(kind: string, value: unknown): ComponentValidation<ComponentDataMap[ComponentKind]> {
  if (!KINDS.has(kind)) return { error: "Unknown dashboard component type." }
  if (!isJson(value) || !object(value)) return { error: "Dashboard component data must be a JSON object with finite values." }
  if (!valid(kind as ComponentKind, value)) return { error: `Invalid ${kind} component data. Check its required fields, values and local links.` }
  return { data: value as ComponentDataMap[ComponentKind] }
}

export function isComponentData<K extends ComponentKind>(kind: K, value: unknown): value is ComponentDataMap[K] {
  return validateComponentData(kind, value).data !== undefined
}
