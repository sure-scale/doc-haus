import { useMemo, useSyncExternalStore } from "react"
import { AIComposer, AboutView, DashboardTabs } from "./DashboardChrome"
import { ChartCard } from "./ChartCards"
import { DemoDataCard, SetupChecklist } from "./SetupCards"
import { TokenBreakdown } from "./TokenBreakdown"
import { ActionTiles, AttentionRows, LibraryCards, MatterRows, SamplePreview, StatCard } from "./WidgetCards"
import { icons, ShapeIcon, type IconName } from "../icons"
import { validateComponentData, type ComponentKind } from "./schema"
import { isJson, readBinding, type DashboardSources, type DataBinding, type JsonValue } from "./source"
import "../../../styles/dashboard.css"
import "../../../styles/dashboard-chart.css"

export type ComponentFrame = { title: string; subtitle?: string; icon?: IconName; tone?: "default" | "bold" | "ink" | "tint-1" | "tint-2" | "tint-3"; size?: "tile" | "wide" | "full" }
export type ComponentDefinition = {
  id: string
  kind: ComponentKind
  props?: { [key: string]: JsonValue }
  // Each binding replaces one top-level prop. Sources can emit the whole card
  // model or just a series/count. JSON Pointer addresses nested source values.
  bindings?: Record<string, DataBinding>
  frame?: ComponentFrame
}
export type ComponentAction = (action: string, componentId: string, payload?: JsonValue) => void
const EMPTY: DashboardSources = Object.freeze({})

export function useDashboardSource(source: import("./source").DashboardSource) {
  return useSyncExternalStore(source.subscribe, source.getSnapshot, source.getSnapshot)
}

function useSources(sources: DashboardSources, ids: string[]) {
  const key = [...new Set(ids)].sort().join("\u0000")
  const store = useMemo(() => {
    const selected = key ? key.split("\u0000").map((id) => Object.hasOwn(sources, id) ? sources[id] : undefined) : []
    let snapshots = selected.map((source) => source?.getSnapshot())
    return {
      getSnapshot() {
        const next = selected.map((source) => source?.getSnapshot())
        if (next.some((snapshot, index) => snapshot !== snapshots[index])) snapshots = next
        return snapshots
      },
      subscribe(listener: () => void) {
        const unsubscribe = selected.flatMap((source) => source ? [source.subscribe(listener)] : [])
        return () => unsubscribe.forEach((stop) => stop())
      },
    }
  }, [sources, key])
  useSyncExternalStore(store.subscribe, store.getSnapshot, store.getSnapshot)
}

export function checkDefinition(value: unknown): value is ComponentDefinition {
  if (!isJson(value) || !value || Array.isArray(value) || typeof value !== "object") return false
  const { id, kind, props, bindings, frame } = value
  if (typeof id !== "string" || !id.trim() || id.length > 100 || typeof kind !== "string" || !["stat", "attention", "matters", "library", "sample", "chart", "setup", "demo", "tabs", "about", "composer", "tokens", "actions"].includes(kind)) return false
  if (props !== undefined && (!props || Array.isArray(props) || typeof props !== "object")) return false
  if (bindings !== undefined) {
    if (!bindings || Array.isArray(bindings) || typeof bindings !== "object" || Object.keys(bindings).length > 30) return false
    for (const binding of Object.values(bindings)) {
      if (!binding || Array.isArray(binding) || typeof binding !== "object" || typeof binding.source !== "string" || !binding.source.trim() || binding.source.includes("\u0000") || (binding.path !== undefined && (typeof binding.path !== "string" || (binding.path !== "" && (!binding.path.startsWith("/") || /~(?![01])/.test(binding.path)))))) return false
    }
  }
  if (frame !== undefined) {
    if (!frame || Array.isArray(frame) || typeof frame !== "object" || typeof frame.title !== "string" || !frame.title.trim()) return false
    if (frame.subtitle !== undefined && typeof frame.subtitle !== "string") return false
    if (frame.icon !== undefined && (typeof frame.icon !== "string" || !Object.hasOwn(icons, frame.icon))) return false
    if (frame.tone !== undefined && (typeof frame.tone !== "string" || !["default", "bold", "ink", "tint-1", "tint-2", "tint-3"].includes(frame.tone))) return false
    if (frame.size !== undefined && (typeof frame.size !== "string" || !["tile", "wide", "full"].includes(frame.size))) return false
  }
  return true
}

export function LiveComponent({ definition, sources = EMPTY, onAction }: { definition: ComponentDefinition; sources?: DashboardSources; onAction?: ComponentAction }) {
  // Invalid JSON must not become executable props or partially rendered data.
  const valid = checkDefinition(definition)
  const bindings = valid ? definition.bindings ?? {} : {}
  useSources(sources, Object.values(bindings).map(({ source }) => source))
  let error = valid ? "" : "This component configuration is invalid."
  let loading = false
  const props: Record<string, unknown> = valid ? { ...definition.props } : {}
  const updates: number[] = []
  for (const [prop, binding] of Object.entries(bindings)) {
    const source = Object.hasOwn(sources, binding.source) ? sources[binding.source] : undefined
    const snapshot = source?.getSnapshot()
    if (!snapshot) { error = `Data source “${binding.source}” is not connected.`; continue }
    if (snapshot.updatedAt !== undefined) updates.push(snapshot.updatedAt)
    if (snapshot.status === "error") { error = "Could not update this view. Retry when the data source is available."; continue }
    if (snapshot.status === "loading") { loading = true; continue }
    const value = readBinding(snapshot.data, binding.path)
    if (value === undefined) { error = `The data source has no value at ${binding.path || "the root"}.`; continue }
    props[prop] = value
  }
  const checked = valid && !loading && !error ? validateComponentData(definition.kind, props) : undefined
  if (checked?.error) error = checked.error
  if (valid && definition.kind === "composer" && props.open === true && !onAction) error = "Connect the composer actions before opening this dialog."
  const action = onAction ? (id: string) => onAction(id, definition.id) : undefined
  const rendered = !error && !loading && checked?.data ? (() => {
    // Each branch narrows through its own validator. No source values are ever
    // spread onto DOM elements or interpreted as HTML, URLs or expressions.
    switch (definition.kind) {
      case "tokens": { const view = validateComponentData("tokens", props); return view.data ? <TokenBreakdown {...view.data} /> : null }
      case "actions": { const view = validateComponentData("actions", props); return view.data ? <ActionTiles {...view.data} onAction={action} /> : null }
      case "tabs": { const view = validateComponentData("tabs", props); return view.data ? <DashboardTabs {...view.data} onSelect={(id) => onAction?.("select-tab", definition.id, { id })} onNew={() => onAction?.("new-dashboard", definition.id)} /> : null }
      case "about": { const view = validateComponentData("about", props); return view.data ? <AboutView {...view.data} onRefine={onAction ? () => onAction("refine-view", definition.id) : undefined} /> : null }
      case "composer": { const view = validateComponentData("composer", props); return view.data ? <AIComposer {...view.data} onRequestChange={(value) => onAction?.("change-request", definition.id, { value })} onGenerate={() => onAction?.("generate-dashboard", definition.id)} onCancel={() => onAction?.("cancel-dashboard", definition.id)} onStop={onAction ? () => onAction("stop-dashboard", definition.id) : undefined} /> : null }
      case "chart": { const view = validateComponentData("chart", props); return view.data ? <ChartCard {...view.data} /> : null }
      case "stat": { const view = validateComponentData("stat", props); return view.data ? <StatCard {...view.data} /> : null }
      case "attention": { const view = validateComponentData("attention", props); return view.data ? <AttentionRows {...view.data} onAction={action} /> : null }
      case "matters": { const view = validateComponentData("matters", props); return view.data ? <MatterRows {...view.data} /> : null }
      case "library": { const view = validateComponentData("library", props); return view.data ? <LibraryCards {...view.data} /> : null }
      case "sample": { const view = validateComponentData("sample", props); return view.data ? <SamplePreview {...view.data} /> : null }
      case "setup": { const view = validateComponentData("setup", props); return view.data ? <SetupChecklist {...view.data} onAction={action} /> : null }
      case "demo": { const view = validateComponentData("demo", props); return view.data ? <DemoDataCard {...view.data} onAction={action} /> : null }
    }
  })() : null
  const body = error ? <div className="dash-source-error" role="alert"><p>{error}</p>{Object.values(bindings).some(({ source }) => (Object.hasOwn(sources, source) && sources[source]?.refresh)) && <button type="button" onClick={() => [...new Set(Object.values(bindings).map(({ source }) => source))].forEach((id) => (Object.hasOwn(sources, id) && sources[id]?.refresh?.()))}>Retry</button>}{updates.length > 0 && <p>Last updated {new Date(Math.min(...updates)).toLocaleTimeString()}</p>}</div>
    : loading ? <div className="dash-skeleton" role="status" aria-label="Loading live data"><span /><span /><span /></div> : rendered
  if (!valid || !definition.frame) return <div data-component={valid ? definition.id : undefined} aria-busy={loading}>{body}</div>
  const frame = definition.frame
  return <section data-component={definition.id} className={`dash-card size-${frame.size ?? "wide"}${frame.tone === "bold" ? " is-bold" : frame.tone && frame.tone !== "default" ? ` tone-${frame.tone}` : ""}`} aria-label={frame.title} aria-busy={loading}>
    <header className="dash-card-head">{frame.icon && <ShapeIcon icon={frame.icon} shape="sunny" className="dash-card-badge" />}<div className="dash-card-heading"><h2>{frame.title}</h2>{frame.subtitle && <p>{frame.subtitle}</p>}</div></header>
    <div className="dash-card-body">{body}</div>
  </section>
}
