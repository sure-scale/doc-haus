import { useEffect, useId, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from "react"
import { icons } from "../icons"
import "../../../styles/dashboard-chrome.css"

export type DashboardTabItem = {
  id: string
  label: string
  badge?: { value: number; total?: number; label?: string }
  disabled?: boolean
}

export function DashboardTabs({ items, activeId, onSelect, onNew, newDisabled = false, label = "Dashboards", panelId }: {
  items: DashboardTabItem[]
  activeId: string
  onSelect: (id: string) => void
  onNew: () => void
  newDisabled?: boolean
  label?: string
  panelId?: string
}) {
  const list = useRef<HTMLDivElement>(null)
  useEffect(() => {
    list.current?.querySelector<HTMLButtonElement>('[aria-selected="true"]')?.scrollIntoView({ inline: "nearest", block: "nearest" })
  }, [activeId])
  const enabled = items.filter((item) => !item.disabled)
  return <div className="dash-view-tabs-row">
    <div ref={list} className="dash-view-tabs" role="tablist" aria-label={label} onKeyDown={(event) => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key) || !enabled.length) return
      event.preventDefault()
      const focused = (event.target as HTMLElement).closest<HTMLButtonElement>("button[data-tab-id]")?.dataset.tabId ?? activeId
      const index = Math.max(0, enabled.findIndex((item) => item.id === focused))
      const target = event.key === "Home" ? enabled[0]! : event.key === "End" ? enabled.at(-1)! : enabled[(index + (event.key === "ArrowRight" ? 1 : -1) + enabled.length) % enabled.length]!
      onSelect(target.id)
      Array.from(list.current?.querySelectorAll<HTMLButtonElement>("button[data-tab-id]") ?? []).find((button) => button.dataset.tabId === target.id)?.focus()
    }}>
      {items.map((item) => <button key={item.id} type="button" role="tab" id={panelId ? `${panelId}-tab-${item.id}` : undefined} aria-controls={panelId} data-tab-id={item.id} className="dash-view-tab" disabled={item.disabled} aria-selected={item.id === activeId} tabIndex={item.id === activeId || !enabled.some((tab) => tab.id === activeId) && enabled[0]?.id === item.id ? 0 : -1} onClick={() => onSelect(item.id)}>
        <span>{item.label}</span>{item.badge && <span className="dash-view-tab-badge" aria-label={item.badge.label ?? `${item.badge.value}${item.badge.total === undefined ? "" : ` of ${item.badge.total}`}`}>{item.badge.value}{item.badge.total === undefined ? "" : `/${item.badge.total}`}</span>}
      </button>)}
    </div>
    <button type="button" className="dash-view-new" onClick={onNew} disabled={newDisabled} aria-label="New dashboard" title="New dashboard">{icons.plus}</button>
  </div>
}

export type AIComposerConfig = {
  title?: string
  description?: string
  promptLabel?: string
  placeholder?: string
  suggestions?: string[]
  submitLabel?: string
  footnote?: string
}

const sparkle = <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" aria-hidden="true"><path d="m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5zM20 2v5M17.5 4.5h5" /></svg>

// A controlled native dialog supplies focus containment and an inert background.
// Closing or stopping work remains the caller's decision through callbacks.
export function AIComposer({ open, request, config = {}, working = false, error, disabled = false, onRequestChange, onGenerate, onCancel, onStop }: {
  open: boolean
  request: string
  config?: AIComposerConfig
  working?: boolean
  error?: string
  disabled?: boolean
  onRequestChange: (value: string) => void
  onGenerate: () => void
  onCancel: () => void
  onStop?: () => void
}) {
  const dialog = useRef<HTMLDialogElement>(null)
  const prompt = useRef<HTMLTextAreaElement>(null)
  const body = useRef<HTMLDivElement>(null)
  const before = useRef<HTMLElement | null>(null)
  const titleId = useId(), descriptionId = useId(), promptId = useId()
  useEffect(() => {
    const node = dialog.current
    if (!node) return
    if (open && !node.open) {
      before.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
      node.showModal()
      prompt.current?.focus({ preventScroll: true })
      node.scrollTop = 0
      if (body.current) body.current.scrollTop = 0
    } else if (!open && node.open) {
      node.close()
      if (before.current?.isConnected) before.current.focus()
      before.current = null
    }
  }, [open])
  useEffect(() => {
    const node = dialog.current
    return () => {
      if (!node?.open) return
      node.close()
      if (before.current?.isConnected) before.current.focus()
      before.current = null
    }
  }, [])
  return <dialog ref={dialog} className="dash-ai-composer" aria-labelledby={titleId} aria-describedby={descriptionId} onCancel={(event) => { event.preventDefault(); onCancel() }}>
    <form onSubmit={(event) => { event.preventDefault(); if (request.trim() && !working && !disabled) onGenerate() }}>
      <header className="dash-ai-composer-head"><span className="dash-ai-composer-icon">{sparkle}</span><div><h2 id={titleId}>{config.title ?? "New dashboard"}</h2><p id={descriptionId}>{config.description ?? "Tell your AI what you want to see. It will choose widgets and arrange a view for you."}</p></div><button type="button" className="dash-ai-composer-close" onClick={onCancel} aria-label="Close dashboard composer">{icons.close}</button></header>
      <div ref={body} className="dash-ai-composer-body">
      <div className="dash-ai-composer-field"><label htmlFor={promptId}>{config.promptLabel ?? "Describe what you want"}</label><textarea ref={prompt} id={promptId} value={request} onChange={(event) => onRequestChange(event.target.value)} placeholder={config.placeholder ?? "For example, show spend by client and conversations over time…"} rows={3} maxLength={2000} disabled={working || disabled} /></div>
      {!!config.suggestions?.length && <div className="dash-ai-composer-ideas"><span>Try an idea</span><div>{config.suggestions.map((suggestion) => <button key={suggestion} type="button" disabled={working || disabled} onClick={() => { onRequestChange(suggestion); prompt.current?.focus({ preventScroll: true }) }}>{suggestion}</button>)}</div></div>}
      {working && <p className="dash-ai-composer-status" role="status">Your AI is choosing widgets and chart settings…</p>}
      {error && <p className="dash-ai-composer-error" role="alert">{error}</p>}
      </div>
      <footer className="dash-ai-composer-foot"><p>{config.footnote ?? "Pinned widgets stay in place. Figures come from your matter data."}</p><div>{working && onStop ? <button type="button" className="dash-ai-composer-cancel" onClick={() => { onStop(); requestAnimationFrame(() => { prompt.current?.focus({ preventScroll: true }); if (dialog.current) dialog.current.scrollTop = 0; if (body.current) body.current.scrollTop = 0 }) }}>Stop</button> : <button type="button" className="dash-ai-composer-cancel" onClick={onCancel}>Cancel</button>}<button type="submit" className="dash-ai-composer-generate" disabled={!request.trim() || working || disabled} aria-busy={working}>{sparkle}{working ? "Generating…" : config.submitLabel ?? "Generate dashboard"}</button></div></footer>
    </form>
  </dialog>
}

export type AdaptiveGridItem = { id: string; size?: "tile" | "wide" | "full" }

// Measured row spans keep natural card heights while retaining one DOM sequence.
// The 1px row unit avoids the large empty tracks beside a taller neighboring card.
export function AdaptiveGrid({ items, renderItem, mode = "adaptive", label }: { items: AdaptiveGridItem[]; renderItem: (id: string) => ReactNode; mode?: "adaptive" | "catalog"; label?: string }) {
  const grid = useRef<HTMLDivElement>(null)
  const [heights, setHeights] = useState<Record<string, number>>(() => Object.create(null))
  const keys = items.map((item) => item.id).join("\u0000")
  useLayoutEffect(() => {
    const node = grid.current
    if (!node || mode === "catalog") return
    let frame = 0
    const measure = () => {
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => {
        const next: Record<string, number> = Object.create(null)
        node.querySelectorAll<HTMLElement>(".dash-adaptive-natural").forEach((item) => { next[item.dataset.gridItem!] = item.offsetHeight + 16 })
        setHeights((previous) => Object.keys(next).length === Object.keys(previous).length && Object.entries(next).every(([key, height]) => previous[key] === height) ? previous : next)
      })
    }
    const observer = new ResizeObserver(measure)
    node.querySelectorAll<HTMLElement>(".dash-adaptive-natural").forEach((item) => observer.observe(item))
    measure()
    return () => { observer.disconnect(); cancelAnimationFrame(frame) }
  }, [keys, mode])
  return <div className={`dash-adaptive-grid is-${mode}`} ref={grid} aria-label={label}>
    {items.map((item) => <div key={item.id} className={`dash-adaptive-item size-${item.size ?? "wide"}`} style={{ "--dash-item-height": heights[item.id] ?? 1 } as CSSProperties}><div className="dash-adaptive-natural" data-grid-item={item.id}>{renderItem(item.id)}</div></div>)}
  </div>
}

export type AboutViewConfig = { title?: string; text: string; actionLabel?: string }
export function AboutView({ config, onRefine }: { config: AboutViewConfig; onRefine?: () => void }) {
  return <div className="dash-about-view"><span className="dash-about-view-icon" aria-hidden="true">{icons.question}</span><p><strong>{config.title ?? "About this view"}:</strong> {config.text}</p>{onRefine && <button type="button" onClick={onRefine}>{config.actionLabel ?? "Refine view"}</button>}</div>
}
