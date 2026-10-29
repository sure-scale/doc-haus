import { useEffect, useMemo, useRef, useState, type FormEvent } from "react"
import { createPortal } from "react-dom"
import type { DraftingPreferences } from "../api/ingest/preferences"
import type { SetupProgress } from "../api/ingest/setup"
import { askDashboard } from "../api/opencode/dashboard"
import PageBar from "../components/PageBar"
import { addDays, attention, PERIOD_LABELS, PERIODS, startOfDay, summarize, useDashboardData, type DashboardData, type Period } from "../components/dashboard/data"
import { ChartWidget } from "../components/dashboard/chart-widget"
import { icons, ShapeIcon } from "../components/dashboard/icons"
import { applyPlan, loadDashboards, PERIODIC, saveDashboards, tabWidgets, WIDGETS, type Dashboards, type DashboardTab, type WidgetInstance, type WidgetId } from "../components/dashboard/layout"
import { caption, LOOKS, RENDER, setupComplete, stageOf, type WidgetContext } from "../components/dashboard/widgets"

function greeting(now: number) {
  const hour = new Date(now).getHours()
  return hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening"
}

const sparkle = <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" aria-hidden="true"><path d="m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5zM20 2v5M17.5 4.5h5" /></svg>
const pin = <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" aria-hidden="true"><path d="m8 3 8 0-1 7 4 4H5l4-4zM12 14v7" /></svg>
const suggestions = ["Show my spend by client", "Give me a weekly overview", "Focus on drafts that need me"]
const titleOf = ({ widget }: WidgetInstance) => widget.type === "chart" ? widget.title : WIDGETS[widget.type].title

// Setup is re-read by App whenever Settings or onboarding closes.
export default function Dashboard({ setup, refresh = 0 }: { setup: SetupProgress | null; refresh?: number }) {
  const { data, loading, reload } = useDashboardData(setup, refresh)
  const [workspace, setWorkspace] = useState<Dashboards>(loadDashboards)
  const [promptOpen, setPromptOpen] = useState(false)
  const [request, setRequest] = useState("")
  const [working, setWorking] = useState(false)
  const [error, setError] = useState("")
  const [storageFailed, setStorageFailed] = useState(false)
  const [announcement, setAnnouncement] = useState("")
  const [undoState, setUndoState] = useState<{ text: string; value: Dashboards }>()
  const [rename, setRename] = useState("")
  const [renaming, setRenaming] = useState(false)
  const [deleting, setDeleting] = useState(false)
  const askButton = useRef<HTMLButtonElement>(null)
  const prompt = useRef<HTMLTextAreaElement>(null)
  const tabMenu = useRef<HTMLDetailsElement>(null)
  const controller = useRef<AbortController>(null)
  const workspaceRef = useRef(workspace)
  workspaceRef.current = workspace
  const active = workspace.dashboards.find((tab) => tab.id === workspace.active) ?? workspace.dashboards[0]!
  const notes = active.notes ?? []
  const now = useMemo(() => Date.now(), [data])
  useEffect(() => {
    const timer = setTimeout(reload, addDays(startOfDay(now), 1) - Date.now() + 1000)
    return () => clearTimeout(timer)
  }, [now])
  useEffect(() => () => controller.current?.abort(), [])
  useEffect(() => {
    const outside = (event: PointerEvent) => { if (tabMenu.current && !tabMenu.current.contains(event.target as Node)) tabMenu.current.open = false }
    document.addEventListener("pointerdown", outside)
    return () => document.removeEventListener("pointerdown", outside)
  }, [])
  useEffect(() => { if (promptOpen) prompt.current?.focus() }, [promptOpen])
  useEffect(() => {
    document.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]')?.scrollIntoView({ block: "nearest", inline: "nearest" })
  }, [active.id])

  const [saved, setSaved] = useState<{ data?: DashboardData; preferences?: DraftingPreferences }>({})
  const latest = useRef(data)
  latest.current = data
  const shownData = data && saved.data === data ? { ...data, preferences: saved.preferences } : data
  const ready = data && !data.error ? data : undefined
  const context: WidgetContext | undefined = shownData && ready
    ? { data: shownData, setup, period: active.period, now, summary: summarize(shownData.matters, active.period, now), onPreferences: (preferences) => setSaved({ data: latest.current, preferences }), onRetry: reload }
    : undefined
  const setupDone = useMemo(() => !!ready && setupComplete(ready), [ready])
  const stage = ready ? stageOf(ready) : "starting"
  const widgets = tabWidgets(active, stage, setupDone)

  function update(next: Dashboards) {
    // A later edit starts a new undo boundary, so Undo never erases newer work.
    setUndoState(undefined)
    workspaceRef.current = next
    setWorkspace(next)
    setStorageFailed(!saveDashboards(next))
  }
  function changeTab(next: DashboardTab) {
    update({ ...workspace, dashboards: workspace.dashboards.map((tab) => tab.id === next.id ? next : tab) })
  }
  function select(id: string) {
    if (id === active.id) return
    controller.current?.abort()
    setWorking(false)
    setError("")
    setRenaming(false)
    setDeleting(false)
    update({ ...workspace, active: id })
  }
  function create() {
    if (workspace.dashboards.length >= 12) return
    const tab: DashboardTab = { id: crypto.randomUUID(), name: "New dashboard", period: active.period, stages: {}, widgets: [] }
    update({ active: tab.id, dashboards: [...workspace.dashboards, tab] })
    controller.current?.abort()
    setWorking(false)
    setRequest("")
    setError("")
    setRenaming(false)
    setDeleting(false)
    setPromptOpen(true)
    requestAnimationFrame(() => prompt.current?.focus())
  }
  function setWidgets(next: WidgetInstance[]) {
    changeTab(active.widgets ? { ...active, widgets: next } : { ...active, stages: { ...active.stages, [stage]: next } })
  }
  function togglePin(key: string) {
    const item = widgets.find((widget) => widget.key === key)!
    setWidgets(widgets.map((widget) => widget.key === key ? { ...widget, pinned: !widget.pinned } : widget))
    setAnnouncement(`${titleOf(item)} ${item.pinned ? "unpinned" : "pinned"}`)
  }
  function remove(key: string) {
    const index = widgets.findIndex((widget) => widget.key === key)
    const item = widgets[index]!
    setWidgets(widgets.filter((widget) => widget.key !== key))
    setUndoState({ text: `${titleOf(item)} removed`, value: workspace })
    setAnnouncement(`${titleOf(item)} removed`)
    requestAnimationFrame(() => {
      const buttons = document.querySelectorAll<HTMLButtonElement>(".dash-pin")
      ;(buttons[Math.min(index, buttons.length - 1)] ?? askButton.current)?.focus()
    })
  }
  function undo() {
    if (!undoState) return
    controller.current?.abort()
    setWorking(false)
    update(undoState.value)
    setUndoState(undefined)
    setAnnouncement("Dashboard change undone")
    requestAnimationFrame(() => askButton.current?.focus())
  }
  function deleteTab() {
    const dashboards = workspace.dashboards.filter((tab) => tab.id !== active.id)
    controller.current?.abort()
    setWorking(false)
    update({ active: dashboards[0]!.id, dashboards })
    setUndoState({ text: `${active.name} deleted`, value: workspace })
    setDeleting(false)
    setError("")
    setAnnouncement("Dashboard deleted")
    requestAnimationFrame(() => document.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]')?.focus())
  }
  async function generate(event: FormEvent) {
    event.preventDefault()
    if (!request.trim() || !ready || working) return
    const abort = new AbortController()
    controller.current?.abort()
    controller.current = abort
    const tabId = active.id
    setWorking(true)
    setError("")
    try {
      const own = ready.matters.filter((row) => !row.matter.sample)
      const plan = await askDashboard({ request: request.trim(), current: widgets.map(({ widget, pinned }) => ({ widget, pinned })), clients: [...new Set(own.flatMap((row) => row.matter.client ? [row.matter.client] : []))], matters: own.map((row) => row.matter.title), signal: abort.signal })
      if (abort.signal.aborted || workspaceRef.current.active !== tabId) return
      const current = workspaceRef.current
      const tab = current.dashboards.find((entry) => entry.id === tabId)
      if (!tab) return
      const next = { ...tab, name: tab.name === "New dashboard" && plan.name ? plan.name : tab.name, widgets: applyPlan(tabWidgets(tab, stage, setupDone), plan), notes: plan.unsupported }
      update({ ...current, dashboards: current.dashboards.map((entry) => entry.id === tabId ? next : entry) })
      setUndoState({ text: "Dashboard updated", value: current })
      setAnnouncement("Dashboard updated")
    } catch (cause) {
      if (!abort.signal.aborted) setError(cause instanceof Error ? cause.message : "Could not update the dashboard. Try again.")
    } finally {
      if (controller.current === abort) { controller.current = null; setWorking(false) }
    }
  }

  const periodic = !data?.error && widgets.some(({ widget }) => widget.type === "chart" ? widget.period === "dashboard" : PERIODIC.has(widget.type))
  const unread = ready?.matters.filter((row) => row.unread).length ?? 0
  const bold: WidgetId | undefined = ready && widgets.some(({ widget }) => widget.type === "attention") && attention(ready.matters).length ? "attention" : widgets.some(({ widget }) => widget.type === "journey") ? "journey" : undefined
  const name = shownData?.preferences?.attorney.trim()

  return <>
    <PageBar title="Dashboard" description={`${greeting(now)}${name ? `, ${name}` : ""}. What needs you, your drafts and what your matters cost.`} />
    <div className="dash">
      <div className="dash-tabs-row">
        <div className="dash-tabs" role="tablist" aria-label="Dashboards" onKeyDown={(event) => {
          if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return
          event.preventDefault()
          const index = workspace.dashboards.findIndex((tab) => tab.id === active.id)
          const next = event.key === "Home" ? 0 : event.key === "End" ? workspace.dashboards.length - 1 : (index + (event.key === "ArrowRight" ? 1 : -1) + workspace.dashboards.length) % workspace.dashboards.length
          select(workspace.dashboards[next]!.id)
          requestAnimationFrame(() => document.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]')?.focus())
        }}>
          {workspace.dashboards.map((tab) => <button type="button" role="tab" id={`dash-tab-${tab.id}`} aria-controls="dash-content" aria-selected={active.id === tab.id} tabIndex={active.id === tab.id ? 0 : -1} key={tab.id} onClick={() => select(tab.id)}>{tab.name}</button>)}
        </div>
        <button type="button" className="dash-new" onClick={create} disabled={workspace.dashboards.length >= 12} aria-label="New dashboard" title="New dashboard">{icons.plus}</button>
        <details ref={tabMenu} className="viewer-more-actions dash-tab-menu" key={active.id} onKeyDown={(event) => {
          if (event.key === "Escape" && event.currentTarget.open) { event.preventDefault(); event.currentTarget.open = false; event.currentTarget.querySelector("summary")?.focus() }
        }} onBlur={(event) => { if (event.relatedTarget && !event.currentTarget.contains(event.relatedTarget as Node)) event.currentTarget.open = false }}>
          <summary aria-label={`Options for ${active.name} dashboard`} title="Dashboard options">{icons.edit}</summary>
          <div className="viewer-more-menu">
            <button type="button" onClick={(event) => { event.currentTarget.closest("details")!.open = false; setRename(active.name); setRenaming(true); setDeleting(false) }}>Rename</button>
            <button type="button" disabled={workspace.dashboards.length === 1} onClick={(event) => { event.currentTarget.closest("details")!.open = false; setDeleting(true); setRenaming(false); requestAnimationFrame(() => document.querySelector<HTMLButtonElement>(".dash-delete-confirm button")?.focus()) }}>Delete dashboard</button>
          </div>
        </details>
      </div>
      {renaming && <form className="dash-tab-form" onSubmit={(event) => { event.preventDefault(); if (rename.trim()) { changeTab({ ...active, name: rename.trim() }); setRenaming(false); setAnnouncement("Dashboard renamed"); requestAnimationFrame(() => tabMenu.current?.querySelector("summary")?.focus()) } }}>
        <label htmlFor="dash-name">Dashboard name</label><input id="dash-name" value={rename} maxLength={30} onChange={(event) => setRename(event.target.value)} autoFocus required />
        <button type="submit" disabled={!rename.trim()}>Save</button><button type="button" className="ghost" onClick={() => { setRenaming(false); requestAnimationFrame(() => tabMenu.current?.querySelector("summary")?.focus()) }}>Cancel</button>
      </form>}
      {deleting && <div className="dash-delete-confirm" role="group" aria-label="Delete dashboard">
        <p>Delete “{active.name}”? Your matters stay as they are.</p><button type="button" onClick={deleteTab}>Delete dashboard</button><button type="button" className="ghost" onClick={() => { setDeleting(false); requestAnimationFrame(() => tabMenu.current?.querySelector("summary")?.focus()) }}>Cancel</button>
      </div>}
      <div className="dash-toolbar">
        {periodic && <div className="settings-seg dash-period" role="group" aria-label="Period">{PERIODS.map((period: Period) => <button key={period} type="button" aria-pressed={active.period === period} className={active.period === period ? "active" : ""} onClick={() => changeTab({ ...active, period })}>
          {active.period === period && <svg className="settings-seg-check" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12" /></svg>}{PERIOD_LABELS[period]}
        </button>)}</div>}
        <div className="dash-controls">
          <button type="button" className="icon-btn dash-refresh" onClick={reload} disabled={loading} aria-busy={loading} aria-label="Refresh" title="Refresh">{icons.refresh}</button>
          <button ref={askButton} type="button" className={`dash-ask${promptOpen ? " is-open" : ""}`} onClick={() => { if (promptOpen && working) { controller.current?.abort(); setWorking(false); setAnnouncement("Dashboard update stopped") } setPromptOpen(!promptOpen) }} disabled={!ready} aria-expanded={promptOpen} aria-controls="dash-ai-panel">{sparkle}Ask AI</button>
        </div>
      </div>
      {promptOpen && <form id="dash-ai-panel" className="dash-ai-panel" onSubmit={generate} onKeyDown={(event) => {
        if (event.key !== "Escape") return
        event.preventDefault()
        controller.current?.abort()
        setWorking(false)
        setPromptOpen(false)
        if (working) setAnnouncement("Dashboard update stopped")
        requestAnimationFrame(() => askButton.current?.focus())
      }}>
        <div className="dash-ai-heading"><span className="dash-ai-mark">{sparkle}</span><div><label htmlFor="dash-request">Tell AI what you want from this dashboard</label><p>Your AI arranges the widgets. Your app works out the figures.</p></div></div>
        <div className="dash-ai-entry"><textarea ref={prompt} id="dash-request" value={request} onChange={(event) => setRequest(event.target.value)} placeholder="For example, show my spend by client and the drafts that need me…" rows={2} maxLength={2000} disabled={working} /><button type="submit" className="dash-ai-submit" disabled={!request.trim() || working || !ready}>{sparkle}{working ? "Building dashboard…" : "Update dashboard"}</button></div>
        <div className="dash-ai-suggestions" aria-label="Dashboard ideas">{suggestions.map((suggestion) => <button type="button" key={suggestion} disabled={working} onClick={() => { setRequest(suggestion); prompt.current?.focus() }}>{suggestion}</button>)}</div>
        <div className="dash-ai-foot"><p>Pinned widgets stay in place. Uses your fast AI model.</p>{working && <button type="button" className="dash-text-button" onClick={() => { controller.current?.abort(); setWorking(false); setAnnouncement("Dashboard update stopped"); requestAnimationFrame(() => prompt.current?.focus()) }}>Stop</button>}</div>
        {working && <p role="status" className="dash-ai-working">Your AI is choosing widgets and chart settings…</p>}
        {error && <p role="alert" className="dash-ai-error">{error}</p>}
      </form>}
      {notes.length > 0 && <div className="dash-ai-notes" role="status"><strong>About this dashboard</strong><ul>{notes.map((note) => <li key={note}>{note}</li>)}</ul></div>}
      {storageFailed && <p className="dash-ai-error" role="status">Your browser could not save this layout. It will last while this page stays open.</p>}
      {unread > 0 && <div className="dash-banner" role="status">{icons.warning}<p>Some matters could not be read, so their spend and drafts are missing here.</p><button type="button" className="btn-sm" onClick={reload} disabled={loading}>Retry</button></div>}
      <div id="dash-content" role="tabpanel" aria-labelledby={`dash-tab-${active.id}`} tabIndex={0}>
        {data?.error ? <div className="dash-error" role="alert"><ShapeIcon icon="warning" shape="flower" /><div><p className="dash-error-title">Could not load your matters</p><p>Your figures will show once the matter list loads.</p></div><button type="button" onClick={reload} disabled={loading}>Retry</button></div>
          : widgets.length === 0 ? <div className="dash-blank"><ShapeIcon icon="plus" shape="sunny" className="dash-blank-badge" /><p className="dash-blank-title">A dashboard for the way you work</p><p>Tell your AI what you want to keep an eye on.</p><button type="button" disabled={!ready} onClick={() => { setPromptOpen(true); requestAnimationFrame(() => prompt.current?.focus()) }}>{sparkle}Ask AI to build it</button></div>
            : <div className="dash-grid" aria-busy={!context}>{widgets.map((item) => <Widget key={item.key} item={item} context={context} period={active.period} bold={item.widget.type === bold} onPin={() => togglePin(item.key)} onRemove={() => remove(item.key)} />)}</div>}
      </div>
    </div>
    <p className="sr-only" role="status">{announcement}</p>
    {undoState && <Snackbar key={undoState.text} text={undoState.text} onUndo={undo} onDismiss={() => { if (document.activeElement?.closest(".dash-snackbar")) askButton.current?.focus(); setUndoState(undefined) }} />}
  </>
}

function Widget({ item, context, period, bold, onPin, onRemove }: { item: WidgetInstance; context?: WidgetContext; period: Period; bold: boolean; onPin: () => void; onRemove: () => void }) {
  const { widget } = item
  const chart = widget.type === "chart"
  const id = chart ? undefined : widget.type
  const meta = id ? WIDGETS[id] : undefined
  const look = id ? LOOKS[id] : { icon: "trend" as const, shape: "sunny" as const, tone: undefined }
  const title = titleOf(item)
  const heading = `dash-${encodeURIComponent(item.key)}`
  const note = id ? caption(id, period) : "From your matter data"
  return <section className={`dash-card size-${meta?.size ?? (chart && widget.form === "stat" ? "tile" : chart && widget.form === "calendar" ? "full" : "wide")}${look.tone ? ` tone-${look.tone}` : ""}${bold ? " is-bold" : ""}${item.pinned ? " is-pinned" : ""}`} aria-labelledby={heading}>
    <header className="dash-card-head"><ShapeIcon icon={look.icon} shape={look.shape} className="dash-card-badge" /><div className="dash-card-heading"><h2 id={heading}>{title}</h2>{note && <p>{note}</p>}</div><div className="dash-widget-actions">
      <button type="button" className="dash-pin" aria-label={`${item.pinned ? "Unpin" : "Pin"} ${title}`} aria-pressed={item.pinned} title={item.pinned ? "Unpin widget" : "Keep this widget in place"} onClick={onPin}>{pin}</button>
      <button type="button" className="dash-remove" aria-label={`Remove ${title}`} title="Remove widget" onClick={onRemove}>{icons.close}</button>
    </div></header>
    <div className="dash-card-body">{context ? chart ? <ChartWidget chart={widget} data={context.data} period={period} /> : id ? RENDER[id](context) : null : <div className="dash-skeleton" aria-hidden="true"><span /><span /><span /></div>}</div>
  </section>
}

function Snackbar({ text, onUndo, onDismiss }: { text: string; onUndo: () => void; onDismiss: () => void }) {
  const [held, setHeld] = useState(false)
  useEffect(() => { if (held) return; const timer = setTimeout(onDismiss, 8000); return () => clearTimeout(timer) }, [held, onDismiss])
  return createPortal(<div className="dash-snackbar" onPointerEnter={() => setHeld(true)} onPointerLeave={() => setHeld(false)} onFocus={() => setHeld(true)} onBlur={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setHeld(false) }}><span>{text}</span><button type="button" className="dash-snackbar-action" onClick={onUndo}>Undo</button><button type="button" className="dash-snackbar-close" onClick={onDismiss} aria-label="Dismiss" title="Dismiss">{icons.close}</button></div>, document.body)
}
