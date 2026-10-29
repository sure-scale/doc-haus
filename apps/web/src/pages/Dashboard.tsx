import { useEffect, useMemo, useRef, useState } from "react"
import { createPortal } from "react-dom"
import type { DraftingPreferences } from "../api/ingest/preferences"
import type { SetupProgress } from "../api/ingest/setup"
import { askDashboard } from "../api/opencode/dashboard"
import { useDashboardTabs } from "../components/dashboard/useDashboardTabs"
import { SetupTab, setupTabProgress } from "../components/dashboard/SetupTab"
import { AdaptiveGrid, AIComposer, AboutView, DashboardTabs } from "../components/dashboard/visual/DashboardChrome"
import { addDays, attention, PERIOD_LABELS, PERIODS, startOfDay, summarize, useDashboardData, type DashboardData, type Period } from "../components/dashboard/data"
import { ChartWidget } from "../components/dashboard/chart-widget"
import { icons, ShapeIcon } from "../components/dashboard/icons"
import { applyPlan, PERIODIC, tabWidgets, WIDGETS, type Dashboards, type DashboardTab, type WidgetInstance, type WidgetId } from "../components/dashboard/layout"
import { caption, LOOKS, RENDER, setupComplete, stageOf, type WidgetContext } from "../components/dashboard/widgets"

const sparkle = <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" aria-hidden="true"><path d="m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5zM20 2v5M17.5 4.5h5" /></svg>
const pin = <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" aria-hidden="true"><path d="m8 3 8 0-1 7 4 4H5l4-4zM12 14v7" /></svg>
const suggestions = ["Show my spend by client", "Give me a weekly overview", "Focus on drafts that need me"]
const describeRequest = (request: string) => request.replace(/[\s\p{Cc}]+/gu, " ").trim().slice(0, 240)
const titleOf = ({ widget }: WidgetInstance) => widget.type === "chart" ? widget.title : WIDGETS[widget.type].title

// Setup is re-read by App whenever Settings or onboarding closes.
export default function Dashboard({ setup, refresh = 0, onSampleAdded, onSetupRefresh }: { setup: SetupProgress | null; refresh?: number; onSampleAdded?: () => void; onSetupRefresh?: () => void }) {
  const { data, loading, reload } = useDashboardData(setup, refresh)
  const stored = useDashboardTabs()
  const workspace = stored.state.workspace
  // An explicit Setup link can open it without changing the stored visibility.
  const [forcedSetup, setForcedSetup] = useState(() => new URLSearchParams(window.location.search).get("tab") === "setup")
  const setupHidden = stored.state.setupHidden && !forcedSetup
  const showSetup = forcedSetup || stored.state.view === "setup"
  // Keep Overview's saved arrangement for later design work, out of the tab strip.
  const customTabs = workspace.dashboards.filter((tab) => tab.id !== "overview")
  const empty = !showSetup && customTabs.length === 0
  const [newOpen, setNewOpen] = useState(false)
  const [newRequest, setNewRequest] = useState("")
  const [newWorking, setNewWorking] = useState(false)
  const [newError, setNewError] = useState("")
  const newController = useRef<AbortController | null>(null)
  const [promptOpen, setPromptOpen] = useState(false)
  const [request, setRequest] = useState("")
  const [working, setWorking] = useState(false)
  const [error, setError] = useState("")
  const [announcement, setAnnouncement] = useState("")
  const [undoState, setUndoState] = useState<{ text: string; value: Dashboards }>()
  const [rename, setRename] = useState("")
  const [renaming, setRenaming] = useState(false)
  const [deleting, setDeleting] = useState(false)
  const tabMenu = useRef<HTMLDetailsElement>(null)
  const controller = useRef<AbortController>(null)
  const workspaceRef = useRef(workspace)
  workspaceRef.current = workspace
  const active: DashboardTab = customTabs.find((tab) => tab.id === workspace.active) ?? customTabs[0] ?? workspace.dashboards[0] ?? { id: "", name: "Dashboard", period: "month" as const, stages: {} }
  const notes = active.notes ?? []
  const now = useMemo(() => Date.now(), [data])
  useEffect(() => {
    const timer = setTimeout(reload, addDays(startOfDay(now), 1) - Date.now() + 1000)
    return () => clearTimeout(timer)
  }, [now])
  useEffect(() => () => { controller.current?.abort(); newController.current?.abort() }, [])
  useEffect(() => {
    const outside = (event: PointerEvent) => { if (tabMenu.current && !tabMenu.current.contains(event.target as Node)) tabMenu.current.open = false }
    document.addEventListener("pointerdown", outside)
    return () => document.removeEventListener("pointerdown", outside)
  }, [])
  useEffect(() => {
    document.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]')?.scrollIntoView({ block: "nearest", inline: "nearest" })
  }, [active.id, showSetup])

  const [saved, setSaved] = useState<{ data?: DashboardData; preferences?: DraftingPreferences }>({})
  const latest = useRef(data)
  latest.current = data
  const shownData = data && saved.data === data ? { ...data, preferences: saved.preferences } : data
  const ready = data && !data.error ? data : undefined
  const context: WidgetContext | undefined = shownData
    ? { data: shownData, setup, period: active.period, now, summary: summarize(shownData.matters, active.period, now), onPreferences: (preferences) => setSaved({ data: latest.current, preferences }), onRetry: reload }
    : undefined
  const canGenerate = !!ready && setup?.ready === true
  const setupProgress = shownData ? setupTabProgress(shownData, setup) : undefined
  const setupDone = useMemo(() => !!ready && setupComplete(ready), [ready])
  const stage = ready ? stageOf(ready) : "starting"
  const widgets = tabWidgets(active, stage, setupDone)

  function focusDashboardAction() {
    ;(document.querySelector<HTMLButtonElement>(".dash-about-view button, .dash-blank button:not(:disabled)") ?? tabMenu.current?.querySelector<HTMLElement>("summary"))?.focus()
  }
  function stopRefine() {
    controller.current?.abort()
    controller.current = null
    setWorking(false)
    setAnnouncement("Dashboard update stopped")
  }
  function cancelRefine() { stopRefine(); setPromptOpen(false) }

  function update(next: Dashboards, extra: { setupHidden?: boolean; view?: "setup" | "dashboard" } = {}) {
    // A later edit starts a new undo boundary, so Undo never erases newer work.
    setUndoState(undefined)
    workspaceRef.current = next
    stored.change((latest) => ({ ...latest, workspace: next, ...extra }))
  }
  function changeTab(next: DashboardTab) {
    update({ ...workspace, dashboards: workspace.dashboards.map((tab) => tab.id === next.id ? next : tab) })
  }
  function select(id: string) {
    const isSetup = id === "__setup"
    if (isSetup === showSetup && (isSetup || id === active.id)) return
    setForcedSetup(false)
    setPromptOpen(false)
    controller.current?.abort()
    setWorking(false)
    setError("")
    setRenaming(false)
    setDeleting(false)
    const url = new URL(window.location.href)
    if (isSetup) url.searchParams.set("tab", "setup")
    else url.searchParams.delete("tab")
    window.history.replaceState(window.history.state, "", url)
    update(isSetup ? workspace : { ...workspace, active: id }, { view: isSetup ? "setup" : "dashboard", ...(isSetup ? { setupHidden: false } : {}) })
  }
  function closeTab(id: string) {
    if (id !== "__setup") { deleteTab(id); return }
    setForcedSetup(false)
    const url = new URL(window.location.href)
    url.searchParams.delete("tab")
    window.history.replaceState(window.history.state, "", url)
    update(workspace, { setupHidden: true, view: "dashboard" })
    setAnnouncement("Setup tab closed. Your settings are saved.")
    requestAnimationFrame(() => (document.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]') ?? document.querySelector<HTMLButtonElement>('[aria-label="New dashboard"]'))?.focus())
  }
  function create() {
    if (workspaceRef.current.dashboards.length >= 12) return
    controller.current?.abort()
    setWorking(false)
    setNewError("")
    setNewOpen(true)
  }
  function stopNew() {
    newController.current?.abort()
    newController.current = null
    setNewWorking(false)
  }
  function cancelNew() {
    stopNew()
    setNewOpen(false)
  }
  async function generateNew() {
    if (!newRequest.trim() || !ready || !canGenerate || newWorking) return
    if (workspaceRef.current.dashboards.length >= 12) { setNewError("You can have up to 12 dashboard tabs. Close a tab before adding another."); return }
    const abort = new AbortController()
    newController.current?.abort()
    newController.current = abort
    setNewWorking(true)
    setNewError("")
    try {
      const own = ready.matters.filter(row => !row.matter.sample)
      const plan = await askDashboard({ request: newRequest.trim(), current: [], clients: [...new Set(own.flatMap(row => row.matter.client ? [row.matter.client] : []))], matters: own.map(row => row.matter.title), signal: abort.signal })
      if (abort.signal.aborted || newController.current !== abort) return
      const current = workspaceRef.current
      if (current.dashboards.length >= 12) { setNewError("You can have up to 12 dashboard tabs. Close a tab before adding another."); return }
      const tab: DashboardTab = { id: crypto.randomUUID(), name: plan.name || "My dashboard", period: active.period, stages: {}, widgets: applyPlan([], plan), notes: plan.unsupported, description: describeRequest(newRequest) }
      setForcedSetup(false)
      const url = new URL(window.location.href)
      url.searchParams.delete("tab")
      window.history.replaceState(window.history.state, "", url)
      update({ active: tab.id, dashboards: [...current.dashboards, tab] }, { view: "dashboard" })
      setPromptOpen(false)
      setRenaming(false)
      setDeleting(false)
      setNewOpen(false)
      setNewRequest("")
      setAnnouncement("Dashboard created")
    } catch (cause) {
      if (!abort.signal.aborted && newController.current === abort) setNewError(cause instanceof Error ? cause.message : "Could not create the dashboard. Your request is still here. Try again.")
    } finally {
      if (newController.current === abort) { newController.current = null; setNewWorking(false) }
    }
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
      ;(buttons[Math.min(index, buttons.length - 1)])?.focus()
      if (!buttons.length) focusDashboardAction()
    })
  }
  function undo() {
    if (!undoState) return
    controller.current?.abort()
    setWorking(false)
    const restored = undoState.value
    const hasCustom = restored.dashboards.some((tab) => tab.id === restored.active && tab.id !== "overview")
    update(restored, hasCustom ? { view: "dashboard" } : {})
    if (restored.dashboards.some((tab) => tab.id === restored.active && tab.id !== "overview")) {
      setForcedSetup(false)
      const url = new URL(window.location.href)
      url.searchParams.delete("tab")
      window.history.replaceState(window.history.state, "", url)
    }
    setUndoState(undefined)
    setAnnouncement("Dashboard change undone")
    requestAnimationFrame(focusDashboardAction)
  }
  function deleteTab(id = active.id) {
    const removed = workspace.dashboards.find((tab) => tab.id === id)
    if (!removed || id === "overview") return
    const dashboards = workspace.dashboards.filter((tab) => tab.id !== id)
    const remaining = dashboards.filter((tab) => tab.id !== "overview")
    const wasActive = id === active.id && !showSetup
    if (wasActive) {
      controller.current?.abort()
      setWorking(false)
      setPromptOpen(false)
      setError("")
    }
    const nextActive = workspace.active !== id && dashboards.some(tab => tab.id === workspace.active) ? workspace.active : remaining[0]?.id ?? dashboards[0]?.id ?? ""
    update({ active: nextActive, dashboards }, wasActive && !remaining.length ? { view: setupHidden ? "dashboard" : "setup" } : {})
    setUndoState({ text: `${removed.name} deleted`, value: workspace })
    setDeleting(false)
    setAnnouncement("Dashboard deleted")
    requestAnimationFrame(() => document.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]')?.focus())
  }
  async function generate() {
    if (!request.trim() || !ready || !canGenerate || working) return
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
      const next = { ...tab, name: tab.name === "New dashboard" && plan.name ? plan.name : tab.name, widgets: applyPlan(tabWidgets(tab, stage, setupDone), plan), notes: plan.unsupported, description: describeRequest(request) }
      update({ ...current, dashboards: current.dashboards.map((entry) => entry.id === tabId ? next : entry) })
      setUndoState({ text: "Dashboard updated", value: current })
      setPromptOpen(false)
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

  if (!stored.ready) return <>
    <h1 className="sr-only">Dashboard</h1>
    <div className="dash"><div className="dash-error" role={stored.error ? "alert" : "status"}><ShapeIcon icon={stored.error ? "warning" : "setup"} shape="sunny" /><div><p className="dash-error-title">{stored.error ? "Could not load your dashboards" : "Loading your dashboards…"}</p>{stored.error && <p>Check the document service, then retry. Your saved dashboards have not been replaced.</p>}</div>{stored.error && <button type="button" disabled={stored.loading} onClick={() => void stored.retry()}>Retry dashboards</button>}</div></div>
  </>

  return <>
    <h1 className="sr-only">Dashboard</h1>
    <div className="dash" inert={stored.loading}>
      <div className="dash-tabs-row">
        <DashboardTabs items={[...customTabs.map((tab) => ({ id: tab.id, label: tab.name, closable: true })), ...(!setupHidden ? [{ id: "__setup", label: "Setup", closable: true, ...(setupProgress ? { badge: { value: setupProgress.done, total: setupProgress.total, label: `${setupProgress.done} of ${setupProgress.total} setup steps complete` } } : {}) }] : [])]} activeId={showSetup ? "__setup" : empty ? "" : active.id} onSelect={select} onNew={create} onClose={closeTab} newDisabled={workspace.dashboards.length >= 12} panelId="dash-content" />
        {setupHidden && customTabs.length > 0 && <button type="button" className="dash-text-button" onClick={() => select("__setup")}>Show setup</button>}
        {!showSetup && !empty && <details ref={tabMenu} className="viewer-more-actions dash-tab-menu" key={active.id} onKeyDown={(event) => {
          if (event.key === "Escape" && event.currentTarget.open) { event.preventDefault(); event.currentTarget.open = false; event.currentTarget.querySelector("summary")?.focus() }
        }} onBlur={(event) => { if (event.relatedTarget && !event.currentTarget.contains(event.relatedTarget as Node)) event.currentTarget.open = false }}>
          <summary aria-label={`Options for ${active.name} dashboard`} title="Dashboard options">{icons.edit}</summary>
          <div className="viewer-more-menu">
            {periodic && <div className="dash-menu-period" role="group" aria-label="Period"><span>Period</span>{PERIODS.map((period: Period) => <button key={period} type="button" aria-pressed={active.period === period} onClick={() => changeTab({ ...active, period })}>{PERIOD_LABELS[period]}{active.period === period && <span aria-hidden="true"> ✓</span>}</button>)}</div>}
            <button type="button" onClick={() => { if (tabMenu.current) tabMenu.current.open = false; reload() }} disabled={loading} aria-busy={loading}>Refresh</button>
            <button type="button" onClick={(event) => { event.currentTarget.closest("details")!.open = false; setRename(active.name); setRenaming(true); setDeleting(false) }}>Rename</button>
            <button type="button" onClick={(event) => { event.currentTarget.closest("details")!.open = false; setDeleting(true); setRenaming(false); requestAnimationFrame(() => document.querySelector<HTMLButtonElement>(".dash-delete-confirm button")?.focus()) }}>Delete dashboard</button>
          </div>
        </details>}
      </div>
      {renaming && <form className="dash-tab-form" onSubmit={(event) => { event.preventDefault(); if (rename.trim()) { changeTab({ ...active, name: rename.trim() }); setRenaming(false); setAnnouncement("Dashboard renamed"); requestAnimationFrame(() => tabMenu.current?.querySelector("summary")?.focus()) } }}>
        <label htmlFor="dash-name">Dashboard name</label><input id="dash-name" value={rename} maxLength={30} onChange={(event) => setRename(event.target.value)} autoFocus required />
        <button type="submit" disabled={!rename.trim()}>Save</button><button type="button" className="ghost" onClick={() => { setRenaming(false); requestAnimationFrame(() => tabMenu.current?.querySelector("summary")?.focus()) }}>Cancel</button>
      </form>}
      {deleting && <div className="dash-delete-confirm" role="group" aria-label="Delete dashboard">
        <p>Delete “{active.name}”? Your matters stay as they are.</p><button type="button" onClick={() => deleteTab()}>Delete dashboard</button><button type="button" className="ghost" onClick={() => { setDeleting(false); requestAnimationFrame(() => tabMenu.current?.querySelector("summary")?.focus()) }}>Cancel</button>
      </div>}
      {!showSetup && !empty && !setup?.ready && <p className="dash-ai-error" role="status">Connect your AI provider and models to build a dashboard. <button type="button" onClick={() => select("__setup")}>Connect AI in Setup</button></p>}
      {!showSetup && !empty && widgets.length > 0 && <AboutView config={{ text: active.description || `Showing ${widgets.map(titleOf).join(", ")}.` }} onRefine={canGenerate ? () => { setError(""); setPromptOpen(true) } : undefined} />}
      {!showSetup && !empty && notes.length > 0 && <div className="dash-ai-notes" role="status"><strong>About this dashboard</strong><ul>{notes.map((note) => <li key={note}>{note}</li>)}</ul></div>}
      {stored.error && <p className="dash-ai-error" role="alert">{stored.error} {stored.conflict ? <button type="button" onClick={() => { controller.current?.abort(); setWorking(false); setForcedSetup(false); setUndoState(undefined); setRenaming(false); setDeleting(false); setPromptOpen(false); void stored.loadSaved() }}>Load saved dashboards</button> : <button type="button" onClick={() => void stored.retry()}>Retry save</button>}</p>}
      {!showSetup && !empty && unread > 0 && <div className="dash-banner" role="status">{icons.warning}<p>Some matters could not be read, so their spend and drafts are missing here.</p><button type="button" className="btn-sm" onClick={reload} disabled={loading}>Retry</button></div>}
      <div id="dash-content" role={empty ? "region" : "tabpanel"} aria-label={empty ? "Dashboard" : undefined} aria-labelledby={empty ? undefined : `dash-content-tab-${showSetup ? "__setup" : active.id}`} tabIndex={0}>
        {empty ? <div className="dash-blank"><ShapeIcon icon="plus" shape="sunny" className="dash-blank-badge" /><p className="dash-blank-title">A dashboard for the way you work</p><p>Create a dashboard, or return to your setup checklist.</p><div className="dash-blank-actions"><button type="button" onClick={create}>{icons.plus}New dashboard</button><button type="button" className="ghost" onClick={() => select("__setup")}>Show setup</button></div></div> : showSetup ? context ? <>{data?.error && <p className="dash-ai-error" role="status">Your matters could not be loaded. You can still update setup. <button type="button" onClick={reload} disabled={loading}>Retry matters</button></p>}<SetupTab context={context} onSetupRefresh={onSetupRefresh} onSampleAdded={onSampleAdded} onHideSetup={() => closeTab("__setup")} /></> : <div className="dash-error" role={data?.error ? "alert" : "status"}><ShapeIcon icon="setup" shape="sunny" /><div><p className="dash-error-title">{data?.error ? "Could not load setup" : "Loading your setup…"}</p>{data?.error && <p>Your saved settings will show when the services are available.</p>}</div>{data?.error && <button type="button" onClick={reload} disabled={loading}>Retry</button>}</div> : data?.error ? <div className="dash-error" role="alert"><ShapeIcon icon="warning" shape="flower" /><div><p className="dash-error-title">Could not load your matters</p><p>Your figures will show once the matter list loads.</p></div><button type="button" onClick={reload} disabled={loading}>Retry</button></div>
          : widgets.length === 0 ? <div className="dash-blank"><ShapeIcon icon="plus" shape="sunny" className="dash-blank-badge" /><p className="dash-blank-title">A dashboard for the way you work</p><p>Tell your AI what you want to keep an eye on.</p><button type="button" disabled={!canGenerate} onClick={() => { setError(""); setPromptOpen(true) }}>{sparkle}Ask AI to build it</button></div>
            : <div aria-busy={!context}><AdaptiveGrid mode="practice" items={widgets.map((item) => ({ id: item.key, size: item.widget.type === "chart" ? item.widget.form === "stat" ? "tile" : item.widget.form === "calendar" ? "full" : "wide" : WIDGETS[item.widget.type].size }))} renderItem={(id) => { const item = widgets.find((entry) => entry.key === id)!; return <Widget item={item} context={context} period={active.period} bold={item.widget.type === bold} onPin={() => togglePin(item.key)} onRemove={() => remove(item.key)} /> }} /></div>}
      </div>
    </div>
    <AIComposer open={promptOpen && !showSetup && !empty} request={request} working={working} error={error} disabled={!canGenerate} config={{ title: "Refine dashboard", promptLabel: "Tell AI what you want from this dashboard", submitLabel: "Update dashboard", suggestions, footnote: "Pinned widgets stay in place. Figures come from your live matter data." }} onRequestChange={setRequest} onGenerate={() => void generate()} onCancel={cancelRefine} onStop={stopRefine} />
    <AIComposer open={newOpen} request={newRequest} working={newWorking} error={newError || (!setup?.ready ? "Connect your AI provider and models in Setup before generating a dashboard." : !ready ? "Your matter data could not be loaded. Retry loading it before generating a dashboard." : workspace.dashboards.length >= 12 ? "You can have up to 12 dashboard tabs. Close a tab before adding another." : undefined)} disabled={!canGenerate || workspace.dashboards.length >= 12} config={{ suggestions, footnote: "Your AI chooses widgets. Figures come from your live matter data." }} onRequestChange={setNewRequest} onGenerate={() => void generateNew()} onCancel={cancelNew} onStop={stopNew} />
    <p className="sr-only" role="status">{announcement}</p>
    {undoState && <Snackbar key={undoState.text} text={undoState.text} onUndo={undo} onDismiss={() => { if (document.activeElement?.closest(".dash-snackbar")) focusDashboardAction(); setUndoState(undefined) }} />}
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
