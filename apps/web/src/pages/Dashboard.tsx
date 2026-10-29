import { useEffect, useLayoutEffect, useMemo, useRef, useState, type SyntheticEvent } from "react"
import type { DraftingPreferences } from "../api/ingest/preferences"
import type { SetupProgress } from "../api/ingest/setup"
import PageBar from "../components/PageBar"
import ModalHeader from "../components/ModalHeader"
import backdropClose from "../components/backdropClose"
import useExit from "../components/useExit"
import { addDays, attention, PERIOD_LABELS, PERIODS, startOfDay, summarize, useDashboardData, type DashboardData, type Period } from "../components/dashboard/data"
import { icons, ShapeIcon } from "../components/dashboard/icons"
import { available, DEFAULT_WIDGETS, loadLayout, PERIODIC, saveLayout, visibleWidgets, WIDGETS, type Arrangement, type Layout, type Stage, type WidgetId } from "../components/dashboard/layout"
import { caption, LOOKS, RENDER, setupComplete, stageOf, type WidgetContext } from "../components/dashboard/widgets"

function greeting(now: number) {
  const hour = new Date(now).getHours()
  return hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening"
}

const EMPTY: Arrangement = { hidden: [], added: [] }

// `setup` and `refresh` come from App: setup is re-read whenever Settings or
// onboarding closes, and refresh changes when a matter is added elsewhere.
export default function Dashboard({ setup, refresh = 0 }: { setup: SetupProgress | null; refresh?: number }) {
  const { data, loading, reload } = useDashboardData(setup, refresh)
  const [layout, setLayout] = useState<Layout>(loadLayout)
  const [editing, setEditing] = useState(false)
  const [hidden, setHidden] = useState<{ id: WidgetId; stage: Stage }>()
  const [announcement, setAnnouncement] = useState("")
  const editButton = useRef<HTMLButtonElement>(null)
  const added = useRef<WidgetId>(undefined)
  // One clock per load, so every widget measures the same window.
  const now = useMemo(() => Date.now(), [data])
  // A page left open overnight re-reads at midnight, so "today" moves on.
  useEffect(() => {
    const timer = setTimeout(reload, addDays(startOfDay(now), 1) - Date.now() + 1000)
    return () => clearTimeout(timer)
  }, [now])

  // Settings saved from a widget show at once, until the next load reads them.
  // A save that lands after a reload applies to the data now on screen.
  const [saved, setSaved] = useState<{ data?: DashboardData; preferences?: DraftingPreferences }>({})
  const latest = useRef(data)
  latest.current = data
  const shownData = data && saved.data === data ? { ...data, preferences: saved.preferences } : data
  const ready = data && !data.error ? data : undefined
  const context: WidgetContext | undefined = shownData && ready
    ? { data: shownData, setup, period: layout.period, now, summary: summarize(shownData.matters, layout.period, now), onPreferences: (preferences) => setSaved({ data: latest.current, preferences }), onRetry: reload }
    : undefined
  // Setup is judged once per load, so its card stays put while being filled in.
  const setupDone = useMemo(() => !!ready && setupComplete(ready), [ready])
  const stage = ready ? stageOf(ready) : layout.stage ?? "starting"
  const arrangement = layout.stages[stage]
  // While it loads, the widgets last shown hold the places.
  const widgets = ready ? visibleWidgets(stage, setupDone, arrangement) : layout.shown ?? DEFAULT_WIDGETS[stage]
  const shown = context && `${stage}:${widgets.join()}`
  useEffect(() => {
    if (shown && shown !== `${layout.stage}:${layout.shown?.join()}`) update({ ...layout, stage, shown: widgets })
  }, [shown])

  function update(next: Layout) {
    setLayout(next)
    saveLayout(next)
  }
  // Arranging changes only one stage's dashboard; none resets the default.
  function arrange(next: Arrangement | undefined, of: Stage = stage) {
    const { [of]: _, ...others } = layout.stages
    update({ ...layout, stages: next ? { ...others, [of]: next } : others })
  }
  const focusMenu = (id: WidgetId) => document.querySelector<HTMLElement>(`[aria-labelledby="dash-${id}"] .dash-menu > summary`)
  function move(id: WidgetId, by: -1 | 1) {
    const next = [...widgets]
    const from = next.indexOf(id)
    const to = from + by
    if (from < 0 || to < 0 || to >= next.length) return
    ;[next[from], next[to]] = [next[to], next[from]]
    arrange({ ...(arrangement ?? EMPTY), order: next })
    setAnnouncement(`${WIDGETS[id].title} moved to position ${to + 1} of ${next.length}`)
    requestAnimationFrame(() => focusMenu(id)?.focus())
  }
  function hide(id: WidgetId) {
    const index = widgets.indexOf(id)
    const neighbour = widgets[index + 1] ?? widgets[index - 1]
    const current = arrangement ?? EMPTY
    arrange({
      ...current,
      hidden: DEFAULT_WIDGETS[stage].includes(id) ? [...current.hidden.filter((other) => other !== id), id] : current.hidden,
      added: current.added.filter((other) => other !== id),
    })
    setHidden({ id, stage })
    setAnnouncement(`${WIDGETS[id].title} hidden`)
    // Focus moves to the next card's menu, or to Edit dashboard once none is left.
    requestAnimationFrame(() => ((neighbour && focusMenu(neighbour)) || editButton.current)?.focus())
  }
  // Undo reverses only the hide, keeping any move or add made since. The
  // widget's place in a saved order was kept, so it returns where it was.
  function undo() {
    if (!hidden) return
    const { id, stage: of } = hidden
    const current = layout.stages[of] ?? EMPTY
    arrange(DEFAULT_WIDGETS[of].includes(id)
      ? { ...current, hidden: current.hidden.filter((other) => other !== id) }
      : { ...current, added: [...current.added.filter((other) => other !== id), id] }, of)
    setHidden(undefined)
    setAnnouncement(`${WIDGETS[id].title} restored`)
    requestAnimationFrame(() => focusMenu(id)?.focus())
  }
  function dismiss() {
    // The snackbar is about to go; focus it held moves to Edit dashboard.
    if (document.activeElement?.closest(".dash-snackbar")) editButton.current?.focus()
    setHidden(undefined)
  }
  function add(id: WidgetId) {
    const current = arrangement ?? EMPTY
    arrange({
      ...current,
      hidden: current.hidden.filter((other) => other !== id),
      added: DEFAULT_WIDGETS[stage].includes(id) ? current.added : [...current.added, id],
    })
    added.current = id
    setAnnouncement(`${WIDGETS[id].title} added`)
  }
  function reset() {
    arrange(undefined)
    setAnnouncement("Dashboard reset to the default")
    // From the empty state, whose button goes with it.
    if (!editing) requestAnimationFrame(() => editButton.current?.focus())
  }
  // Once the dialog has closed and handed focus back, an added widget takes it.
  useEffect(() => {
    if (editing || !added.current) return
    const menu = focusMenu(added.current)
    added.current = undefined
    const still = window.matchMedia("(prefers-reduced-motion: reduce)").matches
    menu?.closest("section")?.scrollIntoView({ block: "nearest", behavior: still ? "auto" : "smooth" })
    menu?.focus({ preventScroll: true })
  }, [editing])

  const name = shownData?.preferences?.attorney.trim()
  // The period is offered only while a widget on the dashboard follows it.
  const periodic = !data?.error && widgets.some((id) => PERIODIC.has(id))
  const unread = ready?.matters.filter((row) => row.unread).length ?? 0
  // One bold block, on what to do next: what is waiting, else the first steps.
  const bold: WidgetId | undefined = ready && widgets.includes("attention") && attention(ready.matters).length ? "attention" : widgets.includes("journey") ? "journey" : undefined
  const offered = ready ? available(stage, setupDone).filter((id) => !widgets.includes(id)) : []

  return (
    <>
      <PageBar title="Dashboard" description={`${greeting(now)}${name ? `, ${name}` : ""}. What needs you, your drafts and what your matters cost.`} />
      <div className="dash">
        <div className="dash-toolbar">
          {periodic && <div className="settings-seg dash-period" role="group" aria-label="Period">
            {PERIODS.map((period: Period) => (
              <button key={period} type="button" aria-pressed={layout.period === period} className={layout.period === period ? "active" : ""} onClick={() => update({ ...layout, period })}>
                {layout.period === period && (
                  <svg className="settings-seg-check" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12" /></svg>
                )}
                {PERIOD_LABELS[period]}
              </button>
            ))}
          </div>}
          <div className="dash-controls">
            <button type="button" className="icon-btn dash-refresh" onClick={reload} disabled={loading} aria-busy={loading} aria-label="Refresh" title="Refresh">
              {icons.refresh}
            </button>
            <button ref={editButton} type="button" className="dash-edit" onClick={() => setEditing(true)} disabled={!ready}>
              {icons.edit}
              Edit dashboard
            </button>
          </div>
        </div>

        {unread > 0 && (
          <div className="dash-banner" role="status">
            {icons.warning}
            <p>Some matters could not be read, so their spend and drafts are missing here.</p>
            <button type="button" className="btn-sm" onClick={reload} disabled={loading}>Retry</button>
          </div>
        )}

        {data?.error ? (
          <div className="dash-error" role="alert">
            <ShapeIcon icon="warning" shape="flower" />
            <div>
              <p className="dash-error-title">Could not load your matters</p>
              <p>Your figures will show once the matter list loads.</p>
            </div>
            <button type="button" onClick={reload} disabled={loading}>Retry</button>
          </div>
        ) : widgets.length === 0 ? (
          <div className="dash-blank">
            <ShapeIcon icon="plus" shape="sunny" className="dash-blank-badge" />
            <p className="dash-blank-title">Your dashboard is empty</p>
            <p>Add the widgets you want to see, or bring back the defaults.</p>
            <div className="dash-blank-actions">
              <button type="button" onClick={() => setEditing(true)}>{icons.edit}Edit dashboard</button>
              <button type="button" onClick={reset}>Reset to default</button>
            </div>
          </div>
        ) : (
          <div className="dash-grid" aria-busy={!context}>
            {widgets.map((id, index) => (
              <Widget
                key={id}
                id={id}
                context={context}
                period={layout.period}
                bold={id === bold}
                first={index === 0}
                last={index === widgets.length - 1}
                onMove={(by) => move(id, by)}
                onHide={() => hide(id)}
              />
            ))}
          </div>
        )}
      </div>
      <p className="sr-only" role="status">{announcement}</p>
      {hidden && <Snackbar key={hidden.id} text={`${WIDGETS[hidden.id].title} hidden`} onUndo={undo} onDismiss={dismiss} />}
      {editing && <EditDashboard offered={offered} arranged={!!arrangement} onAdd={add} onReset={reset} onClose={() => setEditing(false)} />}
    </>
  )
}

function Widget({ id, context, period, bold, first, last, onMove, onHide }: {
  id: WidgetId
  context?: WidgetContext
  period: Period
  bold: boolean
  first: boolean
  last: boolean
  onMove: (by: -1 | 1) => void
  onHide: () => void
}) {
  const meta = WIDGETS[id]
  const look = LOOKS[id]
  const note = caption(id, period)
  return (
    <section className={`dash-card size-${meta.size}${look.tone ? ` tone-${look.tone}` : ""}${bold ? " is-bold" : ""}`} aria-labelledby={`dash-${id}`}>
      <header className="dash-card-head">
        <ShapeIcon icon={look.icon} shape={look.shape} className="dash-card-badge" />
        <div className="dash-card-heading">
          <h2 id={`dash-${id}`}>{meta.title}</h2>
          {note && <p>{note}</p>}
        </div>
        <WidgetMenu title={meta.title} first={first} last={last} onMove={onMove} onHide={onHide} />
      </header>
      <div className="dash-card-body">
        {context ? RENDER[id](context) : <div className="dash-skeleton" aria-hidden="true"><span /><span /><span /></div>}
      </div>
    </section>
  )
}

function closeMenu(element: Element) {
  const menu = element.closest("details")
  if (menu) menu.open = false
}

function WidgetMenu({ title, first, last, onMove, onHide }: { title: string; first: boolean; last: boolean; onMove: (by: -1 | 1) => void; onHide: () => void }) {
  // While the menu is open, a press anywhere outside closes it.
  const outside = useRef<(press: PointerEvent) => void>(undefined)
  useEffect(() => () => { if (outside.current) document.removeEventListener("pointerdown", outside.current) }, [])
  function onToggle(event: SyntheticEvent<HTMLDetailsElement>) {
    const menu = event.currentTarget
    if (outside.current) document.removeEventListener("pointerdown", outside.current)
    outside.current = undefined
    if (!menu.open) return
    outside.current = (press: PointerEvent) => { if (!menu.contains(press.target as Node)) menu.open = false }
    document.addEventListener("pointerdown", outside.current)
  }
  return (
    <details
      className="viewer-more-actions dash-menu"
      onToggle={onToggle}
      onKeyDown={(event) => {
        if (event.key !== "Escape" || !event.currentTarget.open) return
        event.preventDefault()
        event.stopPropagation()
        event.currentTarget.open = false
        event.currentTarget.querySelector("summary")?.focus()
      }}
      onBlur={(event) => {
        const next = event.relatedTarget as Node | null
        if (next && !event.currentTarget.contains(next)) event.currentTarget.open = false
      }}
    >
      <summary aria-label={`Options for ${title}`} title="Options">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
          <circle cx="5" cy="12" r="1.8" />
          <circle cx="12" cy="12" r="1.8" />
          <circle cx="19" cy="12" r="1.8" />
        </svg>
      </summary>
      <div className="viewer-more-menu">
        <button type="button" disabled={first} onClick={(event) => { closeMenu(event.currentTarget); onMove(-1) }}>
          {icons.chevronUp}<span>Move earlier</span>
        </button>
        <button type="button" disabled={last} onClick={(event) => { closeMenu(event.currentTarget); onMove(1) }}>
          {icons.chevronDown}<span>Move later</span>
        </button>
        <button type="button" onClick={(event) => { closeMenu(event.currentTarget); onHide() }}>
          {icons.hide}<span>Hide widget</span>
        </button>
      </div>
    </details>
  )
}

// "X hidden" with Undo, at the foot of the page. It waits while the pointer or
// focus is on it, so there is time to reach Undo.
function Snackbar({ text, onUndo, onDismiss }: { text: string; onUndo: () => void; onDismiss: () => void }) {
  const [held, setHeld] = useState(false)
  useEffect(() => {
    if (held) return
    const timer = setTimeout(onDismiss, 8000)
    return () => clearTimeout(timer)
  }, [held])
  return (
    <div
      className="dash-snackbar"
      onPointerEnter={() => setHeld(true)}
      onPointerLeave={() => setHeld(false)}
      onFocus={() => setHeld(true)}
      onBlur={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setHeld(false) }}
    >
      <span>{text}</span>
      <button type="button" className="dash-snackbar-action" onClick={onUndo}>Undo</button>
      <button type="button" className="dash-snackbar-close" onClick={onDismiss} aria-label="Dismiss" title="Dismiss">{icons.close}</button>
    </div>
  )
}

// Every widget this dashboard can show but does not, each one click from being
// added. Reset goes back to the default.
function EditDashboard({ offered, arranged, onAdd, onReset, onClose }: { offered: WidgetId[]; arranged: boolean; onAdd: (id: WidgetId) => void; onReset: () => void; onClose: () => void }) {
  const dialog = useRef<HTMLDialogElement>(null)
  // Reset disables itself, so focus moves on to Done.
  const done = useRef<HTMLButtonElement>(null)
  const [open, setOpen] = useState(true)
  const [shown, closing] = useExit(open)
  useLayoutEffect(() => {
    const element = dialog.current
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    element?.showModal()
    return () => {
      element?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [])
  useLayoutEffect(() => {
    if (!shown) onClose()
  }, [shown, onClose])
  if (!shown) return null
  return (
    <dialog
      ref={dialog}
      className={`picker-panel dash-add-panel${closing ? " is-closing" : ""}`}
      inert={closing}
      aria-labelledby="dash-add-heading"
      onCancel={(event) => { event.preventDefault(); setOpen(false) }}
      {...backdropClose(() => setOpen(false))}
    >
      <ModalHeader title={<span id="dash-add-heading">Edit dashboard</span>} onClose={() => setOpen(false)} />
      <div className="dash-add-body">
        <p className="dash-add-hint">Hide or move a widget from its <span className="sr-only">Options</span><span aria-hidden="true">•••</span> menu. Add one back here.</p>
        {offered.length ? (
          <ul className="dash-add-list">
            {offered.map((id) => (
              <li key={id}>
                <ShapeIcon icon={LOOKS[id].icon} shape={LOOKS[id].shape} className="dash-card-badge" />
                <span className="dash-add-text">
                  <strong>{WIDGETS[id].title}</strong>
                  <span>{WIDGETS[id].description}</span>
                </span>
                <button type="button" className="btn-sm" onClick={() => { onAdd(id); setOpen(false) }} aria-label={`Add ${WIDGETS[id].title}`}>
                  {icons.plus}Add
                </button>
              </li>
            ))}
          </ul>
        ) : (
          <div className="dash-add-none">
            <ShapeIcon icon="check" shape="sunny" />
            <p>Every widget is on your dashboard.</p>
          </div>
        )}
      </div>
      <div className="dash-add-foot">
        <button type="button" className="ghost" disabled={!arranged} onClick={() => { onReset(); done.current?.focus() }}>
          {icons.reset}Reset to default
        </button>
        <button ref={done} type="button" className="primary" onClick={() => setOpen(false)}>Done</button>
      </div>
    </dialog>
  )
}
