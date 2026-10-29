import { useEffect, useLayoutEffect, useMemo, useRef, useState, type SyntheticEvent } from "react"
import type { SetupProgress } from "../api/ingest/setup"
import PageBar from "../components/PageBar"
import ModalHeader from "../components/ModalHeader"
import backdropClose from "../components/backdropClose"
import useExit from "../components/useExit"
import { useToast } from "../components/Toast"
import { addDays, startOfDay, summarize, useDashboardData, type Period } from "../components/dashboard/data"
import { icons, ShapeIcon, SHAPES } from "../components/dashboard/icons"
import { DEFAULT_LAYOUT, loadLayout, saveLayout, WIDGETS, type Layout, type WidgetId } from "../components/dashboard/layout"
import { caption, LOOKS, RENDER, type WidgetContext } from "../components/dashboard/widgets"

const PERIODS: { value: Period; label: string }[] = [
  { value: 7, label: "7 days" },
  { value: 30, label: "30 days" },
  { value: 90, label: "90 days" },
  { value: 0, label: "All time" },
]

const BOLD = new Set(["forest", "indigo", "plum", "ink"])

function greeting(now: number) {
  const hour = new Date(now).getHours()
  return hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening"
}

// `setup` and `refresh` come from App: setup is re-read whenever Settings or
// onboarding closes, and refresh changes when a matter is added elsewhere.
export default function Dashboard({ setup, refresh = 0 }: { setup: SetupProgress | null; refresh?: number }) {
  const { data, loading, reload } = useDashboardData(setup, refresh)
  const [layout, setLayout] = useState<Layout>(loadLayout)
  const [adding, setAdding] = useState(false)
  const addButton = useRef<HTMLButtonElement>(null)
  const notify = useToast()
  // One clock per load, so every widget measures the same window.
  const now = useMemo(() => Date.now(), [data])
  // A page left open overnight re-reads at midnight, so "today" moves on.
  useEffect(() => {
    const timer = setTimeout(reload, addDays(startOfDay(now), 1) - Date.now() + 1000)
    return () => clearTimeout(timer)
  }, [now])

  function update(next: Layout) {
    setLayout(next)
    saveLayout(next)
  }
  function move(id: WidgetId, by: -1 | 1) {
    const widgets = [...layout.widgets]
    const from = widgets.indexOf(id)
    const to = from + by
    if (from < 0 || to < 0 || to >= widgets.length) return
    ;[widgets[from], widgets[to]] = [widgets[to], widgets[from]]
    update({ ...layout, widgets })
  }
  function hide(id: WidgetId) {
    const index = layout.widgets.indexOf(id)
    const neighbour = layout.widgets[index + 1] ?? layout.widgets[index - 1]
    update({ ...layout, widgets: layout.widgets.filter((widget) => widget !== id) })
    notify("info", `${WIDGETS[id].title} hidden. Add it back from Add widgets.`)
    // Focus moves to the next card's menu, or to Add widgets once none is left.
    requestAnimationFrame(() => {
      const next = neighbour && document.querySelector<HTMLElement>(`[aria-labelledby="dash-${neighbour}"] .dash-menu > summary`)
      ;(next || addButton.current)?.focus()
    })
  }

  const context: WidgetContext | undefined = data && !data.error
    ? { data, setup, period: layout.period, now, summary: summarize(data.matters, layout.period, now) }
    : undefined
  const name = data?.preferences?.attorney.trim()
  const blank = Boolean(data?.error) || layout.widgets.length === 0

  return (
    <>
      <PageBar title="Dashboard" description="Your practice at a glance: spend, drafting and what needs you." />
      <div className="dash">
        <div className="dash-toolbar">
          <div className="dash-hello">
            <h2>{greeting(now)}{name ? `, ${name}` : ""}</h2>
            <p>{new Date(now).toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" })}</p>
          </div>
          <div className="dash-controls">
            {!blank && <div className="settings-seg dash-period" role="group" aria-label="Period">
              {PERIODS.map((option) => (
                <button key={option.value} type="button" aria-pressed={layout.period === option.value} className={layout.period === option.value ? "active" : ""} onClick={() => update({ ...layout, period: option.value })}>
                  {layout.period === option.value && (
                    <svg className="settings-seg-check" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12" /></svg>
                  )}
                  {option.label}
                </button>
              ))}
            </div>}
            <button type="button" className="icon-btn dash-refresh" onClick={reload} disabled={loading} aria-busy={loading} aria-label="Refresh" title="Refresh">
              <span className={loading ? "dash-spin" : ""}>{icons.reset}</span>
            </button>
            <button ref={addButton} type="button" className="primary dash-add" onClick={() => setAdding(true)}>
              {icons.plus}
              Add widgets
            </button>
          </div>
        </div>

        {data?.error ? (
          <div className="dash-error" role="alert">
            <ShapeIcon icon="warning" shape="flower" />
            <div>
              <p className="dash-error-title">Could not load your matters</p>
              <p>{data.error}</p>
            </div>
            <button type="button" onClick={reload} disabled={loading}>Retry</button>
          </div>
        ) : layout.widgets.length === 0 ? (
          <div className="dash-blank">
            <ShapeIcon icon="plus" shape="sunny" className="dash-blank-badge" />
            <p className="dash-blank-title">Your dashboard is empty</p>
            <p>Add the widgets you want to see, or bring back the defaults.</p>
            <div className="dash-blank-actions">
              <button type="button" onClick={() => update({ ...layout, widgets: DEFAULT_LAYOUT.widgets })}>Restore defaults</button>
            </div>
          </div>
        ) : (
          <div className="dash-grid" aria-busy={!context}>
            {layout.widgets.map((id, index) => (
              <Widget
                key={id}
                id={id}
                context={context}
                period={layout.period}
                first={index === 0}
                last={index === layout.widgets.length - 1}
                onMove={(by) => move(id, by)}
                onHide={() => hide(id)}
              />
            ))}
          </div>
        )}
      </div>
      {adding && <AddWidgets layout={layout} onChange={update} onClose={() => setAdding(false)} />}
    </>
  )
}

function Widget({ id, context, period, first, last, onMove, onHide }: {
  id: WidgetId
  context?: WidgetContext
  period: Period
  first: boolean
  last: boolean
  onMove: (by: -1 | 1) => void
  onHide: () => void
}) {
  const meta = WIDGETS[id]
  const look = LOOKS[id]
  const note = caption(id, period)
  return (
    <section className={`dash-card size-${meta.size} tone-${look.tone}${BOLD.has(look.tone) ? " is-bold" : ""}`} aria-labelledby={`dash-${id}`}>
      {BOLD.has(look.tone) && <span className="dash-card-art" aria-hidden="true"><svg className="dash-card-deco" viewBox="-1 -1 2 2"><path d={SHAPES[look.shape]} /></svg></span>}
      <header className="dash-card-head">
        <ShapeIcon icon={look.icon} shape={look.shape} className="dash-card-badge" />
        <div className="dash-card-heading">
          <h3 id={`dash-${id}`}>{meta.title}</h3>
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

// Every widget not on the dashboard, each one click from being added (to the
// end). Reset brings back the default set and order.
function AddWidgets({ layout, onChange, onClose }: { layout: Layout; onChange: (layout: Layout) => void; onClose: () => void }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const [open, setOpen] = useState(true)
  const [shown, closing] = useExit(open)
  const hidden = (Object.keys(WIDGETS) as WidgetId[]).filter((id) => !layout.widgets.includes(id))
  const isDefault = layout.widgets.join() === DEFAULT_LAYOUT.widgets.join()
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
      <ModalHeader title={<span id="dash-add-heading">Add widgets</span>} onClose={() => setOpen(false)} />
      <div className="dash-add-body">
        {hidden.length ? (
          <ul className="dash-add-list">
            {hidden.map((id) => (
              <li key={id} className={`tone-${LOOKS[id].tone}`}>
                <ShapeIcon icon={LOOKS[id].icon} shape={LOOKS[id].shape} className="dash-card-badge" />
                <span className="dash-add-text">
                  <strong>{WIDGETS[id].title}</strong>
                  <span>{WIDGETS[id].description}</span>
                </span>
                <button type="button" className="btn-sm" onClick={() => onChange({ ...layout, widgets: [...layout.widgets, id] })} aria-label={`Add ${WIDGETS[id].title}`}>
                  {icons.plus}Add
                </button>
              </li>
            ))}
          </ul>
        ) : (
          <div className="dash-add-none">
            <ShapeIcon icon="check" shape="sunny" />
            <p>Every widget is on your dashboard. Hide one from its menu to make room.</p>
          </div>
        )}
      </div>
      <div className="dash-add-foot">
        <button type="button" className="ghost" disabled={isDefault} onClick={() => onChange({ ...layout, widgets: DEFAULT_LAYOUT.widgets })}>
          {icons.reset}Reset to default
        </button>
        <button type="button" className="primary" onClick={() => setOpen(false)}>Done</button>
      </div>
    </dialog>
  )
}
