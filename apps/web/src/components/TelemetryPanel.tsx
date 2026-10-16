import { useEffect, useId, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"
import { getTelemetry, getTelemetryDetail, type TelemetryDetail, type TelemetryLean } from "../api/ingest/telemetry"
import { formatCost } from "../lib/format"
import ModalHeader from "./ModalHeader"

type Props = { matterId: string; sessionID?: string }

// The status-bar button opens one right-hand drawer for the complete usage ledger.
export default function TelemetryPanel({ matterId, sessionID }: Props) {
  const [lean, setLean] = useState<TelemetryLean>()
  const [detail, setDetail] = useState<TelemetryDetail>()
  const [open, setOpen] = useState(false)
  const [closing, setClosing] = useState(false)
  const dialog = useRef<HTMLDialogElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const panelId = useId()
  const [detailError, setDetailError] = useState(false)
  const [visible, setVisible] = useState(() => !document.hidden)

  useEffect(() => {
    const onVisibility = () => setVisible(!document.hidden)
    document.addEventListener("visibilitychange", onVisibility)
    return () => document.removeEventListener("visibilitychange", onVisibility)
  }, [])

  useEffect(() => {
    setLean(undefined)
    setDetail(undefined)
    setDetailError(false)
    setOpen(false)
    setClosing(false)
  }, [matterId])

  useEffect(() => {
    if (!open) return
    const element = dialog.current
    element?.showModal()
    element?.querySelector<HTMLButtonElement>(".modal-close")?.focus({ preventScroll: true })
    return () => {
      element?.close()
      trigger.current?.focus()
    }
  }, [open])

  useEffect(() => {
    if (!visible) return
    let cancelled = false

    async function load() {
      const next = await getTelemetry(matterId, sessionID).catch(() => undefined)
      if (cancelled) return
      if (next) setLean(next)
      if (!open) return
      const nextDetail = await getTelemetryDetail(matterId).catch(() => undefined)
      if (cancelled) return
      setDetailError(!nextDetail)
      if (!nextDetail) return
      setDetail(nextDetail)
    }

    load()
    const timer = window.setInterval(load, 5000)
    return () => {
      cancelled = true
      window.clearInterval(timer)
    }
  }, [open, matterId, sessionID, visible])

  function close() {
    if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
      setOpen(false)
      return
    }
    setClosing(true)
  }

  const live = Boolean(lean?.current_activity)
  // A model missing from the rate card contributes nothing to the sums, so any
  // total that may include it is a lower bound, never a confident $0.00. When
  // nothing is priced the ledger counts turns only and says so once.
  const unknown = lean?.unknown_models ?? []
  const unpriced = unknown.length > 0 && (lean?.matter_total.cost_usd ?? 0) <= 0
  const money = (value: number) => (unpriced ? "" : formatCost(value, unknown.length > 0))
  const total = !lean ? "—" : unpriced ? `${lean.matter_total.turns} turns` : money(lean.matter_total.cost_usd)

  return (
    <div className="rail-strip-item">
      <button
        ref={trigger}
        className={`rail-strip-btn${open ? " is-active" : ""}`}
        type="button"
        onClick={() => {
          setClosing(false)
          setOpen(true)
        }}
        aria-label={`Usage: ${total}`}
        aria-haspopup="dialog"
        aria-expanded={open}
        aria-controls={open ? panelId : undefined}
        title={live ? "Usage — assistant working now" : "Usage"}
      >
        <span className="rail-strip-cost">{total}</span>
        {live && <span className="rail-dot is-live" aria-hidden />}
      </button>
      {open && createPortal(
        <dialog
          ref={dialog}
          id={panelId}
          className={`telemetry-overlay${closing ? " is-closing" : ""}`}
          autoFocus
          tabIndex={-1}
          aria-labelledby={`${panelId}-title`}
          onCancel={(event) => {
            event.preventDefault()
            close()
          }}
          onClick={(event) => {
            if (event.target === event.currentTarget) close()
          }}
        >
          <div className="telemetry-drawer" onTransitionEnd={(event) => {
            if (closing && event.target === event.currentTarget && event.propertyName === "transform") setOpen(false)
          }}>
            <ModalHeader title={<span id={`${panelId}-title`}>Usage</span>} onClose={close} />
            <div className="telemetry-drawer-body">
              <TelemetrySection title="Overview">
                <dl className="telemetry-ledger">
                  <Metric label="Current activity" cost={lean?.current_activity ? money(lean.current_activity.cost_usd) : "—"} note={lean?.current_activity ? `${lean.current_activity.turns} turns` : "idle"} live={live} />
                  <Metric label="Today" cost={lean ? money(lean.today.cost_usd) : "—"} note={lean ? `${lean.today.turns} turns` : ""} />
                  <Metric label="Matter total" cost={lean ? money(lean.matter_total.cost_usd) : "—"} note={lean ? `${lean.matter_total.turns} turns · ${lean.matter_total.sessions} sessions` : ""} />
                </dl>
                {(lean?.rate_card_as_of || unknown.length > 0) && (
                  <div className="telemetry-foot">
                    {unknown.length > 0 && <p className="telemetry-unpriced">
                      {unpriced ? "Cost not available for the models used." : "Totals leave out usage from models without a published rate."}
                    </p>}
                    {unknown.length > 0 && <p>No published rate: {unknown.join(", ")}</p>}
                    {lean?.rate_card_as_of && <p>Rates as of {lean.rate_card_as_of}</p>}
                  </div>
                )}
              </TelemetrySection>

              {lean?.current_session && <TelemetrySection title={sessionID ? "Current conversation" : "Latest conversation"}>
                <dl className="telemetry-ledger">
                  <Metric label="Estimated input tokens" cost={number(lean.current_session.tokens)} note="" />
                  <Metric label="Agent" cost={lean.current_session.agent} note="" />
                  <Metric label="Model" cost={lean.current_session.model} note="" />
                  <Metric label="Context" cost={{ ok: "Normal usage", warn: "High usage", fail: "Very high usage" }[lean.current_session.status]} note="" />
                </dl>
              </TelemetrySection>}

              {detailError && <p className="telemetry-load-status" role="status">Could not refresh usage details. Retrying automatically.</p>}
              {!detail && !detailError && <p className="telemetry-load-status" role="status">Loading usage details…</p>}
              {detail && <>
                <TelemetrySection title="Spend over time">
                  <Sparkline points={detail.sparkline} />
                </TelemetrySection>

                <TelemetrySection title="Cost split">
                  {detail.by_agent.some((row) => row.cost_usd > 0) ? <CostDonut rows={detail.by_agent} money={money} /> : undefined}
                </TelemetrySection>

                <TelemetrySection title="By model">
                  {detail.by_model.length ? detail.by_model.map((row) => (
                    <BarRow key={row.model} label={row.model} value={unknown.includes(row.model) && row.cost_usd <= 0 ? "Cost unavailable" : formatCost(row.cost_usd, unknown.includes(row.model))} meta={`${row.turns} turns`} fraction={row.cost_usd / maxBy(detail.by_model, (r) => r.cost_usd)} />
                  )) : undefined}
                </TelemetrySection>

                <TelemetrySection title="By tool">
                  {detail.by_tool.length ? detail.by_tool.map((row) => (
                    <BarRow key={row.tool} label={row.tool} value={`${row.calls} calls`} meta={`${number(row.avg_duration_ms)} ms avg · ${number(row.total_duration_ms)} ms total`} fraction={row.calls / maxBy(detail.by_tool, (r) => r.calls)} />
                  )) : undefined}
                </TelemetrySection>

                <TelemetrySection title="By agent">
                  {detail.by_agent.length ? detail.by_agent.map((row) => (
                    <BarRow key={row.agent} label={row.agent} value={money(row.cost_usd)} meta={`${row.turns} turns · ${row.tool_calls} tools`} fraction={row.cost_usd / maxBy(detail.by_agent, (r) => r.cost_usd)} />
                  )) : undefined}
                </TelemetrySection>

                <TelemetrySection title="Recent turns">
                  {detail.recent_turns.length ? <ul className="telemetry-recent">
                    {detail.recent_turns.map((row, index) => <li key={`${row.ts}-${index}`}>
                      <div className="telemetry-bar-head">
                        <span className="telemetry-bar-label">{row.agent}</span>
                        <span className="telemetry-bar-value">{row.cost_usd === null ? "Cost unavailable" : money(row.cost_usd)}</span>
                      </div>
                      <span className="telemetry-bar-meta">{row.model} · {number(row.tokens)} input tokens · {row.status}</span>
                      <time className="telemetry-bar-meta" dateTime={new Date(row.ts).toISOString()}>{new Date(row.ts).toLocaleString()}</time>
                    </li>)}
                  </ul> : undefined}
                </TelemetrySection>

                <TelemetrySection title="Recent conversations">
                  {detail.recent_sessions.length ? <ul className="telemetry-recent">
                    {detail.recent_sessions.map((row) => <li key={row.id}>
                      <div className="telemetry-bar-head">
                        <span className="telemetry-bar-label">{row.agent}</span>
                        <span className="telemetry-bar-value">{money(row.cost_usd)}</span>
                      </div>
                      <span className="telemetry-bar-meta">{row.status}</span>
                      <time className="telemetry-bar-meta" dateTime={new Date(row.last_activity_at).toISOString()}>{new Date(row.last_activity_at).toLocaleString()}</time>
                    </li>)}
                  </ul> : undefined}
                </TelemetrySection>
              </>}
            </div>
          </div>
        </dialog>,
        document.body,
      )}
    </div>
  )
}

function Metric({ label, cost, note, live }: { label: string; cost: string; note: string; live?: boolean }) {
  return (
    <div className={`telemetry-metric${live ? " is-live" : ""}`}>
      <dt className="telemetry-metric-label">{label}</dt>
      <dd className="telemetry-metric-value">
        {cost && <span className="telemetry-metric-cost">{cost}</span>}
        {note && <span className="telemetry-metric-note">{note}</span>}
      </dd>
    </div>
  )
}

function BarRow({ label, value, meta, fraction }: { label: string; value: string; meta: string; fraction: number }) {
  return (
    <div className="telemetry-bar-row">
      <div className="telemetry-bar-head">
        <span className="telemetry-bar-label">{label}</span>
        <span className="telemetry-bar-value">{value}</span>
      </div>
      <div className="telemetry-bar-track">
        <div className="telemetry-bar-fill" style={{ width: `${Math.max(2, Math.min(100, fraction * 100))}%` }} />
      </div>
      <span className="telemetry-bar-meta">{meta}</span>
    </div>
  )
}

const DONUT_COLORS = ["#2d4a3e", "#8a5a1a", "#5d7a6d", "#b98a4a", "#a3b3ab", "#44443f", "#c4d2cb", "#6b6b66"]

function CostDonut({ rows, money }: { rows: Array<{ agent: string; cost_usd: number }>; money: (value: number) => string }) {
  const slices = rows.filter((row) => row.cost_usd > 0).sort((a, b) => b.cost_usd - a.cost_usd)
  const total = slices.reduce((sum, row) => sum + row.cost_usd, 0)
  const radius = 60
  const circumference = 2 * Math.PI * radius
  let offset = 0

  return (
    <div className="telemetry-donut">
      <svg className="telemetry-donut-svg" viewBox="0 0 160 160" role="img" aria-label="Cost split by agent">
        <g transform="rotate(-90 80 80)">
          {slices.map((row, index) => {
            const length = (row.cost_usd / total) * circumference
            const dash = (
              <circle
                key={row.agent}
                cx="80"
                cy="80"
                r={radius}
                fill="none"
                stroke={DONUT_COLORS[index % DONUT_COLORS.length]}
                strokeWidth="24"
                strokeDasharray={`${length} ${circumference - length}`}
                strokeDashoffset={-offset}
              />
            )
            offset += length
            return dash
          })}
        </g>
        <text x="80" y="76" className="telemetry-donut-total">{money(total)}</text>
        <text x="80" y="92" className="telemetry-donut-caption">total</text>
      </svg>
      <ul className="telemetry-donut-legend">
        {slices.map((row, index) => (
          <li key={row.agent}>
            <span className="telemetry-donut-swatch" style={{ background: DONUT_COLORS[index % DONUT_COLORS.length] }} aria-hidden />
            <span className="telemetry-donut-name">{row.agent}</span>
            <span className="telemetry-donut-pct">{Math.round((row.cost_usd / total) * 100)}%</span>
            <span className="telemetry-donut-cost">{money(row.cost_usd)}</span>
          </li>
        ))}
      </ul>
    </div>
  )
}

function Sparkline({ points }: { points: Array<{ ts: number; cost_usd: number }> }) {
  if (points.length === 0) return <div className="telemetry-spark telemetry-spark-empty">No activity yet</div>
  const max = Math.max(...points.map((point) => point.cost_usd), 0)
  return (
    <div className="telemetry-spark" role="img" aria-label="Recent spend">
      {points.map((point, index) => (
        <span key={point.ts ?? index} className="telemetry-spark-bar" style={{ height: `${max === 0 ? 4 : Math.max(8, (point.cost_usd / max) * 100)}%` }} />
      ))}
    </div>
  )
}

function TelemetrySection({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="telemetry-section">
      <h3>{title}</h3>
      {children || <div className="empty-inline">No data yet.</div>}
    </section>
  )
}

function maxBy<T>(rows: T[], pick: (row: T) => number) {
  return Math.max(...rows.map(pick), 0) || 1
}

function number(value: number) {
  return new Intl.NumberFormat().format(Math.round(value))
}
