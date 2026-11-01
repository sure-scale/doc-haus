import { useEffect, useId, useLayoutEffect, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"
import { getTelemetry, getTelemetryDetail, type TelemetryDetail, type TelemetryLean } from "../api/ingest/telemetry"
import { formatCost } from "../lib/format"
import ModalHeader from "./ModalHeader"
import backdropClose from "./backdropClose"
import useExit from "./useExit"

type Props = { matterId: string; sessionID?: string }

// The status-bar button opens one right-hand drawer for the complete usage ledger.
export default function TelemetryPanel({ matterId, sessionID }: Props) {
  const [lean, setLean] = useState<TelemetryLean>()
  const [detail, setDetail] = useState<TelemetryDetail>()
  const [open, setOpen] = useState(false)
  const [shown, closing] = useExit(open)
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
  }, [matterId])

  // Closes, and returns focus, as the exit starts; only the picture lingers.
  useLayoutEffect(() => {
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
    setOpen(false)
  }

  const live = Boolean(lean?.current_activity)
  // A model call whose provider reported no cost adds nothing to the sums, so
  // any total that may include one is a lower bound, never a confident $0.00.
  // When nothing is priced the ledger counts calls only and says so once.
  const unpricedCalls = lean?.unpriced_turns ?? 0
  const unpriced = unpricedCalls > 0 && (lean?.matter_total.cost_usd ?? 0) <= 0
  const money = (value: number) => (unpriced ? "" : formatCost(value, unpricedCalls > 0))
  const total = !lean ? "—" : unpriced ? calls(lean.matter_total.turns) : money(lean.matter_total.cost_usd)

  return (
    <div className="rail-strip-item">
      <button
        ref={trigger}
        className={`rail-strip-btn${open ? " is-active" : ""}`}
        type="button"
        onClick={() => setOpen(true)}
        aria-label={`Usage: ${total}`}
        aria-haspopup="dialog"
        aria-expanded={open}
        aria-controls={open ? panelId : undefined}
        title={live ? "Usage — assistant working now" : "Usage"}
      >
        <span className="rail-strip-cost">{total}</span>
        {live && <span className="rail-dot is-live" aria-hidden />}
      </button>
      {shown && createPortal(
        <dialog
          ref={dialog}
          id={panelId}
          className={`telemetry-overlay${closing ? " is-closing" : ""}`}
          inert={closing}
          autoFocus
          tabIndex={-1}
          aria-labelledby={`${panelId}-title`}
          onCancel={(event) => {
            event.preventDefault()
            close()
          }}
          {...backdropClose(() => close())}
        >
          <div className="telemetry-drawer">
            <ModalHeader title={<span id={`${panelId}-title`}>Usage</span>} onClose={close} />
            <div className="telemetry-drawer-body">
              <TelemetrySection title="Overview">
                <dl className="telemetry-ledger">
                  <Metric label="Current activity" cost={lean?.current_activity ? money(lean.current_activity.cost_usd) : "—"} note={lean?.current_activity ? calls(lean.current_activity.turns) : "idle"} live={live} />
                  <Metric label="Today" cost={lean ? money(lean.today.cost_usd) : "—"} note={lean ? calls(lean.today.turns) : ""} />
                  <Metric label="Matter total" cost={lean ? money(lean.matter_total.cost_usd) : "—"} note={lean ? calls(lean.matter_total.turns) : ""} />
                </dl>
                <div className="telemetry-foot">
                  <p>Cost as reported by your model provider.</p>
                  {unpricedCalls > 0 && <p className="telemetry-unpriced">
                    {unpriced ? "Your provider reported no cost for the models used." : `Totals leave out ${calls(unpricedCalls)} with no reported cost.`}
                  </p>}
                  {lean?.partial && <p className="telemetry-unpriced">Some conversations could not be read, so totals may be low.</p>}
                </div>
              </TelemetrySection>

              {lean?.current_session && <TelemetrySection title={sessionID ? "Current conversation" : "Latest conversation"}>
                <dl className="telemetry-ledger">
                  <Metric label="Estimated input tokens" cost={number(lean.current_session.tokens)} note="" />
                  <Metric label="Assistant" cost={lean.current_session.label} note="" />
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

                <TelemetrySection title="Background agents and your work">
                  {detail.by_group.some((row) => row.cost_usd > 0) ? <CostDonut rows={detail.by_group.map((row) => ({ key: row.group, label: row.label, cost_usd: row.cost_usd }))} money={money} /> : undefined}
                </TelemetrySection>

                <TelemetrySection title="By task">
                  {detail.by_agent.length ? detail.by_group.map((group) => {
                    const rows = detail.by_agent.filter((row) => row.group === group.group)
                    return rows.length ? <div key={group.group} className="telemetry-group">
                      <h4 className="telemetry-group-title">{group.label}</h4>
                      {rows.map((row) => (
                        <BarRow key={row.agent} label={row.label} value={money(row.cost_usd)} meta={[row.turns ? calls(row.turns) : "", row.tool_calls ? `${number(row.tool_calls)} tool ${row.tool_calls === 1 ? "use" : "uses"}` : ""].filter(Boolean).join(" · ")} fraction={row.cost_usd / maxBy(detail.by_agent, (r) => r.cost_usd)} />
                      ))}
                    </div> : null
                  }) : undefined}
                </TelemetrySection>

                <TelemetrySection title="By model">
                  {detail.by_model.length ? detail.by_model.map((row) => (
                    <BarRow key={row.model ?? ""} label={row.model ?? "Model not reported"} value={row.unpriced_turns && row.cost_usd <= 0 ? "Cost not reported" : formatCost(row.cost_usd, row.unpriced_turns > 0)} meta={calls(row.turns)} fraction={row.cost_usd / maxBy(detail.by_model, (r) => r.cost_usd)} />
                  )) : undefined}
                </TelemetrySection>

                <TelemetrySection title="By tool">
                  {detail.by_tool.length ? detail.by_tool.map((row) => (
                    <BarRow key={row.tool} label={row.tool} value={`${row.calls} calls`} meta={`${number(row.avg_duration_ms)} ms avg · ${number(row.total_duration_ms)} ms total`} fraction={row.calls / maxBy(detail.by_tool, (r) => r.calls)} />
                  )) : undefined}
                </TelemetrySection>

                <TelemetrySection title="Recent model calls">
                  {detail.recent_turns.length ? <ul className="telemetry-recent">
                    {detail.recent_turns.map((row, index) => <li key={`${row.ts}-${index}`}>
                      <div className="telemetry-bar-head">
                        <span className="telemetry-bar-label">{row.label}</span>
                        <span className="telemetry-bar-value">{row.cost_usd === null ? "Cost not reported" : money(row.cost_usd)}</span>
                      </div>
                      <span className="telemetry-bar-meta">{row.model ?? "Model not reported"} · {number(row.tokens)} tokens</span>
                      <time className="telemetry-bar-meta" dateTime={new Date(row.ts).toISOString()}>{new Date(row.ts).toLocaleString()}</time>
                    </li>)}
                  </ul> : undefined}
                </TelemetrySection>

                <TelemetrySection title="Recent conversations">
                  {detail.recent_sessions.length ? <ul className="telemetry-recent">
                    {detail.recent_sessions.map((row) => <li key={row.id}>
                      <div className="telemetry-bar-head">
                        <span className="telemetry-bar-label">{row.title || row.label}</span>
                        <span className="telemetry-bar-value">{money(row.cost_usd)}</span>
                      </div>
                      {row.title && <span className="telemetry-bar-meta">{row.label}</span>}
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

// Slices keep the caller's order so each group keeps its colour.
function CostDonut({ rows, money }: { rows: Array<{ key: string; label: string; cost_usd: number }>; money: (value: number) => string }) {
  const slices = rows.filter((row) => row.cost_usd > 0)
  const total = slices.reduce((sum, row) => sum + row.cost_usd, 0)
  const radius = 60
  const circumference = 2 * Math.PI * radius
  let offset = 0

  return (
    <div className="telemetry-donut">
      <svg className="telemetry-donut-svg" viewBox="0 0 160 160" role="img" aria-label="Cost split between background agents and your work">
        <g transform="rotate(-90 80 80)">
          {slices.map((row, index) => {
            const length = (row.cost_usd / total) * circumference
            const dash = (
              <circle
                key={row.key}
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
          <li key={row.key}>
            <span className="telemetry-donut-swatch" style={{ background: DONUT_COLORS[index % DONUT_COLORS.length] }} aria-hidden />
            <span className="telemetry-donut-name">{row.label}</span>
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

function calls(count: number) {
  return `${number(count)} model ${count === 1 ? "call" : "calls"}`
}

function number(value: number) {
  return new Intl.NumberFormat().format(Math.round(value))
}
