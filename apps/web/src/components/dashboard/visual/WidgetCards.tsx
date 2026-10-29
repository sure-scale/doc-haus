import type { ReactNode } from "react"
import "../../../styles/dashboard-widgets.css"
import { Link } from "react-router-dom"
import { icons, ShapeIcon, type IconName } from "../icons"
import { Sparkline, SpreadBars, WeekBars } from "../charts"
import { formatChartValue } from "./ChartCards"

// These view models contain only JSON values. Data sources and subscriptions
// prepare them; the cards render them without fetching or running expressions.
export type StatVisual =
  | { type: "sparkline"; values: number[]; label: string }
  | { type: "spread"; values: number[]; average: number; label: string; format?: "money" | "count" | "percent" }
  | { type: "week"; days: { day: number; cost: number; label: string; today: boolean; detail?: string }[] }
export type StatCardData = { value: string; caption: string; tag?: string; tagTitle?: string; note?: string; visual?: StatVisual }
export type AttentionRowData = {
  id: string
  title: string
  detail: string
  icon: IconName
  tone: "error" | "info" | "warning" | "neutral"
  action: { label: string; href?: string }
}
export type MatterRowData = {
  id: string
  title: string
  detail: string
  href: string
  updated: string
  spend: string
  status?: { icon: IconName; label: string; tone: "neutral" | "warning" | "error" | "good" }
}
export type LibraryCardData = { id: string; icon: IconName; label: string; count: number | null; detail: string; emptyLabel: string; href: string }
export type SamplePreviewData = { id: string; icon: IconName; title: string; detail: string }
const SHAPES = ["cookie", "clover", "flower", "sunny"] as const

function StatVisualChart({ visual }: { visual: StatVisual }) {
  if (visual.type === "sparkline") return <Sparkline values={visual.values} label={visual.label} />
  if (visual.type === "spread") return <SpreadBars values={visual.values} average={visual.average} label={visual.label} format={(value) => formatChartValue(value, visual.format)} />
  return <WeekBars days={visual.days} format={(value) => formatChartValue(value, "money")} />
}

export function StatCard({ value, caption, tag, tagTitle, note, visual, children }: StatCardData & { children?: ReactNode }) {
  return <div className="dash-stat">
    <p className="dash-stat-caption">{caption}</p>
    <div className="dash-stat-line"><p className="dash-stat-value">{value}</p>{tag && <span className="dash-stat-tag" title={tagTitle}>{tag}</span>}</div>
    {note && <p className="dash-stat-note">{note}</p>}
    <div className="dash-stat-chart">{children === undefined && visual ? <StatVisualChart visual={visual} /> : children}</div>
  </div>
}

export function AttentionRows({ rows, onAction }: { rows: AttentionRowData[]; onAction?: (id: string) => void }) {
  return <ul className="dash-attention">{rows.map((row) => <li key={row.id} className={`tone-${row.tone}`}>
    <span className="dash-attention-icon">{icons[row.icon]}</span>
    <span className="dash-attention-text"><strong>{row.title}</strong><span title={row.detail}>{row.detail}</span></span>
    {row.action.href
      ? <Link className="dash-row-action" to={row.action.href}>{row.action.label}{icons.arrowRight}<span className="sr-only"> {row.detail}</span></Link>
      : <button type="button" className="dash-row-action" disabled={!onAction} onClick={() => onAction?.(row.id)}>{row.action.label}<span className="sr-only"> {row.detail}</span></button>}
  </li>)}</ul>
}

export function MatterRows({ rows }: { rows: MatterRowData[] }) {
  const hasMeta = rows.some((row) => row.updated || row.spend)
  return <ul className={`dash-rows dash-recent-rows${hasMeta ? "" : " is-status-list"}`}>{rows.map((row) => <li key={row.id}>
    <Link to={row.href} className="dash-row-main" title={row.title}><strong>{row.title}</strong><span>{row.detail}</span></Link>
    <span className="dash-row-meta"><span>{row.updated}</span><span>{row.spend}</span></span>
    <span className="dash-status-cell">{row.status && <span className={`dash-pill tone-${row.status.tone}`}>{icons[row.status.icon]}{row.status.label}</span>}</span>
  </li>)}</ul>
}

export function LibraryCards({ cards }: { cards: LibraryCardData[] }) {
  return <ul className="dash-library">{cards.map((card, index) => <li key={card.id} className={card.count === 0 ? "is-zero" : undefined}>
    <Link to={card.href}>
      <span className="dash-library-top"><ShapeIcon icon={card.icon} shape={SHAPES[index % SHAPES.length]} /><strong>{card.count ?? "—"}</strong></span>
      <span className="dash-library-text"><span>{card.label}</span>{card.count === 0
        ? <span className="dash-library-first">{card.emptyLabel}{icons.arrowRight}</span>
        : <span className="dash-library-detail">{card.detail}</span>}</span>
    </Link>
  </li>)}</ul>
}

export function SamplePreview({ items }: { items: SamplePreviewData[] }) {
  return <div className="dash-sample-preview" aria-label="Inside the sample matter">{items.map((item) => <div key={item.id}>
    <span className="dash-sample-document" aria-hidden="true">{icons[item.icon]}<i /><i /><i /></span><strong>{item.title}</strong><span>{item.detail}</span>
  </div>)}</div>
}

export type ActionTileData = { id: string; icon: IconName; title: string; detail: string; href?: string }
export function ActionTiles({ items, onAction }: { items: ActionTileData[]; onAction?: (id: string) => void }) {
  return <ul className="dash-actions dash-action-tiles">{items.map((item, index) => {
    const body = <><ShapeIcon icon={item.icon} shape={SHAPES[index % SHAPES.length]} /><strong>{item.title}</strong><span>{item.detail}</span></>
    return <li key={item.id}>{item.href
      ? <Link to={item.href}>{body}</Link>
      : <button type="button" onClick={() => onAction?.(item.id)} disabled={!onAction}>{body}</button>}</li>
  })}</ul>
}
