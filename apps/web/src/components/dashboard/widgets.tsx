import { useEffect, useRef, useState, type FormEvent, type ReactNode } from "react"
import { Link, useNavigate } from "react-router-dom"
import { addSampleMatter, type Jurisdiction } from "../../api/ingest/matters"
import { savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import type { SetupProgress } from "../../api/ingest/setup"
import { formatAgo, formatCost, formatDate } from "../../lib/format"
import { loadPrefs, savePrefs } from "../../prefs"
import JurisdictionFlag, { countryOf } from "../JurisdictionFlag"
import JurisdictionSelect from "../JurisdictionSelect"
import { SUGGESTED_PRACTICE_AREAS } from "../settings/MatterDefaultsTab"
import { useOpenSettings } from "../settings/navigation"
import type { Tab } from "../settings/types"
import { STARTERS } from "../ChatPanel"
import { useToast } from "../Toast"
import { ActivityCalendar, AreaChart, Ring, Sparkline, SpreadBars, StackedBar, WeekBars } from "./charts"
import {
  addDays,
  attention,
  averageDraftCost,
  averageMatterCost,
  bucketed,
  chartStart,
  costByMatter,
  costEntries,
  dailyConversations,
  dailySpend,
  drafts,
  lastDays,
  PERIOD_LABELS,
  periodRange,
  progress,
  recentMatters,
  spendCsv,
  startOfDay,
  type DashboardData,
  type DaySpend,
  type DraftState,
  type Period,
  type Summary,
} from "./data"
import { icons, ShapeIcon, SHAPES, type IconName, type ShapeName } from "./icons"
import type { Stage, WidgetId } from "./layout"

// `onPreferences` takes settings a widget saved, so the page shows them at once.
export type WidgetContext = { data: DashboardData; setup: SetupProgress | null; summary: Summary; period: Period; now: number; onPreferences?: (preferences: DraftingPreferences) => void; onRetry: () => void }

// Each widget's badge and tone. Cards stay on the surface; the three cost
// tiles wear a light tint and This week is ink. The page makes one action
// block bold: see Dashboard.
export const LOOKS: Record<WidgetId, { icon: IconName; shape: ShapeName; tone?: "ink" | "tint-1" | "tint-2" | "tint-3" }> = {
  journey: { icon: "flag", shape: "sunny" },
  setup: { icon: "setup", shape: "sunny" },
  attention: { icon: "attention", shape: "flower" },
  drafts: { icon: "draft", shape: "clover" },
  spend: { icon: "spend", shape: "cookie", tone: "tint-1" },
  "matter-cost": { icon: "matter", shape: "clover", tone: "tint-2" },
  "draft-cost": { icon: "draft", shape: "flower", tone: "tint-3" },
  week: { icon: "week", shape: "clover", tone: "ink" },
  recent: { icon: "recent", shape: "clover" },
  "matter-spend": { icon: "bars", shape: "cookie" },
  activity: { icon: "activity", shape: "sunny" },
  trend: { icon: "trend", shape: "flower" },
  tokens: { icon: "tokens", shape: "cookie" },
  sample: { icon: "sample", shape: "flower" },
  ideas: { icon: "idea", shape: "clover" },
  actions: { icon: "actions", shape: "sunny" },
  library: { icon: "library", shape: "clover" },
}

const SHAPE_ORDER = ["cookie", "clover", "flower", "sunny"] as const

const money = (value: number) => (value > 0 && value < 0.01 ? "< $0.01" : `$${value.toFixed(2)}`)

// Axis labels: whole dollars once amounts reach them.
const axisMoney = (value: number) => (value === 0 ? "$0" : value >= 10 ? `$${Math.round(value)}` : `$${value.toFixed(value >= 1 ? 1 : 2)}`)

function compact(value: number) {
  if (value >= 1e9) return `${(value / 1e9).toFixed(1)}B`
  if (value >= 1e6) return `${(value / 1e6).toFixed(1)}M`
  if (value >= 1e3) return `${(value / 1e3).toFixed(1)}k`
  return String(value)
}

const plural = (count: number, one: string, many = `${one}s`) => `${count} ${count === 1 ? one : many}`

const lower = (period: Period) => PERIOD_LABELS[period].toLowerCase()

const ALL_TIME = { start: -Infinity, end: Infinity }

// The window before, as an amount. A percentage would mislead early in a
// month, set against the whole of the month before.
const BEFORE: Partial<Record<Period, string>> = { month: "Last month", "last-month": "Month before", quarter: "3 months before" }

function Before({ label, value }: { label?: string; value?: number }) {
  if (value === undefined || !label) return null
  return <span className="dash-before">{label} {formatCost(value)}</span>
}

function Empty({ icon, title, children, action }: { icon: IconName; title: string; children?: ReactNode; action?: ReactNode }) {
  return (
    <div className="dash-empty">
      <ShapeIcon icon={icon} shape="cookie" className="dash-empty-badge" />
      <p className="dash-empty-title">{title}</p>
      {children && <p className="dash-empty-text">{children}</p>}
      {action}
    </div>
  )
}

function Stat({ value, caption, note, children }: { value: string; caption: ReactNode; note?: string; children?: ReactNode }) {
  return (
    <div className="dash-stat">
      <p className="dash-stat-value">{value}</p>
      <p className="dash-stat-caption">{caption}</p>
      {note && <p className="dash-stat-note">{note}</p>}
      <div className="dash-stat-chart">{children}</div>
    </div>
  )
}

function SpendTile({ data, summary, period, now }: WidgetContext) {
  const range = periodRange(period, now)
  return (
    <Stat
      value={formatCost(summary.spend)}
      caption={<><span>{plural(summary.conversations, "conversation")}</span><Before label={BEFORE[period]} value={summary.previous} /></>}
      note={summary.unpriced ? "Estimate: some model use reported no price." : undefined}
    >
      <Sparkline values={bucketed(costEntries(data.matters), chartStart(data.matters, range), range.end, 12)} label={`Spend, ${lower(period)}`} />
    </Stat>
  )
}

function MatterCostTile({ data }: WidgetContext) {
  const average = averageMatterCost(data.matters)
  if (!average) return <Stat value="—" caption="Shows once a matter has spend" />
  return (
    <Stat value={formatCost(average.average)} caption={`Across ${plural(average.matters, "matter")}`}>
      <SpreadBars values={costByMatter(data.matters, ALL_TIME).map((row) => row.cost)} average={average.average} format={money} label="Spend per matter; the line marks the average" />
    </Stat>
  )
}

function DraftCostTile({ data }: WidgetContext) {
  const average = averageDraftCost(data.matters)
  if (!average) return <Stat value="—" caption="Shows once a draft is finished" />
  return (
    <Stat value={formatCost(average.average)} caption={`${plural(average.documents, "document")} from ${plural(average.runs, "run")}`}>
      <SpreadBars values={drafts(data.matters).map((draft) => draft.cost / draft.titles.length)} average={average.average} format={money} label="Cost per document in each run; the line marks the average" />
    </Stat>
  )
}

// A day's tooltip line: the matters that spent the most.
function topMatters(day: DaySpend) {
  if (!day.matters.length) return undefined
  const named = day.matters.slice(0, 2).map(({ matter, cost }) => `${matter.title} ${money(cost)}`)
  const more = day.matters.length - named.length
  return [...named, ...(more > 0 ? [plural(more, "other matter")] : [])].join(" · ")
}

function WeekWidget({ data, now }: WidgetContext) {
  const days = dailySpend(data.matters, lastDays(14, now))
  const week = days.slice(7)
  const total = week.reduce((sum, row) => sum + row.cost, 0)
  const before = days.slice(0, 7).reduce((sum, row) => sum + row.cost, 0)
  const today = startOfDay(now)
  return (
    <div className="dash-week-body">
      <div className="dash-week-head">
        <p className="dash-hero">{formatCost(total)}</p>
        <p className="dash-hero-caption">spent in the last seven days</p>
        <Before label="Week before" value={before} />
      </div>
      <WeekBars
        days={week.map((row) => ({ day: row.day, cost: row.cost, today: row.day === today, label: new Date(row.day).toLocaleDateString("en-GB", { weekday: "short" }), detail: topMatters(row) }))}
        format={money}
      />
    </div>
  )
}

// A local date for file names.
function dateStamp(time: number) {
  const date = new Date(time)
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`
}

function MatterSpend({ data, period, now }: WidgetContext) {
  const rows = costByMatter(data.matters, periodRange(period, now))
  if (!rows.length) return <Empty icon="bars" title="No spend in this period">Each matter's share appears once you work in it.</Empty>
  const shown = rows.slice(0, 6)
  const rest = rows.slice(6)
  const max = shown[0].cost
  function download() {
    const url = URL.createObjectURL(new Blob([spendCsv(rows)], { type: "text/csv" }))
    Object.assign(document.createElement("a"), { href: url, download: `spend-${period}-${dateStamp(now)}.csv` }).click()
    setTimeout(() => URL.revokeObjectURL(url))
  }
  return (
    <div className="dash-hbars">
      <ol>
        {shown.map((row) => (
          <li key={row.matter.id}>
            <span className="dash-hbar-label">
              <Link to={`/matter/${row.matter.id}`} className="dash-hbar-name" title={row.matter.title}>{row.matter.title}</Link>
              <span className="dash-hbar-client">{[row.matter.client, row.matter.reference].filter(Boolean).join(" · ") || "No client recorded"}</span>
            </span>
            <span className="dash-hbar-track" aria-hidden="true">
              <span className="dash-hbar-fill" style={{ width: `max(2px, ${(row.cost / max) * 100}%)` }} />
            </span>
            <span className="dash-hbar-value">{formatCost(row.cost)}</span>
          </li>
        ))}
      </ol>
      <div className="dash-card-foot">
        {rest.length > 0 && <p className="dash-footnote">{plural(rest.length, "other matter")} · {formatCost(rest.reduce((sum, row) => sum + row.cost, 0))}</p>}
        <button type="button" className="dash-text-button" onClick={download}>{icons.download}Export CSV</button>
      </div>
    </div>
  )
}

const ATTENTION = {
  provider: { icon: "plug", tone: "error", text: () => "Reconnect your AI provider" },
  failed: { icon: "error", tone: "error", text: () => "Drafting run failed" },
  questions: { icon: "question", tone: "info", text: (count = 1) => `${plural(count, "question")} waiting` },
  stalled: { icon: "pause", tone: "warning", text: () => "Drafting stalled, press Continue" },
  blocked: { icon: "warning", tone: "warning", text: () => "Drafting stopped for a decision" },
  paused: { icon: "pause", tone: "neutral", text: () => "Drafting paused" },
} as const

function NeedsYou({ data, onRetry }: WidgetContext) {
  const openSettings = useOpenSettings()
  const items = attention(data.matters)
  // A matter that could not be read may hold a question, so never claim
  // nothing is waiting. (The page's banner covers a partial list.)
  const unread = data.matters.filter((row) => row.unread).length
  if (!items.length) {
    return unread
      ? <Empty icon="warning" title={`Couldn't check ${plural(unread, "matter")}`} action={<button type="button" className="dash-text-button" onClick={onRetry}>Retry</button>}>Anything waiting there shows once they load.</Empty>
      : <Empty icon="check" title="Nothing needs you">Questions and stopped drafts show up here.</Empty>
  }
  return (
    <div className="dash-attention-body">
      <ul className="dash-attention">
        {items.slice(0, 6).map((item) => {
          const look = ATTENTION[item.kind]
          return (
            <li key={item.matter.id} className={`tone-${look.tone}`}>
              <span className="dash-attention-icon">{icons[look.icon]}</span>
              <span className="dash-attention-text">
                <strong>{look.text(item.count)}</strong>
                <span title={item.matter.title}>{item.matter.title}</span>
              </span>
              {item.kind === "provider"
                ? <button type="button" className="dash-row-action" onClick={() => openSettings("ai")}>Reconnect</button>
                : <Link className="dash-row-action" to={`/matter/${item.matter.id}?session=${encodeURIComponent(item.session)}`}>Open<span className="sr-only"> {item.matter.title}</span></Link>}
            </li>
          )
        })}
      </ul>
      {items.length > 6 && <p className="dash-footnote">{plural(items.length - 6, "more matter")} waiting</p>}
    </div>
  )
}

const DRAFT_STATES: Record<DraftState, { icon: IconName; tone: string; label: (findings?: number) => string }> = {
  ready: { icon: "check", tone: "good", label: () => "Ready for signature" },
  review: { icon: "draft", tone: "neutral", label: () => "Review draft" },
  findings: { icon: "warning", tone: "warning", label: (count) => (count ? `${plural(count, "finding")} to resolve` : "Findings to resolve") },
  changed: { icon: "warning", tone: "warning", label: () => "Changed since final check" },
}

function RecentDrafts({ data, now }: WidgetContext) {
  const rows = drafts(data.matters, 5)
  if (!rows.length) return <Empty icon="draft" title="No drafts yet">Finished drafts show up here, with what each one needs next.</Empty>
  return (
    <ul className="dash-rows">
      {rows.map((row) => {
        const state = DRAFT_STATES[row.state]
        const more = row.titles.length - 1
        return (
          <li key={row.matter.id}>
            <Link className="dash-row-main" to={`/matter/${row.matter.id}?session=${encodeURIComponent(row.session)}`} title={row.titles.join(", ")}>
              <strong>{row.titles[0]}{more > 0 && <span className="dash-row-more"> +{more} more</span>}</strong>
              <span>{row.matter.title} · {formatAgo(row.finished, now)}</span>
            </Link>
            <span className={`dash-pill tone-${state.tone}`}>{icons[state.icon]}{state.label(row.findings)}</span>
          </li>
        )
      })}
    </ul>
  )
}

// The last 53 weeks as a calendar, Monday first, ending with today's week.
function Activity({ data, now }: WidgetContext) {
  const today = startOfDay(now)
  const weekday = (new Date(today).getDay() + 6) % 7
  const first = addDays(today, -(52 * 7 + weekday))
  const counts = dailyConversations(data.matters, first, 52 * 7 + weekday + 1)
  const total = counts.reduce((sum, count) => sum + count, 0)
  return (
    <div className="dash-activity">
      <p className="dash-activity-total">{total ? `${plural(total, "conversation")} in the last year` : "Each day you start a conversation fills a square."}</p>
      <ActivityCalendar
        counts={counts}
        first={first}
        describe={(day, count) => `${count ? plural(count, "conversation") : "No conversations"} on ${formatDate(day)}`}
      />
    </div>
  )
}

const STATUS = {
  running: { icon: "draft", tone: "neutral", label: "Drafting" },
  blocked: { icon: "warning", tone: "warning", label: "Needs you" },
  failed: { icon: "error", tone: "error", label: "Failed" },
  completed: { icon: "check", tone: "good", label: "Drafted" },
} as const

function Recent({ data, now }: WidgetContext) {
  const rows = recentMatters(data.matters)
  if (!rows.length) return <Empty icon="matter" title="No matters yet"><Link to="/?new=1">Open your first matter</Link> to start.</Empty>
  return (
    <ul className="dash-rows">
      {rows.map((row) => {
        const status = row.status && STATUS[row.status]
        return (
          <li key={row.matter.id}>
            <Link to={`/matter/${row.matter.id}`} className="dash-row-main" title={row.matter.title}>
              <strong>{row.matter.title}</strong>
              <span>{[row.matter.client, row.matter.reference].filter(Boolean).join(" · ") || (row.matter.sample ? "Sample matter" : "No client recorded")}</span>
            </Link>
            <span className="dash-row-meta">
              <span>{formatAgo(row.last, now)}</span>
              <span>{formatCost(row.cost)}</span>
            </span>
            {status && <span className={`dash-pill tone-${status.tone}`}>{icons[status.icon]}{status.label}</span>}
          </li>
        )
      })}
    </ul>
  )
}

function Tokens({ summary, period }: WidgetContext) {
  const { input, cached, output, reasoning } = summary.tokens
  if (!input && !cached && !output && !reasoning) return <Empty icon="tokens" title="No tokens used in this period" />
  return (
    <StackedBar
      label={`Tokens, ${lower(period)}`}
      segments={[
        { key: "input", label: "Input", value: input, display: compact(input) },
        { key: "cached", label: "Cached input", value: cached, display: compact(cached) },
        { key: "output", label: "Output", value: output, display: compact(output) },
        { key: "reasoning", label: "Reasoning", value: reasoning, display: compact(reasoning) },
      ]}
    />
  )
}

// Daily spend over the window; long windows group by week.
function Trend({ data, period, now }: WidgetContext) {
  const range = periodRange(period, now)
  const daily = dailySpend(data.matters, { start: chartStart(data.matters, range), end: range.end })
  const weekly = daily.length > 92
  const points = weekly
    ? Array.from({ length: Math.ceil(daily.length / 7) }, (_, index) => {
        const slice = daily.slice(index * 7, index * 7 + 7)
        return { time: slice[0].day, value: slice.reduce((sum, row) => sum + row.cost, 0), label: `week of ${formatDate(slice[0].day)}` }
      })
    : daily.map((row) => ({ time: row.day, value: row.cost, label: new Date(row.day).toLocaleDateString("en-GB", { day: "numeric", month: "short" }) }))
  if (!points.some((row) => row.value > 0)) return <Empty icon="trend" title="No spend in this period" />
  return <AreaChart points={points} format={axisMoney} detail={money} label={`${weekly ? "Weekly" : "Daily"} spend, ${lower(period)}`} />
}

// Adds the sample matter (an existing one is returned as is) and opens it at
// `path`, such as a chat with a question waiting. Only ever on the lawyer's ask.
function useSampleMatter() {
  const navigate = useNavigate()
  const notify = useToast()
  const [adding, setAdding] = useState(false)
  function open(path = "") {
    setAdding(true)
    addSampleMatter().then((matter) => navigate(`/matter/${matter.id}${path}`), (error: unknown) => {
      setAdding(false)
      notify("error", error instanceof Error ? error.message : "Could not add the sample matter")
    })
  }
  return { adding, open }
}

// The chat's first starter, seated in the composer for the lawyer to send.
const ASK = `?view=chat&prompt=${encodeURIComponent(STARTERS[0])}`

type Tile = { icon: IconName; title: string; text: string; to?: string; run?: () => void }

function Tiles({ tiles, busy }: { tiles: Tile[]; busy?: boolean }) {
  return (
    <ul className="dash-actions">
      {tiles.map((tile, index) => {
        const body = <>
          <ShapeIcon icon={tile.icon} shape={SHAPE_ORDER[index % 4]} />
          <strong>{tile.title}</strong>
          <span>{tile.text}</span>
        </>
        return (
          <li key={tile.title}>
            {tile.to ? <Link to={tile.to}>{body}</Link> : <button type="button" onClick={tile.run} disabled={busy} aria-busy={busy}>{body}</button>}
          </li>
        )
      })}
    </ul>
  )
}

function Actions() {
  const sample = useSampleMatter()
  return (
    <Tiles busy={sample.adding} tiles={[
      { icon: "matter", title: "New matter", text: "Open a folder for an engagement", to: "/?new=1" },
      { icon: "template", title: "Add a template", text: "Draft from your own precedent", to: "/templates" },
      { icon: "client", title: "Add a client", text: "Keep parties and details on file", to: "/clients" },
      { icon: "sample", title: "Try the sample", text: "Explore a matter with documents", run: () => sample.open() },
    ]} />
  )
}

// Ways in. Asking and reviewing happen in a matter: the one last worked in,
// else the sample if the lawyer added it, else a new one.
function Ideas({ data }: WidgetContext) {
  const own = progress(data.matters).target
  const matter = own ?? data.matters.find((row) => row.matter.sample)?.matter
  const at = (path: string) => (matter ? `/matter/${matter.id}${path}` : "/?new=1")
  return (
    <Tiles tiles={[
      { icon: "documents", title: "Ask about documents", text: own ? `Key points in ${own.title}` : "Key points, deadlines and gaps", to: at(ASK) },
      { icon: "template", title: "Draft from a template", text: "Start from the firm's own precedent", to: "/templates" },
      { icon: "review", title: "Review across documents", text: "Ask the same questions of every file", to: at("?view=review") },
      { icon: "workflow", title: "Run a workflow", text: "A repeatable routine for a matter", to: "/workflows" },
    ]} />
  )
}

function Library({ data }: WidgetContext) {
  const tiles: { icon: IconName; label: string; count?: number; first: string; to: string }[] = [
    { icon: "matter", label: "Matters", count: data.error ? undefined : data.matters.length, first: "Open your first", to: "/?new=1" },
    { icon: "template", label: "Templates", count: data.templates, first: "Add your first", to: "/templates" },
    { icon: "playbook", label: "Playbooks", count: data.playbooks, first: "Add your first", to: "/playbooks" },
    { icon: "client", label: "Clients", count: data.clients, first: "Add your first", to: "/clients" },
  ]
  return (
    <ul className="dash-library">
      {tiles.map((tile, index) => (
        <li key={tile.label} className={tile.count === 0 ? "is-zero" : undefined}>
          <Link to={tile.count === 0 || tile.label !== "Matters" ? tile.to : "/"}>
            <ShapeIcon icon={tile.icon} shape={SHAPE_ORDER[index]} />
            <span className="dash-library-text">
              <span>{tile.label}</span>
              {tile.count === 0 && <span className="dash-library-first">{tile.first}{icons.arrowRight}</span>}
            </span>
            <strong>{tile.count ?? "—"}</strong>
          </Link>
        </li>
      ))}
    </ul>
  )
}

type Step = { title: string; detail: string; done: boolean; note: string; action?: { label: string; to: string } }

// Four milestones from a new matter to a first draft, each a shape on a track;
// the next one carries its action.
function FirstSteps({ data }: WidgetContext) {
  const sample = useSampleMatter()
  const steps = progress(data.matters)
  const added = data.matters.find((row) => row.matter.sample)?.matter
  const at = steps.target && `/matter/${steps.target.id}`
  const list: Step[] = [
    { title: "Open a matter", detail: "One home for an engagement's documents and chats.", done: steps.matters > 0, note: plural(steps.matters, "matter"), action: { label: "New matter", to: "/?new=1" } },
    { title: "Add documents", detail: "Upload the agreements, emails or notes you work from.", done: steps.documents > 0, note: "Documents added", action: at ? { label: "Add documents", to: `${at}?view=documents` } : undefined },
    { title: "Ask a question", detail: "Ask about the materials in your own words.", done: steps.conversations > 0, note: plural(steps.conversations, "conversation"), action: at ? { label: "Ask a question", to: `${at}${ASK}` } : undefined },
    { title: "Draft a document", detail: "Draft from your instructions and the matter's files.", done: steps.drafted > 0, note: `${plural(steps.drafted, "document")} drafted`, action: at ? { label: "Start drafting", to: `${at}?view=chat&agent=draft` } : undefined },
  ]
  const current = list.findIndex((step) => !step.done)
  const run = steps.run
  return (
    <div className="dash-journey-body">
      <ol className="dash-journey">
        {list.map((step, index) => {
          const drafting = index === 3 && !step.done && run
          return (
            <li key={step.title} className={step.done ? "is-done" : index === current ? "is-current" : ""} aria-current={index === current ? "step" : undefined}>
              <span className="dash-step-node" aria-hidden="true">
                <svg viewBox="-1 -1 2 2"><path d={SHAPES[SHAPE_ORDER[index]]} /></svg>
                <span>{step.done ? icons.check : index + 1}</span>
              </span>
              <span className="dash-step-text">
                <strong>{step.title}</strong>
                <span>{step.detail}</span>
                {step.done ? (
                  <span className="dash-step-note"><span className="sr-only">Done: </span>{step.note}</span>
                ) : drafting ? (
                  <Link className="dash-step-action" to={`/matter/${drafting.matter.id}?session=${encodeURIComponent(drafting.session)}`}>
                    {drafting.waiting ? "Drafting needs you" : "Drafting now"}{icons.arrowRight}
                  </Link>
                ) : index === current && step.action ? (
                  <Link className="dash-step-action" to={step.action.to}>{step.action.label}{icons.arrowRight}</Link>
                ) : null}
                {index === 0 && !step.done && (added
                  ? <Link className="dash-step-alt" to={`/matter/${added.id}${ASK}`}>or open the sample matter</Link>
                  : <button type="button" className="dash-step-alt" onClick={() => sample.open(ASK)} disabled={sample.adding} aria-busy={sample.adding}>or explore a sample matter first</button>)}
              </span>
            </li>
          )
        })}
      </ol>
      <p className="dash-journey-note">
        Each draft arrives as a .docx file with a memo on what it covers and what is still open. Drafting runs on your own AI provider; its estimated cost shows here.
      </p>
    </div>
  )
}

// The fictional sample, with a question already seated when it opens.
function SampleMatter({ data }: WidgetContext) {
  const sample = useSampleMatter()
  const added = data.matters.find((row) => row.matter.sample)?.matter
  return (
    <div className="dash-sample">
      <p className="dash-sample-text">A fictional engagement letter, with cited answers and a legal review ready to look through.</p>
      {added
        ? <Link className="dash-text-button" to={`/matter/${added.id}${ASK}`}>Open the sample matter{icons.arrowRight}</Link>
        : <button type="button" className="dash-text-button" onClick={() => sample.open(ASK)} disabled={sample.adding} aria-busy={sample.adding}>Add the sample matter{icons.arrowRight}</button>}
    </div>
  )
}

type SetupItem = { id: string; icon: IconName; title: string; detail?: string; done: boolean; optional?: boolean; tab?: Tab; to?: string }

// What the lawyer has not told doc.haus yet, from saved settings and the
// library. The essentials shape every draft and answer; extras say so. The AI
// model is never listed (onboarding cannot finish without one), nor are
// document legends, which are the firm's choice and absent by design. A first
// matter and the sample have their own widgets.
// Jurisdictions live in this browser, so the profile card passes its own copy.
export function setupItems({ data }: Pick<WidgetContext, "data">, jurisdictions = loadPrefs().defaultJurisdictions): SetupItem[] {
  const preferences = data.preferences
  const items: (SetupItem | false)[] = [
    !!preferences && { id: "signed", icon: "person", title: "Name and firm", done: !!preferences.attorney.trim() && !!preferences.firm.trim(), tab: "drafting" },
    { id: "jurisdiction", icon: "globe", title: "Jurisdictions", done: jurisdictions.length > 0, tab: "matters" },
    !!preferences && { id: "practice", icon: "practice", title: "Practice areas", done: preferences.practiceAreas.length > 0, tab: "matters" },
    !!preferences && { id: "research", icon: "search", title: "Web research", detail: "A free Exa key adds full results and the source text behind each citation.", done: !!preferences.searchApiKey.trim(), optional: true, tab: "research" },
    // Listed only where this installation can set it up or it is managed elsewhere.
    !!data.word && (data.word.canSetup || ["ready", "external"].includes(data.word.status)) && !["unsupported", "word_missing"].includes(data.word.status) && { id: "word", icon: "word", title: "Word add-in", detail: "Review and redline without leaving Word.", done: data.word.status === "ready" || (data.word.status === "external" && data.word.available), optional: true, tab: "addins" },
    !!preferences && { id: "style", icon: "style", title: "House style notes", detail: "Standing instructions for every draft and edit.", done: !!preferences.houseStyle.trim(), optional: true, tab: "drafting" },
    data.templates !== undefined && { id: "template", icon: "template", title: "A template", detail: "Draft from the firm's own precedents.", done: data.templates > 0, optional: true, to: "/templates" },
  ]
  return items.filter((item): item is SetupItem => !!item)
}

// Every essential is in, so the profile card can step aside on the next load.
export function setupComplete(data: DashboardData) {
  return !!data.preferences && setupItems({ data }).every((item) => item.optional || item.done)
}

const unique = (values: string[]) => values.filter((value, index) => values.indexOf(value) === index)
const toggled = (values: string[], value: string) => (values.includes(value) ? values.filter((other) => other !== value) : [...values, value])

// Packs in the browser's own region, country-wide first: en-GB offers England
// and Wales, Scotland and Northern Ireland.
function nearby(all: Jurisdiction[]) {
  const regions = new Set(navigator.languages.flatMap((tag) => {
    try { return new Intl.Locale(tag).region ?? [] } catch { return [] }
  }))
  return all.filter((pack) => regions.has(countryOf(pack.code)))
    .sort((a, b) => Number(a.code.includes("-")) - Number(b.code.includes("-")))
    .slice(0, 3).map((pack) => pack.code)
}

// A few common areas to tap; the picker holds the full list and the lawyer's own.
const PRACTICE_PICKS = ["Commercial", "Corporate", "Employment", "Litigation", "Real estate", "Private client"]

type Save = (patch: Partial<DraftingPreferences>, failed: string) => Promise<boolean>

// An open essential is a tile to fill in; a done one folds into a quiet line
// (`summary`) with its way to change it.
function ProfileTile({ id, icon, shape, title, summary, children }: { id: string; icon: IconName; shape: ShapeName; title: string; summary?: ReactNode; children?: ReactNode }) {
  if (summary) {
    return (
      <section className="dash-profile-row" aria-labelledby={`dash-profile-${id}`}>
        <span className="dash-profile-check" aria-hidden="true">{icons.check}</span>
        <h4 id={`dash-profile-${id}`}>{title}<span className="sr-only">, done</span></h4>
        {summary}
      </section>
    )
  }
  return (
    <section className="dash-profile-tile" aria-labelledby={`dash-profile-${id}`}>
      <div className="dash-profile-tile-head">
        <ShapeIcon icon={icon} shape={shape} />
        <h4 id={`dash-profile-${id}`}>{title}</h4>
      </div>
      {children}
    </section>
  )
}

// Name and firm, filled in where they show: as who prepared the draft.
function SignedAs({ preferences, save }: { preferences: DraftingPreferences; save: Save }) {
  const complete = !!preferences.attorney.trim() && !!preferences.firm.trim()
  const [editing, setEditing] = useState(!complete)
  const [name, setName] = useState(preferences.attorney)
  const [firm, setFirm] = useState(preferences.firm)
  const [saving, setSaving] = useState(false)
  const editButton = useRef<HTMLButtonElement>(null)
  const nameInput = useRef<HTMLInputElement>(null)
  const changed = name.trim() !== preferences.attorney.trim() || firm.trim() !== preferences.firm.trim()
  function edit() {
    // From what is saved now, which Settings may have changed since.
    setName(preferences.attorney)
    setFirm(preferences.firm)
    setEditing(true)
    requestAnimationFrame(() => nameInput.current?.focus())
  }
  async function submit(event: FormEvent) {
    event.preventDefault()
    setSaving(true)
    const saved = await save({ attorney: name.trim(), firm: firm.trim() }, "Could not save your name and firm. Try again.")
    setSaving(false)
    if (!saved || !name.trim() || !firm.trim()) return
    setEditing(false)
    // The form gives way to the summary line; keep focus in it.
    requestAnimationFrame(() => editButton.current?.focus())
  }
  if (complete && !editing) {
    return (
      <ProfileTile id="signed" icon="person" shape="cookie" title="Name and firm" summary={<>
        <span className="dash-profile-value">{preferences.attorney}, {preferences.firm}</span>
        <button ref={editButton} type="button" className="dash-text-button" onClick={edit}>Edit<span className="sr-only"> name and firm</span></button>
      </>} />
    )
  }
  return (
    <ProfileTile id="signed" icon="person" shape="cookie" title="Name and firm">
      <form className="dash-profile-form" onSubmit={submit}>
        <p className="dash-profile-detail">Signs your comments and tracked changes, and names who prepared each draft.</p>
        <label><span>Your name</span><input ref={nameInput} value={name} onChange={(event) => setName(event.target.value)} placeholder="Your full name" autoComplete="name" disabled={saving} /></label>
        <label><span>Firm</span><input value={firm} onChange={(event) => setFirm(event.target.value)} placeholder="Firm name" autoComplete="organization" disabled={saving} /></label>
        <button type="submit" className="btn-sm" disabled={saving || !changed} aria-busy={saving}>{saving ? "Saving…" : "Save"}</button>
      </form>
    </ProfileTile>
  )
}

function Jurisdictions({ all, codes, setCodes }: { all: Jurisdiction[]; codes: string[]; setCodes: (codes: string[]) => void }) {
  const notify = useToast()
  // Offered to tap until the card is next opened; after that the choice shows.
  const [offer] = useState(() => codes.length === 0)
  const name = (code: string) => all.find((pack) => pack.code === code)?.name ?? code
  function change(next: string[]) {
    try {
      savePrefs({ defaultJurisdictions: next }, { requireStorage: true })
      setCodes(next)
    } catch {
      notify("error", "Could not save your jurisdictions in this browser. Try again.")
    }
  }
  const picker = (trigger: ReactNode, title: string) =>
    all.length > 0 && <JurisdictionSelect jurisdictions={all} selected={codes} onChange={change} trigger={trigger} triggerTitle={title} />
  if (!offer && codes.length) {
    return (
      <ProfileTile id="jurisdictions" icon="globe" shape="clover" title="Jurisdictions" summary={<>
        <span className="dash-profile-value dash-profile-flags-inline">
          {codes.map((code) => <span key={code}><JurisdictionFlag code={code} region />{name(code)}</span>)}
        </span>
        {picker(<span className="dash-text-button">Change<span className="sr-only"> jurisdictions</span></span>, "Change jurisdictions")}
      </>} />
    )
  }
  const shown = unique([...nearby(all), ...codes])
  return (
    <ProfileTile id="jurisdictions" icon="globe" shape="clover" title="Jurisdictions">
      <p className="dash-profile-detail">Which jurisdictions do you work in? Your picks are preselected on every new matter.</p>
      {shown.length > 0 && (
        <ul className="dash-profile-flags">
          {shown.map((code) => (
            <li key={code}>
              <button type="button" aria-pressed={codes.includes(code)} onClick={() => change(toggled(codes, code))}>
                <JurisdictionFlag code={code} region />
                <span>{name(code)}</span>
                <span className="dash-profile-toggle" aria-hidden="true">{codes.includes(code) ? icons.check : icons.plus}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
      {picker(<span className="dash-text-button">All {all.length} jurisdictions</span>, "Choose from every jurisdiction")}
    </ProfileTile>
  )
}

function PracticeAreas({ preferences, save }: { preferences: DraftingPreferences; save: Save }) {
  // The latest change shows while it saves; once the last save settles, the
  // saved areas show, so a failed save falls back to what was kept.
  const [pending, setPending] = useState<string[]>()
  const latest = useRef(0)
  const areas = pending ?? preferences.practiceAreas
  const [offer] = useState(() => preferences.practiceAreas.length === 0)
  function change(next: string[]) {
    const ticket = ++latest.current
    setPending(next)
    void save({ practiceAreas: next }, "Could not save your practice areas. Try again.").then(() => { if (ticket === latest.current) setPending(undefined) })
  }
  const options = unique([...areas, ...SUGGESTED_PRACTICE_AREAS]).map((area) => ({ code: area, name: area }))
  const picker = (trigger: ReactNode, title: string) => <JurisdictionSelect practiceAreas jurisdictions={options} selected={areas} onChange={change} trigger={trigger} triggerTitle={title} />
  if (!offer && areas.length) {
    return (
      <ProfileTile id="practice" icon="practice" shape="flower" title="Practice areas" summary={<>
        <span className="dash-profile-value">{areas.join(", ")}</span>
        {picker(<span className="dash-text-button">Change<span className="sr-only"> practice areas</span></span>, "Change practice areas")}
      </>} />
    )
  }
  return (
    <ProfileTile id="practice" icon="practice" shape="flower" title="Practice areas">
      <p className="dash-profile-detail">What kind of work do you do? It steers drafting and research toward it.</p>
      <ul className="dash-profile-areas">
        {unique([...PRACTICE_PICKS, ...areas]).map((area) => (
          <li key={area}>
            <button type="button" aria-pressed={areas.includes(area)} onClick={() => change(toggled(areas, area))}>{areas.includes(area) && icons.check}{area}</button>
          </li>
        ))}
      </ul>
      {picker(<span className="dash-text-button">More areas, or add your own</span>, "Choose or add practice areas")}
    </ProfileTile>
  )
}

function PracticeProfile(context: WidgetContext) {
  const { data, onPreferences, onRetry } = context
  const openSettings = useOpenSettings()
  const navigate = useNavigate()
  const notify = useToast()
  const [codes, setCodes] = useState(() => loadPrefs().defaultJurisdictions)
  // A refresh re-reads this browser's choice, which another tab may have changed.
  useEffect(() => setCodes(loadPrefs().defaultJurisdictions), [data])
  // One save at a time, in order, so the last change is the one kept.
  const queue = useRef(Promise.resolve(true))
  const save: Save = (patch, failed) => {
    queue.current = queue.current.then(() => savePreferences(patch)).then((saved) => {
      onPreferences?.(saved)
      return true
    }, () => {
      notify("error", failed)
      return false
    })
    return queue.current
  }
  if (data.preferencesFailed) {
    return <Empty icon="warning" title="Could not load your details" action={<button type="button" className="dash-text-button" onClick={onRetry}>Retry</button>}>Your name, firm and practice areas show here once they load.</Empty>
  }
  const items = setupItems(context, codes)
  const essentials = items.filter((item) => !item.optional)
  const extras = items.filter((item) => item.optional)
  const left = essentials.filter((item) => !item.done).length
  const preferences = data.preferences
  return (
    <div className="dash-profile">
      <div className="dash-setup-progress">
        <Ring done={essentials.length - left} total={essentials.length} size={72} />
        <div>
          <p className="dash-setup-title">{left ? `${plural(left, "detail")} to go` : "Your practice is set up"}</p>
          <p className="dash-setup-text">Every draft starts with your name, your law and your kind of work.</p>
        </div>
      </div>
      <div className="dash-profile-tiles">
        {preferences && <SignedAs preferences={preferences} save={save} />}
        <Jurisdictions all={data.jurisdictions ?? []} codes={codes} setCodes={setCodes} />
        {preferences && <PracticeAreas preferences={preferences} save={save} />}
      </div>
      {extras.length > 0 && (
        <div className="dash-profile-extras">
          <span className="dash-profile-extras-label">Extras</span>
          <ul>
            {extras.map((item) => (
              <li key={item.id}>
                <button type="button" className={item.done ? "is-done" : undefined} title={item.detail} onClick={() => (item.tab ? openSettings(item.tab) : item.to && navigate(item.to))}>
                  {item.done ? icons.check : icons[item.icon]}
                  {item.title}
                  {item.done && <span className="sr-only">, done</span>}
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  )
}

// Starting out until a draft of the lawyer's own has finished.
export const stageOf = (data: DashboardData): Stage => (progress(data.matters).drafted ? "practice" : "starting")

export const RENDER: Record<WidgetId, (context: WidgetContext) => ReactNode> = {
  journey: (context) => <FirstSteps {...context} />,
  setup: (context) => <PracticeProfile {...context} />,
  attention: (context) => <NeedsYou {...context} />,
  drafts: (context) => <RecentDrafts {...context} />,
  spend: (context) => <SpendTile {...context} />,
  "matter-cost": (context) => <MatterCostTile {...context} />,
  "draft-cost": (context) => <DraftCostTile {...context} />,
  week: (context) => <WeekWidget {...context} />,
  recent: (context) => <Recent {...context} />,
  "matter-spend": (context) => <MatterSpend {...context} />,
  activity: (context) => <Activity {...context} />,
  trend: (context) => <Trend {...context} />,
  tokens: (context) => <Tokens {...context} />,
  sample: (context) => <SampleMatter {...context} />,
  ideas: (context) => <Ideas {...context} />,
  actions: () => <Actions />,
  library: (context) => <Library {...context} />,
}

// Captions under a widget's title, where the figure needs its basis stated.
// Every card that shares a row carries a subtitle, so titles line up.
export function caption(id: WidgetId, period: Period) {
  switch (id) {
    case "spend": case "matter-spend": case "tokens": case "trend":
      return PERIOD_LABELS[period]
    case "matter-cost":
      return "All time"
    case "draft-cost":
      return "All time, latest run per matter"
    case "week":
      return "Last seven days"
    case "attention":
      return "Drafting runs waiting on you"
    case "drafts":
      return "Latest first, with what each needs"
    case "recent":
      return "Latest first, with all-time spend"
    case "activity":
      return "Conversations started each day"
    case "actions":
      return "Start something new"
    case "library":
      return "Templates, playbooks, clients and matters"
    case "journey":
      return "From a new matter to your first draft"
    case "setup":
      return "Used in every draft and answer"
    case "sample":
      return "Fictional, safe to explore"
    case "ideas":
      return "A few ways to start"
  }
}
