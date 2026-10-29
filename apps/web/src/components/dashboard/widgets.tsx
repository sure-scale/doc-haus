import { useState, type ReactNode } from "react"
import { Link, useNavigate } from "react-router-dom"
import { addSampleMatter } from "../../api/ingest/matters"
import type { SetupProgress } from "../../api/ingest/setup"
import { formatAgo, formatCost, formatDate } from "../../lib/format"
import { loadPrefs } from "../../prefs"
import { useOpenSettings } from "../settings/navigation"
import type { Tab } from "../settings/types"
import { useToast } from "../Toast"
import { ActivityCalendar, AreaChart, Ring, Sparkline, SpreadBars, StackedBar, WeekBars } from "./charts"
import {
  addDays,
  attention,
  bucketed,
  costByMatter,
  dailyConversations,
  dailySpend,
  dayIndex,
  draftedRuns,
  periodStart,
  previousSpend,
  recentMatters,
  sessionCosts,
  startOfDay,
  windowStart,
  type DashboardData,
  type Period,
  type Summary,
} from "./data"
import { icons, ShapeIcon, type IconName, type ShapeName } from "./icons"
import type { WidgetId } from "./layout"

export type WidgetContext = { data: DashboardData; setup: SetupProgress | null; summary: Summary; period: Period; now: number }

// How each widget looks: its badge icon and shape, and its tone. Bold tones
// are full-colour blocks with light text; the rest stay on white.
export const LOOKS: Record<WidgetId, { icon: IconName; shape: ShapeName; tone: string }> = {
  setup: { icon: "setup", shape: "sunny", tone: "mint" },
  spend: { icon: "spend", shape: "cookie", tone: "forest" },
  "matter-cost": { icon: "matter", shape: "clover", tone: "indigo" },
  "draft-cost": { icon: "draft", shape: "flower", tone: "plum" },
  documents: { icon: "documents", shape: "cookie", tone: "sky" },
  week: { icon: "week", shape: "clover", tone: "ink" },
  "matter-spend": { icon: "bars", shape: "cookie", tone: "plain" },
  attention: { icon: "attention", shape: "flower", tone: "plain" },
  activity: { icon: "activity", shape: "sunny", tone: "plain" },
  recent: { icon: "recent", shape: "clover", tone: "plain" },
  tokens: { icon: "tokens", shape: "cookie", tone: "plain" },
  trend: { icon: "trend", shape: "flower", tone: "plain" },
  actions: { icon: "actions", shape: "sunny", tone: "plain" },
  library: { icon: "library", shape: "clover", tone: "plain" },
}


export function periodLabel(period: Period) {
  return period === 0 ? "all time" : `the last ${period} days`
}

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

// Change against the window before, as an arrow and a percentage. Direction
// is not judged good or bad: more spend can mean more work done.
function Change({ now, before, versus }: { now: number; before?: number; versus: string }) {
  if (before === undefined || (before <= 0 && now <= 0)) return null
  if (before <= 0) return <span className="dash-change"><span className="dash-change-icon">{icons.arrowUp}</span>New {versus}</span>
  const change = Math.round(((now - before) / before) * 100)
  return (
    <span className="dash-change">
      <span className="dash-change-icon">{change >= 0 ? icons.arrowUp : icons.arrowDown}</span>
      {Math.abs(change)}% {versus}
    </span>
  )
}

function Empty({ icon, title, children }: { icon: IconName; title: string; children?: ReactNode }) {
  return (
    <div className="dash-empty">
      <ShapeIcon icon={icon} shape="cookie" className="dash-empty-badge" />
      <p className="dash-empty-title">{title}</p>
      {children && <p className="dash-empty-text">{children}</p>}
    </div>
  )
}

function Stat({ value, caption, children }: { value: string; caption: ReactNode; children?: ReactNode }) {
  return (
    <div className="dash-stat">
      <p className="dash-stat-value">{value}</p>
      <p className="dash-stat-caption">{caption}</p>
      <div className="dash-stat-chart">{children}</div>
    </div>
  )
}

function SpendTile({ data, summary, period, now }: WidgetContext) {
  const start = windowStart(data.matters, period, now)
  return (
    <Stat value={formatCost(summary.spend)} caption={<><Change now={summary.spend} before={previousSpend(data.matters, period, now)} versus="vs previous period" /> <span>{plural(summary.conversations, "conversation")}</span></>}>
      <Sparkline values={bucketed(sessionCosts(data.matters), start, now, 12)} label={`Spend over ${periodLabel(period)}`} />
    </Stat>
  )
}

function MatterCostTile({ data, summary, period, now }: WidgetContext) {
  const rows = costByMatter(data.matters, period, now)
  if (summary.averageMatter === undefined) return <Stat value="—" caption="No matter has used the assistant yet" />
  return (
    <Stat value={formatCost(summary.averageMatter)} caption={`Across ${plural(summary.activeMatters, "matter")}`}>
      <SpreadBars values={rows.map((row) => row.cost)} average={summary.averageMatter} format={money} label={`Spend per matter; the line marks the average`} />
    </Stat>
  )
}

function DraftCostTile({ data, summary, period, now }: WidgetContext) {
  const runs = draftedRuns(data.matters, period, now)
  if (summary.averageDocument === undefined) return <Stat value="—" caption="Appears when a drafting run delivers" />
  return (
    <Stat value={formatCost(summary.averageDocument)} caption={`${plural(summary.documents, "document")} from ${plural(runs.length, "run")}`}>
      <SpreadBars values={runs.map((run) => run.cost / run.documents)} average={summary.averageDocument} format={money} label="Cost per document in each run; the line marks the average" />
    </Stat>
  )
}

function DocumentsTile({ data, summary, period, now }: WidgetContext) {
  const runs = draftedRuns(data.matters, period, now)
  const start = windowStart(data.matters, period, now)
  return (
    <Stat value={String(summary.documents)} caption={summary.documents ? `From ${plural(runs.length, "drafting run")}` : "None delivered yet"}>
      <Sparkline values={bucketed(runs.map((run) => ({ time: run.finished, value: run.documents })), start, now, 12)} label={`Documents drafted over ${periodLabel(period)}`} />
    </Stat>
  )
}

function WeekWidget({ data, now }: WidgetContext) {
  const days = dailySpend(data.matters, 14, now)
  const week = days.slice(7)
  const total = week.reduce((sum, row) => sum + row.cost, 0)
  const before = days.slice(0, 7).reduce((sum, row) => sum + row.cost, 0)
  const today = startOfDay(now)
  return (
    <div className="dash-week-body">
      <div className="dash-week-head">
        <div>
          <p className="dash-hero">{formatCost(total)}</p>
          <p className="dash-hero-caption">spent in the last seven days</p>
        </div>
        <Change now={total} before={before} versus="vs week before" />
      </div>
      <WeekBars
        days={week.map((row) => ({ day: row.day, cost: row.cost, today: row.day === today, label: new Date(row.day).toLocaleDateString("en-GB", { weekday: "short" }) }))}
        format={money}
      />
    </div>
  )
}

function MatterSpend({ data, period, now }: WidgetContext) {
  const rows = costByMatter(data.matters, period, now)
  if (!rows.length) return <Empty icon="bars" title="No spend yet">Each matter's share appears once you work in it.</Empty>
  const shown = rows.slice(0, 6)
  const rest = rows.slice(6)
  const max = shown[0].cost
  return (
    <div className="dash-hbars">
      <ol>
        {shown.map((row) => (
          <li key={row.matter.id}>
            <Link to={`/matter/${row.matter.id}`} className="dash-hbar-name">{row.matter.title}</Link>
            <span className="dash-hbar-track" aria-hidden="true">
              <span className="dash-hbar-fill" style={{ width: `${Math.max(2, (row.cost / max) * 100)}%` }} />
            </span>
            <span className="dash-hbar-value">{formatCost(row.cost)}</span>
            <span className="dash-hbar-meta">{plural(row.conversations, "conversation")}</span>
          </li>
        ))}
      </ol>
      {rest.length > 0 && <p className="dash-footnote">{plural(rest.length, "other matter")} · {formatCost(rest.reduce((sum, row) => sum + row.cost, 0))}</p>}
    </div>
  )
}

const ATTENTION = {
  provider: { icon: "plug", tone: "error", text: () => "Reconnect your AI provider" },
  failed: { icon: "error", tone: "error", text: () => "Drafting run failed" },
  questions: { icon: "question", tone: "info", text: (count = 1) => `${plural(count, "question")} waiting` },
  blocked: { icon: "warning", tone: "warning", text: () => "Drafting stopped for a decision" },
  paused: { icon: "pause", tone: "neutral", text: () => "Drafting paused" },
} as const

function NeedsYou({ data }: WidgetContext) {
  const openSettings = useOpenSettings()
  const items = attention(data.matters)
  if (!items.length) return <Empty icon="check" title="Nothing needs you">Questions and stopped drafts show up here.</Empty>
  return (
    <ul className="dash-attention">
      {items.slice(0, 6).map((item) => {
        const look = ATTENTION[item.kind]
        return (
          <li key={item.matter.id} className={`tone-${look.tone}`}>
            <span className="dash-attention-icon">{icons[look.icon]}</span>
            <span className="dash-attention-text">
              <strong>{look.text(item.count)}</strong>
              <span>{item.matter.title}</span>
            </span>
            {item.kind === "provider"
              ? <button type="button" className="btn-sm" onClick={() => openSettings("ai")}>Reconnect</button>
              : <Link className="dash-link-button" to={`/matter/${item.matter.id}?session=${encodeURIComponent(item.session)}`}>Open</Link>}
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
  let longest = 0, run = 0
  for (const count of counts) {
    run = count > 0 ? run + 1 : 0
    longest = Math.max(longest, run)
  }
  // A streak holds through today until the day ends without a conversation.
  let current = 0
  for (let index = counts.length - (counts[counts.length - 1] ? 1 : 2); index >= 0 && counts[index] > 0; index--) current++
  const byWeekday = new Array<number>(7).fill(0)
  counts.forEach((count, index) => { byWeekday[index % 7] += count })
  const busiest = byWeekday.indexOf(Math.max(...byWeekday))
  const names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
  return (
    <div className="dash-activity">
      {total > 0 ? (
        <ul className="dash-chips">
          <li className="chip-forest"><strong>{total}</strong> {total === 1 ? "conversation" : "conversations"} this year</li>
          <li className="chip-indigo"><strong>{current}</strong> day streak now</li>
          <li className="chip-plum"><strong>{longest}</strong> day longest streak</li>
          <li className="chip-sky">Busiest on <strong>{names[busiest]}s</strong></li>
        </ul>
      ) : (
        <ul className="dash-chips"><li className="chip-forest">Each day you start a conversation fills a square</li></ul>
      )}
      <ActivityCalendar
        counts={counts}
        first={first}
        describe={(day, count) => `${count ? plural(count, "conversation") : "No conversations"} on ${formatDate(day)}`}
      />
    </div>
  )
}

function Recent({ data }: WidgetContext) {
  const rows = recentMatters(data.matters)
  if (!rows.length) return <Empty icon="matter" title="No matters yet"><Link to="/">Open your first matter</Link> to start.</Empty>
  const STATUS = { running: "Drafting", blocked: "Needs you", failed: "Failed", completed: "Drafted" } as const
  return (
    <ul className="dash-recent">
      {rows.map((row, index) => (
        <li key={row.matter.id}>
          <ShapeIcon icon="matter" shape={(["cookie", "clover", "flower", "sunny"] as const)[index % 4]} className={`dash-recent-badge hue-${index % 4}`} />
          <Link to={`/matter/${row.matter.id}`} className="dash-recent-name">
            <strong>{row.matter.title}</strong>
            <span>{[row.matter.client, row.matter.reference].filter(Boolean).join(" · ") || (row.matter.sample ? "Sample matter" : "No client recorded")}</span>
          </Link>
          <span className="dash-recent-meta">
            <span>{formatAgo(row.last, Date.now())}</span>
            <span>{formatCost(row.cost)}{row.documents ? ` · ${plural(row.documents, "doc")}` : ""}</span>
          </span>
          {row.status && <span className={`dash-pill status-${row.status}`}>{STATUS[row.status]}</span>}
        </li>
      ))}
    </ul>
  )
}

function Tokens({ summary, period }: WidgetContext) {
  const { input, cached, output, reasoning } = summary.tokens
  if (!input && !cached && !output && !reasoning) return <Empty icon="tokens" title="No tokens used yet" />
  return (
    <StackedBar
      label={`Tokens over ${periodLabel(period)}`}
      segments={[
        { key: "input", label: "Input", value: input, display: compact(input) },
        { key: "cached", label: "Cached input", value: cached, display: compact(cached) },
        { key: "output", label: "Output", value: output, display: compact(output) },
        { key: "reasoning", label: "Reasoning", value: reasoning, display: compact(reasoning) },
      ]}
    />
  )
}

// Daily spend over the window; all time and long windows group by week.
function Trend({ data, period, now }: WidgetContext) {
  const start = windowStart(data.matters, period, now)
  const days = Math.max(1, dayIndex(start, now) + 1)
  const daily = dailySpend(data.matters, days, now)
  const weekly = days > 92
  const points = weekly
    ? Array.from({ length: Math.ceil(daily.length / 7) }, (_, index) => {
        const slice = daily.slice(index * 7, index * 7 + 7)
        return { time: slice[0].day, value: slice.reduce((sum, row) => sum + row.cost, 0), label: `week of ${formatDate(slice[0].day)}` }
      })
    : daily.map((row) => ({ time: row.day, value: row.cost, label: new Date(row.day).toLocaleDateString("en-GB", { day: "numeric", month: "short" }) }))
  if (!points.some((row) => row.value > 0)) return <Empty icon="trend" title="No spend in this period" />
  return <AreaChart points={points} format={axisMoney} detail={money} label={`${weekly ? "Weekly" : "Daily"} spend over ${periodLabel(period)}`} />
}

function Actions() {
  const navigate = useNavigate()
  const notify = useToast()
  const [adding, setAdding] = useState(false)
  const tiles: { icon: IconName; title: string; text: string; hue: number; to?: string; run?: () => void }[] = [
    { icon: "matter", title: "New matter", text: "Open a folder for an engagement", hue: 0, to: "/" },
    { icon: "template", title: "Add a template", text: "Draft from your own precedent", hue: 1, to: "/templates" },
    { icon: "client", title: "Add a client", text: "Keep parties and details on file", hue: 2, to: "/clients" },
    { icon: "sample", title: "Try the sample", text: "Explore a matter with documents", hue: 3, run: () => {
      setAdding(true)
      addSampleMatter().then((matter) => navigate(`/matter/${matter.id}`), (error: unknown) => {
        setAdding(false)
        notify("error", error instanceof Error ? error.message : "Could not add the sample matter")
      })
    } },
  ]
  return (
    <ul className="dash-actions">
      {tiles.map((tile) => {
        const body = <>
          <ShapeIcon icon={tile.icon} shape={(["cookie", "clover", "flower", "sunny"] as const)[tile.hue]} />
          <strong>{tile.title}</strong>
          <span>{tile.text}</span>
        </>
        return (
          <li key={tile.title} className={`hue-${tile.hue}`}>
            {tile.to ? <Link to={tile.to}>{body}</Link> : <button type="button" onClick={tile.run} disabled={adding} aria-busy={adding}>{body}</button>}
          </li>
        )
      })}
    </ul>
  )
}

function Library({ data }: WidgetContext) {
  const tiles: { icon: IconName; label: string; count?: number; to: string }[] = [
    { icon: "matter", label: "Matters", count: data.error ? undefined : data.matters.length, to: "/" },
    { icon: "template", label: "Templates", count: data.templates, to: "/templates" },
    { icon: "playbook", label: "Playbooks", count: data.playbooks, to: "/playbooks" },
    { icon: "client", label: "Clients", count: data.clients, to: "/clients" },
  ]
  return (
    <ul className="dash-library">
      {tiles.map((tile, index) => (
        <li key={tile.label} className={`hue-${index}`}>
          <Link to={tile.to}>
            <span className="dash-library-icon">{icons[tile.icon]}</span>
            <strong>{tile.count ?? "—"}</strong>
            <span>{tile.label}</span>
          </Link>
        </li>
      ))}
    </ul>
  )
}

type SetupItem = { id: string; icon: IconName; title: string; detail: string; done: boolean; optional?: boolean; tab?: Tab; to?: string; sample?: boolean }

// What the firm has not set up yet, from saved settings and the library.
// Optional items say so; document legends are the firm's choice and absent by
// design, so they are never listed here.
export function setupItems({ data, setup }: Pick<WidgetContext, "data" | "setup">): SetupItem[] {
  const preferences = data.preferences
  const own = data.matters.filter((row) => !row.matter.sample)
  const items: (SetupItem | false)[] = [
    { id: "models", icon: "model", title: "Connect an AI model", detail: "Drafting, review and research run on it.", done: !!setup?.ready, tab: "ai" },
    !!preferences && { id: "name", icon: "person", title: "Add your name", detail: "It signs your comments and tracked changes.", done: !!preferences.attorney.trim(), tab: "drafting" },
    !!preferences && { id: "firm", icon: "firm", title: "Add your firm", detail: "Used in letters, drafts and document headers.", done: !!preferences.firm.trim(), tab: "drafting" },
    !!preferences && { id: "practice", icon: "practice", title: "Choose practice areas", detail: "Steers drafting toward your kind of work.", done: preferences.practiceAreas.length > 0, tab: "drafting" },
    { id: "jurisdiction", icon: "globe", title: "Set a default jurisdiction", detail: "Prefilled on every new matter.", done: loadPrefs().defaultJurisdictions.length > 0, tab: "matters" },
    !data.error && { id: "matter", icon: "matter", title: "Open your first matter", detail: "A home for one engagement's documents and chats.", done: own.length > 0, to: "/" },
    !!preferences && { id: "research", icon: "search", title: "Connect legal research", detail: "Web and case-law search for cited answers.", done: !!(preferences.searchApiKey || preferences.caseLawApiKey || preferences.courtlistenerMcpEnabled), optional: true, tab: "research" },
    // Listed only where this installation can set it up or it is managed elsewhere.
    !!data.word && (data.word.canSetup || ["ready", "external"].includes(data.word.status)) && !["unsupported", "word_missing"].includes(data.word.status) && { id: "word", icon: "word", title: "Set up the Word add-in", detail: "Review and redline without leaving Word.", done: data.word.status === "ready" || (data.word.status === "external" && data.word.available), optional: true, tab: "addins" },
    !!preferences && { id: "style", icon: "style", title: "Add house style notes", detail: "Standing instructions for every draft and edit.", done: !!preferences.houseStyle.trim(), optional: true, tab: "drafting" },
    data.templates !== undefined && { id: "template", icon: "template", title: "Add a template", detail: "Draft from the firm's own precedents.", done: data.templates > 0, optional: true, to: "/templates" },
    !data.error && { id: "sample", icon: "sample", title: "Try the sample matter", detail: "A fictional engagement to explore safely.", done: data.matters.some((row) => row.matter.sample), optional: true, sample: true },
  ]
  return items.filter((item): item is SetupItem => !!item)
}

function SetupChecklist(context: WidgetContext) {
  const openSettings = useOpenSettings()
  const navigate = useNavigate()
  const notify = useToast()
  const [adding, setAdding] = useState(false)
  const items = setupItems(context)
  const done = items.filter((item) => item.done)
  const open = items.filter((item) => !item.done).sort((a, b) => Number(!!a.optional) - Number(!!b.optional))
  const essentials = open.filter((item) => !item.optional).length
  function act(item: SetupItem) {
    if (item.tab) return openSettings(item.tab)
    if (item.to) return navigate(item.to)
    setAdding(true)
    addSampleMatter().then((matter) => navigate(`/matter/${matter.id}`), (error: unknown) => {
      setAdding(false)
      notify("error", error instanceof Error ? error.message : "Could not add the sample matter")
    })
  }
  if (!open.length)
    return (
      <div className="dash-setup is-done">
        <ShapeIcon icon="party" shape="sunny" className="dash-setup-party" />
        <div>
          <p className="dash-setup-title">You're all set</p>
          <p className="dash-setup-text">Everything is in place. You can hide this card from its menu.</p>
        </div>
      </div>
    )
  return (
    <div className="dash-setup">
      <div className="dash-setup-progress">
        <Ring done={done.length} total={items.length} />
        <div>
          <p className="dash-setup-title">{essentials ? `${plural(essentials, "essential")} left` : "Essentials done"}</p>
          <p className="dash-setup-text">{essentials ? "Each takes a minute. Optional extras follow." : "A few optional extras remain."}</p>
        </div>
      </div>
      <ul className="dash-setup-items">
        {open.slice(0, 6).map((item, index) => (
          <li key={item.id} className={`hue-${index % 4}`}>
            <ShapeIcon icon={item.icon} shape={(["cookie", "clover", "flower", "sunny"] as const)[index % 4]} />
            <span className="dash-setup-item-text">
              <strong>{item.title}{item.optional && <span className="dash-optional">Optional</span>}</strong>
              <span>{item.detail}</span>
            </span>
            <button type="button" className={`btn-sm${index === 0 ? " primary" : ""}`} disabled={item.sample && adding} aria-busy={item.sample && adding} onClick={() => act(item)}>
              {item.sample ? "Add" : item.tab ? "Set up" : "Open"}
            </button>
          </li>
        ))}
      </ul>
      {done.length > 0 && (
        <ul className="dash-setup-done" aria-label="Done">
          {done.map((item) => <li key={item.id}><span aria-hidden="true">{icons.check}</span>{item.title}</li>)}
        </ul>
      )}
    </div>
  )
}

export const RENDER: Record<WidgetId, (context: WidgetContext) => ReactNode> = {
  setup: (context) => <SetupChecklist {...context} />,
  spend: (context) => <SpendTile {...context} />,
  "matter-cost": (context) => <MatterCostTile {...context} />,
  "draft-cost": (context) => <DraftCostTile {...context} />,
  documents: (context) => <DocumentsTile {...context} />,
  week: (context) => <WeekWidget {...context} />,
  "matter-spend": (context) => <MatterSpend {...context} />,
  attention: (context) => <NeedsYou {...context} />,
  activity: (context) => <Activity {...context} />,
  recent: (context) => <Recent {...context} />,
  tokens: (context) => <Tokens {...context} />,
  trend: (context) => <Trend {...context} />,
  actions: () => <Actions />,
  library: (context) => <Library {...context} />,
}

// Captions under a widget's title, where the figure needs its basis stated.
// Every card that shares a row carries a subtitle, so titles line up.
export function caption(id: WidgetId, period: Period) {
  const window = `${periodLabel(period)[0].toUpperCase()}${periodLabel(period).slice(1)}`
  switch (id) {
    case "spend": case "matter-cost": case "matter-spend": case "tokens": case "trend":
      return window
    case "draft-cost": case "documents":
      return `${window}, latest run per matter`
    case "week":
      return "Each conversation counted on the day it started"
    case "attention":
      return "Drafting runs waiting on you"
    case "activity":
      return "Conversations started each day"
    case "recent":
      return "Latest first, with all-time spend"
    case "actions":
      return "Start something new"
    case "library":
      return "Templates, playbooks, clients and matters"
  }
}

export { periodStart }
