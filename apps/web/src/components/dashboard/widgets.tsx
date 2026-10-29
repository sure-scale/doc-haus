import { useRef, useState, type FormEvent, type ReactNode } from "react"
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
  bucketed,
  costByMatter,
  dailyConversations,
  dailySpend,
  dayIndex,
  draftedRuns,
  periodStart,
  previousSpend,
  progress,
  recentMatters,
  sessionCosts,
  startOfDay,
  windowStart,
  type DashboardData,
  type Period,
  type Summary,
} from "./data"
import { icons, ShapeIcon, SHAPES, type IconName, type ShapeName } from "./icons"
import { PRACTICE_WIDGETS, STARTING_WIDGETS, type Stage, type WidgetId } from "./layout"

// `onPreferences` takes settings a widget saved, so the page shows them at once.
export type WidgetContext = { data: DashboardData; setup: SetupProgress | null; summary: Summary; period: Period; now: number; onPreferences?: (preferences: DraftingPreferences) => void }

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
  journey: { icon: "flag", shape: "sunny", tone: "indigo" },
  sample: { icon: "sample", shape: "flower", tone: "plum" },
  ideas: { icon: "idea", shape: "clover", tone: "plain" },
}

const SHAPE_ORDER = ["cookie", "clover", "flower", "sunny"] as const


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
          <ShapeIcon icon="matter" shape={SHAPE_ORDER[index % 4]} className={`dash-recent-badge hue-${index % 4}`} />
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

// Adds the sample matter (an existing one is returned as is) and opens it at
// `path`, such as a chat with a question waiting.
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
          <li key={tile.title} className={`hue-${index % 4}`}>
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
      { icon: "matter", title: "New matter", text: "Open a folder for an engagement", to: "/" },
      { icon: "template", title: "Add a template", text: "Draft from your own precedent", to: "/templates" },
      { icon: "client", title: "Add a client", text: "Keep parties and details on file", to: "/clients" },
      { icon: "sample", title: "Try the sample", text: "Explore a matter with documents", run: () => sample.open() },
    ]} />
  )
}

// Ways in for someone new. Asking and reviewing happen in a matter: the one
// last worked in, else the sample.
function Ideas({ data }: WidgetContext) {
  const sample = useSampleMatter()
  const own = progress(data.matters).target
  const matter = own ?? data.matters.find((row) => row.matter.sample)?.matter
  const inMatter = (path: string) => (matter ? { to: `/matter/${matter.id}${path}` } : { run: () => sample.open(path) })
  return (
    <Tiles busy={sample.adding} tiles={[
      { icon: "documents", title: "Ask about documents", text: own ? `Key points in ${own.title}` : "Key points, deadlines and gaps", ...inMatter(ASK) },
      { icon: "template", title: "Draft from a template", text: "Start from the firm's own precedent", to: "/templates" },
      { icon: "review", title: "Review across documents", text: "Ask the same questions of every file", ...inMatter("?view=review") },
      { icon: "workflow", title: "Run a workflow", text: "A repeatable routine for a matter", to: "/workflows" },
    ]} />
  )
}

function Library({ data }: WidgetContext) {
  const tiles: { icon: IconName; label: string; count?: number; first: string; to: string }[] = [
    { icon: "matter", label: "Matters", count: data.error ? undefined : data.matters.length, first: "Open your first", to: "/" },
    { icon: "template", label: "Templates", count: data.templates, first: "Add your first", to: "/templates" },
    { icon: "playbook", label: "Playbooks", count: data.playbooks, first: "Add your first", to: "/playbooks" },
    { icon: "client", label: "Clients", count: data.clients, first: "Add your first", to: "/clients" },
  ]
  return (
    <ul className="dash-library">
      {tiles.map((tile, index) => (
        <li key={tile.label} className={`hue-${index}${tile.count === 0 ? " is-zero" : ""}`}>
          <Link to={tile.to}>
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
  const steps = progress(data.matters)
  const at = steps.target && `/matter/${steps.target.id}`
  const list: Step[] = [
    { title: "Open a matter", detail: "One home for an engagement's documents and chats.", done: steps.matters > 0, note: plural(steps.matters, "matter"), action: { label: "New matter", to: "/" } },
    { title: "Add documents", detail: "Upload the agreements, emails or notes you work from.", done: steps.documents > 0, note: "Documents added", action: at ? { label: "Add documents", to: `${at}?view=documents` } : undefined },
    { title: "Ask a question", detail: "Ask about the materials in your own words.", done: steps.conversations > 0, note: plural(steps.conversations, "conversation"), action: at ? { label: "Ask a question", to: `${at}${ASK}` } : undefined },
    { title: "Draft a document", detail: "Draft from your instructions and the matter's files.", done: steps.drafted > 0, note: `${plural(steps.drafted, "document")} drafted`, action: at ? { label: "Start drafting", to: `${at}?view=chat&agent=draft` } : undefined },
  ]
  const current = list.findIndex((step) => !step.done)
  const run = steps.run
  return (
    <div className="dash-journey-body">
      {current < 0 && (
        <p className="dash-journey-done"><span aria-hidden="true">{icons.party}</span>First draft delivered. Your figures fill the dashboard as you work.</p>
      )}
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
              </span>
            </li>
          )
        })}
      </ol>
    </div>
  )
}

// The fictional sample, with a question already seated when it opens.
function SampleMatter({ data }: WidgetContext) {
  const sample = useSampleMatter()
  const added = data.matters.find((row) => row.matter.sample)?.matter
  return (
    <div className="dash-sample">
      <div className="dash-sample-art" aria-hidden="true">
        <span className="dash-sample-page">
          <i /><i /><i className="is-marked" /><i /><i />
        </span>
        <span className="dash-sample-comment">{icons.question}</span>
      </div>
      <p className="dash-sample-title">{added ? "Pick up where you left off" : "Explore before your own"}</p>
      <p className="dash-sample-text">A fictional engagement letter, with cited answers and a legal review ready to look through.</p>
      <ul className="dash-sample-tags">
        <li>Engagement letter</li>
        <li>Cited answers</li>
        <li>Legal review</li>
      </ul>
      {added
        ? <Link className="dash-sample-action" to={`/matter/${added.id}${ASK}`}>Open the sample matter{icons.arrowRight}</Link>
        : <button type="button" className="dash-sample-action" onClick={() => sample.open(ASK)} disabled={sample.adding} aria-busy={sample.adding}>Add the sample matter{icons.arrowRight}</button>}
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
    !!preferences && { id: "name", icon: "person", title: "Your name", done: !!preferences.attorney.trim(), tab: "drafting" },
    !!preferences && { id: "firm", icon: "firm", title: "Your firm", done: !!preferences.firm.trim(), tab: "drafting" },
    { id: "jurisdiction", icon: "globe", title: "Jurisdictions", done: jurisdictions.length > 0, tab: "matters" },
    !!preferences && { id: "practice", icon: "practice", title: "Practice areas", done: preferences.practiceAreas.length > 0, tab: "matters" },
    !!preferences && { id: "research", icon: "search", title: "Web research", done: !!preferences.searchApiKey.trim(), tab: "research" },
    // Listed only where this installation can set it up or it is managed elsewhere.
    !!data.word && (data.word.canSetup || ["ready", "external"].includes(data.word.status)) && !["unsupported", "word_missing"].includes(data.word.status) && { id: "word", icon: "word", title: "Word add-in", detail: "Review and redline without leaving Word.", done: data.word.status === "ready" || (data.word.status === "external" && data.word.available), optional: true, tab: "addins" },
    !!preferences && { id: "style", icon: "style", title: "House style notes", detail: "Standing instructions for every draft and edit.", done: !!preferences.houseStyle.trim(), optional: true, tab: "drafting" },
    data.templates !== undefined && { id: "template", icon: "template", title: "A template", detail: "Draft from the firm's own precedents.", done: data.templates > 0, optional: true, to: "/templates" },
  ]
  return items.filter((item): item is SetupItem => !!item)
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

function ProfileTile({ id, icon, shape, hue, title, done, children }: { id: string; icon: IconName; shape: ShapeName; hue: number; title: string; done: boolean; children: ReactNode }) {
  return (
    <section className={`dash-profile-tile hue-${hue}${done ? " is-done" : ""}`} aria-labelledby={`dash-profile-${id}`}>
      <div className="dash-profile-tile-head">
        <ShapeIcon icon={icon} shape={shape} />
        <h4 id={`dash-profile-${id}`}>{title}</h4>
        {done && <span className="dash-profile-check">{icons.check}<span className="sr-only">Done</span></span>}
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
  const changed = name.trim() !== preferences.attorney.trim() || firm.trim() !== preferences.firm.trim()
  async function submit(event: FormEvent) {
    event.preventDefault()
    setSaving(true)
    const saved = await save({ attorney: name.trim(), firm: firm.trim() }, "Could not save your name and firm. Try again.")
    setSaving(false)
    if (!saved || !name.trim() || !firm.trim()) return
    setEditing(false)
    // The form gives way to the letterhead; keep focus in the tile.
    requestAnimationFrame(() => editButton.current?.focus())
  }
  return (
    <ProfileTile id="signed" icon="person" shape="cookie" hue={0} title="Name and firm" done={complete}>
      {editing ? (
        <form className="dash-profile-form" onSubmit={submit}>
          <p className="dash-profile-detail">Signs your comments and tracked changes, and names who prepared each draft.</p>
          <label><span>Your name</span><input value={name} onChange={(event) => setName(event.target.value)} placeholder="Jane Doe" autoComplete="name" disabled={saving} /></label>
          <label><span>Firm</span><input value={firm} onChange={(event) => setFirm(event.target.value)} placeholder="Doe & Partners LLP" autoComplete="organization" disabled={saving} /></label>
          <button type="submit" className="btn-sm primary" disabled={saving || !changed} aria-busy={saving}>{saving ? "Saving…" : "Save"}</button>
        </form>
      ) : (
        <>
          <p className="dash-profile-letterhead">
            <span>Prepared by</span>
            <strong>{preferences.attorney}</strong>
            <span>{preferences.firm}</span>
          </p>
          <button ref={editButton} type="button" className="dash-profile-link" onClick={() => setEditing(true)}>Edit</button>
        </>
      )}
    </ProfileTile>
  )
}

function Jurisdictions({ all, codes, setCodes }: { all: Jurisdiction[]; codes: string[]; setCodes: (codes: string[]) => void }) {
  const notify = useToast()
  // Offered to tap until the card is next opened; after that the choice shows.
  const [offer] = useState(() => codes.length === 0)
  const shown = offer ? unique([...nearby(all), ...codes]) : codes
  const name = (code: string) => all.find((pack) => pack.code === code)?.name ?? code
  const others = unique(all.map((pack) => countryOf(pack.code))).filter((country) => !shown.some((code) => countryOf(code) === country)).slice(0, 4)
  function change(next: string[]) {
    try {
      savePrefs({ defaultJurisdictions: next }, { requireStorage: true })
      setCodes(next)
    } catch {
      notify("error", "Could not save your jurisdictions in this browser. Try again.")
    }
  }
  return (
    <ProfileTile id="jurisdictions" icon="globe" shape="clover" hue={1} title="Jurisdictions" done={codes.length > 0}>
      <p className="dash-profile-detail">{codes.length ? "Preselected on every new matter." : "Where do you practice? Your picks are preselected on every new matter."}</p>
      {shown.length > 0 && (
        <ul className="dash-profile-flags">
          {shown.map((code) => (
            <li key={code}>
              {offer ? (
                <button type="button" aria-pressed={codes.includes(code)} onClick={() => change(toggled(codes, code))}>
                  <JurisdictionFlag code={code} region />
                  <span>{name(code)}</span>
                  <span className="dash-profile-toggle" aria-hidden="true">{codes.includes(code) ? icons.check : icons.plus}</span>
                </button>
              ) : (
                <span className="dash-profile-flag"><JurisdictionFlag code={code} region /><span>{name(code)}</span></span>
              )}
            </li>
          ))}
        </ul>
      )}
      {all.length > 0 && (
        <JurisdictionSelect
          jurisdictions={all}
          selected={codes}
          onChange={change}
          trigger={
            <span className="dash-profile-more">
              {others.length > 0 && <span className="dash-flag-stack">{others.map((country) => <JurisdictionFlag key={country} code={all.find((pack) => countryOf(pack.code) === country)!.code} />)}</span>}
              {offer ? `All ${all.length} jurisdictions` : "Change"}
            </span>
          }
        />
      )}
    </ProfileTile>
  )
}

function PracticeAreas({ preferences, save }: { preferences: DraftingPreferences; save: Save }) {
  const [areas, setAreas] = useState(preferences.practiceAreas)
  const [offer] = useState(() => preferences.practiceAreas.length === 0)
  const shown = offer ? unique([...PRACTICE_PICKS, ...areas]) : areas
  function change(next: string[]) {
    const before = areas
    setAreas(next)
    void save({ practiceAreas: next }, "Could not save your practice areas. Try again.").then((saved) => { if (!saved) setAreas(before) })
  }
  const options = unique([...areas, ...SUGGESTED_PRACTICE_AREAS]).map((area) => ({ code: area, name: area }))
  return (
    <ProfileTile id="practice" icon="practice" shape="flower" hue={2} title="Practice areas" done={areas.length > 0}>
      <p className="dash-profile-detail">{areas.length ? "Steers drafting and research toward your work." : "What kind of work do you do? It steers drafting and research toward it."}</p>
      <ul className="dash-profile-areas">
        {shown.map((area) => (
          <li key={area}>
            {offer
              ? <button type="button" aria-pressed={areas.includes(area)} onClick={() => change(toggled(areas, area))}>{areas.includes(area) && icons.check}{area}</button>
              : <span>{area}</span>}
          </li>
        ))}
      </ul>
      <JurisdictionSelect practiceAreas jurisdictions={options} selected={areas} onChange={change} trigger={<span className="dash-profile-more">{offer ? "More areas, or add your own" : "Change"}</span>} />
    </ProfileTile>
  )
}

function WebResearch({ preferences }: { preferences: DraftingPreferences }) {
  const openSettings = useOpenSettings()
  const connected = !!preferences.searchApiKey.trim()
  return (
    <ProfileTile id="research" icon="search" shape="sunny" hue={3} title="Web research" done={connected}>
      <p className="dash-profile-detail">
        {connected
          ? `Exa connected, searching ${preferences.webResearch === "official" ? "approved sources only" : "the entire web"}.`
          : "Search runs on limited keyless access. A free Exa key adds full results and the source text behind each citation."}
      </p>
      <button type="button" className={connected ? "dash-profile-link" : "btn-sm primary"} onClick={() => openSettings("research")}>
        {connected ? "Settings" : "Add an Exa key"}
      </button>
    </ProfileTile>
  )
}

function PracticeProfile(context: WidgetContext) {
  const { data, onPreferences } = context
  const openSettings = useOpenSettings()
  const navigate = useNavigate()
  const notify = useToast()
  const [codes, setCodes] = useState(() => loadPrefs().defaultJurisdictions)
  const items = setupItems(context, codes)
  const essentials = items.filter((item) => !item.optional)
  const extras = items.filter((item) => item.optional)
  const left = essentials.filter((item) => !item.done).length
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
  const preferences = data.preferences
  return (
    <div className="dash-profile">
      <div className="dash-setup-progress">
        <Ring done={essentials.length - left} total={essentials.length} size={96} />
        <div>
          <p className="dash-setup-title">{left ? `${plural(left, "detail")} to go` : "Your practice is set up"}</p>
          <p className="dash-setup-text">
            {left ? "Fill them in right here. Every draft then starts with your name, your law and your kind of work." : "Every draft starts with your name, your law and your kind of work."}
          </p>
        </div>
      </div>
      <div className="dash-profile-tiles">
        {preferences && <SignedAs preferences={preferences} save={save} />}
        <Jurisdictions all={data.jurisdictions ?? []} codes={codes} setCodes={setCodes} />
        {preferences && <PracticeAreas preferences={preferences} save={save} />}
        {preferences && <WebResearch preferences={preferences} />}
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

export const stageOf = (context: Pick<WidgetContext, "data">): Stage => (progress(context.data.matters).conversations ? "practice" : "starting")

// The dashboard shown until the lawyer arranges their own: see layout.ts.
export function defaultWidgets(context: Pick<WidgetContext, "data">): WidgetId[] {
  const open = setupItems(context).filter((item) => !item.done)
  if (stageOf(context) === "starting") return ["journey", ...(open.length ? ["setup" as const] : []), ...STARTING_WIDGETS]
  return [...(progress(context.data.matters).drafted ? [] : ["journey" as const]), ...(open.some((item) => !item.optional) ? ["setup" as const] : []), ...PRACTICE_WIDGETS]
}

export const RENDER: Record<WidgetId, (context: WidgetContext) => ReactNode> = {
  setup: (context) => <PracticeProfile {...context} />,
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
  journey: (context) => <FirstSteps {...context} />,
  sample: (context) => <SampleMatter {...context} />,
  ideas: (context) => <Ideas {...context} />,
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
    case "journey":
      return "From a new matter to your first draft"
    case "sample":
      return "Fictional, safe to explore"
    case "ideas":
      return "A few ways to start"
  }
}

export { periodStart }
