import { StackedBar } from "../charts"

export type TokenBreakdownData = {
  segments: { key: string; label: string; value: number }[]
  // If supplied, this equals the sum of the segments; the validator checks it.
  total?: number
  period?: string
  notes?: string[]
}

const compact = (value: number) => new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(value)

// Values arrive from a data source. The compact headline and every displayed
// share are computed from the same segment sum; no cache hit-rate is inferred.
export function TokenBreakdown({ segments, period, notes = [] }: TokenBreakdownData) {
  const total = segments.reduce((sum, segment) => sum + segment.value, 0)
  return <div className="dash-token-body dash-token-breakdown">
    <div className="dash-token-summary"><div><strong title={`${new Intl.NumberFormat("en-GB").format(total)} tokens`}>{compact(total)}</strong><span>total tokens processed</span></div>{period && <span className="dash-token-period">{period}</span>}</div>
    <StackedBar label={`Token breakdown${period ? ` · ${period}` : ""}`} segments={segments.map((segment) => ({ ...segment, display: compact(segment.value) }))} />
    {notes.length > 0 && <ul className="dash-chart-notes">{notes.map((note, index) => <li key={`${index}-${note}`}>{note}</li>)}</ul>}
  </div>
}
