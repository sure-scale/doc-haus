// One row of the reviewer's structured issues register, as document-reviewer.md
// emits it: a fenced ```issues code block holding a JSON array, one element per
// checklist area / missing-protection / playbook clause. The reviewer is a prose
// subagent (no tool emits issues), so the register travels in the answer text;
// parseIssues lifts it out and IssuesView cards it the way CitationView and
// RedlineView card a turn's other structured output.
export type Issue = {
  area: string
  status: "present" | "absent" | "deviating"
  severity: "critical" | "high" | "medium" | "low"
  rule: string
  why: string
  clause: string
  redline: string
}

// Severity drives the row's colored left rail and the count pill, highest first —
// critical/high block a silent sign-off, so they sort to the top.
const SEVERITY = {
  critical: { className: "issue-critical", rank: 0, label: "Critical" },
  high: { className: "issue-high", rank: 1, label: "High" },
  medium: { className: "issue-medium", rank: 2, label: "Medium" },
  low: { className: "issue-low", rank: 3, label: "Low" },
} as const

function severityOf(s: string) {
  return SEVERITY[s as keyof typeof SEVERITY] ?? SEVERITY.low
}

// Lift the reviewer's register out of the answer text: the single fenced
// ```issues block holding a JSON array. The combined legal-review report passes
// it through verbatim, so the same parse works for a lone reviewer turn and for
// the orchestrated report. Returns [] when there is no block or it does not parse
// to an array of issue-shaped rows — a malformed block must never break the
// bubble, only fall back to the prose.
export function parseIssues(text: string): Issue[] {
  const block = text.match(/```issues\s*([\s\S]*?)```/)
  if (!block) return []
  try {
    const parsed = JSON.parse(block[1].trim())
    if (!Array.isArray(parsed)) return []
    return parsed.filter((r): r is Issue => typeof r?.area === "string" && typeof r?.severity === "string")
  } catch {
    return []
  }
}

// Strip the parsed ```issues block from the answer text so the prose findings
// render without the raw JSON showing beneath them — the register renders as the
// card instead.
export function stripIssues(text: string) {
  return text.replace(/```issues\s*[\s\S]*?```/, "").trim()
}

export default function IssuesView({ issues }: { issues: Issue[] }) {
  if (issues.length === 0) return null
  const sorted = [...issues].sort((a, b) => severityOf(a.severity).rank - severityOf(b.severity).rank)
  const blocking = sorted.filter((i) => i.severity === "critical" || i.severity === "high").length
  return (
    <details className="issues" open={blocking > 0}>
      <summary>
        Issues ({issues.length}
        {blocking > 0 ? `, ${blocking} blocking` : ""})
      </summary>
      <ul className="issue-list">
        {sorted.map((issue, i) => {
          const s = severityOf(issue.severity)
          return (
            <li className={`issue-card ${s.className}`} key={i}>
              <div className="issue-head">
                <span className="issue-area">{issue.area}</span>
                <span className="issue-severity">{s.label}</span>
                <span className="issue-status">{issue.status}</span>
              </div>
              {issue.rule && <div className="issue-rule">{issue.rule}</div>}
              {issue.why && <div className="issue-why">{issue.why}</div>}
              {issue.clause && <div className="issue-clause">{issue.clause}</div>}
              {issue.redline && <div className="redline-new issue-redline">{issue.redline}</div>}
            </li>
          )
        })}
      </ul>
    </details>
  )
}
