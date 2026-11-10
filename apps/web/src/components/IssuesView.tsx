import type { Issue } from "./chat/issue-report"

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
                <span className="issue-status">{issue.status === "absent" ? "Missing" : issue.status === "deviating" ? "Deviates" : issue.status === "present" ? "Present" : issue.status}</span>
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
