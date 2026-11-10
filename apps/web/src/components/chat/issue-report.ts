import { unified } from "unified"
import remarkParse from "remark-parse"

export type Issue = {
  area: string
  status: "present" | "absent" | "deviating"
  severity: "critical" | "high" | "medium" | "low"
  rule: string
  why: string
  clause: string
  redline: string
}

const parser = unified().use(remarkParse)
const fields = ["area", "status", "severity", "rule", "why", "clause", "redline"] as const
function issue(value: unknown): value is Issue {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as Record<string, unknown>
  return Object.keys(row).length === fields.length
    && fields.every(key => typeof row[key] === "string")
    && ["present", "absent", "deviating"].includes(row.status as string)
    && ["critical", "high", "medium", "low"].includes(row.severity as string)
}

// Only the review producers promise this transport. Markdown owns literal and
// container context; source offsets remove exactly the one admitted register.
// Partial streaming, malformed rows and ambiguous registers remain visible.
export function issueReport(text: string, agent: string | undefined): { text: string; issues: Issue[] } {
  const unchanged = { text, issues: [] }
  if (agent !== "legal-review" && agent !== "document-reviewer") return unchanged
  const blocks = parser.parse(text).children.filter(node => node.type === "code" && node.lang === "issues")
  if (blocks.length !== 1) return unchanged
  const block = blocks[0]!
  if (block.type !== "code" || block.meta || !block.position) return unchanged
  const start = block.position.start.offset, end = block.position.end.offset
  if (start === undefined || end === undefined) return unchanged
  const source = text.slice(start, end), lines = source.split(/\r?\n/)
  const open = /^(\x60{3,}|~{3,})issues[ \t]*$/.exec(lines[0]!)
  if (!open || lines.length < 2) return unchanged
  const close = /^[ \t]{0,3}(\x60{3,}|~{3,})[ \t]*$/.exec(lines.at(-1)!)
  if (!close || close[1]![0] !== open[1]![0] || close[1]!.length < open[1]!.length) return unchanged
  try {
    const rows: unknown = JSON.parse(block.value)
    if (!Array.isArray(rows) || !rows.every(issue)) return unchanged
    return { text: text.slice(0, start) + text.slice(end), issues: rows }
  } catch {
    return unchanged
  }
}
