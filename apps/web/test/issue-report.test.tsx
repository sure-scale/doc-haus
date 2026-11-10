import { describe, expect, test } from "bun:test"
import { renderToStaticMarkup } from "react-dom/server"
import IssuesView from "../src/components/IssuesView"
import { issueReport } from "../src/components/chat/issue-report"

const row = { area: "Liability [cap]", status: "deviating", severity: "high", rule: "Source §2", why: "Keep the exception \"unless agreed\".", clause: "§2", redline: "" } as const
const register = "```issues\n" + JSON.stringify([row]) + "\n```"

describe("review issues transport", () => {
  test("one admitted register supplies cards and preserves all surrounding bytes", () => {
    const report = issueReport("  Findings\r\n\r\n" + register + "\r\n\r\nQualifications.  ", "legal-review")
    expect(report.text).toBe("  Findings\r\n\r\n\r\n\r\nQualifications.  ")
    expect(report.issues).toEqual([row])
    const html = renderToStaticMarkup(<IssuesView issues={report.issues} />)
    expect(html).toContain("1 blocking")
    expect(html).toContain("Liability [cap]")
    expect(html).toContain("unless agreed")
  })

  test("unrelated agents retain even valid register-shaped source", () => {
    for (const agent of [undefined, "quick-draft", "client-summary", "general"]) {
      expect(issueReport(register, agent)).toEqual({ text: register, issues: [] })
    }
  })

  test("quoted, nested, inline and indented examples remain literal", () => {
    const examples = ["````markdown\n" + register + "\n````", register.split("\n").map(line => "> " + line).join("\n"), register.split("\n").map(line => "    " + line).join("\n"), "Example `" + register + "`", "- Example:\n" + register.split("\n").map(line => "  " + line).join("\n")]
    for (const text of examples) expect(issueReport(text, "document-reviewer")).toEqual({ text, issues: [] })
  })

  test("malformed, incomplete, unsupported and multiple registers preserve the original", () => {
    const invalidRows = [{ ...row, severity: "urgent" }, { ...row, status: "approved" }, { ...row, why: {} }, { ...row, extra: true }, { area: "Only area", severity: "high" }]
    const examples = ["```issues\nnot JSON\n```", "```issues\n" + JSON.stringify([row]), register + "\n\n" + register, register.replace("```issues", "```issues example"), ...invalidRows.map(value => "```issues\n" + JSON.stringify([row, value]) + "\n```")]
    for (const text of examples) expect(issueReport(text, "legal-review")).toEqual({ text, issues: [] })
  })

  test("streaming and persisted projections agree only when the complete block arrives", () => {
    for (const [text, count] of [[register.slice(0, -3), 0], [register, 1], [register + "\nA material caveat.", 1]] as const) {
      const live = issueReport(text, "legal-review")
      const stored = JSON.parse(JSON.stringify({ text, agent: "legal-review" }))
      expect(issueReport(stored.text, stored.agent)).toEqual(live)
      expect(live.issues.length).toBe(count)
    }
    expect(issueReport("~~~issues\n[]\n~~~", "document-reviewer")).toEqual({ text: "", issues: [] })
  })
})
