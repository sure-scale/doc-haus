import { describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import type { Part } from "../src/api/session-view"
import StepsPanel from "../src/components/chat/StepsPanel"
import { toTurns } from "../src/components/chat/turn-model"
import Markdown, { DocMarkdown } from "../src/components/markdown/Markdown"
import MarkdownContent from "../src/components/markdown/MarkdownContent"
import { parseStructuredReportMarkdown } from "../src/components/markdown/structured-report"
import { parseStructuredToolOutput } from "../src/components/markdown/tool-output"

const renderContent = (text: string) => renderToStaticMarkup(createElement(MarkdownContent, { children: text }))
const renderAnswer = (text: string) => renderToStaticMarkup(createElement(Markdown, { children: text }))
const renderDoc = (text: string) => renderToStaticMarkup(createElement(DocMarkdown, { children: text }))

function renderToolStep(tool: string, output: string, extra?: { metadata?: Record<string, unknown>; pool?: Part[] }) {
  const part: Part = {
    type: "tool",
    id: "call",
    sessionID: "session",
    messageID: "message",
    tool,
    state: { status: "completed", input: {}, title: tool, output, ...(extra?.metadata ? { metadata: extra.metadata } : {}) },
  }
  const turns = toTurns([{ info: { id: "message", role: "assistant" }, parts: [part] }], "", extra?.pool)
  return renderToStaticMarkup(createElement(StepsPanel, { steps: turns[0]!.steps, busy: false, answered: true }))
}

describe("parseStructuredReportMarkdown", () => {
  test("turns known result tags into markdown sections and removes closing tags", () => {
    const report = parseStructuredReportMarkdown(
      `Now I have a solid picture of the matter. Stage 1 complete.<result><files>
- **consent-to-adoption.docx** — Rendered consent
</files><findings> **Discrepancies flagged:** 1. **Father contact-right conflict** — Later consent controls.</findings><notes>
The memo carries unresolved research.
</notes> </result>`,
    )

    expect(report.structured).toBe(true)
    expect(report.preamble).toBe("Now I have a solid picture of the matter. Stage 1 complete.")
    expect(report.sections.map((section) => section.tag)).toEqual(["files", "findings", "notes"])
    expect(report.sections[1].markdown).toBe(
      "**Discrepancies flagged:**\n\n1. **Father contact-right conflict** — Later consent controls.",
    )
    expect(report.sections.map((section) => section.markdown).join("\n")).not.toContain("</notes>")
    expect(report.sections.map((section) => section.markdown).join("\n")).not.toContain("</result>")
  })

  test("collapses operational stage packet blocks by default", () => {
    const report = parseStructuredReportMarkdown(
      `<stage_packet><inputs>consent-to-adoption.md</inputs><outputs>consent-to-adoption.docx</outputs></stage_packet>`,
    )

    expect(report.sections).toEqual([
      {
        tag: "stage_packet",
        title: "Stage Packet",
        markdown: "**Inputs**\n\nconsent-to-adoption.md\n\n**Outputs**\n\nconsent-to-adoption.docx",
        collapsed: true,
      },
    ])
  })

  test("unwraps raw and escaped task result envelopes before formatting report tags", () => {
    const raw = parseStructuredReportMarkdown(`<task id="ses_0ef8de502ffeYvclQONHB044fk" state="completed">
<task_result>
<result>
Wrote \`drafting-strategy.md\`. Classification: **Consent to Adoption** under **Maryland** law.
</result>
<files>
- drafting-strategy.md (created)
</files>
<stage_packet>
  <inputs>source-index.json</inputs>
  <outputs>drafting-strategy.md</outputs>
</stage_packet>
</task_result>
</task>`, true)

    expect(raw.structured).toBe(true)
    expect(raw.preamble).toBe("")
    expect(raw.sections.map((section) => section.tag)).toEqual(["result", "files", "stage_packet"])
    expect(raw.sections[2]).toMatchObject({ tag: "stage_packet", collapsed: true })

    const escaped = parseStructuredReportMarkdown(`&lt;task id="ses_0ef80d180ffepEsQk1sy5R21Wd" state="completed"&gt;
&lt;task_result&gt;
&lt;result&gt;
Source reconciliation complete. \`register.md\` written with 13 material rows.
&lt;/result&gt;
&lt;findings&gt;
- **Confirmed-consistent identities**: Marcus &amp; Diana Holloway.
&lt;/findings&gt;
&lt;/task_result&gt;
&lt;/task&gt;`, true)

    expect(escaped.structured).toBe(true)
    expect(escaped.sections.map((section) => section.tag)).toEqual(["result", "findings"])
    expect(escaped.sections[1].markdown).toBe("- **Confirmed-consistent identities**: Marcus & Diana Holloway.")
  })
})

describe("parseStructuredToolOutput", () => {
  test("formats matter listings, escaped tags, numbered/clipped files, and JSON reads without leaking XML markers", () => {
    const listing = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter</path>
<type>directory</type>
<entries>
.dochaus/
client-intake-memo.docx
source-chunks/

(16 entries)
</entries>`)
    expect(listing.structured).toBe(true)
    expect(listing.fields.map((f) => [f.tag, f.text])).toEqual([
      ["path", "/synthetic/workspace/sample-matter"],
      ["type", "directory"],
      ["entries", ".dochaus/\nclient-intake-memo.docx\nsource-chunks/\n\n(16 entries)"],
    ])
    expect(listing.fields.map((field) => field.text).join("\n")).not.toContain("<entries>")

    expect(parseStructuredToolOutput(`&lt;path&gt;/tmp/matter.json&lt;/path&gt;\n&lt;type&gt;file&lt;/type&gt;`)).toEqual({
      structured: true,
      fields: [
        { tag: "path", title: "Path", text: "/tmp/matter.json" },
        { tag: "type", title: "Type", text: "file" },
      ],
    })

    const clipped = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter/source-map.md</path>
<type>file</type>
<content>
73: | \`client-intake-memo-c005\` | \`source-chunks/client-intake-memo-c005.txt\` | 45-45 | 49 |
74: | \`client-intake-memo-c006\` | \`source-chunks/client-intake-memo-c006.txt\` | 47-61 | 1446 |

[tool-output-preview] read output clipped at 6000 characters for context budget.`)
    expect(clipped.fields.map((f) => [f.tag, f.text])).toEqual([
      ["path", "/synthetic/workspace/sample-matter/source-map.md"],
      ["type", "file"],
      ["lines", "73-74"],
      ["content", "| `client-intake-memo-c005` | `source-chunks/client-intake-memo-c005.txt` | 45-45 | 49 |\n| `client-intake-memo-c006` | `source-chunks/client-intake-memo-c006.txt` | 47-61 | 1446 |"],
      ["tool-output-preview", "read output clipped at 6000 characters for context budget."],
    ])

    const jsonRead = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter/matter.json</path>
<type>file</type>
<content>
1: {
2:   "id": "draft-consent-to-adoption-mnp123-060d50",
3:   "title": "draft-consent-to-adoption-mnp123"
4: }

(End of file - total 4 lines)
</content>`)
    expect(jsonRead.fields.find((f) => f.tag === "lines")?.text).toBe("1-4")
    expect(JSON.parse(jsonRead.fields.find((f) => f.tag === "content")!.text)).toEqual({
      id: "draft-consent-to-adoption-mnp123-060d50",
      title: "draft-consent-to-adoption-mnp123",
    })
    expect(jsonRead.fields.find((f) => f.tag === "file-note")?.text).toBe("End of file - total 4 lines")

    expect(parseStructuredToolOutput("Wrote drafting-workflow-state.md to this matter directory.")).toEqual({
      structured: false,
      fields: [],
    })
  })
})

describe("DocMarkdown and MarkdownContent literal escaping and GFM boundaries", () => {
  const table = "| Item | Term |\n| :--- | ---: |\n| Notice | 30 days |"

  test("escapes instruction tags in DocMarkdown while preserving fenced code and surrounding Markdown", () => {
    const example = "<task>\n\nKeep the `tick` and [source: later].\n\n</task>"
    const fenced = renderDoc("````text\n" + example + "\n````")
    expect(fenced).toContain("<pre><code")
    expect(fenced).toContain("&lt;task&gt;\n\nKeep the `tick` and [source: later].\n\n&lt;/task&gt;\n</code></pre>")

    const standalone = renderDoc("<task>\n\n**Keep qualifications.**\n\n</task>")
    expect(standalone).toContain("&lt;task&gt;")
    expect(standalone).toContain("&lt;/task&gt;")
    expect(standalone).toContain("<strong>Keep qualifications.</strong>")
  })

  test("renders tables with alignment and prose order, leaving fenced, indented, and escaped tables literal", () => {
    const html = renderContent(`Before.\n\n${table}\n\nAfter.`)
    expect(html).toContain('<th style="text-align:left">Item</th>')
    expect(html).toContain('<td style="text-align:right">30 days</td>')
    expect(html.indexOf("Before.")).toBeLessThan(html.indexOf("<table"))
    expect(html.indexOf("</table>")).toBeLessThan(html.indexOf("After."))

    for (const fence of ["```text", "~~~~text"]) {
      const close = fence.startsWith("`") ? "```" : "~~~~"
      const fenced = renderContent(`${fence}\n${table}\n${close}`)
      expect(fenced).toContain("<pre><code")
      expect(fenced).toContain("| Notice | 30 days |")
      expect(fenced).not.toContain("<table")
    }

    const indented = renderContent(table.split("\n").map((line) => `    ${line}`).join("\n") + "\n\n\\| Literal \\| text \\|")
    expect(indented).toContain("<pre><code>")
    expect(indented).toContain("| Literal | text |")
    expect(indented).not.toContain("<table")
  })

  test("preserves escaped pipes, inline formatting, references, container nesting, and disables unrelated GFM syntax", () => {
    const html = renderContent("| Literal | Content |\n|---|---|\n| A\\|B | **échéance** [party] [source: missing] [1, 2] |\n| Code | `x\\|y` |")
    expect(html.match(/<tr>/g)).toHaveLength(3)
    expect(html).toContain("A|B")
    expect(html).toContain("<strong>échéance</strong> [party] [source: missing] [1, 2]")
    expect(html).toContain("<code>x|y</code>")
    expect(html).not.toContain("<button")

    const refHtml = renderContent('| Item | Source |\n|---|---|\n| Notice | [clause][evidence] |\n\nOutside [also][evidence].\n\n[evidence]: /documents/terms "Terms"')
    expect(refHtml.match(/href="\/documents\/terms" title="Terms"/g)).toHaveLength(2)

    const quoted = renderContent(table.split("\n").map((line) => `> ${line}`).join("\n"))
    expect(quoted.indexOf("<blockquote>")).toBeLessThan(quoted.indexOf("<table"))
    const listed = renderContent("- A term:\n\n" + table.split("\n").map((line) => `  ${line}`).join("\n"))
    expect(listed.indexOf("<li>")).toBeLessThan(listed.indexOf("<table"))
    expect(listed.indexOf("</table>")).toBeLessThan(listed.indexOf("</li>"))

    const plain = renderContent("~~literal~~ [ ] task www.example.invalid person@example.invalid")
    expect(plain).toContain("~~literal~~ [ ] task www.example.invalid person@example.invalid")
    expect(plain).not.toContain("<del>")
    expect(plain).not.toContain("<a ")
    expect(plain).not.toContain("checkbox")

    const fencedThenTable = `\`\`\`text\n${table}\n\`\`\`\n\n${table}`
    expect(renderContent(fencedThenTable.slice(0, fencedThenTable.indexOf("\n```", 3)))).not.toContain("<table")
    const completed = renderContent(fencedThenTable)
    expect(completed.match(/<table\b/g)).toHaveLength(1)
    expect(renderContent(JSON.parse(JSON.stringify({ text: fencedThenTable })).text)).toBe(completed)
  })
})

describe("lossless report and tool-output rendering", () => {
  test("preserves qualifications, warnings, and nested result text in order without executing literal markup", () => {
    const html = renderAnswer(
      "Before\n<findings>Thirty days.</findings>\nCheck the missing exhibit.\n<notes>A qualification.</notes>\nDo not rely on this yet.",
    )
    expect(html.indexOf("Thirty days.")).toBeLessThan(html.indexOf("Check the missing exhibit."))
    expect(html.indexOf("Check the missing exhibit.")).toBeLessThan(html.indexOf("A qualification."))
    expect(html.indexOf("A qualification.")).toBeLessThan(html.indexOf("Do not rely on this yet."))

    const warning = "Do not rely on this report until the source access gate is resolved. The original exhibit is missing."
    expect(renderAnswer(`${warning}\n<findings>Thirty days.</findings>`)).toContain(warning)

    const nested = renderAnswer("<result>First.<findings>Second.</findings>Third.</result>")
    expect(nested.indexOf("First.")).toBeLessThan(nested.indexOf("Second."))
    expect(nested.indexOf("Second.")).toBeLessThan(nested.indexOf("Third."))

    for (const literal of [
      "The source quotes `<task_result>Thirty days.</task_result>`. The actual period is sixty days.",
      "The source quotes `<findings>Thirty days.</findings>`. The actual period is sixty days.",
      "The source quotes ``<findings>`Thirty days.`</findings>``. The actual period is sixty days.",
      "The source quotes &lt;findings&gt;Thirty days.&lt;/findings&gt;. The actual period is sixty days.",
      "The source quotes \\<findings>Thirty days.\\</findings>. The actual period is sixty days.",
      "```xml\n<findings>Thirty days.</findings>\n```\nThe actual period is sixty days.",
      "~~~xml\n<findings>Thirty days.</findings>\n~~~\nThe actual period is sixty days.",
    ]) {
      const rendered = renderAnswer(literal)
      expect(rendered).toContain("The actual period is sixty days.")
      expect(rendered).toContain("&lt;")
      expect(rendered).not.toContain("Findings")
    }

    const codeInsideReport = renderAnswer("<findings>Use this example:\n\n```xml\n<notes>Literal only.</notes>\n```\n\nActual finding.</findings>")
    expect(codeInsideReport).toContain("Findings")
    expect(codeInsideReport).toContain("&lt;notes&gt;Literal only.&lt;/notes&gt;")
    expect(codeInsideReport).toContain("Actual finding.")
    expect(codeInsideReport).not.toContain(">Notes<")
  })

  test("tool steps keep task qualifications, child timelines, warnings, and opaque file/search content lossless", () => {
    const rawTask = '<task id="s" state="completed"><task_result><findings>Thirty days.</findings></task_result></task>'
    for (const text of [rawTask, rawTask.replaceAll("<", "&lt;").replaceAll(">", "&gt;")]) {
      const html = renderToolStep("task", text)
      expect(html).toContain("Findings")
      expect(html).toContain("Thirty days.")
    }

    const childOutput =
      '<task id="child" state="completed"><task_result><findings>Thirty days.</findings>Do not rely on this until the missing exhibit is checked.</task_result></task>'
    const childRead: Part = {
      type: "tool",
      id: "child-read",
      sessionID: "child",
      messageID: "child-message",
      tool: "read",
      state: {
        status: "completed",
        input: {},
        title: "Read exhibit",
        output: "<path>/synthetic/exhibit.txt</path><content>1: Child source evidence.</content>",
      },
    }
    const withChild = renderToolStep("task", childOutput, { metadata: { sessionId: "child" }, pool: [childRead] })
    expect(withChild).toContain("Child source evidence.")
    expect(withChild).toContain("Thirty days.")
    expect(withChild).toContain("Do not rely on this until the missing exhibit is checked.")

    const wrappedTask = renderToolStep("task", `UNVERIFIED.\n${rawTask}\nCheck the exhibit.`)
    expect(wrappedTask.indexOf("UNVERIFIED.")).toBeLessThan(wrappedTask.indexOf("Thirty days."))
    expect(wrappedTask.indexOf("Thirty days.")).toBeLessThan(wrappedTask.indexOf("Check the exhibit."))

    const rawRead = "<path>/synthetic/contract.txt</path><content>1: Thirty days.</content>"
    const warnedRead = renderToolStep("read", `UNVERIFIED EXTRACTION.\n${rawRead}\nCheck the exhibit.`)
    expect(warnedRead).toContain("UNVERIFIED EXTRACTION.")
    expect(warnedRead).toContain("Check the exhibit.")
    expect(warnedRead).toContain("Thirty days.")

    for (const literal of ["<findings>Thirty days.</findings>", "&lt;findings&gt;Thirty days.&lt;/findings&gt;"]) {
      const readHtml = renderToolStep("read", `<path>/synthetic/contract.txt</path><content>1: ${literal}\n2: The actual period is sixty days.</content>`)
      expect(readHtml).toContain("The actual period is sixty days.")
      expect(readHtml).not.toContain(">Findings<")
    }

    const literalClose = renderToolStep(
      "read",
      "<path>/synthetic/contract.txt</path><content>1: Example </content> remains literal.\n2: The actual period is sixty days.</content>",
    )
    expect(literalClose).toContain("Example &lt;/content&gt; remains literal.")
    expect(literalClose).toContain("The actual period is sixty days.")

    const searchHtml = renderToolStep("search-document", "Literal <findings>Thirty days.</findings>. The actual period is sixty days.")
    expect(searchHtml).toContain("The actual period is sixty days.")
    expect(searchHtml).toContain("&lt;findings&gt;")
    expect(searchHtml).not.toContain(">Findings<")
  })
})
