import { describe, expect, test } from "bun:test"
import { splitMarkdownTableBlocks } from "../src/components/markdown/MarkdownContent"
import { parseStructuredReportMarkdown } from "../src/components/markdown/structured-report"
import { parseStructuredToolOutput } from "../src/components/markdown/tool-output"

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

  test("unwraps task result envelopes before formatting report tags", () => {
    const report = parseStructuredReportMarkdown(`<task id="ses_0ef8de502ffeYvclQONHB044fk" state="completed">
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
</task>`)

    expect(report.structured).toBe(true)
    expect(report.preamble).toBe("")
    expect(report.sections.map((section) => section.tag)).toEqual(["result", "files", "stage_packet"])
    expect(report.sections[0].markdown).toBe(
      "Wrote `drafting-strategy.md`. Classification: **Consent to Adoption** under **Maryland** law.",
    )
    expect(report.sections[2]).toEqual({
      tag: "stage_packet",
      title: "Stage Packet",
      markdown: "**Inputs**\n\nsource-index.json\n\n**Outputs**\n\ndrafting-strategy.md",
      collapsed: true,
    })
  })

  test("unwraps escaped task result envelopes from rendered pre text", () => {
    const report = parseStructuredReportMarkdown(`&lt;task id="ses_0ef80d180ffepEsQk1sy5R21Wd" state="completed"&gt;
&lt;task_result&gt;
&lt;result&gt;
Source reconciliation complete. \`register.md\` written with 13 material rows.
&lt;/result&gt;
&lt;findings&gt;
- **Confirmed-consistent identities**: Marcus &amp; Diana Holloway.
&lt;/findings&gt;
&lt;/task_result&gt;
&lt;/task&gt;`)

    expect(report.structured).toBe(true)
    expect(report.preamble).toBe("")
    expect(report.sections.map((section) => section.tag)).toEqual(["result", "findings"])
    expect(report.sections[1].markdown).toBe("- **Confirmed-consistent identities**: Marcus & Diana Holloway.")
  })
})

describe("splitMarkdownTableBlocks", () => {
  test("extracts complete markdown tables between prose blocks", () => {
    const parts = splitMarkdownTableBlocks(`**Six retry-target register rows now present in the Carry-Forward Register:**

| # | Anchor | Type | Priority | Status |
|---|---|---|---|---|
| 1 | Client request for immediate irrevocability | \`instruction\` | critical | open |
| 2 | Father contact right | \`conflict\` | critical | open |

**Compact binding decisions:** Court caption.`)

    expect(parts).toEqual([
      {
        kind: "markdown",
        markdown: "**Six retry-target register rows now present in the Carry-Forward Register:**",
      },
      {
        kind: "table",
        header: ["#", "Anchor", "Type", "Priority", "Status"],
        rows: [
          ["1", "Client request for immediate irrevocability", "`instruction`", "critical", "open"],
          ["2", "Father contact right", "`conflict`", "critical", "open"],
        ],
      },
      { kind: "markdown", markdown: "**Compact binding decisions:** Court caption." },
    ])
  })
})

describe("parseStructuredToolOutput", () => {
  test("formats matter listing tags without leaking raw XML markers", () => {
    const output = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter</path>
<type>directory</type>
<entries>
.dochaus/
client-intake-memo.docx
source-chunks/

(16 entries)
</entries>`)

    expect(output.structured).toBe(true)
    expect(output.fields).toEqual([
      {
        tag: "path",
        title: "Path",
        text: "/synthetic/workspace/sample-matter",
      },
      { tag: "type", title: "Type", text: "directory" },
      { tag: "entries", title: "Entries", text: ".dochaus/\nclient-intake-memo.docx\nsource-chunks/\n\n(16 entries)" },
    ])
    expect(output.fields.map((field) => field.text).join("\n")).not.toContain("<entries>")
  })

  test("formats escaped file-output tags from rendered pre text", () => {
    const output = parseStructuredToolOutput(`&lt;path&gt;/tmp/matter.json&lt;/path&gt;
&lt;type&gt;file&lt;/type&gt;`)

    expect(output).toEqual({
      structured: true,
      fields: [
        { tag: "path", title: "Path", text: "/tmp/matter.json" },
        { tag: "type", title: "Type", text: "file" },
      ],
    })
  })

  test("formats numbered file content with line range and preview notes", () => {
    const output = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter/matter-identifiers.json</path>
<type>file</type>
<content>
1: {
2:   "schema_version": 1,
128:
</content>

[tool-output-preview] read output clipped at 6000 characters for context budget.`)

    expect(output.structured).toBe(true)
    expect(output.fields).toEqual([
      {
        tag: "path",
        title: "Path",
        text: "/synthetic/workspace/sample-matter/matter-identifiers.json",
      },
      { tag: "type", title: "Type", text: "file" },
      { tag: "lines", title: "Lines", text: "1-128" },
      { tag: "content", title: "Content", text: '{\n  "schema_version": 1,' },
      {
        tag: "tool-output-preview",
        title: "Output Preview",
        text: "read output clipped at 6000 characters for context budget.",
      },
    ])
  })

  test("formats clipped markdown table content even when the content tag is unclosed", () => {
    const output = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter/source-map.md</path>
<type>file</type>
<content>
73: | \`client-intake-memo-c005\` | \`source-chunks/client-intake-memo-c005.txt\` | 45-45 | 49 |
74: | \`client-intake-memo-c006\` | \`source-chunks/client-intake-memo-c006.txt\` | 47-61 | 1446 |

[tool-output-preview] read output clipped at 6000 characters for context budget.`)

    expect(output.structured).toBe(true)
    expect(output.fields).toEqual([
      {
        tag: "path",
        title: "Path",
        text: "/synthetic/workspace/sample-matter/source-map.md",
      },
      { tag: "type", title: "Type", text: "file" },
      { tag: "lines", title: "Lines", text: "73-74" },
      {
        tag: "content",
        title: "Content",
        text: "| `client-intake-memo-c005` | `source-chunks/client-intake-memo-c005.txt` | 45-45 | 49 |\n| `client-intake-memo-c006` | `source-chunks/client-intake-memo-c006.txt` | 47-61 | 1446 |",
      },
      {
        tag: "tool-output-preview",
        title: "Output Preview",
        text: "read output clipped at 6000 characters for context budget.",
      },
    ])
  })

  test("pretty-prints complete numbered JSON file reads", () => {
    const output = parseStructuredToolOutput(`<path>/synthetic/workspace/sample-matter/matter.json</path>
<type>file</type>
<content>
1: {
2:   "id": "draft-consent-to-adoption-mnp123-060d50",
3:   "title": "draft-consent-to-adoption-mnp123",
4:   "jurisdictions": [],
5:   "dir": "/synthetic/workspace/sample-matter",
6:   "created_at": 1782687284129
7: }

(End of file - total 7 lines)
</content>`)

    expect(output.structured).toBe(true)
    expect(output.fields).toEqual([
      {
        tag: "path",
        title: "Path",
        text: "/synthetic/workspace/sample-matter/matter.json",
      },
      { tag: "type", title: "Type", text: "file" },
      { tag: "lines", title: "Lines", text: "1-7" },
      {
        tag: "content",
        title: "Content",
        text: JSON.stringify(
          {
            id: "draft-consent-to-adoption-mnp123-060d50",
            title: "draft-consent-to-adoption-mnp123",
            jurisdictions: [],
            dir: "/synthetic/workspace/sample-matter",
            created_at: 1782687284129,
          },
          null,
          2,
        ),
      },
      { tag: "file-note", title: "File Note", text: "End of file - total 7 lines" },
    ])
  })

  test("leaves ordinary tool text unstructured", () => {
    expect(parseStructuredToolOutput("Wrote drafting-workflow-state.md to this matter directory.")).toEqual({
      structured: false,
      fields: [],
    })
  })
})
