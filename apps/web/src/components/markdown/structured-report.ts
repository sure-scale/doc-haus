import { unified } from "unified"
import remarkParse from "remark-parse"

const REPORT_TAGS = {
  result: { title: "Result", kind: "wrapper" },
  files: { title: "Files", kind: "section" },
  findings: { title: "Findings", kind: "section" },
  gaps: { title: "Gaps", kind: "section" },
  unresolved: { title: "Unresolved", kind: "section" },
  notes: { title: "Notes", kind: "section" },
  stage_packet: { title: "Stage Packet", kind: "operational" },
  inputs: { title: "Inputs", kind: "operational" },
  outputs: { title: "Outputs", kind: "operational" },
  open_items: { title: "Open Items", kind: "operational" },
  handoff_status: { title: "Handoff Status", kind: "operational" },
  missing_files: { title: "Missing Files", kind: "operational" },
  output_issues: { title: "Output Issues", kind: "operational" },
  instruction: { title: "Instruction", kind: "operational" },
} as const

type ReportTag = keyof typeof REPORT_TAGS
type ReportNode = { tag: ReportTag; children: ReportChild[] }
type ReportChild = string | ReportNode

export type StructuredReportSection = {
  tag: string
  title: string
  markdown: string
  collapsed: boolean
}

export type StructuredReport = {
  structured: boolean
  preamble: string
  sections: StructuredReportSection[]
  blocks: ({ kind: "markdown"; markdown: string } | { kind: "section"; section: StructuredReportSection })[]
}

// Subagent results arrive wrapped by the task tool as
// <task id="..." state="completed"><task_result>…markdown…</task_result></task>.
// Only a complete task-tool envelope is transport markup. A mention in an
// answer, or extra text around an envelope, must retain the original content.
export function unwrapTaskResult(text: string, taskOutput = false) {
  if (!taskOutput) return text
  const encoded = /^\s*&lt;task\s/i.test(text)
  const source = encoded ? decodeMarkupEntities(text) : text
  const match = source.match(/^\s*<task\s[^>]*>\s*<task_result>([\s\S]*)<\/task_result>\s*<\/task>\s*$/i)
  return match ? match[1].trim() : text
}

export function parseStructuredReportMarkdown(text: string, taskOutput = false): StructuredReport {
  const children = parseReportChildren(unwrapTaskResult(text, taskOutput))
  const structured = children.some((child) => typeof child !== "string")
  if (!structured) return { structured: false, preamble: "", sections: [], blocks: [] }

  const sections: StructuredReportSection[] = []
  const blocks: StructuredReport["blocks"] = []
  const preamble: string[] = []
  for (const child of children) {
    if (typeof child === "string") {
      if (!sections.length && child.trim()) preamble.push(child)
      const markdown = child.trim()
      if (markdown) blocks.push({ kind: "markdown", markdown })
      continue
    }
    appendReportNode(child, sections, blocks)
  }

  return {
    structured: true,
    preamble: normalizeReportMarkdown(preamble.join("")),
    sections,
    blocks,
  }
}

export function normalizeReportMarkdown(text: string) {
  return text
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+(\d+\.\s+\*\*)/g, "\n\n$1")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
}

function parseReportChildren(text: string) {
  const root: { children: ReportChild[] } = { children: [] }
  const stack: ({ children: ReportChild[] } | ReportNode)[] = [root]
  const pattern = /<\/?([a-z][a-z0-9_-]*)\b[^>]*>/gi
  const matches = [...text.matchAll(pattern)]
  if (!matches.some(match => isReportTag(match[1].toLowerCase()))) return [text]
  const code = markdownCodeRanges(text)
  let cursor = 0
  for (const match of matches) {
    if (code.some(([start, end]) => match.index >= start && match.index < end) || escapedAt(text, match.index)) continue
    appendText(stack.at(-1)!, text.slice(cursor, match.index))
    cursor = match.index + match[0].length

    const tag = match[1].toLowerCase()
    if (!isReportTag(tag)) {
      appendText(stack.at(-1)!, match[0])
      continue
    }

    if (match[0].startsWith("</")) {
      const openIndex = findOpenReportTag(stack, tag)
      if (openIndex > 0) stack.length = openIndex
      else appendText(stack.at(-1)!, match[0])
      continue
    }

    const node: ReportNode = { tag, children: [] }
    stack.at(-1)!.children.push(node)
    if (!match[0].endsWith("/>")) stack.push(node)
  }
  appendText(stack.at(-1)!, text.slice(cursor))
  return root.children
}

// Let the same Markdown grammar used by react-markdown identify code. Disable
// HTML here so legacy report tags do not hide the Markdown inside their body.
const markdownParser = unified().use(remarkParse).data("micromarkExtensions", [{ disable: { null: ["htmlFlow", "htmlText"] } }])
function markdownCodeRanges(text: string): [number, number][] {
  const ranges: [number, number][] = []
  const visit = (node: ReturnType<typeof markdownParser.parse>["children"][number] | ReturnType<typeof markdownParser.parse>) => {
    if ((node.type === "code" || node.type === "inlineCode") && node.position?.start.offset !== undefined && node.position.end.offset !== undefined) {
      ranges.push([node.position.start.offset, node.position.end.offset])
    } else if ("children" in node) for (const child of node.children) visit(child)
  }
  visit(markdownParser.parse(text))
  return ranges
}

function escapedAt(text: string, index: number) {
  let slashes = 0
  while (index > 0 && text[--index] === "\\") slashes++
  return slashes % 2 === 1
}

function appendText(parent: { children: ReportChild[] }, text: string) {
  if (!text) return
  const last = parent.children.at(-1)
  if (typeof last === "string") parent.children[parent.children.length - 1] = last + text
  else parent.children.push(text)
}

function findOpenReportTag(stack: ({ children: ReportChild[] } | ReportNode)[], tag: ReportTag) {
  for (let i = stack.length - 1; i > 0; i -= 1) {
    const node = stack[i]
    if ("tag" in node && node.tag === tag) return i
  }
  return -1
}

function appendReportNode(node: ReportNode, sections: StructuredReportSection[], blocks: StructuredReport["blocks"]) {
  const meta = REPORT_TAGS[node.tag]
  if (meta.kind === "wrapper") {
    for (const child of node.children) {
      if (typeof child !== "string") appendReportNode(child, sections, blocks)
      else {
        const markdown = normalizeReportMarkdown(child)
        if (!markdown) continue
        const section = { tag: node.tag, title: meta.title, markdown, collapsed: false }
        sections.push(section)
        blocks.push({ kind: "section", section })
      }
    }
    return
  }

  const section = {
    tag: node.tag,
    title: meta.title,
    markdown: nodeChildrenToMarkdown(node.children),
    collapsed: meta.kind === "operational",
  }
  sections.push(section)
  blocks.push({ kind: "section", section })
}

function nodeChildrenToMarkdown(children: ReportChild[]): string {
  return normalizeReportMarkdown(
    children
      .map((child) => {
        if (typeof child === "string") return child
        const meta = REPORT_TAGS[child.tag]
        if (meta.kind === "wrapper") return nodeChildrenToMarkdown(child.children)
        return `\n\n**${meta.title}**\n\n${nodeChildrenToMarkdown(child.children)}`
      })
      .join(""),
  )
}

function isReportTag(tag: string): tag is ReportTag {
  return Object.hasOwn(REPORT_TAGS, tag)
}

export function decodeMarkupEntities(text: string) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
}
