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
}

// Subagent results arrive wrapped by the task tool as
// <task id="..." state="completed"><task_result>…markdown…</task_result></task>.
// Strip that envelope so only the model's markdown reaches the renderer.
export function unwrapTaskResult(text: string) {
  const decoded = decodeMarkupEntities(text)
  const match = decoded.match(/<task_result>([\s\S]*?)<\/task_result>/)
  return (match ? match[1] : decoded).trim()
}

export function parseStructuredReportMarkdown(text: string): StructuredReport {
  const children = parseReportChildren(unwrapTaskResult(text))
  const structured = children.some((child) => typeof child !== "string")
  if (!structured) return { structured: false, preamble: "", sections: [] }

  const sections: StructuredReportSection[] = []
  const preamble: string[] = []
  for (const child of children) {
    if (typeof child === "string") {
      if (!sections.length && child.trim()) preamble.push(child)
      continue
    }
    appendReportNode(child, sections)
  }

  return {
    structured: true,
    preamble: normalizeReportMarkdown(preamble.join("")),
    sections,
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
  let cursor = 0
  for (const match of text.matchAll(pattern)) {
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
      continue
    }

    const node: ReportNode = { tag, children: [] }
    stack.at(-1)!.children.push(node)
    if (!match[0].endsWith("/>")) stack.push(node)
  }
  appendText(stack.at(-1)!, text.slice(cursor))
  return root.children
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

function appendReportNode(node: ReportNode, sections: StructuredReportSection[]) {
  const meta = REPORT_TAGS[node.tag]
  if (meta.kind === "wrapper") {
    const directText = normalizeReportMarkdown(
      node.children.filter((child): child is string => typeof child === "string").join(""),
    )
    if (directText) sections.push({ tag: node.tag, title: meta.title, markdown: directText, collapsed: false })
    for (const child of node.children) if (typeof child !== "string") appendReportNode(child, sections)
    return
  }

  sections.push({
    tag: node.tag,
    title: meta.title,
    markdown: nodeChildrenToMarkdown(node.children),
    collapsed: meta.kind === "operational",
  })
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
  return tag in REPORT_TAGS
}

export function decodeMarkupEntities(text: string) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
}
