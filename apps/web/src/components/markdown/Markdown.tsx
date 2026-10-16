import ReactMarkdown from "react-markdown"
import MarkdownContent from "./MarkdownContent"
import StructuredReportView from "./StructuredReportView"
import { parseStructuredReportMarkdown, unwrapTaskResult } from "./structured-report"

export default function Markdown({ children }: { children: string }) {
  const text = unwrapTaskResult(children)
  const report = parseStructuredReportMarkdown(text)
  if (report.structured) return <StructuredReportView report={report} />

  return (
    <div className="markdown">
      <MarkdownContent>{text}</MarkdownContent>
    </div>
  )
}

// Skill bodies and agent instructions carry XML-ish section tags (<task>,
// <rules>, ...) that ReactMarkdown would otherwise swallow as unknown HTML.
// Re-emit each tag as inline code on its own line so the sections stay visible
// and the markdown between them (bold, lists) still renders. Only tags alone on
// a line are rewritten, so inline-code mentions like `<section>` stay untouched.
export function DocMarkdown({ children }: { children: string }) {
  const escaped = children.replace(/^[ \t]*<(\/?[a-z][a-z0-9-]*)>[ \t]*$/gm, "\n\n`<$1>`\n\n")
  return (
    <div className="markdown">
      <ReactMarkdown>{escaped}</ReactMarkdown>
    </div>
  )
}
