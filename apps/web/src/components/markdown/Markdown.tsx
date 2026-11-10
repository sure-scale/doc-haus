import ReactMarkdown from "react-markdown"
import type { Processor } from "unified"
import MarkdownContent from "./MarkdownContent"
import StructuredReportView from "./StructuredReportView"
import { parseStructuredReportMarkdown } from "./structured-report"

export default function Markdown({ children }: { children: string }) {
  const text = children
  const report = parseStructuredReportMarkdown(text)
  if (report.structured) return <StructuredReportView report={report} />

  return (
    <div className="markdown">
      <MarkdownContent>{text}</MarkdownContent>
    </div>
  )
}

// Instruction tags are literal documentation, including inside code examples.
// Disable HTML tokenization in the maintained parser instead of rewriting lines
// before Markdown has identified their context.
function literalTags(this: Processor) {
  this.data("micromarkExtensions", [{ disable: { null: ["htmlFlow", "htmlText"] } }])
}
const documentationPlugins = [literalTags]

export function DocMarkdown({ children }: { children: string }) {
  return (
    <div className="markdown">
      <ReactMarkdown remarkPlugins={documentationPlugins}>{children}</ReactMarkdown>
    </div>
  )
}
