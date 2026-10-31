import MarkdownContent from "./MarkdownContent"
import type { StructuredReport } from "./structured-report"

export default function StructuredReportView({ report }: { report: StructuredReport }) {
  return (
    <div className="markdown structured-report">
      {report.blocks.map((block, index) =>
        block.kind === "markdown" ? (
          <MarkdownContent key={`text-${index}`}>{block.markdown}</MarkdownContent>
        ) : block.section.collapsed ? (
          <details key={`${block.section.tag}-${index}`} className="report-section report-section-collapsed">
            <summary>{block.section.title}</summary>
            <div className="report-section-body">
              <MarkdownContent>{block.section.markdown}</MarkdownContent>
            </div>
          </details>
        ) : (
          <section
            key={`${block.section.tag}-${index}`}
            className={`report-section report-section-${block.section.tag}`}
          >
            <h3 className="report-section-title">{block.section.title}</h3>
            <div className="report-section-body">
              <MarkdownContent>{block.section.markdown}</MarkdownContent>
            </div>
          </section>
        ),
      )}
    </div>
  )
}
