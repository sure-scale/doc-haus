import MarkdownContent from "./MarkdownContent"
import type { StructuredReport } from "./structured-report"

export default function StructuredReportView({ report }: { report: StructuredReport }) {
  const collapsePreamble =
    report.preamble.length > 120 || /\b(Stage|Now I|checkpoint|pipeline|dispatch|gate|retry|continuing)\b/i.test(report.preamble)

  return (
    <div className="markdown structured-report">
      {report.preamble &&
        (collapsePreamble ? (
          <details className="report-narration">
            <summary>Run narration</summary>
            <MarkdownContent>{report.preamble}</MarkdownContent>
          </details>
        ) : (
          <MarkdownContent>{report.preamble}</MarkdownContent>
        ))}
      {report.sections.map((section, index) =>
        section.collapsed ? (
          <details key={`${section.tag}-${index}`} className="report-section report-section-collapsed">
            <summary>{section.title}</summary>
            <div className="report-section-body">
              <MarkdownContent>{section.markdown}</MarkdownContent>
            </div>
          </details>
        ) : (
          <section key={`${section.tag}-${index}`} className={`report-section report-section-${section.tag}`}>
            <h3 className="report-section-title">{section.title}</h3>
            <div className="report-section-body">
              <MarkdownContent>{section.markdown}</MarkdownContent>
            </div>
          </section>
        ),
      )}
    </div>
  )
}
