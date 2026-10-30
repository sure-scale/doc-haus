import type { ReviewStarterDefinition } from "../review-starters"
import "./StarterSummary.css"

// What a review-table starter sets up, read before it is used: the Workflows
// page's View and the table's starter chooser show the same definition.
// `purposeId` lets a chooser describe its selected option with the purpose.
export default function StarterSummary({ starter, questions = false, purposeId }: { starter: ReviewStarterDefinition; questions?: boolean; purposeId?: string }) {
  return (
    <div className="starter-summary">
      <p id={purposeId}>{starter.purpose}</p>
      <dl>
        <dt>Inputs</dt><dd>{starter.inputs}</dd>
        <dt>Output</dt><dd>{starter.output}</dd>
      </dl>
      {questions && (
        <section aria-label="Questions">
          <h3>Questions</h3>
          <ol>{starter.questions.map(({ text }) => <li key={text}>{text}</li>)}</ol>
        </section>
      )}
      <p className="starter-summary-limits">{starter.limits}</p>
    </div>
  )
}
