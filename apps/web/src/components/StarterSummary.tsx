import type { ReviewStarterDefinition } from "../review-starters"
import "./StarterSummary.css"

// What a review-table starter sets up, read before it is used: the Workflows
// page's View and the composer's Workflows dialog show the same definition.
export default function StarterSummary({ starter, questions = false }: { starter: ReviewStarterDefinition; questions?: boolean }) {
  return (
    <div className="starter-summary">
      <p>{starter.purpose}</p>
      <dl>
        <dt>Inputs</dt><dd>{starter.inputs}</dd>
        <dt>Output</dt><dd>{starter.output}</dd>
      </dl>
      {questions && (
        <section aria-label="Questions">
          <h3>Questions</h3>
          <ol>{starter.questions.map((question) => <li key={question}>{question}</li>)}</ol>
        </section>
      )}
      <p className="starter-summary-limits">{starter.limits}</p>
    </div>
  )
}
