import { useId, useState } from "react"
import type { StarterQuestion } from "../review-starters"
import type { StarterExampleAnswer, StarterExample } from "../review-starters-example"
import { MARKER_LOOK, MarkerGlyph } from "./ReviewMarker"
import ReviewSetupDialog from "./ReviewSetupDialog"

// A starter's fictional worked example as a small read-only review table. Each
// answer opens a read-only detail laid out as the table's own cell detail: its
// marker and reason, the answer and the cited passage. No model, no matter, no writes.

const marker = (answer: StarterExampleAnswer) =>
  !answer.citation ? { state: "not_addressed" as const, word: "Not found", reason: "The document does not appear to address this question." }
  : answer.comparison === "differs" ? { state: "differs" as const, word: "Differs", reason: answer.reason } : undefined

export default function ReviewExample({ questions, example }: { questions: readonly StarterQuestion[]; example: StarterExample }) {
  const ids = useId()
  const [open, setOpen] = useState<StarterExampleAnswer>()
  const openMark = open && marker(open)
  const openQuestion = open && questions[open.question]?.text
  return (
    <div className="review-example">
      <p className="review-example-note">{example.note}</p>
      <p className="review-example-hint">In this example, terms without a marker agree with the other agreements. Select an answer to see its source.</p>
      <div className="review-scroll review-example-scroll" role="region" aria-label="Example review table" tabIndex={0}>
        <table className="review-table review-example-table" style={{ minWidth: 180 + questions.length * 168 }}>
          <thead>
            <tr>
              <th className="review-doc-col" scope="col">Document</th>
              {questions.map(({ text }) => <th key={text} scope="col">{text}</th>)}
            </tr>
          </thead>
          <tbody>
            {example.documents.map((document) => (
              <tr key={document.name}>
                <th className="review-doc-col" scope="row">{document.name}</th>
                {questions.map(({ text }, index) => {
                  const answer = example.answers.find((item) => item.document === document.name && item.question === index)
                  if (!answer) return <td key={text} />
                  const mark = marker(answer)
                  return (
                    <td key={text} className="review-example-cell">
                      <button type="button" aria-haspopup="dialog" onClick={() => setOpen(answer)}>
                        <span className="review-example-answer">{answer.answer}</span>
                        {answer.citation && <span className="review-cite">§ {answer.citation.section}</span>}
                        {mark && <span className={`review-marker is-${MARKER_LOOK[mark.state].tone}`}><MarkerGlyph state={mark.state} />{mark.word}</span>}
                      </button>
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && openQuestion && (
        <ReviewSetupDialog
          title={`Example answer: ${openQuestion}`}
          heading={<span className="viewer-title-text"><span className="review-detail-question">{openQuestion}</span><span className="viewer-title-meta">Fictional example · {open.document}</span></span>}
          variant="detail"
          onClose={() => setOpen(undefined)}
        >
          {openMark && (
            <div className={`review-walk-why is-${MARKER_LOOK[openMark.state].tone}`}>
              <p className="review-walk-word"><MarkerGlyph state={openMark.state} />{openMark.word}{openMark.state === "differs" && <span className="review-ai-tag">AI</span>}</p>
              <p>{openMark.reason}</p>
            </div>
          )}
          <section className="review-detail-section" aria-labelledby={`${ids}-answer`}>
            <h3 className="review-detail-label" id={`${ids}-answer`}>Answer</h3>
            <p className="review-detail-answer">{open.answer}</p>
          </section>
          {open.citation && (
            <section className="review-detail-section" aria-labelledby={`${ids}-source`}>
              <h3 className="review-detail-label" id={`${ids}-source`}>Source</h3>
              <div className="citation review-source">
                <div className="ref">[{open.document} § {open.citation.section}]</div>
                <div className="excerpt">{open.citation.excerpt}</div>
              </div>
            </section>
          )}
        </ReviewSetupDialog>
      )}
    </div>
  )
}
