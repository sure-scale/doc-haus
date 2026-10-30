import { useId, useLayoutEffect, useRef, useState } from "react"
import type { StarterExampleAnswer, StarterExample } from "../review-starters-example"
import { MarkerGlyph } from "./ReviewMarker"

// A starter's fictional worked example as a small read-only review table. The
// table shows each answer and its marker; one selected answer at a time opens
// its comparison and cited passage below. No model, no matter, no writes.

const marker = (answer: StarterExampleAnswer) =>
  !answer.citation ? { state: "no_citation" as const, word: "Not found" } : answer.comparison === "differs" ? { state: "differs" as const, word: "Differs" } : undefined

export default function ReviewExample({ questions, example }: { questions: readonly string[]; example: StarterExample }) {
  const detailId = useId()
  const [selected, setSelected] = useState<StarterExampleAnswer>()
  const detail = useRef<HTMLElement>(null)
  useLayoutEffect(() => { if (selected) detail.current?.scrollIntoView({ block: "nearest" }) }, [selected])
  const shown = selected && questions[selected.question]
  return (
    <div className="review-example">
      <p className="review-example-note">{example.note}</p>
      <div className="review-scroll review-example-scroll" role="region" aria-label="Example review table" tabIndex={0}>
        <table className="review-table review-example-table" style={{ minWidth: 180 + questions.length * 168 }}>
          <thead>
            <tr>
              <th className="review-doc-col" scope="col">Document</th>
              {questions.map((question) => <th key={question} scope="col">{question}</th>)}
            </tr>
          </thead>
          <tbody>
            {example.documents.map((document) => (
              <tr key={document.name}>
                <th className="review-doc-col" scope="row">{document.name}</th>
                {questions.map((question, index) => {
                  const answer = example.answers.find((item) => item.document === document.name && item.question === index)
                  if (!answer) return <td key={question} />
                  const mark = marker(answer)
                  const open = selected === answer
                  return (
                    <td key={question} className={`review-example-cell${open ? " is-selected" : ""}`}>
                      <button type="button" aria-expanded={open} aria-controls={detailId} onClick={() => setSelected(open ? undefined : answer)}>
                        <span className="review-example-answer">{answer.answer}</span>
                        {answer.citation && <span className="review-cite">§ {answer.citation.section}</span>}
                        {mark && <span className={`review-marker is-${mark.state === "differs" ? "info" : "muted"}`}><MarkerGlyph state={mark.state} />{mark.word}</span>}
                      </button>
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <section ref={detail} id={detailId} className="review-example-detail" aria-label="Selected answer" aria-live="polite">
        {selected && shown ? (
          <>
            <p className="review-walk-doc">{selected.document} · {shown}</p>
            <p className="review-detail-answer">{selected.answer}</p>
            <p className="review-example-comparison">
              {!selected.citation ? <><strong>Not found.</strong> The agreement does not appear to address this question.</>
                : selected.comparison === "differs" ? <><strong>AI comparison: Differs</strong> — {selected.reason}</>
                : <><strong>AI comparison: Consistent</strong> with the other agreements.</>}
            </p>
            {selected.citation && (
              <div className="citation review-source">
                <div className="ref">[{selected.document} § {selected.citation.section}]</div>
                <div className="excerpt">{selected.citation.excerpt}</div>
              </div>
            )}
          </>
        ) : <p className="review-example-hint">Answers without a marker are consistent across the agreements. Select an answer to see its source.</p>}
      </section>
    </div>
  )
}
