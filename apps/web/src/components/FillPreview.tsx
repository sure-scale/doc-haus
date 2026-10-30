import ReviewSetupDialog from "./ReviewSetupDialog"
import "./ReviewStarter.css"

type Target = { docName: string; column: { id: string; question: string; compare?: boolean } }

export default function FillPreview({ targets, checking, onCancel, onConfirm }: {
  targets: Target[]
  checking: boolean
  onCancel: () => void
  onConfirm: () => void
}) {
  const documents = [...new Set(targets.map((target) => target.docName))]
  // A question marked not to compare is filled but never compared.
  const compared = new Set(targets.filter((target) => target.column.compare !== false).map((target) => target.column.id)).size
  const close = () => { if (!checking) onCancel() }
  return (
    <ReviewSetupDialog title="Review fill" onClose={close}>
          <p><strong>{targets.length} {targets.length === 1 ? "cell" : "cells"}</strong> will be processed across {documents.length} {documents.length === 1 ? "document" : "documents"}.</p>
          <p>Fill uses AI for each cell and may use AI again to re-check an answer.{compared > 0 && ` After filling, AI compares the cited answers across documents for ${compared} ${compared === 1 ? "question" : "questions"}, one AI call per question.`} Review the answers and their sources.</p>
          {documents.map((document) => (
            <section key={document}>
              <h3>{document}</h3>
              <ul>{targets.filter((target) => target.docName === document).map((target) => <li key={target.column.id}>{target.column.question}</li>)}</ul>
            </section>
          ))}
          {!targets.length && <p>There are no eligible cells to fill.</p>}
          <div className="compare-actions">
            <button type="button" className="ghost" disabled={checking} onClick={close}>Cancel</button>
            <button type="button" className="primary" disabled={checking || !targets.length} onClick={onConfirm}>{checking ? "Checking…" : "Start fill"}</button>
          </div>
    </ReviewSetupDialog>
  )
}
