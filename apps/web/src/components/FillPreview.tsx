import ReviewSetupDialog from "./ReviewSetupDialog"
import "./ReviewStarter.css"

type Target = { docName: string; column: { id: string; question: string } }

export default function FillPreview({ targets, checking, onCancel, onConfirm }: {
  targets: Target[]
  checking: boolean
  onCancel: () => void
  onConfirm: () => void
}) {
  const documents = [...new Set(targets.map((target) => target.docName))]
  const close = () => { if (!checking) onCancel() }
  return (
    <ReviewSetupDialog title="Review fill" onClose={close}>
          <p><strong>{targets.length} {targets.length === 1 ? "cell" : "cells"}</strong> will be processed across {documents.length} {documents.length === 1 ? "document" : "documents"}.</p>
          <p>Fill uses AI for each cell and may use AI again to re-check an answer. After filling, AI compares each question's cited answers across documents, one AI call per question. Review the answers and their sources.</p>
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
