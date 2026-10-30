import { MAX_COMPARE_DOCUMENTS } from "../api/review-compare"
import ReviewSetupDialog from "./ReviewSetupDialog"
import "./ReviewStarter.css"

type Target = { docName: string; column: { id: string; question: string; compare?: boolean } }

// `documentCount` is the table's: when the targets are every one of its
// documents for the same questions, the questions are listed once.
export default function FillPreview({ targets, documentCount, checking, onCancel, onConfirm }: {
  targets: Target[]
  documentCount: number
  checking: boolean
  onCancel: () => void
  onConfirm: () => void
}) {
  const documents = [...new Set(targets.map((target) => target.docName))]
  const questions = [...new Map(targets.map((target) => [target.column.id, target.column])).values()]
  const grid = documents.length === documentCount && targets.length === documents.length * questions.length
  // A question marked not to compare is filled but never compared; it is
  // marked only beside questions that are.
  const compared = documentCount > 1 ? questions.filter((question) => question.compare !== false).length : 0
  const marked = compared > 0 && compared < questions.length
  const item = (column: Target["column"]) => (
    <li key={column.id}>{column.question}{marked && column.compare === false && <>{" "}<span className="fill-preview-note">Not compared</span></>}</li>
  )
  const count = (n: number, word: string) => `${n} ${n === 1 ? word : `${word}s`}`
  const close = () => { if (!checking) onCancel() }
  return (
    <ReviewSetupDialog title={targets.length ? `Fill ${count(targets.length, "cell")}?` : "Nothing to fill"} onClose={close}>
          {grid ? (
            <section>
              <h3>{count(questions.length, "question")} × {count(documents.length, "document")}</h3>
              <ul>{questions.map(item)}</ul>
            </section>
          ) : documents.map((document) => (
            <section key={document}>
              <h3>{document}</h3>
              <ul>{targets.filter((target) => target.docName === document).map((target) => item(target.column))}</ul>
            </section>
          ))}
          {!targets.length && <p>There are no eligible cells to fill.</p>}
          {targets.length > 0 && <p>Fill uses AI for each cell and may use AI again to re-check an answer.{compared > 0 && ` After filling, AI compares the cited answers across documents: one AI call for each question with ${documentCount > MAX_COMPARE_DOCUMENTS ? `two to ${MAX_COMPARE_DOCUMENTS}` : "two or more"} cited answers.`}{marked && " It skips questions marked Not compared."} Review the answers and their sources.</p>}
          <div className="compare-actions">
            <button type="button" className="ghost" disabled={checking} onClick={close}>Cancel</button>
            <button type="button" className="primary" disabled={checking || !targets.length} onClick={onConfirm}>{checking ? "Checking…" : "Start fill"}</button>
          </div>
    </ReviewSetupDialog>
  )
}
