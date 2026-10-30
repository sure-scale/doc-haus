import { MAX_COMPARE_DOCUMENTS } from "../api/review-compare"
import { InfoTip } from "./InfoTip"
import ReviewSetupDialog from "./ReviewSetupDialog"
import { Tooltip } from "./Tooltip"
import "./ReviewStarter.css"

type Target = { docName: string; column: { id: string; question: string; compare?: boolean } }
const NOT_COMPARED = "Answers to this question are expected to differ by document, so AI does not compare them."

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
    <li key={column.id}>{column.question}{marked && column.compare === false && <>{" "}<Tooltip label={NOT_COMPARED}><span className="fill-preview-note" tabIndex={0}>Not compared</span></Tooltip></>}</li>
  )
  // The compare step in one plain sentence; how many calls, and what it skips, sit behind the "i".
  const limit = documentCount > MAX_COMPARE_DOCUMENTS
  const compareDetail = `One AI call for each question with ${limit ? `two to ${MAX_COMPARE_DOCUMENTS}` : "two or more"} cited answers. Questions with ${limit ? "fewer or more" : "fewer"} are skipped${marked ? ", and so are questions marked Not compared" : ""}.`
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
          {targets.length > 0 && <p>Fill uses AI for each cell and may use AI again to re-check an answer.{compared > 0 && <> After filling, AI compares each question's answers across documents.<InfoTip label={compareDetail} /></>} Review the answers and their sources.</p>}
          <div className="compare-actions">
            <button type="button" className="ghost" disabled={checking} onClick={close}>Cancel</button>
            <button type="button" className="primary" disabled={checking || !targets.length} onClick={onConfirm}>{checking ? "Checking…" : "Start fill"}</button>
          </div>
    </ReviewSetupDialog>
  )
}
