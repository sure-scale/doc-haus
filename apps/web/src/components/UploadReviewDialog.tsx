import { useId, useLayoutEffect, useRef, type RefObject } from "react"
import type { ExistingSourceChoice } from "../api/ingest/documents"
import useModalFocus from "./useModalFocus"

// One recognized file stays in this dialog through acknowledgement and a
// possible filename conflict. Other uploads can continue behind the review.
export default function UploadReviewDialog({
  reviewId, fileName, pending, existing, choice, collision, waiting,
  onAccept, onChoose, onSkip, returnFocus,
}: {
  reviewId: number
  fileName: string
  pending?: "queued" | "uploading" | "reading"
  existing?: { name: string; role: "source" | "deliverable" }
  choice?: ExistingSourceChoice
  collision: boolean
  waiting: number
  onAccept: () => void
  onChoose: (choice: ExistingSourceChoice) => void
  onSkip: () => void
  returnFocus: RefObject<HTMLElement | null>
}) {
  const dialog = useRef<HTMLDialogElement>(null)
  const heading = useRef<HTMLHeadingElement>(null)
  const titleId = useId()
  const descriptionId = useId()
  useModalFocus(dialog)
  useLayoutEffect(() => {
    const element = dialog.current
    element?.showModal()
    return () => {
      element?.close()
      returnFocus.current?.focus({ preventScroll: true })
    }
  }, [returnFocus])
  // A button can disappear or become disabled while its request is queued.
  // Keep focus on the review context rather than leaving it on the page.
  useLayoutEffect(() => {
    heading.current?.focus({ preventScroll: true })
  }, [reviewId, pending, collision])
  return (
    <dialog ref={dialog} className="upload-review-dialog" aria-labelledby={titleId} aria-describedby={descriptionId}
      onCancel={(event) => {
        event.preventDefault()
        event.stopPropagation()
        // Once submitted, a save cannot be cancelled by closing its dialog.
        if (!pending) onSkip()
      }}
      onKeyDown={(event) => {
        // The phone Documents drawer also handles Escape. Only this dialog
        // should answer it, including while the upload is in flight.
        if (event.key !== "Escape") return
        event.preventDefault()
        event.stopPropagation()
        if (!pending) onSkip()
      }}>
      <div className="upload-review-content">
        <h2 ref={heading} id={titleId} tabIndex={-1}>Add this chronology?</h2>
        <p className="upload-review-name">{fileName}</p>
        <p id={descriptionId}>This file is a chronology export. Add it as a supplied account, not independent corroboration. It will not replace the chronology workbench.</p>
        {collision && !pending && (
          <div className="upload-review-collision">
            <p>{existing?.role === "deliverable"
              ? "A document created by doc.haus has this name. Keep both to add a separate file."
              : "A source with this name is already in this matter. Save this file as a new version, or keep both."}</p>
            {existing?.name && existing.name !== fileName && <p className="upload-review-name">{existing.name}</p>}
            {existing?.role === "source" && <p className="upload-review-detail">Earlier versions stay in its history.</p>}
          </div>
        )}
        {!collision && !pending && choice && <p className="upload-review-detail">Filename choice: {choice === "replace" ? "save as a new version" : "keep both files"}.</p>}
        {pending && <p className="upload-review-progress" role="status">{pending === "queued" ? "Waiting for other uploads…" : pending === "reading" ? "Reading this file…" : "Adding this file…"}</p>}
        {waiting > 0 && <p className="upload-review-detail">{waiting} more file{waiting === 1 ? "" : "s"} awaiting review. This choice applies only to this file.</p>}
      </div>
      <div className="upload-review-actions">
        <button type="button" className="icon-btn" disabled={Boolean(pending)} onClick={onSkip}>Skip this file</button>
        {collision && !pending ? <>
          {existing?.role === "source" && <button type="button" className="icon-btn" onClick={() => onChoose("replace")}>Save as new version</button>}
          <button type="button" className="icon-btn upload-choice-lead" onClick={() => onChoose("keep-both")}>Keep both</button>
        </> : <button type="button" className="icon-btn upload-choice-lead" disabled={Boolean(pending)} onClick={onAccept}>{pending ? "Adding…" : "Add as supplied chronology"}</button>}
      </div>
    </dialog>
  )
}
