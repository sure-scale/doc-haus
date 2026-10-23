import { useContext, useRef, useState } from "react"
import { flushSync } from "react-dom"
import { Closing } from "../useExit"
import { conflictsChip, recordConflicts, type Conflicts } from "../../api/ingest/conflicts"
import { humanizeError } from "../../api/ingest/errors"
import type { Matter } from "../../api/ingest/matters"
import ModalHeader from "../ModalHeader"
import backdropClose from "../backdropClose"
import ConflictsFields, { conflictsDraft, draftInput, draftProblems, type ConflictsDraft } from "./ConflictsFields"

// Record how the matter's conflicts check stands. A new record replaces the one
// shown (the matter keeps only the latest), a bypass included.
export default function ConflictsDialog({
  matterId,
  current,
  attorney,
  onSaved,
  onClose,
}: {
  matterId: string
  current?: Conflicts
  // The firm's attorney name from preferences, offered as the default "by".
  attorney?: string
  onSaved: (matter: Matter) => void
  onClose: () => void
}) {
  const closing = useContext(Closing)
  const [draft, setDraft] = useState<ConflictsDraft>(() => conflictsDraft(current, "cleared", attorney))
  // Missing names are shown once Record has been tried, and clear as they are typed.
  const [tried, setTried] = useState(false)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string>()
  const form = useRef<HTMLFormElement>(null)
  const problems = tried ? draftProblems(draft) : {}

  async function save() {
    const input = draftInput(draft)
    if (!input) return
    if (Object.keys(draftProblems(draft)).length) {
      flushSync(() => setTried(true))
      form.current?.querySelector<HTMLElement>("[aria-invalid]")?.focus()
      return
    }
    setSaving(true)
    setError(undefined)
    try {
      onSaved(await recordConflicts(matterId, input))
      onClose()
    } catch (cause) {
      setError(humanizeError(cause, "Could not record the conflicts check."))
      setSaving(false)
    }
  }

  return (
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => !saving && onClose())}>
      <div className="picker-panel compare-dialog" role="dialog" aria-modal="true" aria-labelledby="conflicts-dialog-title" onClick={(e) => e.stopPropagation()}>
        {/* Closing mid-save would hide a failed save; Close and Escape wait for it, as the backdrop does. */}
        <ModalHeader title={<span id="conflicts-dialog-title">Conflicts check</span>} onClose={() => !saving && onClose()} />
        <form
          ref={form}
          className="picker-body compare-form"
          noValidate
          onSubmit={(event) => {
            event.preventDefault()
            if (!saving) void save()
          }}
        >
          {current && <p className="compare-hint">Now: {conflictsChip(current).text.replace(/^Conflicts: /, "")}. Saving replaces it.</p>}
          <ConflictsFields
            draft={draft}
            onChange={setDraft}
            problems={problems}
            legend="Where the check stands"
            fieldClassName="conflicts-field"
            attorney={attorney}
            autoFocus
          />
          <p className="compare-hint">It applies from the next message you send the assistant.</p>
          {error && (
            <p className="compare-error" role="alert">
              {error}
            </p>
          )}
          <div className="compare-actions">
            <button type="button" onClick={onClose} disabled={saving}>
              Cancel
            </button>
            <button type="submit" className="primary" disabled={saving}>
              {saving ? "Saving…" : "Record"}
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}
