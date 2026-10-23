import { useContext, useState } from "react"
import { Closing } from "../useExit"
import { conflictsChip, recordConflicts, type Conflicts } from "../../api/ingest/conflicts"
import { humanizeError } from "../../api/ingest/errors"
import type { Matter } from "../../api/ingest/matters"
import ModalHeader from "../ModalHeader"
import backdropClose from "../backdropClose"

const STATES: { value: Conflicts["status"]; label: string }[] = [
  { value: "cleared", label: "Cleared" },
  { value: "pending", label: "Pending" },
  { value: "bypassed", label: "Bypassed" },
]

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
  const [status, setStatus] = useState<Conflicts["status"]>(current?.status ?? "cleared")
  const [by, setBy] = useState(current?.by ?? attorney ?? "")
  const [owner, setOwner] = useState(current?.owner ?? attorney ?? "")
  const [due, setDue] = useState(current?.due ?? "")
  const [reason, setReason] = useState(current?.reason ?? "")
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string>()

  async function save() {
    setSaving(true)
    setError(undefined)
    try {
      onSaved(await recordConflicts(matterId, { status, by, owner, due, reason }))
      onClose()
    } catch (cause) {
      setError(humanizeError(cause, "Could not record the conflicts check."))
      setSaving(false)
    }
  }

  return (
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => !saving && onClose())}>
      <div className="picker-panel compare-dialog" role="dialog" aria-modal="true" aria-labelledby="conflicts-dialog-title" onClick={(e) => e.stopPropagation()}>
        <ModalHeader title={<span id="conflicts-dialog-title">Conflicts check</span>} onClose={onClose} />
        <form
          className="picker-body compare-form"
          onSubmit={(event) => {
            event.preventDefault()
            if (!saving) void save()
          }}
        >
          {current && <p className="compare-hint">Now: {conflictsChip(current).text.replace(/^Conflicts: /, "")}. Saving replaces it.</p>}
          <fieldset className="conflicts-states">
            <legend>Where the check stands</legend>
            {STATES.map((row) => (
              <label key={row.value}>
                <input type="radio" name="conflicts-status" value={row.value} checked={status === row.value} onChange={() => setStatus(row.value)} />
                {row.label}
              </label>
            ))}
          </fieldset>
          {status === "cleared" && (
            <label className="conflicts-field">
              Checked by
              <input autoFocus value={by} onChange={(e) => setBy(e.target.value)} placeholder="Name" />
            </label>
          )}
          {status === "pending" && (
            <>
              <label className="conflicts-field">
                Owner
                <input autoFocus value={owner} onChange={(e) => setOwner(e.target.value)} placeholder="Who will complete the check" />
              </label>
              <label className="conflicts-field">
                Due (optional)
                <input type="date" value={due} onChange={(e) => setDue(e.target.value)} />
              </label>
              <p className="conflicts-warning" role="note">
                Pending blocks redlines and drafting tools on this matter until it is cleared or bypassed.
              </p>
            </>
          )}
          {status === "bypassed" && (
            <>
              <label className="conflicts-field">
                Bypassed by
                <input autoFocus value={by} onChange={(e) => setBy(e.target.value)} placeholder="Name" />
              </label>
              <label className="conflicts-field">
                Reason
                <input value={reason} onChange={(e) => setReason(e.target.value)} placeholder="Why the check is bypassed" />
              </label>
              <p className="compare-hint">Reviews will state who bypassed the check and why.</p>
            </>
          )}
          <p className="compare-hint">This records that a check happened. doc.haus does not check for conflicts. It applies from the next message you send the assistant.</p>
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
