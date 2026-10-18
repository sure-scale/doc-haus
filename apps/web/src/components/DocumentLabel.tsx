import { useState } from "react"
import { setDocumentLabel, type DocumentLabel } from "../api/ingest/documents"
import { humanizeError } from "../api/ingest/errors"
import { useToast } from "./Toast"

export default function DocumentLabelBadge({ matterId, name, label, onChanged }: {
  matterId: string
  name: string
  /** "combined" names a saved combined file the lawyer has not labelled, and
   * "earlier_combined" one from an earlier run; each is shown, never saved,
   * since a label the lawyer chooses is one of the four. */
  label: DocumentLabel | "combined" | "earlier_combined"
  onChanged: () => void
}) {
  const [saving, setSaving] = useState(false)
  const toast = useToast()
  return (
    <span className={`document-label-badge is-${label}`} aria-busy={saving || undefined}
      title={label === "earlier_combined" ? "The combined file from an earlier run" : undefined}>
      <select
        style={{ color: "#176b50", WebkitTextFillColor: "#176b50" }}
        aria-label={`Label for ${name}`}
        value={label}
        disabled={saving}
        onChange={async (event) => {
          const next = event.target.value as DocumentLabel
          setSaving(true)
          try {
            await setDocumentLabel(matterId, name, next)
            onChanged()
          } catch (error) {
            toast("error", humanizeError(error, "Could not save the document label."))
          } finally {
            setSaving(false)
          }
        }}
      >
        {label === "combined" && <option value="combined">Combined</option>}
        {label === "earlier_combined" && <option value="earlier_combined">Earlier combined</option>}
        <option value="draft">Draft</option>
        <option value="redline">Redline</option>
        <option value="notes">Notes</option>
        <option value="final">Final</option>
      </select>
      <svg width="10" height="10" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" aria-hidden><path d="m4 6 4 4 4-4" /></svg>
    </span>
  )
}
