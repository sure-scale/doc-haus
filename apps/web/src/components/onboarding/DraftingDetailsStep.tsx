import { useState, type Dispatch, type ReactNode, type SetStateAction } from "react"
import { savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import type { Jurisdiction } from "../../api/ingest/matters"

type DateFormat = DraftingPreferences["dateFormat"]

export default function DraftingDetailsStep({
  drafting,
  setDrafting,
  setStep,
  onClose,
  busyLabel,
  disabled = false,
}: {
  drafting: DraftingPreferences | null
  setDrafting: Dispatch<SetStateAction<DraftingPreferences | null>>
  jurisdictions: Jurisdiction[]
  selectedJx: string[]
  setSelectedJx: (codes: string[]) => void
  dateFormatChoice: DateFormat | null
  setDateFormatChoice: (dateFormat: DateFormat) => void
  setStep: (step: "setup" | "details" | "research") => void
  onClose: () => Promise<void>
  /** Shown after the drafting details, before the actions. */
  extra?: ReactNode
  /** Replaces "Saving…" while finishing does more than save. */
  busyLabel?: string
  disabled?: boolean
}) {
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  return (
    <>
      {error && <p role="alert" className="settings-warning">{error}</p>}
      <section className="settings-section">
        <h3>Attorney and firm</h3>
        <p className="settings-hint muted">
          Used in signature blocks, notice clauses and letterhead.
        </p>
        <div className="onboarding-details-grid">
          <label className="onboarding-field" htmlFor="onboarding-attorney">Attorney
          <input
            id="onboarding-attorney"
            autoComplete="name"
            disabled={busy || !drafting}
            placeholder="Jane Doe"
            value={drafting?.attorney ?? ""}
            onChange={(e) => setDrafting((d) => d && { ...d, attorney: e.target.value })}
          /></label>
          <label className="onboarding-field" htmlFor="onboarding-firm">Firm
          <input
            id="onboarding-firm"
            autoComplete="organization"
            disabled={busy || !drafting}
            placeholder="Doe & Partners LLP"
            value={drafting?.firm ?? ""}
            onChange={(e) => setDrafting((d) => d && { ...d, firm: e.target.value })}
          /></label>
        </div>
      </section>

      <div className="onboarding-actions">
        <button type="button" className="onboarding-back" disabled={busy} onClick={() => setStep("research")}>Back to research</button>
        <button
          className="primary btn-lg"
          disabled={busy || !drafting || disabled}
          onClick={async () => {
            if (!drafting || busy) return
            setBusy(true); setError("")
            try {
              await savePreferences({ attorney: drafting.attorney, firm: drafting.firm })
              await onClose()
            } catch { setError("Could not save drafting details. Retry.") }
            finally { setBusy(false) }
          }}
        >
          {busy ? busyLabel ?? "Saving…" : "Save and finish"}{!busy && <span aria-hidden="true"> ✓</span>}
        </button>
      </div>
    </>
  )
}
