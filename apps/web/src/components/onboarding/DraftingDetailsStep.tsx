import { useState, type Dispatch, type SetStateAction } from "react"
import { savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import type { Jurisdiction } from "../../api/ingest/matters"
import { savePrefs } from "../../prefs"
import JurisdictionSelect from "../JurisdictionSelect"

export default function DraftingDetailsStep({
  drafting,
  setDrafting,
  jurisdictions,
  selectedJx,
  setSelectedJx,
  setStep,
  onClose,
}: {
  drafting: DraftingPreferences | null
  setDrafting: Dispatch<SetStateAction<DraftingPreferences | null>>
  jurisdictions: Jurisdiction[]
  selectedJx: string[]
  setSelectedJx: (codes: string[]) => void
  setStep: (step: "setup" | "details" | "research") => void
  onClose: () => Promise<void>
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

      <section className="settings-section">
        <h3><label htmlFor="onboarding-jurisdictions">Default jurisdictions</label></h3>
        <p id="onboarding-jurisdictions-hint" className="settings-hint muted">
          Preselected on every new matter — most firms work in one or two. You can still change them per
          matter.
        </p>
        <div className="row settings-row">
          <JurisdictionSelect id="onboarding-jurisdictions" aria-describedby="onboarding-jurisdictions-hint" jurisdictions={jurisdictions} selected={selectedJx} onChange={setSelectedJx} />
        </div>
      </section>

      <div className="onboarding-actions">
        <button disabled={busy} onClick={() => setStep("research")}>Back to research</button>
        <button
          className="primary"
          disabled={busy || !drafting}
          onClick={async () => {
            if (!drafting || busy) return
            setBusy(true); setError("")
            try {
              await savePreferences({ attorney: drafting.attorney, firm: drafting.firm })
              savePrefs({ defaultJurisdictions: selectedJx })
              await onClose()
            } catch { setError("Could not save drafting details. Retry.") }
            finally { setBusy(false) }
          }}
        >
          {busy ? "Saving…" : "Save and finish"}{!busy && <span aria-hidden="true"> ✓</span>}
        </button>
      </div>
    </>
  )
}
