import { useState, type Dispatch, type SetStateAction } from "react"
import { savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import type { Jurisdiction } from "../../api/ingest/matters"
import { savePrefs } from "../../prefs"
import JurisdictionSelect from "../JurisdictionSelect"
import { LegendField, PageLayoutPreview } from "../settings/PageLayout"
import { ADVICE_LEGENDS, AGREEMENT_LEGENDS, suggestedLegends } from "../settings/page-layout"

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
  const suggested = suggestedLegends(selectedJx)
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

      <section className="settings-section">
        <h3>Legends</h3>
        <p className="settings-hint muted">
          A legend has legal effect, so doc.haus adds none unless you choose one. Each appears only on the documents
          it names.{suggested.agreement.length > 0 && " Suggestions follow common practice in your default jurisdictions."}
        </p>
        <LegendField
          id="onboarding-agreement-legend"
          label="Agreement legend"
          hint="At the top of every page of an agreement."
          presets={AGREEMENT_LEGENDS}
          suggested={suggested.agreement}
          disabled={busy || !drafting}
          value={drafting?.agreementLegend ?? ""}
          onChange={(agreementLegend) => setDrafting((d) => d && { ...d, agreementLegend })}
        />
        <LegendField
          id="onboarding-advice-legend"
          label="Advice legend"
          hint="At the top of every page of advice, such as a memo, opinion or advice letter."
          presets={ADVICE_LEGENDS}
          suggested={suggested.advice}
          disabled={busy || !drafting}
          value={drafting?.adviceLegend ?? ""}
          onChange={(adviceLegend) => setDrafting((d) => d && { ...d, adviceLegend })}
        />
        {drafting && <PageLayoutPreview prefs={drafting} />}
        <p className="settings-hint muted">Page numbers and the draft line can be changed in Settings → Drafting.</p>
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
              await savePreferences({ attorney: drafting.attorney, firm: drafting.firm, agreementLegend: drafting.agreementLegend, adviceLegend: drafting.adviceLegend })
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
