import { useEffect, useState } from "react"
import { listJurisdictions, type Jurisdiction } from "../../api/ingest/matters"
import { loadPrefs, savePrefs } from "../../prefs"
import JurisdictionSelect from "../JurisdictionSelect"
import { FieldRow } from "./controls"

// Defaults applied when a new matter is created. Per-browser (prefs.ts); the
// new-matter form reads them as its starting jurisdiction selection.
export default function MatterDefaultsTab() {
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [selected, setSelected] = useState<string[]>(() => loadPrefs().defaultJurisdictions)
  useEffect(() => {
    listJurisdictions().then(setJurisdictions)
  }, [])
  return (
    <section className="settings-section">
      <h3>Matter defaults</h3>
      <p className="settings-hint muted">
        Jurisdictions preselected on every new matter. Most firms work in one or two — set them once here instead of
        picking them on each matter. You can still change them per matter.
      </p>
      <FieldRow
        id="matter-default-jurisdictions"
        label="Jurisdictions"
        hint="Preselected for new matters in this browser. Changes save immediately."
      >
        <JurisdictionSelect
          id="matter-default-jurisdictions"
          aria-describedby="matter-default-jurisdictions-hint"
          jurisdictions={jurisdictions}
          selected={selected}
          onChange={(codes) => {
            setSelected(codes)
            savePrefs({ defaultJurisdictions: codes })
          }}
        />
      </FieldRow>
    </section>
  )
}
