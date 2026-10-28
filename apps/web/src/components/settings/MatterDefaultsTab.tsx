import { useEffect, useState } from "react"
import { addSampleMatter, listJurisdictions, listMatters, type Jurisdiction } from "../../api/ingest/matters"
import { humanizeError } from "../../api/ingest/errors"
import { loadPrefs, savePrefs } from "../../prefs"
import JurisdictionSelect from "../JurisdictionSelect"
import { FieldRow } from "./controls"
import { SettingsFooter } from "./SettingsFooter"

// Defaults applied when a new matter is created. Per-browser (prefs.ts); the
// new-matter form reads them as its starting jurisdiction selection.
export default function MatterDefaultsTab({ onSampleAdded, onError, onSaved }: { onSampleAdded: () => void; onError: (message: string) => void; onSaved: () => void }) {
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [selected, setSelected] = useState<string[]>(() => loadPrefs().defaultJurisdictions)
  // Unknown until the matter list loads; the action stays disabled meanwhile.
  const [sampleExists, setSampleExists] = useState<boolean>()
  const [adding, setAdding] = useState(false)
  useEffect(() => {
    listJurisdictions().then(setJurisdictions)
    listMatters()
      .then((matters) => setSampleExists(matters.some((matter) => matter.sample)))
      .catch((error) => onError(humanizeError(error, "Could not check for the sample matter.")))
  }, [])
  async function addSample() {
    setAdding(true)
    try {
      await addSampleMatter()
      setSampleExists(true)
      onSampleAdded()
    } catch (error) {
      onError(humanizeError(error, "Could not add the sample matter."))
    } finally {
      setAdding(false)
    }
  }
  return (
    <>
    <section className="settings-section">
      <p className="settings-hint muted">
        Jurisdictions preselected on every new matter. Most firms work in one or two — set them once here instead of
        picking them on each matter. You can still change them per matter.
      </p>
      <FieldRow
        id="matter-default-jurisdictions"
        label="Jurisdictions"
        hint="Preselected for new matters in this browser."
      >
        <JurisdictionSelect
          id="matter-default-jurisdictions"
          aria-describedby="matter-default-jurisdictions-hint"
          jurisdictions={jurisdictions}
          selected={selected}
          onChange={setSelected}
        />
      </FieldRow>
    </section>
    <section className="settings-section">
      <h3>Sample matter</h3>
      <p id="sample-matter-hint" className="settings-hint muted">
        {sampleExists
          ? "The sample matter is in your matters list. Delete it there like any matter when you are done with it."
          : "A fictional engagement letter, ready for cited questions and a legal review. Delete it like any matter when you are done."}
      </p>
      <div className="settings-actions settings-actions-end">
        <button disabled={sampleExists !== false || adding} aria-describedby="sample-matter-hint" onClick={() => void addSample()}>
          {adding ? "Adding sample matter…" : sampleExists ? "Sample matter added" : "Add sample matter"}
        </button>
      </div>
    </section>
    <SettingsFooter tab="matters"><div className="settings-actions"><button type="button" className="primary" onClick={() => {
      try { savePrefs({ defaultJurisdictions: selected }, { requireStorage: true }); onSaved() }
      catch (error) { onError(humanizeError(error, "Could not save matter defaults.")) }
    }}>Save</button></div></SettingsFooter>
    </>
  )
}
