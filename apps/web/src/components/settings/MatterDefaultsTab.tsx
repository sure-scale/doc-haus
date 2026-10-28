import { useEffect, useState } from "react"
import { addSampleMatter, listJurisdictions, listMatters, type Jurisdiction } from "../../api/ingest/matters"
import { humanizeError } from "../../api/ingest/errors"
import { getPreferences, savePreferences } from "../../api/ingest/preferences"
import { loadPrefs, savePrefs } from "../../prefs"
import JurisdictionSelect from "../JurisdictionSelect"
import { FieldRow } from "./controls"
import { SettingsFooter } from "./SettingsFooter"

const SUGGESTED_PRACTICE_AREAS = [
  "Banking and finance", "Commercial", "Competition", "Construction", "Corporate",
  "Data protection and privacy", "Dispute resolution", "Employment", "Energy", "Family",
  "Immigration", "Insolvency and restructuring", "Insurance", "Intellectual property",
  "Litigation", "Private client", "Real estate", "Regulatory", "Tax", "Technology",
]

// Defaults applied when a new matter is created. Per-browser (prefs.ts); the
// new-matter form reads them as its starting jurisdiction selection.
export default function MatterDefaultsTab({ onSampleAdded, onError, onSaved, onBusy }: { onSampleAdded: () => void; onError: (message: string) => void; onSaved: () => void; onBusy: (busy: boolean) => void }) {
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [selected, setSelected] = useState<string[]>(() => loadPrefs().defaultJurisdictions)
  const [practiceAreas, setPracticeAreas] = useState<string[] | null>(null)
  const [loadedPracticeAreas, setLoadedPracticeAreas] = useState<string[] | null>(null)
  const [practiceError, setPracticeError] = useState("")
  const [practiceSaving, setPracticeSaving] = useState(false)
  const [practiceReload, setPracticeReload] = useState(0)
  // Unknown until the matter list loads; the action stays disabled meanwhile.
  const [sampleExists, setSampleExists] = useState<boolean>()
  const [adding, setAdding] = useState(false)
  useEffect(() => {
    listJurisdictions().then(setJurisdictions)
    listMatters()
      .then((matters) => setSampleExists(matters.some((matter) => matter.sample)))
      .catch((error) => onError(humanizeError(error, "Could not check for the sample matter.")))
  }, [])
  useEffect(() => {
    let active = true
    setPracticeError("")
    getPreferences().then((prefs) => {
      if (active) {
        setPracticeAreas(prefs.practiceAreas ?? [])
        setLoadedPracticeAreas(prefs.practiceAreas ?? [])
      }
    }).catch(() => {
      if (active) setPracticeError("Could not load practice areas.")
    })
    return () => { active = false }
  }, [practiceReload])
  async function saveMatterSettings() {
    if (practiceAreas === null) return
    const names = practiceAreas
    if (names.length > 50 || names.some((name) => name.length > 100 || /[\p{Cc}\u2028\u2029]/u.test(name))) {
      setPracticeError("Enter up to 50 practice areas, each no longer than 100 characters.")
      return
    }
    const practiceEdited = JSON.stringify(names) !== JSON.stringify(loadedPracticeAreas)
    setPracticeSaving(true)
    onBusy(true)
    setPracticeError("")
    try {
      if (practiceEdited) {
        const saved = await savePreferences({ practiceAreas: names })
        setPracticeAreas(saved.practiceAreas)
        setLoadedPracticeAreas(saved.practiceAreas)
      }
      try {
        savePrefs({ defaultJurisdictions: selected }, { requireStorage: true })
        onSaved()
      } catch {
        setPracticeError(practiceEdited ? "Practice areas were saved, but jurisdiction defaults could not be saved in this browser. Your selections are still here. Try again." : "Jurisdiction defaults could not be saved in this browser. Your selections are still here. Try again.")
      }
    } catch {
      setPracticeError("Could not confirm that practice areas were saved. Your changes are still here and jurisdiction defaults were not saved. Try again.")
    } finally {
      setPracticeSaving(false)
      onBusy(false)
    }
  }
  // Include saved custom names and keep their spelling when they coincide with a suggestion.
  const practiceOptions = [...(practiceAreas ?? []), ...SUGGESTED_PRACTICE_AREAS]
    .filter((name, index, names) => names.findIndex((other) => other.toLocaleLowerCase() === name.toLocaleLowerCase()) === index)
    .map((name) => ({ code: name, name }))
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
          disabled={practiceSaving}
        />
      </FieldRow>
    </section>
    <section className="settings-section">
      <h3>Practice areas</h3>
      <p className="settings-hint muted">Set your practice areas to provide general context for assistants and library builders.</p>
      {practiceAreas !== null ? <>
        <FieldRow id="matter-practice-areas" label="Practice areas" hint="Select or add your own. Shared by everyone using this installation.">
          <JurisdictionSelect id="matter-practice-areas" aria-describedby="matter-practice-areas-hint" practiceAreas disabled={practiceSaving} jurisdictions={practiceOptions} selected={practiceAreas} onChange={setPracticeAreas} />
        </FieldRow>
      </> : practiceError ? <button type="button" onClick={() => setPracticeReload((value) => value + 1)}>Try again</button> : <p className="muted">Loading practice areas…</p>}
      {practiceAreas === null && practiceError && <p className="settings-error" role="alert">{practiceError}</p>}
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
    <SettingsFooter tab="matters">{practiceAreas !== null && practiceError && <p className="settings-error" role="alert">{practiceError}</p>}<div className="settings-actions"><button type="button" className="primary" disabled={practiceAreas === null || practiceSaving} onClick={() => void saveMatterSettings()}>{practiceSaving ? "Saving…" : "Save"}</button></div></SettingsFooter>
    </>
  )
}
