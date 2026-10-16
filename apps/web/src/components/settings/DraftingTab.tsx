import { useEffect, useRef, useState } from "react"
import { getPreferences, savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { FieldRow, SegRow } from "./controls"

// Firm-wide drafting preferences. The engine renders the saved structured
// settings into standing instructions for each drafting or explanation turn.
export default function DraftingTab({ onSaved }: { onSaved: (text: string) => void }) {
  const [prefs, setPrefs] = useState<DraftingPreferences | null>(null)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState("")
  const [reload, setReload] = useState(0)
  const mounted = useRef(false)
  useEffect(() => {
    mounted.current = true
    let active = true
    setError("")
    getPreferences().then((loaded) => {
      if (active) setPrefs(loaded)
    }).catch(() => {
      if (active) setError("Could not load drafting preferences.")
    })
    return () => {
      active = false
      mounted.current = false
    }
  }, [reload])
  if (!prefs) return (
    <section className="settings-section">
      <h3>Drafting preferences</h3>
      {error ? (
        <>
          <p className="settings-error" role="alert">{error}</p>
          <button type="button" onClick={() => setReload((value) => value + 1)}>Try again</button>
        </>
      ) : <p className="muted">Loading preferences...</p>}
    </section>
  )
  const set = (patch: Partial<DraftingPreferences>) => setPrefs({ ...prefs, ...patch })
  return (
    <section className="settings-section">
      <h3>Drafting preferences</h3>
      <p className="settings-hint muted">
        How doc.haus drafts, redlines, and explains its work across all matters. Changes apply to new replies
        and when drafting resumes or completed work is revisited.
      </p>
      <fieldset className="settings-drafting-fields" disabled={saving} aria-label="Drafting preferences">
        <FieldRow id="drafting-attorney" label="Attorney" hint="The lawyer identified as preparing documents.">
          <input
            id="drafting-attorney"
            aria-describedby="drafting-attorney-hint"
            placeholder="Jane Doe"
            value={prefs.attorney}
            onChange={(e) => set({ attorney: e.target.value })}
          />
        </FieldRow>
        <FieldRow id="drafting-firm" label="Firm" hint="The firm identified as preparing documents.">
          <input
            id="drafting-firm"
            aria-describedby="drafting-firm-hint"
            placeholder="Doe & Partners LLP"
            value={prefs.firm}
            onChange={(e) => set({ firm: e.target.value })}
          />
        </FieldRow>
        <SegRow
          label="Posture"
          hint="How hard negotiated terms lean toward your client."
          value={prefs.posture}
          onChange={(posture) => set({ posture: posture as DraftingPreferences["posture"] })}
          options={[
            { value: "client-favorable", label: "Client-favorable" },
            { value: "balanced", label: "Balanced" },
            { value: "conservative", label: "Conservative" },
          ]}
        />
        <SegRow
          label="Style"
          hint="Formal legal drafting, or plain language where precision allows."
          value={prefs.formality}
          onChange={(formality) => set({ formality: formality as DraftingPreferences["formality"] })}
          options={[
            { value: "formal", label: "Formal" },
            { value: "plain", label: "Plain language" },
          ]}
        />
        <SegRow
          label="Explanations"
          hint="How much reasoning the assistant shows for its suggestions."
          value={prefs.detail}
          onChange={(detail) => set({ detail: detail as DraftingPreferences["detail"] })}
          options={[
            { value: "concise", label: "Concise" },
            { value: "detailed", label: "Detailed" },
          ]}
        />
        <SegRow
          label="Dates"
          hint="The date format used in documents."
          value={prefs.dateFormat}
          onChange={(dateFormat) => set({ dateFormat: dateFormat as DraftingPreferences["dateFormat"] })}
          options={[
            { value: "month-day-year", label: "June 11, 2026" },
            { value: "day-month-year", label: "11 June 2026" },
            { value: "iso", label: "2026-06-11" },
          ]}
        />
        <SegRow
          label="Numbers"
          hint="Use plain numerals or spell out numbers followed by numerals in operative text."
          value={prefs.numberStyle}
          onChange={(numberStyle) => set({ numberStyle: numberStyle as DraftingPreferences["numberStyle"] })}
          options={[
            { value: "words-and-numerals", label: "thirty (30) days" },
            { value: "numerals", label: "30 days" },
          ]}
        />
        <FieldRow
          id="drafting-house-style"
          label="House style notes"
          hint="Additional standing instructions for drafting and editing across all matters."
        >
          <textarea
            id="drafting-house-style"
            aria-describedby="drafting-house-style-hint"
            className="settings-textarea"
            rows={4}
            placeholder={'Anything else the assistants should always honor, e.g. "Define parties as Customer and Provider, never Licensee/Licensor."'}
            value={prefs.houseStyle}
            onChange={(e) => set({ houseStyle: e.target.value })}
          />
        </FieldRow>
      </fieldset>
      {error && <p className="settings-error" role="alert">{error}</p>}
      <div className="settings-actions">
        <button
          className={`primary${saving ? " btn-loading" : ""}`}
          disabled={saving}
          onClick={async () => {
            setSaving(true)
            setError("")
            try {
              const saved = await savePreferences(prefs)
              if (!mounted.current) return
              setPrefs(saved)
              onSaved("Drafting preferences saved. They apply to new replies and resumed drafting.")
            } catch {
              if (mounted.current) setError("Could not save drafting preferences. Your changes are still here. Try again.")
            } finally {
              if (mounted.current) setSaving(false)
            }
          }}
        >
          Save
        </button>
      </div>
    </section>
  )
}
