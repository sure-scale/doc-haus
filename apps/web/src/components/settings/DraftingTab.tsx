import { useEffect, useRef, useState } from "react"
import { getPreferences, POSTURES, savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { FieldRow, SegRow } from "./controls"
import { CheckRow, LegendField, PageLayoutPreview } from "./PageLayout"
import { ADVICE_LEGENDS, AGREEMENT_LEGENDS, DATE_FORMATS, NOTES_MARKING, draftLineText } from "./page-layout"

// Firm-wide drafting preferences. The engine renders the saved structured
// settings into standing instructions for each drafting or explanation turn;
// the page layout settings are read when a document renders instead.
export default function DraftingTab({ onSaved }: { onSaved: (text: string) => void }) {
  const [prefs, setPrefs] = useState<DraftingPreferences | null>(null)
  const [today] = useState(() => new Date())
  const [saving, setSaving] = useState(false)
  // Bumped on each save, so the legend choices start again from the text the
  // server kept (it trims and caps a legend; blank custom text is no legend).
  const [saves, setSaves] = useState(0)
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
    <>
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
            hint="How hard negotiated terms lean toward your client. A matter can set its own."
            value={prefs.posture}
            onChange={(posture) => set({ posture: posture as DraftingPreferences["posture"] })}
            options={POSTURES}
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
            options={DATE_FORMATS}
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
      </section>
      <section className="settings-section">
        <h3>Page layout</h3>
        <p className="settings-hint muted">
          Page numbers, the draft line and legends doc.haus adds to the documents it drafts. Changes apply to
          documents rendered after you save; uploaded, counterparty and template files keep their own.
        </p>
        <p className="settings-hint muted">
          Legends have legal effect, so they are your firm's choice: none is added unless you choose one, and each
          appears only on the documents it names.
        </p>
        <fieldset className="settings-drafting-fields" disabled={saving} aria-label="Page layout">
          <SegRow
            label="Page numbers"
            hint="In the footer of each page."
            value={prefs.pageNumbers}
            onChange={(pageNumbers) => set({ pageNumbers: pageNumbers as DraftingPreferences["pageNumbers"] })}
            options={[
              { value: "off", label: "Off" },
              { value: "plain", label: "2" },
              { value: "page", label: "Page 2" },
              { value: "dashed", label: "- 2 -" },
            ]}
          />
          {prefs.pageNumbers !== "off" && (
            <>
              <SegRow
                label="Page number position"
                value={prefs.pageNumberPosition}
                onChange={(pageNumberPosition) => set({ pageNumberPosition: pageNumberPosition as DraftingPreferences["pageNumberPosition"] })}
                options={[
                  { value: "center", label: "Center" },
                  { value: "right", label: "Right" },
                ]}
              />
              <CheckRow
                id="drafting-number-first-page"
                label="Number the first page"
                hint="Letters always start numbering on page 2."
                checked={prefs.numberFirstPage}
                onChange={(numberFirstPage) => set({ numberFirstPage })}
              />
            </>
          )}
          <SegRow
            label="Draft line"
            hint={`On agreements, advice and letters: "${draftLineText(prefs, today)}", dated when the draft is first made.`}
            value={prefs.draftLine ? "on" : "off"}
            onChange={(value) => set({ draftLine: value === "on" })}
            options={[
              { value: "off", label: "Off" },
              { value: "on", label: "On" },
            ]}
          />
          {prefs.draftLine && !!prefs.firm.trim() && (
            <CheckRow
              id="drafting-draft-line-firm"
              label="Include the firm name"
              checked={prefs.draftLineFirm}
              onChange={(draftLineFirm) => set({ draftLineFirm })}
            />
          )}
          <LegendField
            key={`agreement-${saves}`}
            id="drafting-agreement-legend"
            label="Agreement legend"
            hint="At the top of every page of an agreement."
            presets={AGREEMENT_LEGENDS}
            value={prefs.agreementLegend}
            onChange={(agreementLegend) => set({ agreementLegend })}
          />
          <LegendField
            key={`advice-${saves}`}
            id="drafting-advice-legend"
            label="Advice legend"
            hint="At the top of every page of advice, such as a memo, opinion or advice letter."
            presets={ADVICE_LEGENDS}
            value={prefs.adviceLegend}
            onChange={(adviceLegend) => set({ adviceLegend })}
          />
          <SegRow
            label="Drafting notes marking"
            hint={`"${NOTES_MARKING}" at the top of the internal drafting notes.`}
            value={prefs.notesMarking ? "on" : "off"}
            onChange={(value) => set({ notesMarking: value === "on" })}
            options={[
              { value: "off", label: "Off" },
              { value: "on", label: "On" },
            ]}
          />
          <div className="settings-form-row">
            <div className="settings-field-copy"><span className="settings-label">Preview</span></div>
            <div className="settings-field-control"><PageLayoutPreview prefs={prefs} /></div>
          </div>
        </fieldset>
      </section>
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
              setSaves(count => count + 1)
              onSaved("Drafting preferences saved. They apply to new replies, resumed drafting and documents rendered from now on.")
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
    </>
  )
}
