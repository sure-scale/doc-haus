import { useEffect, useState } from "react"
import type { DraftingPreferences } from "../../api/ingest/preferences"
import { LEGEND_MAX_LENGTH, agreementFurniture } from "./page-layout"
import "./PageLayout.css"

// One legend setting: none, a preset or the firm's own text. Radios rather than
// a segmented control: the presets are long, and a suggestion must stay visible.
export function LegendField({
  id,
  label,
  hint,
  presets,
  suggested = [],
  value,
  onChange,
  disabled,
}: {
  id: string
  label: string
  hint: string
  presets: readonly string[]
  suggested?: string[]
  value: string
  onChange: (value: string) => void
  disabled?: boolean
}) {
  const isCustom = (text: string) => text !== "" && !presets.includes(text)
  // Custom stays chosen while its text is empty or matches a preset, and its
  // text is kept while another choice is tried.
  const [custom, setCustom] = useState(() => isCustom(value))
  const [text, setText] = useState(() => (isCustom(value) ? value : ""))
  useEffect(() => {
    if (isCustom(value)) setText(value)
  }, [value])
  const choice = custom || isCustom(value) ? "custom" : value
  const options = [{ value: "", label: "None" }, ...presets.map(preset => ({ value: preset, label: preset })), { value: "custom", label: "Custom text" }]
  return (
    <div className="settings-form-row">
      <div className="settings-field-copy">
        <span id={`${id}-label`} className="settings-label">{label}</span>
      </div>
      <div className="settings-field-control">
        <div className="page-layout-choices" role="radiogroup" aria-labelledby={`${id}-label`} aria-describedby={`${id}-hint`}>
          {options.map(option => (
            <label key={option.value} className="page-layout-choice">
              <input
                type="radio"
                name={id}
                disabled={disabled}
                checked={choice === option.value}
                onChange={() => {
                  setCustom(option.value === "custom")
                  onChange(option.value === "custom" ? text : option.value)
                }}
              />
              <span>{option.label}</span>
              {suggested.includes(option.value) && <span className="page-layout-suggested">Suggested</span>}
            </label>
          ))}
        </div>
        {choice === "custom" && (
          <input
            id={`${id}-custom`}
            aria-label={`${label} text`}
            aria-describedby={`${id}-hint`}
            disabled={disabled}
            maxLength={LEGEND_MAX_LENGTH}
            placeholder="Your firm's legend"
            value={text}
            onChange={event => {
              const next = event.target.value.slice(0, LEGEND_MAX_LENGTH)
              setText(next)
              onChange(next)
            }}
          />
        )}
        <p id={`${id}-hint`} className="settings-hint muted">
          {[hint, choice === "custom" && `One line, up to ${LEGEND_MAX_LENGTH} characters (${text.length} used).`].filter(Boolean).join(" ")}
        </p>
      </div>
    </div>
  )
}

// A small picture of an agreement's first page: the header lines and the page
// number the renderer adds under these settings, dated today as an example.
export function PageLayoutPreview({ prefs }: { prefs: Pick<DraftingPreferences, "firm" | "dateFormat" | "pageNumbers" | "pageNumberPosition" | "draftLine" | "agreementLegend"> }) {
  const [today] = useState(() => new Date())
  const { header, pageNumber } = agreementFurniture(prefs, today)
  return (
    <figure className="page-layout-preview" aria-label="Agreement page preview">
      <div className="page-layout-sheet">
        <div className="page-layout-header" data-preview="header">
          {header.length ? header.map(line => <p key={line}>{line}</p>) : <p className="page-layout-empty">No header</p>}
        </div>
        <div className="page-layout-body" aria-hidden="true"><span /><span /><span /><span /></div>
        <div className={`page-layout-footer is-${pageNumber?.align ?? "center"}`} data-preview="footer">
          {pageNumber ? pageNumber.text : <span className="page-layout-empty">No page number</span>}
        </div>
      </div>
      <figcaption className="settings-hint muted">An agreement's first page. The draft line carries the date the draft is first made.</figcaption>
    </figure>
  )
}
