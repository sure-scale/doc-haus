import { useId, useState, type ReactNode } from "react"

export function FieldRow({ id, label, labelAction, hint, children }: { id: string; label: string; labelAction?: ReactNode; hint?: string; children: ReactNode }) {
  return (
    <div className="settings-form-row">
      <div className="settings-field-copy">
        <label className="settings-label" htmlFor={id}>{label}</label>
        {labelAction}
      </div>
      <div className="settings-field-control">
        {children}
        {hint && <p id={`${id}-hint`} className="settings-hint muted">{hint}</p>}
      </div>
    </div>
  )
}

// A masked text input with a show/hide toggle, for API keys. Defaults to hidden;
// the toggle flips the input type so the key can be checked without re-typing.
export function PasswordInput({
  value,
  placeholder,
  disabled,
  onChange,
  id,
  "aria-describedby": describedBy,
}: {
  value: string
  placeholder?: string
  disabled?: boolean
  onChange: (value: string) => void
  id?: string
  "aria-describedby"?: string
}) {
  const [show, setShow] = useState(false)
  return (
    <div className="settings-password">
      <input
        id={id}
        aria-describedby={describedBy}
        type={show ? "text" : "password"}
        autoComplete="off"
        spellCheck={false}
        placeholder={placeholder}
        value={value}
        disabled={disabled}
        onChange={(e) => onChange(e.target.value)}
      />
      <button
        type="button"
        className="btn-sm ghost settings-reveal"
        disabled={disabled}
        aria-pressed={show}
        onClick={() => setShow((s) => !s)}
      >
        {show ? "Hide" : "Show"}
      </button>
    </div>
  )
}

export function ToggleRow({
  title,
  description,
  on,
  onChange,
}: {
  title: string
  description: string
  on: boolean
  onChange: (on: boolean) => void
}) {
  return (
    <div className="settings-conn settings-toggle-row">
      <div className="settings-toggle-text">
        <span className="settings-conn-name">{title}</span>
        <span className="muted">{description}</span>
      </div>
      <span className="switch-state muted">{on ? "Auto-approve" : "Ask first"}</span>
      <button
        role="switch"
        aria-checked={on}
        aria-label={title}
        className={`switch${on ? " on" : ""}`}
        onClick={() => onChange(!on)}
      />
    </div>
  )
}

// A labeled segmented control row — the pick-one primitive the preference tabs
// are built from.
export function SegRow({
  label,
  hint,
  value,
  options,
  onChange,
  disabled,
}: {
  label: string
  hint?: string
  value: string
  options: { value: string; label: string }[]
  onChange: (value: string) => void
  disabled?: boolean
}) {
  const id = useId()
  return (
    <div className="settings-form-row">
      <div className="settings-field-copy">
        <span id={`${id}-label`} className="settings-label">{label}</span>
      </div>
      <div className="settings-field-control">
        <div className="settings-seg" role="group" aria-labelledby={`${id}-label`} aria-describedby={hint ? `${id}-hint` : undefined}>
          {options.map((o) => (
            <button type="button" key={o.value} disabled={disabled} aria-pressed={value === o.value} className={value === o.value ? "active" : ""} onClick={() => onChange(o.value)}>
              {value === o.value && (
                <svg className="settings-seg-check" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <polyline points="20 6 9 17 4 12" />
                </svg>
              )}
              {o.label}
            </button>
          ))}
        </div>
        {hint && <p id={`${id}-hint`} className="settings-hint muted">{hint}</p>}
      </div>
    </div>
  )
}
