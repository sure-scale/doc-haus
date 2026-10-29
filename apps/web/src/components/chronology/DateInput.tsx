import { useEffect, useId, useRef, useState } from "react"
import { getPreferences } from "../../api/ingest/preferences"
import {
  dateInputPlaceholder,
  formatDateInput,
  maskDateInput,
  parseDateInput,
  type DateInputFormat,
  type DateInputPrecision,
} from "./date-input"

export function useChronologyDateFormat(): DateInputFormat {
  const [format, setFormat] = useState<DateInputFormat>("iso")
  useEffect(() => {
    let live = true
    getPreferences().then(
      (preferences) => {
        if (live && ["day-month-year", "month-day-year", "iso"].includes(preferences.dateFormat))
          setFormat(preferences.dateFormat)
      },
      () => {},
    )
    return () => {
      live = false
    }
  }, [])
  return format
}

export default function DateInput({
  label,
  value,
  onChange,
  format,
  precision = "day",
  required = false,
  disabled = false,
  onValidityChange,
}: {
  label: string
  value?: string
  onChange: (value: string | undefined) => void
  format: DateInputFormat
  precision?: DateInputPrecision
  required?: boolean
  disabled?: boolean
  onValidityChange?: (valid: boolean) => void
}) {
  const id = useId()
  const input = useRef<HTMLInputElement>(null)
  const calendar = useRef<HTMLInputElement>(null)
  const calendarAvailable =
    typeof HTMLInputElement !== "undefined" && typeof HTMLInputElement.prototype.showPicker === "function"
  const dirty = useRef(false)
  const previous = useRef({ value, precision })
  const [text, setText] = useState(() => formatDateInput(value, format, precision))
  const [inputFormat, setInputFormat] = useState(format)
  useEffect(() => {
    if (!dirty.current || value !== previous.current.value || precision !== previous.current.precision) {
      setText(formatDateInput(value, format, precision))
      setInputFormat(format)
      dirty.current = false
    }
    previous.current = { value, precision }
  }, [value, format, precision])
  const parsed = parseDateInput(text, inputFormat, precision)
  const valid = parsed !== undefined && (!required || Boolean(parsed))
  const message = valid ? "" : `Enter a complete valid date (${dateInputPlaceholder(inputFormat, precision)}).`
  useEffect(() => {
    input.current?.setCustomValidity(disabled ? "" : message)
    onValidityChange?.(valid)
  }, [message, valid, disabled, onValidityChange])
  return (
    <div>
      <label htmlFor={id}>{label}</label>
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <input
          ref={input}
          id={id}
          type="text"
          inputMode="numeric"
          autoComplete="off"
          placeholder={dateInputPlaceholder(inputFormat, precision)}
          value={text}
          required={required}
          disabled={disabled}
          aria-invalid={!valid || undefined}
          aria-describedby={!valid ? `${id}-error` : undefined}
          style={{ minWidth: 0, flex: 1 }}
          onBlur={() => {
            if (!valid) return
            dirty.current = false
            setInputFormat(format)
            setText(formatDateInput(parsed || undefined, format, precision))
          }}
          onChange={(event) => {
            const next = maskDateInput(event.target.value, inputFormat, precision)
            setText(next)
            dirty.current = true
            const result = parseDateInput(next, inputFormat, precision)
            if (result !== undefined && (result || !required)) onChange(result || undefined)
          }}
        />
        {precision !== "year" && calendarAvailable && (
          <span style={{ position: "relative", display: "inline-flex", flex: "none" }}>
            <button
              type="button"
              disabled={disabled}
              aria-label={`Choose ${label.toLowerCase()} from calendar`}
              onClick={() => {
                const field = calendar.current
                if (!field) return
                if (typeof field.showPicker === "function") field.showPicker()
                else field.focus()
              }}
            >
              <svg
                width="17"
                height="17"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.5"
                aria-hidden="true"
              >
                <rect x="4" y="5" width="16" height="16" rx="2" />
                <path d="M8 3v4m8-4v4M4 10h16" />
              </svg>
            </button>
            <input
              ref={calendar}
              type={precision === "month" ? "month" : "date"}
              tabIndex={-1}
              aria-label={`${label} calendar`}
              aria-hidden="true"
              disabled={disabled}
              value={value ?? ""}
              style={{
                position: "absolute",
                width: 1,
                height: 1,
                visibility: "hidden",
                pointerEvents: "none",
                padding: 0,
              }}
              onChange={(event) => {
                const next = event.target.value
                dirty.current = !next && required
                setInputFormat(format)
                setText(formatDateInput(next, format, precision))
                if (next || !required) onChange(next || undefined)
              }}
            />
          </span>
        )}
      </div>
      {!valid && (
        <p id={`${id}-error`} className="chronology-warning">
          {message}
        </p>
      )}
    </div>
  )
}
