import { useId, useLayoutEffect, useRef, useState, type CSSProperties, type InputHTMLAttributes, type KeyboardEvent } from "react"
import { createPortal } from "react-dom"
import { suggestClients } from "../lib/clients"

type InputProps = Omit<InputHTMLAttributes<HTMLInputElement>, "value" | "onChange" | "role" | "list">

// The matter Client text field, offering clients already used on other matters
// as you type (an ARIA 1.2 list-autocomplete combobox). Nothing is preselected,
// so Enter on free text still means that text; ArrowDown then Enter, or a click,
// picks a suggestion. The caller's own keys (Enter to save, Escape to cancel)
// apply whenever the list is closed.
export default function ClientCombobox({
  value,
  onChange,
  clients,
  onKeyDown,
  ...input
}: InputProps & { value: string; onChange: (value: string) => void; clients: string[] }) {
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(-1)
  const [pos, setPos] = useState<CSSProperties>()
  const field = useRef<HTMLInputElement>(null)
  const listId = useId()
  const options = open ? suggestClients(clients, value) : []
  const shown = options.length > 0

  // A fixed list in the input's dialog (or the body) so a scrolling dialog or
  // header cannot clip it; offsets convert out of the appearance body zoom.
  useLayoutEffect(() => {
    if (!shown || !field.current) return
    const zoom = Number.parseFloat(getComputedStyle(document.body).zoom) || 1
    const rect = field.current.getBoundingClientRect()
    setPos({ left: rect.left / zoom, top: rect.bottom / zoom + 4, width: Math.max(rect.width / zoom, 200) })
    const close = (event: Event) => {
      if (!(event.target instanceof Node && document.getElementById(listId)?.contains(event.target))) setOpen(false)
    }
    window.addEventListener("scroll", close, true)
    window.addEventListener("resize", close)
    return () => {
      window.removeEventListener("scroll", close, true)
      window.removeEventListener("resize", close)
    }
  }, [shown, listId, value])

  function pick(name: string) {
    onChange(name)
    setOpen(false)
    setActive(-1)
  }

  function keyDown(event: KeyboardEvent<HTMLInputElement>) {
    // Keys that confirm an IME composition are not list navigation; the
    // caller's own keys behave as before.
    if (event.nativeEvent.isComposing) return onKeyDown?.(event)
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault()
      const list = shown ? options : suggestClients(clients, value)
      if (list.length === 0) return
      setOpen(true)
      if (!shown || event.altKey) return setActive(-1)
      setActive((index) => event.key === "ArrowDown" ? (index + 1) % list.length : index <= 0 ? list.length - 1 : index - 1)
      return
    }
    if (shown && event.key === "Enter" && active >= 0) {
      event.preventDefault()
      return pick(options[active]!)
    }
    if (shown && event.key === "Escape") {
      // Closing the list must not also cancel the edit or close the dialog.
      event.preventDefault()
      event.stopPropagation()
      setOpen(false)
      setActive(-1)
      return
    }
    if (event.key === "Tab") setOpen(false)
    onKeyDown?.(event)
  }

  const activeId = shown && active >= 0 ? `${listId}-${active}` : undefined
  return (
    <>
      <input
        {...input}
        ref={field}
        role="combobox"
        aria-autocomplete="list"
        aria-expanded={shown}
        aria-controls={shown ? listId : undefined}
        aria-activedescendant={activeId}
        autoComplete="off"
        value={value}
        onChange={(event) => {
          onChange(event.target.value)
          setOpen(true)
          setActive(-1)
        }}
        onBlur={(event) => {
          setOpen(false)
          input.onBlur?.(event)
        }}
        onKeyDown={keyDown}
      />
      {shown &&
        createPortal(
          <ul
            id={listId}
            role="listbox"
            aria-label="Existing clients"
            className="client-suggestions"
            style={pos}
            // Keep focus in the field, so a click or a scrollbar drag never
            // fires the blur that saves or closes the edit.
            onMouseDown={(event) => event.preventDefault()}
          >
            {options.map((name, index) => (
              <li
                key={name}
                id={`${listId}-${index}`}
                role="option"
                aria-selected={index === active}
                className={`client-suggestion${index === active ? " is-active" : ""}`}
                onClick={() => pick(name)}
              >
                {name}
              </li>
            ))}
          </ul>,
          field.current?.closest("dialog") ?? document.body,
        )}
    </>
  )
}
