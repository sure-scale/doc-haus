import { useEffect, useId, useLayoutEffect, useRef, useState, type CSSProperties, type InputHTMLAttributes, type KeyboardEvent } from "react"
import { createPortal, flushSync } from "react-dom"
import { listClients, type StoredClient } from "../api/ingest/clients"
import { clientKey, suggestClients, withStoredClients } from "../lib/clients"

type InputProps = Omit<InputHTMLAttributes<HTMLInputElement>, "value" | "onChange" | "role" | "list">

// The matter Client text field, offering clients already used on other matters
// as you type (an ARIA 1.2 list-autocomplete combobox). Nothing is preselected,
// so Enter on free text still means that text; ArrowDown then Enter, or a click,
// picks a suggestion. Text that is no client yet gets a last option, Add "X" as
// a client, which puts that name in the field. Nothing is saved here: a new
// client exists once a saved matter names it. Clients added on the Clients page
// are suggested alongside the caller's names. Text left in the field (Tab or
// click away, or Enter to save) that is a spelling of a known client becomes
// that client's name, before the caller's own blur or Enter handler sees it.
// The caller's own keys (Enter to save, Escape to cancel) apply whenever the
// list is closed.
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
  const [stored, setStored] = useState<StoredClient[]>([])
  const field = useRef<HTMLInputElement>(null)
  const composing = useRef(false)
  // The caller's handlers from the latest render, so they run after a snap.
  const latest = useRef({ onBlur: input.onBlur, onKeyDown })
  latest.current = { onBlur: input.onBlur, onKeyDown }
  const listId = useId()
  const names = withStoredClients(clients, stored)
  const typed = value.replace(/\s+/g, " ").trim()
  const candidates = suggestClients(names, value)
  const canAdd = clientKey(typed) !== "" && !names.some((name) => clientKey(name) === clientKey(typed))
  const options = open ? candidates : []
  const addable = open && canAdd
  const count = options.length + (addable ? 1 : 0)
  const shown = count > 0

  useEffect(() => {
    let live = true
    listClients()
      .then((rows) => live && setStored(rows))
      // Suggestions are a convenience; typing still works without them.
      .catch(() => undefined)
    return () => {
      live = false
    }
  }, [])

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

  // A spelling of a known client becomes that client's name.
  function snap() {
    const key = clientKey(typed)
    if (!key || composing.current) return
    const known = names.find((name) => clientKey(name) === key)
    if (known && known !== value) flushSync(() => onChange(known))
  }

  function choose(index: number) {
    if (index < options.length) pick(options[index]!)
    else pick(typed)
  }

  function keyDown(event: KeyboardEvent<HTMLInputElement>) {
    // Keys that confirm an IME composition are not list navigation; the
    // caller's own keys behave as before.
    if (event.nativeEvent.isComposing) return latest.current.onKeyDown?.(event)
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault()
      const total = shown ? count : candidates.length + (canAdd ? 1 : 0)
      if (total === 0) return
      setOpen(true)
      if (!shown || event.altKey) return setActive(-1)
      setActive((index) => event.key === "ArrowDown" ? (index + 1) % total : index <= 0 ? total - 1 : index - 1)
      return
    }
    if (shown && event.key === "Enter" && active >= 0) {
      event.preventDefault()
      return choose(active)
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
    if (event.key === "Enter") snap()
    latest.current.onKeyDown?.(event)
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
        onCompositionStart={() => { composing.current = true }}
        onCompositionEnd={() => { composing.current = false }}
        onBlur={(event) => {
          setOpen(false)
          snap()
          latest.current.onBlur?.(event)
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
            {addable && (
              <li
                id={`${listId}-${options.length}`}
                role="option"
                aria-selected={active === options.length}
                className={`client-suggestion client-suggestion-add${active === options.length ? " is-active" : ""}`}
                onClick={() => pick(typed)}
              >
                Add &ldquo;{typed}&rdquo; as a client
              </li>
            )}
          </ul>,
          field.current?.closest("dialog") ?? document.body,
        )}
    </>
  )
}
