import { useContext, useEffect, useRef, useState } from "react"
import { createPortal } from "react-dom"
import { Link } from "react-router-dom"
import { listTemplates, type Template } from "../../api/ingest/library"
import { humanizeError } from "../../api/ingest/errors"
import ModalHeader from "../ModalHeader"
import backdropClose from "../backdropClose"
import { Closing } from "../useExit"
import { templateFieldCounts, templateFieldSummary } from "../../lib/template-use"
import "./TemplateStartDialog.css"

// "Start from a template" inside a matter: pick a library template, then
// draft one document with the assistant or fill many from a spreadsheet. The
// matter is already known, so both continue into the same flows the
// Templates page uses, without asking for it. A template with nothing to
// fill goes straight to the one-document draft.
const SEARCH_FROM = 7
const stem = (name: string) => name.replace(/\.docx$/i, "")

export default function TemplateStartDialog({ onDraft, onBatch, onClose }: {
  onDraft: (template: Template) => void
  onBatch: (template: Template) => void
  onClose: () => void
}) {
  const closing = useContext(Closing)
  const [templates, setTemplates] = useState<Template[]>()
  const [error, setError] = useState<string>()
  const [query, setQuery] = useState("")
  const [chosen, setChosen] = useState<Template>()

  useEffect(() => {
    let live = true
    listTemplates()
      .then((res) => live && setTemplates([...res.templates].sort((a, b) => a.name.localeCompare(b.name))))
      .catch((cause) => live && setError(humanizeError(cause, "Could not load templates.")))
    return () => { live = false }
  }, [])

  // "Choose another template" unmounts its own button; put focus back in the
  // list so Escape and the keyboard still reach the dialog.
  const body = useRef<HTMLDivElement>(null)
  const returning = useRef(false)
  useEffect(() => {
    if (chosen || !returning.current) return
    returning.current = false
    body.current?.querySelector<HTMLElement>("input, button")?.focus()
  }, [chosen])

  function pick(template: Template) {
    if (templateFieldCounts(template).total === 0) onDraft(template)
    else setChosen(template)
  }

  const words = query.trim().toLowerCase().split(/\s+/).filter(Boolean)
  const shown = templates?.filter((t) => words.every((word) => `${t.name} ${t.description}`.toLowerCase().includes(word)))

  return createPortal(
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
      <div className="picker-panel template-start" onClick={(event) => event.stopPropagation()}>
        <ModalHeader title="Start from a template" onClose={onClose} />
        <div className="picker-body template-start-body" ref={body}>
          {chosen ? <>
            <div className="template-start-chosen">
              <span className="template-start-name">{stem(chosen.name)}</span>
              <button type="button" className="linklike" onClick={() => { returning.current = true; setChosen(undefined) }}>Choose another template</button>
            </div>
            <button type="button" className="assistant-option" autoFocus onClick={() => onDraft(chosen)}>
              <span className="assistant-name">One document</span>
              <span className="assistant-desc">The assistant drafts it, asking you for anything the matter does not cover.</span>
            </button>
            <button type="button" className="assistant-option" onClick={() => onBatch(chosen)}>
              <span className="assistant-name">Many from a spreadsheet</span>
              <span className="assistant-desc">One document for each row of an Excel or CSV file.</span>
            </button>
          </> : error ? <p className="template-batch-error" role="alert">{error}</p>
            : !templates ? null
            : templates.length === 0 ? <p className="muted">No templates yet. <Link to="/templates" onClick={onClose}>Add one on the Templates page</Link>.</p>
            : <>
              {templates.length >= SEARCH_FROM && <input className="template-start-search" type="search" aria-label="Find a template" autoComplete="off" placeholder="Find a template" value={query} onChange={(event) => setQuery(event.target.value)} />}
              {shown?.length === 0 && <p className="muted">No template matches “{query.trim()}”.</p>}
              {shown?.map((template) => <button key={template.name} type="button" className="assistant-option" onClick={() => pick(template)}>
                <span className="assistant-name">{stem(template.name)}</span>
                <span className="assistant-desc">{template.description || templateFieldSummary(template)}</span>
              </button>)}
            </>}
        </div>
      </div>
    </div>,
    document.body,
  )
}
