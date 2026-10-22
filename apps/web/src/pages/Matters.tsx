import { useEffect, useLayoutEffect, useRef, useState } from "react"
import { Link, useNavigate } from "react-router-dom"
import {
  createMatter,
  deleteMatter,
  listMatters,
  listJurisdictions,
  renameMatter,
  type Matter,
  type Jurisdiction,
} from "../api/ingest/matters"
import JurisdictionSelect from "../components/JurisdictionSelect"
import { useToast } from "../components/Toast"
import ModalHeader from "../components/ModalHeader"
import { humanizeError } from "../api/ingest/errors"
import { POSTURES, type Posture } from "../api/ingest/preferences"
import { formatDate } from "../lib/format"
import { loadPrefs } from "../prefs"

export default function Matters() {
  const [matters, setMatters] = useState<Matter[]>([])
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [loading, setLoading] = useState(true)
  // Set when the matter list could not be fetched, so an unreachable service is
  // never mistaken for an empty practice.
  const [loadError, setLoadError] = useState<string>()
  const [title, setTitle] = useState("")
  const [reference, setReference] = useState("")
  const [client, setClient] = useState("")
  // "" leaves the matter on the firm's Posture (Settings → Drafting).
  const [posture, setPosture] = useState<Posture | "">("")
  // New matters start from the firm's default jurisdictions (Settings → Matter
  // defaults) instead of empty.
  const [newJurisdictions, setNewJurisdictions] = useState<string[]>(() => loadPrefs().defaultJurisdictions)
  const [filter, setFilter] = useState("")
  const [busy, setBusy] = useState(false)
  const [creating, setCreating] = useState(false)
  const [confirmId, setConfirmId] = useState<string | null>(null)
  // The matter being edited in the modal, plus its draft fields. null when closed.
  const [editing, setEditing] = useState<Matter | null>(null)
  const [draftTitle, setDraftTitle] = useState("")
  const [draftReference, setDraftReference] = useState("")
  const [draftClient, setDraftClient] = useState("")
  const [draftPosture, setDraftPosture] = useState<Posture | "">("")
  const [draftJurisdictions, setDraftJurisdictions] = useState<string[]>([])
  const editDialog = useRef<HTMLDialogElement>(null)
  const toast = useToast()
  const navigate = useNavigate()

  function load() {
    setLoading(true)
    setLoadError(undefined)
    listMatters()
      .then(setMatters)
      .catch((e) => setLoadError(humanizeError(e)))
      .finally(() => setLoading(false))
  }

  useEffect(() => {
    load()
    listJurisdictions()
      .then(setJurisdictions)
      .catch(() => undefined)
  }, [])

  useEffect(() => {
    if (!confirmId) return
    const dismiss = () => setConfirmId(null)
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmId])

  useLayoutEffect(() => {
    if (!editing) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const dialog = editDialog.current
    dialog?.showModal()
    return () => {
      dialog?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [editing])

  async function onCreate() {
    if (!title.trim()) return
    setBusy(true)
    try {
      const matter = await createMatter(
        title.trim(),
        reference.trim() || undefined,
        newJurisdictions.length ? newJurisdictions : undefined,
        undefined,
        client.trim() || undefined,
        posture || undefined,
      )
      toast("success", `Created matter "${matter.title}".`)
      // A new matter is opened straight away: the next step is always adding its
      // documents or asking a question, and both happen inside it.
      navigate(`/matter/${matter.id}`)
    } catch (e) {
      toast("error", humanizeError(e, "Could not create the matter."))
      setBusy(false)
    }
  }

  async function onDelete(m: Matter) {
    setConfirmId(null)
    try {
      await deleteMatter(m.id)
      setMatters((prev) => prev.filter((x) => x.id !== m.id))
      toast("success", `Deleted matter "${m.title}".`)
    } catch (e) {
      toast("error", humanizeError(e, `Could not delete "${m.title}".`))
    }
  }

  function openEdit(m: Matter) {
    setEditing(m)
    setDraftTitle(m.title)
    setDraftReference(m.reference ?? "")
    setDraftClient(m.client ?? "")
    setDraftPosture(m.posture ?? "")
    setDraftJurisdictions(m.jurisdictions ?? [])
  }

  async function onSaveEdit() {
    if (!editing || !draftTitle.trim()) return
    setBusy(true)
    try {
      const updated = await renameMatter(
        editing.id,
        draftTitle.trim(),
        draftReference.trim() || undefined,
        draftJurisdictions.length ? draftJurisdictions : undefined,
        editing.playbook,
        // Sent only when edited, so a stale row never reverts the stored client or Posture.
        draftClient.trim() === (editing.client ?? "") ? undefined : draftClient.trim(),
        draftPosture === (editing.posture ?? "") ? undefined : draftPosture,
      )
      setMatters((prev) => prev.map((m) => (m.id === updated.id ? updated : m)))
      setEditing(null)
      toast("success", `Updated matter "${updated.title}".`)
    } catch (e) {
      toast("error", humanizeError(e, "Could not save the matter."))
    } finally {
      setBusy(false)
    }
  }

  const term = filter.trim().toLowerCase()
  const visible = [...matters]
    .sort((a, b) => b.created_at - a.created_at)
    .filter((m) => !term || m.title.toLowerCase().includes(term) || (m.reference ?? "").toLowerCase().includes(term))

  return (
    <>
      <header className="page-head library-head matters-head">
        <div>
          <span className="page-eyebrow">Your workspace</span>
          <h1>Matters</h1>
          <p className="page-sub">A clear home for every engagement, document, and conversation.</p>
        </div>
      </header>

      <div className="library-layout matters-layout">
        <section className="library-builder matter-create-panel" aria-labelledby="new-matter-heading">
          <div className="library-heading matter-create-heading">
            <h2 id="new-matter-heading">
              <span className="matter-create-label">Start a new matter</span>
              <button
                type="button"
                className="matter-create-toggle"
                aria-expanded={creating}
                aria-controls="new-matter-fields"
                onClick={() => setCreating((open) => !open)}
              >
                New matter
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
                  <path d={creating ? "m6 15 6-6 6 6" : "m6 9 6 6 6-6"} />
                </svg>
              </button>
            </h2>
          </div>
          <form
            id="new-matter-fields"
            className={`matter-create-row${creating ? " is-open" : ""}`}
            onSubmit={(event) => { event.preventDefault(); onCreate() }}
          >
            <p className="matter-create-intro">Add the essentials now. You can update the details later.</p>
            <label className="matter-form-field matter-create-title">
              <span>Matter title <span className="field-required">Required</span></span>
              <input
                placeholder="e.g. Acme MSA review"
                aria-label="Matter title"
                required
                value={title}
                onChange={(e) => setTitle(e.target.value)}
              />
            </label>
            <label className="matter-form-field matter-create-client">
              <span>Client <span className="field-optional">Optional</span></span>
              <input
                placeholder="Whom you act for"
                aria-label="Client you act for (optional)"
                maxLength={200}
                value={client}
                onChange={(e) => setClient(e.target.value)}
              />
            </label>
            <label className="matter-form-field matter-create-reference">
              <span>Matter number <span className="field-optional">Optional</span></span>
              <input
                placeholder="Your reference"
                aria-label="Matter number (optional)"
                value={reference}
                onChange={(e) => setReference(e.target.value)}
              />
            </label>
            <fieldset className="matter-create-preferences">
              <legend>Drafting preferences</legend>
              <label className="matter-form-field">
                <span>Posture</span>
                <select
                  aria-label="Posture for this matter"
                  aria-describedby="new-matter-posture-help"
                  value={posture}
                  onChange={(e) => setPosture(e.target.value as Posture | "")}
                >
                  <option value="">Firm default</option>
                  {POSTURES.map((row) => (
                    <option key={row.value} value={row.value}>{row.label}</option>
                  ))}
                </select>
                <small id="new-matter-posture-help">How firmly drafts favour your client.</small>
              </label>
              <div className="matter-form-field">
                <label htmlFor="new-matter-jurisdictions">Jurisdictions</label>
                <JurisdictionSelect
                  id="new-matter-jurisdictions"
                  jurisdictions={jurisdictions}
                  selected={newJurisdictions}
                  onChange={setNewJurisdictions}
                  align="right"
                />
              </div>
            </fieldset>
            <button type="submit" className={`primary matter-create-submit${busy ? " btn-loading" : ""}`} disabled={busy || !title.trim()} aria-busy={busy}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>
              Create matter
            </button>
          </form>
        </section>
        <section className="library-collection" aria-labelledby="your-matters-heading">
          <div className="library-heading">
            <h2 id="your-matters-heading">Your matters</h2>
            {!loading && !loadError && <span className="matter-count">{matters.length}</span>}
          </div>
          {matters.length > 0 && (
            <div className="matters-list-toolbar">
              <label className="matter-search">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 4.5 4.5" /></svg>
                <input
                  type="search"
                  className="page-filter"
                  placeholder="Search by title or number"
                  aria-label="Filter matters"
                  value={filter}
                  onChange={(e) => setFilter(e.target.value)}
                />
              </label>
              <span className="matter-sort-label">Newest first</span>
            </div>
          )}
          {loading ? (
            <div className="skeleton-list">
              <div className="skeleton-row" />
              <div className="skeleton-row" />
              <div className="skeleton-row" />
            </div>
          ) : loadError ? (
            <div className="load-error" role="alert">
              <h3>Couldn&apos;t load your matters</h3>
              <p>{loadError}</p>
              <button type="button" onClick={load}>
                Try again
              </button>
            </div>
          ) : matters.length === 0 ? (
            <div className="empty-state">
              <svg className="empty-glyph" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" aria-hidden="true">
                <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
              </svg>
              <h3>No matters yet</h3>
              <p>A matter holds an engagement&apos;s documents, reviews, and conversations. Create a new matter to get started.</p>
            </div>
          ) : visible.length === 0 ? (
            <p className="muted">No matters match "{filter}".</p>
          ) : (
            <ul className="matter-list matters-index-list">
              {visible.map((m) => (
                <li key={m.id} className="list-row">
                  <Link to={`/matter/${m.id}`} className="list-row-main">
                    <span className="list-row-title">
                      {m.title}
                    </span>
                    <span className="list-row-meta">
                      {m.client && <span className="matter-client">{m.client}</span>}
                      {m.reference && <span className="matter-reference">{m.reference}</span>}
                      <span className="list-row-date">Opened {formatDate(m.created_at)}</span>
                      {m.jurisdictions?.map((code) => (
                        <span key={code} className="matter-jurisdiction">
                          {jurisdictions.find((j) => j.code === code)?.name ?? code}
                        </span>
                      ))}
                    </span>
                  </Link>
                  <div className="list-row-actions">
                    <button className="icon-btn" title={`Edit ${m.title}`} aria-label={`Edit ${m.title}`} onClick={() => openEdit(m)}>
                      Edit
                    </button>
                    {confirmId === m.id ? (
                      <button
                        className="icon-btn danger"
                        title="Confirm delete"
                        aria-label={`Confirm delete ${m.title}`}
                        onClick={(e) => {
                          e.stopPropagation()
                          onDelete(m)
                        }}
                      >
                        Confirm
                      </button>
                    ) : (
                      <button
                        className="icon-btn"
                        title={`Delete ${m.title}`}
                        aria-label={`Delete ${m.title}`}
                        onClick={(e) => {
                          e.stopPropagation()
                          setConfirmId(m.id)
                        }}
                      >
                        Delete
                      </button>
                    )}
                  </div>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>

      {editing && (
        <dialog
          ref={editDialog}
          className="picker-panel matter-edit"
          aria-labelledby="edit-matter-heading"
          onCancel={(event) => { event.preventDefault(); setEditing(null) }}
          onClick={(event) => { if (event.target === event.currentTarget) setEditing(null) }}
        >
            <ModalHeader title={<span id="edit-matter-heading">Edit matter</span>} onClose={() => setEditing(null)} />
            <div className="matter-edit-body">
              <label className="matter-edit-field">
                <span>Title</span>
                <input
                  autoFocus
                  value={draftTitle}
                  onChange={(e) => setDraftTitle(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && onSaveEdit()}
                />
              </label>
              <label className="matter-edit-field">
                <span>Matter number</span>
                <input
                  value={draftReference}
                  onChange={(e) => setDraftReference(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && onSaveEdit()}
                />
              </label>
              <label className="matter-edit-field">
                <span>Client</span>
                <input
                  placeholder="Whom you act for (optional)"
                  maxLength={200}
                  value={draftClient}
                  onChange={(e) => setDraftClient(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && onSaveEdit()}
                />
              </label>
              <label className="matter-edit-field">
                <span>Posture</span>
                <select value={draftPosture} onChange={(e) => setDraftPosture(e.target.value as Posture | "")}>
                  <option value="">Firm default</option>
                  {POSTURES.map((row) => (
                    <option key={row.value} value={row.value}>
                      {row.label}
                    </option>
                  ))}
                </select>
              </label>
              <div className="matter-edit-field">
                <label htmlFor="edit-matter-jurisdictions">Jurisdictions</label>
                <JurisdictionSelect
                  id="edit-matter-jurisdictions"
                  jurisdictions={jurisdictions}
                  selected={draftJurisdictions}
                  onChange={setDraftJurisdictions}
                />
              </div>
            </div>
            <div className="viewer-bar matter-edit-foot">
              <button onClick={() => setEditing(null)}>Cancel</button>
              <button className={`primary${busy ? " btn-loading" : ""}`} onClick={onSaveEdit} disabled={busy || !draftTitle.trim()}>
                Save changes
              </button>
            </div>
        </dialog>
      )}
    </>
  )
}
