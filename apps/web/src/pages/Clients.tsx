import { useEffect, useRef, useState } from "react"
import { flushSync } from "react-dom"
import { Link, useSearchParams } from "react-router-dom"
import { listMatters, type Matter } from "../api/ingest/matters"
import { createClient, deleteClient, listClients, type StoredClient } from "../api/ingest/clients"
import { humanizeError } from "../api/ingest/errors"
import { clientKey, groupClients, type ClientGroup } from "../lib/clients"
import { formatDate } from "../lib/format"
import { useToast } from "../components/Toast"
import PageBar from "../components/PageBar"

// Every client: those added here (or from a matter's Client field) and every
// client named on a matter, with its matters. Laid out like Matters, with the
// add form on the right. A client named on matters is renamed or removed by
// editing the Client on those matters; an added client with no matters can be
// deleted here. ?client=<name> narrows the page to that client's matters.
export default function Clients() {
  const [matters, setMatters] = useState<Matter[]>([])
  const [stored, setStored] = useState<StoredClient[]>([])
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState<string>()
  const [name, setName] = useState("")
  const [reference, setReference] = useState("")
  const [nameMissing, setNameMissing] = useState(false)
  const [filter, setFilter] = useState("")
  const [busy, setBusy] = useState(false)
  const [creating, setCreating] = useState(false)
  const [confirmId, setConfirmId] = useState<number | null>(null)
  const nameInput = useRef<HTMLInputElement>(null)
  const [params] = useSearchParams()
  const toast = useToast()

  function load() {
    setLoading(true)
    setLoadError(undefined)
    Promise.all([listMatters(), listClients()])
      .then(([rows, clients]) => {
        setMatters(rows)
        setStored(clients)
      })
      .catch((e) => setLoadError(humanizeError(e)))
      .finally(() => setLoading(false))
  }

  useEffect(load, [])

  useEffect(() => {
    if (confirmId === null) return
    const dismiss = () => setConfirmId(null)
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmId])

  async function onCreate() {
    if (!name.trim()) {
      // Commit the error before focus lands so the field is announced with it.
      flushSync(() => setNameMissing(true))
      nameInput.current?.focus()
      return
    }
    setBusy(true)
    try {
      const { client, created } = await createClient(name.trim(), reference.trim() || undefined)
      setStored((rows) => [...rows.filter((row) => row.id !== client.id), client])
      setName("")
      setReference("")
      toast(created ? "success" : "info", created ? `Added client "${client.name}".` : `"${client.name}" is already a client.`)
    } catch (e) {
      toast("error", humanizeError(e, "Could not add the client."))
    } finally {
      setBusy(false)
    }
  }

  async function onDelete(client: StoredClient) {
    setConfirmId(null)
    try {
      await deleteClient(client.id)
      setStored((rows) => rows.filter((row) => row.id !== client.id))
      toast("success", `Deleted client "${client.name}".`)
    } catch (e) {
      toast("error", humanizeError(e, `Could not delete "${client.name}".`))
    }
  }

  // On narrow screens the add form is collapsed; open it before focusing.
  function focusCreateName() {
    flushSync(() => setCreating(true))
    nameInput.current?.focus()
    nameInput.current?.scrollIntoView({ block: "center" })
  }

  const clients = groupClients(matters, stored)
  // An empty or blank ?client= is no filter.
  const chosen = params.get("client")?.trim() || null
  const selected = chosen === null ? undefined : clients.find((group) => group.key === clientKey(chosen))
  const term = clientKey(filter)
  const visible = clients.filter((group) => !term || group.key.includes(term) || (group.stored?.reference ?? "").toLowerCase().includes(term))
  const matterCount = (count: number) => count === 0 ? "No matters yet" : `${count} ${count === 1 ? "matter" : "matters"}`

  const row = (group: ClientGroup) => {
    // Only an added client with no matters can go: matters keep their own Client.
    const deletable = group.matters.length === 0 ? group.stored : undefined
    return (
      <li key={group.key} className="list-row">
        <div className="list-row-main">
          <Link to={`/clients?${new URLSearchParams({ client: group.name })}`} className="list-row-title">
            {group.name}
          </Link>
          <span className="list-row-meta">
            {group.stored?.reference && <span className="matter-reference">{group.stored.reference}</span>}
            <span>{matterCount(group.matters.length)}</span>
          </span>
          {group.matters.length > 0 && (
            <ul className="client-matters" aria-label={`Matters for ${group.name}`}>
              {group.matters.map((m) => (
                <li key={m.id}><Link to={`/matter/${m.id}`}>{m.title}</Link></li>
              ))}
            </ul>
          )}
        </div>
        {deletable && (
          <div className="list-row-actions">
            {confirmId === deletable.id ? (
              <button
                className="icon-btn danger"
                title="Confirm delete"
                aria-label={`Confirm delete ${group.name}`}
                onClick={(e) => {
                  e.stopPropagation()
                  onDelete(deletable)
                }}
              >
                Confirm
              </button>
            ) : (
              <button
                className="icon-btn"
                title={`Delete ${group.name}`}
                aria-label={`Delete ${group.name}`}
                onClick={(e) => {
                  e.stopPropagation()
                  setConfirmId(deletable.id)
                }}
              >
                Delete
              </button>
            )}
          </div>
        )}
      </li>
    )
  }

  return (
    <>
      <PageBar title="Clients" description="Everyone you act for, added here or named on a matter." />
      <div className="library-layout matters-layout clients-layout">
        <section className="library-builder matter-create-panel" aria-labelledby="new-client-heading">
          <div className="library-heading matter-create-heading">
            <h2 id="new-client-heading">
              <span className="matter-create-label">Add a new client</span>
              <button
                type="button"
                className="matter-create-toggle"
                aria-expanded={creating}
                aria-controls="new-client-fields"
                onClick={() => setCreating((open) => !open)}
              >
                New client
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true">
                  <path d={creating ? "m6 15 6-6 6 6" : "m6 9 6 6 6-6"} />
                </svg>
              </button>
            </h2>
          </div>
          <form
            id="new-client-fields"
            className={`matter-create-row${creating ? " is-open" : ""}`}
            onSubmit={(event) => { event.preventDefault(); onCreate() }}
          >
            <p className="matter-create-intro">Add a client now and open their matters whenever you are ready.</p>
            <div className="matter-form-field matter-create-title">
              <label htmlFor="new-client-name">Client name <span className="field-required">Required</span></label>
              <input
                id="new-client-name"
                placeholder="e.g. Northwind Traders Ltd"
                aria-label="Client name"
                aria-required="true"
                aria-invalid={nameMissing || undefined}
                aria-describedby={nameMissing ? "new-client-name-error" : undefined}
                maxLength={200}
                ref={nameInput}
                value={name}
                onChange={(e) => { setName(e.target.value); setNameMissing(false) }}
              />
              {nameMissing && <small id="new-client-name-error" className="field-error" role="alert">Name the client to add it.</small>}
            </div>
            <label className="matter-form-field">
              <span>Client number <span className="field-optional">Optional</span></span>
              <input
                placeholder="Your reference"
                aria-label="Client number (optional)"
                value={reference}
                onChange={(e) => setReference(e.target.value)}
              />
            </label>
            <button type="submit" className={`primary btn-lg matter-create-submit${busy ? " btn-loading" : ""}`} disabled={busy} aria-busy={busy}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>
              Add client
            </button>
          </form>
        </section>
        <section className="library-collection" aria-labelledby="clients-heading">
          {chosen !== null && !loading && !loadError && (
            <Link to="/clients" className="clients-filter-back">← All clients</Link>
          )}
          <div className="library-heading">
            <h2 id="clients-heading">{selected ? selected.name : "Your clients"}</h2>
            {!loading && !loadError && (
              <span className="matter-count">{selected ? selected.matters.length : clients.length}</span>
            )}
          </div>
          {chosen === null && clients.length > 0 && (
            <div className="matters-list-toolbar">
              <label className="matter-search">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 4.5 4.5" /></svg>
                <input
                  type="search"
                  className="page-filter"
                  placeholder="Search by name or number"
                  aria-label="Filter clients"
                  value={filter}
                  onChange={(e) => setFilter(e.target.value)}
                />
              </label>
              <span className="matter-sort-label">A to Z</span>
            </div>
          )}
          {loading ? (
            <div className="skeleton-list">
              <div className="skeleton-row" />
              <div className="skeleton-row" />
            </div>
          ) : loadError ? (
            <div className="load-error" role="alert">
              <h3>Couldn&apos;t load your clients</h3>
              <p>{loadError}</p>
              <button type="button" onClick={load}>Try again</button>
            </div>
          ) : chosen !== null ? (
            !selected ? (
              <p className="muted">No client is named "{chosen}".</p>
            ) : selected.matters.length === 0 ? (
              <p className="muted">No matters for {selected.name} yet.</p>
            ) : (
              <ul className="matter-list matters-index-list" aria-label={`Matters for ${selected.name}`}>
                {selected.matters.map((m) => (
                  <li key={m.id} className="list-row">
                    <Link to={`/matter/${m.id}`} className="list-row-main">
                      <span className="list-row-title">{m.title}</span>
                      <span className="list-row-meta">
                        {m.reference && <span className="matter-reference">{m.reference}</span>}
                        <span className="list-row-date">Opened {formatDate(m.created_at)}</span>
                      </span>
                    </Link>
                  </li>
                ))}
              </ul>
            )
          ) : clients.length === 0 ? (
            <div className="empty-state">
              <svg className="empty-glyph" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true">
                <circle cx="12" cy="8" r="4" />
                <path d="M4 20a8 8 0 0 1 16 0" />
              </svg>
              <h3>No clients yet</h3>
              <p>Add the people and companies you act for here, or name a client on a matter.</p>
              <button type="button" onClick={focusCreateName}>Add a client</button>
            </div>
          ) : visible.length === 0 ? (
            <p className="muted">No clients match "{filter}".</p>
          ) : (
            <ul className="matter-list matters-index-list clients-index-list" aria-label="Clients">
              {visible.map(row)}
            </ul>
          )}
        </section>
      </div>
    </>
  )
}
