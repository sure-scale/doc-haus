import { useEffect, useState } from "react"
import { Link, useSearchParams } from "react-router-dom"
import { listMatters, type Matter } from "../api/ingest/matters"
import { humanizeError } from "../api/ingest/errors"
import { clientKey, groupClients } from "../lib/clients"
import { formatDate } from "../lib/format"
import PageBar from "../components/PageBar"

// Every client named on a matter, with its matters. Read-only: a client is
// added, renamed or removed by editing the Client on its matters.
// ?client=<name> narrows the page to that client's matters.
export default function Clients() {
  const [matters, setMatters] = useState<Matter[]>([])
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState<string>()
  const [params] = useSearchParams()

  function load() {
    setLoading(true)
    setLoadError(undefined)
    listMatters()
      .then(setMatters)
      .catch((e) => setLoadError(humanizeError(e)))
      .finally(() => setLoading(false))
  }

  useEffect(load, [])

  const clients = groupClients(matters)
  // An empty or blank ?client= is no filter.
  const chosen = params.get("client")?.trim() || null
  const selected = chosen === null ? undefined : clients.find((group) => group.key === clientKey(chosen))
  const matterCount = (count: number) => `${count} ${count === 1 ? "matter" : "matters"}`

  return (
    <>
      <PageBar title="Clients" description="Everyone you act for, gathered from the Client on each matter." />
      <div className="library-layout matters-layout clients-layout">
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
              <p className="muted">No matters name the client "{chosen}".</p>
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
              <p>Add a client on a matter.</p>
              <Link to="/">Go to matters</Link>
            </div>
          ) : (
            <ul className="matter-list matters-index-list clients-index-list" aria-label="Clients">
              {clients.map((group) => (
                <li key={group.key} className="list-row">
                  <div className="list-row-main">
                    <Link to={`/clients?${new URLSearchParams({ client: group.name })}`} className="list-row-title">
                      {group.name}
                    </Link>
                    <span className="list-row-meta">{matterCount(group.matters.length)}</span>
                    <ul className="client-matters" aria-label={`Matters for ${group.name}`}>
                      {group.matters.map((m) => (
                        <li key={m.id}><Link to={`/matter/${m.id}`}>{m.title}</Link></li>
                      ))}
                    </ul>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>
    </>
  )
}
