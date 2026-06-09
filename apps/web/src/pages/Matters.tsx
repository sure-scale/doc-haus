import { useEffect, useState } from "react"
import { Link } from "react-router-dom"
import { createMatter, deleteMatter, listMatters, type Matter } from "../api/ingest"
import { useToast } from "../components/Toast"

export default function Matters() {
  const [matters, setMatters] = useState<Matter[]>([])
  const [loading, setLoading] = useState(true)
  const [title, setTitle] = useState("")
  const [reference, setReference] = useState("")
  const [filter, setFilter] = useState("")
  const [busy, setBusy] = useState(false)
  const toast = useToast()

  useEffect(() => {
    listMatters()
      .then(setMatters)
      .finally(() => setLoading(false))
  }, [])

  async function onCreate() {
    if (!title.trim()) return
    setBusy(true)
    const matter = await createMatter(title.trim(), reference.trim() || undefined)
    setMatters((prev) => [...prev, matter])
    setTitle("")
    setReference("")
    setBusy(false)
    toast("success", `Created matter "${matter.title}".`)
  }

  async function onDelete(m: Matter) {
    if (!confirm(`Delete matter "${m.title}"? This removes its documents and cannot be undone.`)) return
    await deleteMatter(m.id)
    setMatters((prev) => prev.filter((x) => x.id !== m.id))
    toast("success", `Deleted matter "${m.title}".`)
  }

  const term = filter.trim().toLowerCase()
  const visible = [...matters]
    .sort((a, b) => b.created_at - a.created_at)
    .filter((m) => !term || m.title.toLowerCase().includes(term) || (m.reference ?? "").toLowerCase().includes(term))

  return (
    <>
      <div className="card">
        <h2>New matter</h2>
        <div className="row">
          <input
            placeholder="Matter title, e.g. Acme MSA review"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && onCreate()}
            style={{ flex: 1 }}
          />
          <input
            placeholder="Matter ID (optional)"
            value={reference}
            onChange={(e) => setReference(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && onCreate()}
            style={{ width: 160 }}
          />
          <button className="primary" onClick={onCreate} disabled={busy || !title.trim()}>
            Create matter
          </button>
        </div>
      </div>

      <div className="card">
        <div className="row" style={{ justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <h2 style={{ margin: 0 }}>Matters</h2>
          {matters.length > 0 && (
            <input
              placeholder="Filter by title or ID"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              style={{ width: 220 }}
            />
          )}
        </div>
        {loading ? (
          <p className="muted">Loading matters...</p>
        ) : matters.length === 0 ? (
          <p className="muted">No matters yet. Create one to begin.</p>
        ) : visible.length === 0 ? (
          <p className="muted">No matters match "{filter}".</p>
        ) : (
          <ul className="matter-list">
            {visible.map((m) => (
              <li key={m.id}>
                {m.reference && <span className="matter-ref">{m.reference}</span>}
                <Link to={`/matter/${m.id}`} style={{ flex: 1 }}>
                  {m.title}
                </Link>
                <span className="muted">{new Date(m.created_at).toLocaleDateString()}</span>
                <button className="icon-btn" title="Delete matter" onClick={() => onDelete(m)}>
                  Delete
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
    </>
  )
}
