import { useEffect, useState } from "react"
import { Link, NavLink, useMatch, useSearchParams } from "react-router-dom"
import { getMatter } from "../api/ingest"
import { listSessions, matterClient } from "../api/opencode"
import logo from "../assets/dochaus-logo.svg"

type Convo = { id: string; title: string; updated: number }

// Left rail. Holds the brand, primary nav, the open matter's conversation
// history, and Settings. Collapses to an icons-only strip; the choice persists
// per browser in localStorage (a single UI preference — no datastore needed).
export default function Sidebar({ onOpenSettings }: { onOpenSettings: () => void }) {
  const [collapsed, setCollapsed] = useState(() => localStorage.getItem("dh.sidebar") === "1")
  const matterId = useMatch("/matter/:id")?.params.id
  const [params] = useSearchParams()
  const activeSession = params.get("session")
  const [convos, setConvos] = useState<Convo[]>([])

  useEffect(() => {
    localStorage.setItem("dh.sidebar", collapsed ? "1" : "0")
  }, [collapsed])

  // When a matter is open, list its top-level chats (engine-persisted; we only
  // read them). Refetch when the active session changes so a new chat shows up.
  useEffect(() => {
    if (!matterId) return setConvos([])
    let live = true
    getMatter(matterId)
      .then((m) => listSessions(matterClient(m.dir)))
      .then((list) => {
        if (!live) return
        setConvos(
          list
            .filter((s) => !s.parentID)
            .map((s) => ({ id: s.id, title: s.title.replace(/^doc\.haus\s+/, ""), updated: s.time.updated }))
            .sort((a, b) => b.updated - a.updated),
        )
      })
    return () => {
      live = false
    }
  }, [matterId, activeSession])

  return (
    <aside className={`sidebar${collapsed ? " collapsed" : ""}`}>
      <div className="sidebar-brand">
        <Link to="/" title="doc.haus">
          <img src={logo} className="app-logo" alt="" />
          {!collapsed && <span className="wordmark">Doc.Haus</span>}
        </Link>
      </div>

      <nav className="sidebar-nav">
        <NavLink to="/" end className={({ isActive }) => `nav-item${isActive ? " active" : ""}`} title="Matters">
          <IconMatters />
          {!collapsed && <span>Matters</span>}
        </NavLink>
      </nav>

      {matterId && !collapsed && (
        <div className="sidebar-convos">
          <div className="sidebar-section-head">
            <span>Conversations</span>
            <Link to={`/matter/${matterId}`} className="icon-btn" title="New chat">
              New
            </Link>
          </div>
          {convos.length === 0 ? (
            <p className="muted sidebar-empty">No conversations yet.</p>
          ) : (
            <ul className="convo-list">
              {convos.map((c) => (
                <li key={c.id}>
                  <Link
                    to={`/matter/${matterId}?session=${c.id}`}
                    className={`convo-item${c.id === activeSession ? " active" : ""}`}
                  >
                    {c.title}
                  </Link>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}

      <div className="sidebar-foot">
        <button className="nav-item" onClick={onOpenSettings} title="Settings">
          <IconSettings />
          {!collapsed && <span>Settings</span>}
        </button>
        <button
          className="nav-item"
          onClick={() => setCollapsed((c) => !c)}
          title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
        >
          <IconChevron right={collapsed} />
          {!collapsed && <span>Collapse</span>}
        </button>
      </div>
    </aside>
  )
}

function IconMatters() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
    </svg>
  )
}

function IconSettings() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
    </svg>
  )
}

function IconChevron({ right }: { right: boolean }) {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <polyline points={right ? "9 18 15 12 9 6" : "15 18 9 12 15 6"} />
    </svg>
  )
}
