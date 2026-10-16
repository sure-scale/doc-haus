import { readBrowserValue, writeBrowserValue } from "../browser-storage"
import { useEffect, useState, type ReactNode } from "react"
import { Link, NavLink, useMatch, useNavigate, useSearchParams } from "react-router-dom"
import { deleteMatterSession, getMatter } from "../api/ingest/matters"
import { listSessions } from "../api/opencode/sessions"
import { matterClient } from "../api/opencode/client"
import logo from "../assets/dochaus-logo.svg"
import { toggleDocumentsDrawer, useDocumentsDrawer } from "./matter/documents-drawer"
import { useToast } from "./Toast"
import { Tooltip } from "./Tooltip"

type Convo = { id: string; title: string; updated: number }

// Left rail. Global library navigation gives way to the active matter when one
// is open: Documents (the slide-out drawer), Review table, then the
// conversation list.
// Collapses to an icons-only strip; the choice persists per browser in
// localStorage.
export default function Sidebar({
  onOpenSettings,
  sessionsVersion,
  updateButton,
}: {
  onOpenSettings: () => void
  sessionsVersion: number
  updateButton?: ReactNode
}) {
  const [collapsed, setCollapsed] = useState(() => {
    const saved = readBrowserValue("dh.sidebar")
    return saved === null ? window.matchMedia("(max-width: 960px)").matches : saved === "1"
  })
  const matterId = useMatch("/matter/:id")?.params.id
  const [params] = useSearchParams()
  const navigate = useNavigate()
  const notify = useToast()
  const activeSession = params.get("session")
  const view = params.get("view") ?? "chat"
  // Undefined until the matter's conversations have loaded, so the list shows a
  // placeholder rather than a false "No conversations yet".
  const [convos, setConvos] = useState<Convo[]>()
  const [matterTitle, setMatterTitle] = useState("")
  const drawer = useDocumentsDrawer()
  const [matterDir, setMatterDir] = useState("")
  // The conversation whose delete is armed. Clicking the trash slides out an
  // inline Confirm pill instead of blocking on window.confirm; any click
  // elsewhere disarms it.
  const [confirmId, setConfirmId] = useState<string | null>(null)

  useEffect(() => {
    writeBrowserValue("dh.sidebar", collapsed ? "1" : "0")
  }, [collapsed])

  useEffect(() => {
    if (!confirmId) return
    const dismiss = () => setConfirmId(null)
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmId])

  useEffect(() => setConvos(undefined), [matterId])

  // When a matter is open, list its top-level chats (engine-persisted; we only
  // read them). Refetch when the active session changes so a new chat shows up.
  useEffect(() => {
    if (!matterId) return setConvos([])
    let live = true
    getMatter(matterId)
      .then((m) => {
        setMatterTitle(m.title)
        setMatterDir(m.dir)
        return listSessions(matterClient(m.dir))
      })
      .then((list) => {
        if (!live) return
        setConvos(
          list
            // Chats only: drop subagent runs (parentID), sessions the runtime
            // owns (drafting stages and the narrator carry metadata.dochaus_role
            // and have no chat to show), and our system-titled sessions —
            // workflow runs and legacy chats carry a "doc.haus" title; real chats
            // are titled from the user's first message. A drafting run's own
            // conversation stays: it shows the request and the run's narration.
            .filter((s) => !s.parentID && (draftingRun(s.metadata) || !runtimeOwned(s.metadata)) && !/^doc\.haus/i.test(s.title ?? ""))
            .map((s) => ({
              id: s.id,
              title: draftingRun(s.metadata) ? "Drafting run" : s.title?.trim() || "Untitled conversation",
              updated: s.time.updated,
            }))
            .sort((a, b) => b.updated - a.updated),
        )
      })
      // A matter that is gone or unreachable lists no conversations; the matter
      // page and the connection banner say why.
      .catch(() => live && setConvos([]))
    return () => {
      live = false
    }
  }, [matterId, activeSession, sessionsVersion])

  // Delete a conversation from the engine, then drop it from the list. If it was
  // the open one, fall back to a fresh chat so the canvas isn't left on a dead id.
  async function removeConvo(c: Convo) {
    if (!matterDir || !matterId) return
    setConfirmId(null)
    try {
      await deleteMatterSession(matterId, c.id)
      setConvos((list) => list?.filter((x) => x.id !== c.id))
      if (c.id === activeSession) navigate(`/matter/${matterId}?view=chat`)
      notify("success", "Conversation deleted")
    } catch (error) {
      notify("error", error instanceof Error ? error.message : "Could not delete conversation")
    }
  }

  return (
    <aside className={`sidebar${collapsed ? " collapsed" : ""}${matterId ? " sidebar-in-matter" : ""}`}>
      <div className="sidebar-brand">
        <Link to="/" title="doc.haus">
          <img src={logo} className="app-logo" alt="" />
          <span className="wordmark">Doc.Haus</span>
        </Link>
      </div>
      <button
        className="sidebar-icon-button sidebar-collapse"
        onClick={() => setCollapsed((c) => !c)}
        title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
        aria-label={collapsed ? "Expand sidebar" : "Collapse sidebar"}
        aria-expanded={!collapsed}
      >
        <IconChevron right={false} />
      </button>

      {matterId ? (
        <nav className="sidebar-nav sidebar-primary sidebar-back" aria-label="Main navigation">
          <Link to="/" className="nav-item" title="All matters" aria-label="All matters">
            <IconChevron right={false} />
            <span>All matters</span>
          </Link>
        </nav>
      ) : (
        <>
          <nav className="sidebar-nav sidebar-primary" aria-label="Main navigation">
            <NavLink
              to="/"
              end
              className={({ isActive }) => `nav-item${isActive ? " active" : ""}`}
              title="Matters"
              aria-label="Matters"
            >
              <IconMatters />
              <span>Matters</span>
            </NavLink>
          </nav>
          <nav className="sidebar-nav sidebar-library" aria-label="Library">
            <div className="sidebar-library-label">Library</div>
            <NavLink
              to="/templates"
              className={({ isActive }) => `nav-item${isActive ? " active" : ""}`}
              title="Templates"
              aria-label="Templates"
            >
              <IconTemplates />
              <span>Templates</span>
            </NavLink>
            <NavLink
              to="/playbooks"
              className={({ isActive }) => `nav-item${isActive ? " active" : ""}`}
              title="Playbooks"
              aria-label="Playbooks"
            >
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" aria-hidden="true">
                <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
              </svg>
              <span>Playbooks</span>
            </NavLink>
            <NavLink
              to="/workflows"
              className={({ isActive }) => `nav-item${isActive ? " active" : ""}`}
              title="Workflows"
              aria-label="Workflows"
            >
              <IconWorkflows />
              <span>Workflows</span>
            </NavLink>
            <NavLink
              to="/skills"
              className={({ isActive }) => `nav-item${isActive ? " active" : ""}`}
              title="Skills"
              aria-label="Skills"
            >
              <IconSkills />
              <span>Skills</span>
            </NavLink>
            <NavLink
              to="/agents"
              className={({ isActive }) => `nav-item${isActive ? " active" : ""}`}
              title="Agents"
              aria-label="Agents"
            >
              <IconAgents />
              <span>Agents</span>
            </NavLink>
          </nav>
        </>
      )}

      {matterId && matterTitle && (
        <div className="sidebar-matter-card">
          <div className="sidebar-matter" title={matterTitle}>{matterTitle}</div>
        </div>
      )}

      {matterId && (
        <nav className="sidebar-surfaces" aria-label="Matter">
          <button
            type="button"
            className={`nav-item${drawer.open ? " active" : ""}`}
            onClick={toggleDocumentsDrawer}
            title="Documents"
            aria-label="Documents"
            aria-expanded={drawer.open}
            aria-controls="documents-drawer"
          >
            <IconDocsNav />
            <span>Documents</span>
            {drawer.count !== undefined && <span className="sidebar-nav-count">{drawer.count}</span>}
          </button>
          <Link
            to={`/matter/${matterId}?${new URLSearchParams({ ...Object.fromEntries(params), view: "review" })}`}
            className={`nav-item${view === "review" ? " active" : ""}`}
            title="Review table"
            aria-label="Review table"
            aria-current={view === "review" ? "page" : undefined}
          >
            <IconReview />
            <span>Review table</span>
          </Link>
        </nav>
      )}

      {matterId && (
        <div className="sidebar-matter-sections">
          <section className="sidebar-conversations" aria-labelledby="sidebar-conversations-title">
            <div className="sidebar-section-head">
              <h2 id="sidebar-conversations-title">
                <Link to={`/matter/${matterId}?view=chat`} className={`sidebar-section-link${view === "chat" ? " active" : ""}`} aria-current={view === "chat" ? "page" : undefined}>
                  <IconChat />
                  <span>Conversations</span>
                </Link>
              </h2>
              <span className="sidebar-history-count">{convos?.length ?? ""}</span>
              <Link to={`/matter/${matterId}?view=chat`} className="sidebar-icon-button sidebar-new" title="New conversation" aria-label="New conversation">
                <IconPlus />
              </Link>
            </div>
            <div className="sidebar-history-body" tabIndex={0} role="region" aria-label="Conversations">
              {!convos ? (
                <div className="skeleton-list sidebar-skeleton" aria-label="Loading conversations">
                  <div className="skeleton-line" />
                  <div className="skeleton-line w-40" />
                </div>
              ) : convos.length === 0 ? (
                <p className="muted sidebar-empty">No conversations yet.</p>
              ) : (
                <ul className="convo-list">
                  {convos.map((c) => (
                    <li key={c.id} className="convo-row">
                      <Tooltip label={c.title}>
                        <Link
                          to={`/matter/${matterId}?session=${c.id}`}
                          className={`convo-item${c.id === activeSession ? " active" : ""}`}
                        >
                          {c.title}
                        </Link>
                      </Tooltip>
                      {confirmId === c.id ? (
                        <button
                          className="convo-confirm"
                          title="Confirm delete"
                          onClick={(e) => {
                            e.stopPropagation()
                            removeConvo(c)
                          }}
                        >
                          Confirm
                        </button>
                      ) : (
                        <button
                          className="convo-del"
                          title="Delete conversation"
                          onClick={(e) => {
                            e.stopPropagation()
                            setConfirmId(c.id)
                          }}
                        >
                          <IconTrash />
                        </button>
                      )}
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </section>
        </div>
      )}
      <div className="sidebar-footer">
        {updateButton}
        {matterId && (
          <Link to={`/matter/${matterId}?view=chat`} className="sidebar-icon-button sidebar-compact-new" title="New conversation" aria-label="New conversation">
            <IconPlus />
          </Link>
        )}
        <button className="nav-item sidebar-settings" onClick={onOpenSettings} title="Settings" aria-label="Settings">
          <IconSettings />
          <span>Settings</span>
        </button>
      </div>
    </aside>
  )
}

function runtimeOwned(metadata: unknown) {
  return typeof metadata === "object" && metadata !== null && "dochaus_role" in metadata
}

function draftingRun(metadata: unknown) {
  return typeof metadata === "object" && metadata !== null && "dochaus_role" in metadata && metadata.dochaus_role === "lean-drafting-owner"
}

function IconPlus() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" aria-hidden>
      <path d="M12 5v14M5 12h14" />
    </svg>
  )
}

function IconChat() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
    </svg>
  )
}

function IconReview() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="M3 9h18M3 15h18M9 3v18" />
    </svg>
  )
}

function IconDocsNav() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
    </svg>
  )
}

function IconMatters() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
    </svg>
  )
}

function IconTemplates() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="4" y="3" width="16" height="18" rx="2" />
      <path d="M8 7h8M8 11h8M8 15h5" />
    </svg>
  )
}

function IconWorkflows() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="5" cy="12" r="2" />
      <circle cx="19" cy="5" r="2" />
      <circle cx="19" cy="19" r="2" />
      <path d="M7 12h4l3-5" />
      <path d="M11 12l3 5" />
    </svg>
  )
}

function IconSkills() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
      <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
    </svg>
  )
}

function IconAgents() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="5" y="8" width="14" height="11" rx="2" />
      <path d="M12 8V5" />
      <circle cx="12" cy="4" r="1" />
      <path d="M9 13h.01M15 13h.01" />
      <path d="M5 12H3M21 12h-2" />
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

function IconTrash() {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <polyline points="3 6 5 6 21 6" />
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
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
