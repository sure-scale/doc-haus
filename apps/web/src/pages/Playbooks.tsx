import { useEffect, useRef, useState } from "react"
import { useNavigate, useSearchParams } from "react-router-dom"
import { deletePlaybook, listPlaybookLibrary, savePlaybook, uploadPlaybookSource } from "../api/ingest/library"
import { renameMatter, type Matter, type Playbook } from "../api/ingest/matters"
import { humanizeError } from "../api/ingest/errors"
import { PLAYBOOK_BUILDER } from "../agents"
import { useToast } from "../components/Toast"
import ChatPanel from "../components/ChatPanel"
import Dropzone from "../components/Dropzone"
import LibraryLayout from "../components/LibraryLayout"
import MatterPicker from "../components/MatterPicker"
import ModalHeader from "../components/ModalHeader"
import { DocMarkdown } from "../components/markdown/Markdown"
import "../styles/playbooks.css"

export default function Playbooks() {
  const [params, setParams] = useSearchParams()
  const session = params.get("session") ?? undefined
  const [dir, setDir] = useState<string>()
  const [playbooks, setPlaybooks] = useState<Playbook[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string>()
  const [busy, setBusy] = useState(false)
  const [assistantBusy, setAssistantBusy] = useState(false)
  const working = busy || assistantBusy
  const [confirmName, setConfirmName] = useState<string>()
  const [viewing, setViewing] = useState<Playbook>()
  const [editing, setEditing] = useState<{ name: string; description: string; content: string; create: boolean }>()
  const [using, setUsing] = useState<Playbook>()
  const [reimporting, setReimporting] = useState<Playbook>()
  const createdRef = useRef<string | undefined>(undefined)
  const fileRef = useRef<HTMLInputElement>(null)
  const navigate = useNavigate()
  const toast = useToast()

  async function refresh() {
    const res = await listPlaybookLibrary()
    setDir(res.dir)
    setPlaybooks(res.playbooks)
    setError(undefined)
  }

  useEffect(() => {
    refresh()
      .catch((e) => setError(humanizeError(e)))
      .finally(() => setLoading(false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (!confirmName) return
    const dismiss = () => setConfirmName(undefined)
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmName])

  async function onFiles(files: File[], target?: Playbook) {
    if (!files.length || working) return
    setBusy(true)
    try {
      const sources = []
      for (const file of files) sources.push(await uploadPlaybookSource(file))
      const prompt = `${target ? `Re-import the existing playbook ${JSON.stringify(target.name)} from` : "Import new firm playbooks from"} these uploaded sources: ${sources.map((s) => `${JSON.stringify(s.name)} at ${JSON.stringify(s.source)}`).join(", ")}. Read each complete source. Preserve approved clause text verbatim and do not invent firm positions.${target ? " Update that exact entry rather than creating a duplicate." : " List the existing library first; ask before replacing a matching entry."}`
      setParams({ prompt })
      toast("success", "Source uploaded. Review and send the import request in the playbook assistant.")
    } catch (e) {
      toast("error", humanizeError(e))
    } finally {
      setBusy(false)
      setReimporting(undefined)
    }
  }

  async function onDelete(p: Playbook) {
    if (working) return
    setConfirmName(undefined)
    setBusy(true)
    try {
      await deletePlaybook(p.name)
      setPlaybooks((prev) => prev.filter((entry) => entry.name !== p.name))
      toast("success", `Deleted "${p.name}".`)
    } catch (e) {
      toast("error", humanizeError(e))
    } finally {
      setBusy(false)
    }
  }

  async function onSave() {
    if (!editing || working) return
    setBusy(true)
    try {
      await savePlaybook(editing, editing.create)
      setEditing(undefined)
      await refresh()
      toast("success", "Saved playbook.")
    } catch (e) {
      toast("error", humanizeError(e))
    } finally {
      setBusy(false)
    }
  }

  async function usePlaybook(p: Playbook, m: Matter) {
    if (working) return
    setBusy(true)
    try {
      await renameMatter(m.id, m.title, m.reference, m.jurisdictions, p.name)
      setUsing(undefined)
      navigate(`/matter/${m.id}?view=chat`)
      toast("success", `Assigned "${p.name}" to "${m.title}".`)
    } catch (e) {
      toast("error", humanizeError(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <LibraryLayout
        title="Playbooks"
        description="Firm positions and approved clause text shared across matters."
        libraryTitle="Saved playbooks"
        count={loading || error ? undefined : playbooks.length}
        active={Boolean(session)}
        upload={
          <>
            <Dropzone
              accept=".docx,.md,.txt"
              busy={working || !dir}
              onFiles={onFiles}
              title={busy ? "Preparing playbook…" : "Import a playbook"}
              hint="Drop a Word or text file. The assistant prepares it for your approval."
            />
            <button
              className="icon-btn playbook-create"
              disabled={working || !dir}
              onClick={() => setEditing({ name: "playbook-", description: "", content: "", create: true })}
            >
              Create playbook
            </button>
          </>
        }
        assistant={
          dir && (
            <ChatPanel
              key={session ?? params.get("prompt") ?? "new"}
              directory={dir}
              sessionID={session}
              initialPrompt={params.get("prompt") ?? undefined}
              created={createdRef.current === session}
              agent={PLAYBOOK_BUILDER.name}
              available={new Set([PLAYBOOK_BUILDER.name])}
              pinned={PLAYBOOK_BUILDER}
              title="Create or edit a playbook"
              emptyHint="Import your firm's guide, or describe the positions you want to capture."
              starters={["What playbooks do we have?", "Help me create a playbook from our firm positions."]}
              composerPlaceholder="Describe your firm positions or an update to a playbook"
              onAgentChange={() => {}}
              onViewDocument={() => {}}
              onBusyChange={setAssistantBusy}
              onSessionCreated={(sid) => {
                createdRef.current = sid
                setParams({ session: sid }, { replace: true })
              }}
              onDocumentsChanged={() => {
                refresh().catch((e) => toast("error", humanizeError(e)))
              }}
            />
          )
        }
      >
        {loading ? (
          <div className="skeleton-list">
            <div className="skeleton-row" />
            <div className="skeleton-row" />
          </div>
        ) : error ? (
          <div role="alert">
            <p>{error}</p>
            <button className="icon-btn" onClick={() => refresh().catch((e) => setError(humanizeError(e)))}>
              Try again
            </button>
          </div>
        ) : playbooks.length === 0 ? (
          <p className="library-empty">Import a firm guide or create your first playbook.</p>
        ) : (
          <ul className="matter-list">
            {playbooks.map((p) => (
              <li key={p.name} className="list-row">
                <div className="list-row-main">
                  <span className="list-row-title">
                    {p.name.replace(/^playbook-/, "").replace(/-/g, " ")}
                    {p.builtin && <span className="badge badge-quiet">Built-in · read-only</span>}
                  </span>
                  <span className="list-row-meta">
                    <span className="list-row-desc">{p.description || "No description yet"}</span>
                    <span className="list-row-date">
                      {p.version && `Version ${p.version} · `}
                      {p.last_reviewed && `Reviewed ${p.last_reviewed} · `}
                      {p.matters.length} matter{p.matters.length === 1 ? "" : "s"}
                    </span>
                  </span>
                </div>
                <div className="list-row-actions">
                  <button className="icon-btn is-primary" disabled={working} onClick={() => setUsing(p)}>
                    Use in a matter
                  </button>
                  <button className="icon-btn" onClick={() => setViewing(p)}>
                    View
                  </button>
                  {!p.builtin && (
                    <>
                      <button
                        className="icon-btn"
                        disabled={working}
                        onClick={() => setEditing({ ...p, create: false })}
                      >
                        Edit
                      </button>
                      <button
                        className="icon-btn"
                        disabled={working}
                        onClick={() => {
                          setReimporting(p)
                          fileRef.current?.click()
                        }}
                      >
                        Re-import
                      </button>
                      <button
                        className={`icon-btn${confirmName === p.name ? " danger" : ""}`}
                        disabled={working || p.matters.length > 0}
                        title={
                          p.matters.length
                            ? `Unbind from these matters before deleting: ${p.matters.join(", ")}`
                            : "Delete playbook"
                        }
                        onClick={(e) => {
                          e.stopPropagation()
                          if (confirmName === p.name) onDelete(p)
                          else setConfirmName(p.name)
                        }}
                      >
                        {confirmName === p.name ? "Confirm delete" : "Delete"}
                      </button>
                    </>
                  )}
                </div>
              </li>
            ))}
          </ul>
        )}
      </LibraryLayout>
      <input
        ref={fileRef}
        type="file"
        accept=".docx,.md,.txt"
        hidden
        onChange={(e) => {
          const files = Array.from(e.target.files ?? [])
          e.target.value = ""
          onFiles(files, reimporting)
        }}
      />
      {viewing && (
        <div className="viewer-overlay" onClick={() => setViewing(undefined)}>
          <div
            className="picker-panel doc-panel"
            role="dialog"
            aria-modal="true"
            aria-label={viewing.name}
            onClick={(e) => e.stopPropagation()}
          >
            <ModalHeader title={viewing.name} onClose={() => setViewing(undefined)} />
            <div className="picker-body">
              <p className="muted">{viewing.description}</p>
              {viewing.matters.length > 0 && <p className="muted">Used in: {viewing.matters.join(", ")}</p>}
              <DocMarkdown>{viewing.content}</DocMarkdown>
            </div>
          </div>
        </div>
      )}
      {editing && (
        <div
          className="viewer-overlay"
          onClick={() => {
            if (!working) setEditing(undefined)
          }}
        >
          <div
            className="picker-panel doc-panel"
            role="dialog"
            aria-modal="true"
            aria-label={editing.create ? "Create playbook" : "Edit playbook"}
            onClick={(e) => e.stopPropagation()}
          >
            <ModalHeader
              title={editing.create ? "Create playbook" : `Edit ${editing.name}`}
              onClose={() => {
                if (!working) setEditing(undefined)
              }}
            />
            <form
              className="picker-body playbook-editor"
              onSubmit={(e) => {
                e.preventDefault()
                onSave()
              }}
            >
              <label>
                Name
                <input
                  required
                  pattern="playbook-[a-z0-9][a-z0-9-]*"
                  disabled={!editing.create || working}
                  value={editing.name}
                  onChange={(e) => setEditing({ ...editing, name: e.target.value })}
                />
              </label>
              <label>
                Description
                <input
                  value={editing.description}
                  disabled={working}
                  onChange={(e) => setEditing({ ...editing, description: e.target.value })}
                />
              </label>
              <label>
                Positions and approved clauses
                <textarea
                  required
                  rows={16}
                  value={editing.content}
                  disabled={working}
                  onChange={(e) => setEditing({ ...editing, content: e.target.value })}
                />
              </label>
              <p className="muted">
                Use a heading for each clause type, then Preferred, Fallbacks, Unacceptable and Rationale. Put approved
                clause text in an approved code block. Edits apply to every matter using this playbook.
              </p>
              <div className="playbook-editor-actions">
                <button type="button" className="icon-btn" disabled={working} onClick={() => setEditing(undefined)}>
                  Cancel
                </button>
                <button className="icon-btn is-primary" disabled={working || !editing.content.trim()}>
                  {busy ? "Saving…" : "Save playbook"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      {using && (
        <MatterPicker
          title={`Use "${using.name}" in a matter`}
          onPick={(m) => usePlaybook(using, m)}
          onClose={() => {
            if (!working) setUsing(undefined)
          }}
        />
      )}
    </>
  )
}
