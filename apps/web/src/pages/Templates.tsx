import { useCallback, useEffect, useRef, useState } from "react"
import { Link, useNavigate, useSearchParams } from "react-router-dom"
import {
  deleteTemplate,
  listTemplates,
  updateTemplateDescription,
  uploadTemplate,
  type Template,
} from "../api/ingest/library"
import type { Matter } from "../api/ingest/matters"
import { deleteTemplateDraft, listTemplateDrafts, type TemplateDraft } from "../api/ingest/template-drafts"
import { useToast } from "../components/Toast"
import Dropzone from "../components/Dropzone"
import { humanizeError } from "../api/ingest/errors"
import TemplateViewer from "../components/TemplateViewer"
import { markerKind } from "../components/document-markers"
import MatterPicker from "../components/MatterPicker"
import { Exit } from "../components/useExit"
import ChatPanel from "../components/ChatPanel"
import LibraryLayout from "../components/LibraryLayout"
import TemplateDraftBuilder from "../components/TemplateDraftBuilder"
import { TEMPLATE_BUILDER } from "../agents"

// The firm's global template library: the drafting bases shared across every
// matter. The library lists templates with an editable one-line description and a
// placeholder count, a View button that opens the template in a read-only viewer,
// an inline Confirm delete, and a .docx-only dropzone to add one. Beside it the
// Template Builder chat composes and maintains templates conversationally — it
// runs in the library directory (there is no matter to scope to), pinned to the
// template-builder agent. Templates are also created by draft mode; this
// page manages the uploaded and seeded ones directly.
export default function Templates() {
  const [params, setParams] = useSearchParams()
  const session = params.get("session") ?? undefined
  const draft = params.get("draft") ?? undefined
  const [dir, setDir] = useState<string>()
  const [templates, setTemplates] = useState<Template[]>([])
  const [drafts, setDrafts] = useState<TemplateDraft[]>([])
  const [loading, setLoading] = useState(true)
  const [busy, setBusy] = useState(false)
  const [confirmName, setConfirmName] = useState<string | null>(null)
  const [confirmDraft, setConfirmDraft] = useState<string>()
  const [deletingDraft, setDeletingDraft] = useState<string>()
  const [editing, setEditing] = useState<string | null>(null)
  const [viewing, setViewing] = useState<Template>()
  // "Use" flow: the template being assembled into a matter the lawyer picks.
  const [using, setUsing] = useState<Template>()
  const navigate = useNavigate()
  // The session id the composer just minted, so the remount that follows the first
  // send keeps the cursor seated (same pattern as MatterDetail).
  const createdRef = useRef<string | undefined>(undefined)
  const deletedDrafts = useRef(new Set<string>())
  const deletionPending = useRef(false)
  const toast = useToast()
  const onDraftChanged = useCallback((next: TemplateDraft) => {
    if (deletedDrafts.current.has(next.id)) return
    setDrafts((previous) => [next, ...previous.filter((item) => item.id !== next.id)].sort((a, b) => b.createdAt - a.createdAt))
  }, [])

  function refresh() {
    return Promise.all([
      listTemplates().then((res) => {
        setDir(res.dir)
        setTemplates(res.templates)
      }),
      listTemplateDrafts().then((next) => setDrafts(next.filter((item) => !deletedDrafts.current.has(item.id)))),
    ])
  }

  useEffect(() => {
    refresh()
      .catch((e) => toast("error", humanizeError(e)))
      .finally(() => setLoading(false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (!confirmName && !confirmDraft) return
    const dismiss = () => { setConfirmName(null); setConfirmDraft(undefined) }
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmName, confirmDraft])

  async function onFiles(files: File[]) {
    const docx = files.filter((f) => f.name.toLowerCase().endsWith(".docx"))
    if (!docx.length) {
      toast("error", "Templates must be .docx files.")
      return
    }
    setBusy(true)
    for (const file of docx) {
      try {
        const template = await uploadTemplate(file)
        setTemplates((prev) => [...prev.filter((t) => t.name !== template.name), template])
        toast("success", `Added template "${template.name}".`)
      } catch (e) {
        toast("error", humanizeError(e, `Could not add ${file.name}.`))
      }
    }
    setBusy(false)
  }

  async function onDelete(t: Template) {
    setConfirmName(null)
    try {
      await deleteTemplate(t.name)
      setTemplates((prev) => prev.filter((x) => x.name !== t.name))
      toast("success", `Deleted template "${t.name}".`)
    } catch (e) {
      toast("error", humanizeError(e, `Could not delete "${t.name}".`))
    }
  }

  async function onDeleteDraft(item: TemplateDraft) {
    if (deletionPending.current) return
    deletionPending.current = true
    setDeletingDraft(item.id)
    try {
      await deleteTemplateDraft(item.id)
      deletedDrafts.current.add(item.id)
      setDrafts((previous) => previous.filter((row) => row.id !== item.id))
      setConfirmDraft(undefined)
      const current = new URL(window.location.href)
      if (current.pathname === "/templates" && current.searchParams.get("draft") === item.id) {
        createdRef.current = undefined
        navigate("/templates", { replace: true })
      }
      toast("success", "Deleted template conversation.")
    } catch (error) {
      toast("error", humanizeError(error, "Could not delete this template conversation."))
    } finally {
      deletionPending.current = false
      setDeletingDraft(undefined)
    }
  }

  async function saveDescription(t: Template, description: string) {
    setEditing(null)
    if (description === t.description) return
    try {
      const updated = await updateTemplateDescription(t.name, description)
      setTemplates((prev) => prev.map((x) => (x.name === t.name ? { ...x, description: updated.description } : x)))
      toast("success", `Updated description for "${t.name}".`)
    } catch {
      toast("error", `Could not update description for "${t.name}".`)
    }
  }

  // Picking a matter lands in its chat with draft mode pinned and the assembly
  // prompt prefilled — the lawyer reviews and sends; nothing auto-runs.
  function useTemplate(t: Template, m: Matter) {
    setUsing(undefined)
    const prompt = `Draft a new document from the template "${t.name}". Gather what you can from this matter's documents and our conversation first, then interview me for the remaining terms and any optional clauses before assembling it.`
    navigate(`/matter/${m.id}?view=chat&agent=draft&prompt=${encodeURIComponent(prompt)}`)
  }

  const visible = [...templates].sort((a, b) => a.name.localeCompare(b.name))

  return (
    <>
      <LibraryLayout
        title="Templates"
        description="Drafting bases shared across every matter."
        libraryTitle="Saved templates"
        count={loading ? undefined : templates.length}
        active={Boolean(session || draft)}
        upload={
          <Dropzone
            accept=".docx"
            busy={busy}
            onFiles={onFiles}
            title={busy ? "Adding template…" : "Upload a template"}
            hint="Drop a Word file (.docx) or browse files."
          />
        }
        assistant={
          draft ? <TemplateDraftBuilder key={draft} id={draft} onSaved={refresh} onDraftChanged={onDraftChanged} /> : dir && (
            <ChatPanel
              key={session ?? "new"}
              directory={dir}
              sessionID={session}
              created={createdRef.current === session}
              agent={TEMPLATE_BUILDER.name}
              available={new Set([TEMPLATE_BUILDER.name])}
              pinned={TEMPLATE_BUILDER}
              title="Create or edit a template"
              emptyHint="Describe what you need, or ask about your saved templates."
              starters={[
                "What templates do we have?",
                "Create a template for a consulting agreement.",
                "Create a template for an employment offer letter.",
              ]}
              composerPlaceholder="e.g. Create a template for a mutual NDA"
              onAgentChange={() => {}}
              onSessionCreated={(sid) => {
                createdRef.current = sid
                setParams({ session: sid }, { replace: true })
              }}
              onViewDocument={() => {}}
              onDocumentsChanged={refresh}
            />
          )
        }
      >
        {loading ? (
          <div className="skeleton-list">
            <div className="skeleton-row" />
            <div className="skeleton-row" />
            <div className="skeleton-row" />
          </div>
        ) : templates.length === 0 ? (
          <p className="library-empty">Your saved templates will appear here.</p>
        ) : (
          <ul className="matter-list">
            {visible.map((t) => {
              // Count by the viewer's marker patterns: other bracketed text,
              // such as a [REDACTED] mark, is not a placeholder to fill.
              const markers = t.placeholders.filter((p) => markerKind(p.text) === "placeholder")
              const optional = markers.filter((p) => /^\[optional:/i.test(p.text)).length
              const fills = markers.length - optional
              return (
                <li
                  key={t.name}
                  className="list-row list-row-clickable"
                  title="View template"
                  onClick={() => setViewing(t)}
                >
                  <div className="list-row-main">
                    <span className="list-row-title"><button type="button" className="library-item-title" onClick={(e) => { e.stopPropagation(); setViewing(t) }}>{t.name}</button></span>
                    <span className="list-row-meta">
                      {editing === t.name ? (
                        <input
                          className="template-desc-input"
                          aria-label={`Description for ${t.name}`}
                          autoFocus
                          defaultValue={t.description}
                          placeholder="Add description"
                          onClick={(e) => e.stopPropagation()}
                          onBlur={(e) => saveDescription(t, e.target.value.trim())}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") e.currentTarget.blur()
                            if (e.key === "Escape") setEditing(null)
                          }}
                        />
                      ) : (
                        <button
                          type="button"
                          className="list-row-desc template-desc"
                          aria-label={`${t.description || "Add description"} — edit description for ${t.name}`}
                          title="Click to edit description"
                          onClick={(e) => {
                            e.stopPropagation()
                            setEditing(t.name)
                          }}
                        >
                          {t.description || "Add description"}
                        </button>
                      )}
                      <span className="list-row-date">
                        {fills} placeholder{fills === 1 ? "" : "s"}
                        {optional > 0 && `, ${optional} optional clause${optional === 1 ? "" : "s"}`}
                      </span>
                    </span>
                  </div>
                  <div className="list-row-actions">
                    <button
                      className="icon-btn is-primary"
                      title="Draft a document from this template into a matter"
                      onClick={(e) => {
                        e.stopPropagation()
                        setUsing(t)
                      }}
                    >
                      Use
                    </button>
                    <button
                      className="icon-btn"
                      title="View template"
                      onClick={(e) => {
                        e.stopPropagation()
                        setViewing(t)
                      }}
                    >
                      View
                    </button>
                    {confirmName === t.name ? (
                      <button
                        className="icon-btn danger"
                        title="Confirm delete"
                        onClick={(e) => {
                          e.stopPropagation()
                          onDelete(t)
                        }}
                      >
                        Confirm
                      </button>
                    ) : (
                      <button
                        className="icon-btn"
                        title="Delete template"
                        onClick={(e) => {
                          e.stopPropagation()
                          setConfirmDraft(undefined)
                          setConfirmName(t.name)
                        }}
                      >
                        Delete
                      </button>
                    )}
                  </div>
                </li>
              )
            })}
          </ul>
        )}
        {drafts.length > 0 && <section className="template-conversations" aria-label="Template conversations">
          <h2>Template conversations</h2>
          <ul>{drafts.map((item) => <li key={item.id}>
            <div className="template-conversation-row">
              <Link to={`/templates?draft=${encodeURIComponent(item.id)}`} aria-current={draft === item.id ? "page" : undefined}>
                <strong>{item.source.name}</strong>
                <span>{item.status === "published" ? `Saved as ${item.publishedName ?? item.name}` : item.revision > 0 ? "Ready for review" : "In progress"}</span>
              </Link>
              {confirmDraft === item.id ? <button type="button" className="icon-btn danger" title="Confirm delete" aria-label={`Confirm delete conversation for ${item.source.name}`} disabled={Boolean(deletingDraft)} onClick={(event) => { event.stopPropagation(); void onDeleteDraft(item) }}>{deletingDraft === item.id ? "Deleting…" : "Confirm"}</button>
                : <button type="button" className="icon-btn" aria-label={`Delete conversation for ${item.source.name}`} disabled={Boolean(deletingDraft)} onClick={(event) => { event.stopPropagation(); setConfirmName(null); setConfirmDraft(item.id) }}>Delete</button>}
            </div>
          </li>)}</ul>
        </section>}
      </LibraryLayout>

      <Exit when={viewing}>{(viewing) => (
        <TemplateViewer name={viewing.name} onClose={() => setViewing(undefined)} />
      )}</Exit>

      <Exit when={using}>{(using) => (
        <MatterPicker
          title={`Use "${using.name}" in a matter`}
          onPick={(m) => useTemplate(using, m)}
          onClose={() => setUsing(undefined)}
        />
      )}</Exit>
    </>
  )
}
