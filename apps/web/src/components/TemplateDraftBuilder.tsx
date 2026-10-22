import { useCallback, useEffect, useRef, useState } from "react"
import { Link, useSearchParams } from "react-router-dom"
import { TEMPLATE_BUILDER } from "../agents"
import { humanizeError } from "../api/ingest/errors"
import { fetchTemplateDraftBytes, getTemplateDraft, publishTemplateDraft, type TemplateDraft } from "../api/ingest/template-drafts"
import { boundedRequest, matterClient } from "../api/opencode/client"
import { listSessions } from "../api/opencode/sessions"
import ChatPanel from "./ChatPanel"
import TemplateViewer from "./TemplateViewer"
import { useToast } from "./Toast"
import { Exit } from "./useExit"

// A source document belongs to one isolated template conversation. The saved
// library is written only by the explicit Save button in the rendered review.
export default function TemplateDraftBuilder({ id, onSaved, onDraftChanged }: { id: string; onSaved: () => Promise<unknown>; onDraftChanged: (draft: TemplateDraft) => void }) {
  const [params, setParams] = useSearchParams()
  const session = params.get("session") ?? undefined
  const [draft, setDraft] = useState<TemplateDraft>()
  const [error, setError] = useState("")
  const [sessionError, setSessionError] = useState("")
  const [checkedSession, setCheckedSession] = useState<string>()
  const [resumeDirectory, setResumeDirectory] = useState<string>()
  const [busy, setBusy] = useState(false)
  const [checkingDraft, setCheckingDraft] = useState(false)
  const draftChecks = useRef(0)
  const [saving, setSaving] = useState(false)
  const savingRef = useRef(false)
  const [saveError, setSaveError] = useState("")
  const [preview, setPreview] = useState<"source" | "proposal">()
  const [reviewedRevision, setReviewedRevision] = useState<number>()
  const [name, setName] = useState("")
  const [description, setDescription] = useState("")
  const editedName = useRef(false)
  const editedDescription = useRef(false)
  const createdRef = useRef<string | undefined>(undefined)
  const toast = useToast()

  const refresh = useCallback(async () => {
    draftChecks.current += 1
    setCheckingDraft(true)
    try {
      const next = await getTemplateDraft(id)
      setDraft((current) => current && (current.revision > next.revision || current.status === "published" && next.status !== "published") ? current : next)
      onDraftChanged(next)
      if (!editedName.current) setName(next.name)
      if (!editedDescription.current) setDescription(next.description)
      setError("")
      return next
    } catch (error) {
      setError(humanizeError(error, "Could not load this template conversation."))
    } finally {
      draftChecks.current -= 1
      setCheckingDraft(draftChecks.current > 0)
    }
  }, [id, onDraftChanged])

  useEffect(() => { void refresh() }, [refresh])
  useEffect(() => {
    if (!busy) return
    const timer = window.setInterval(() => void refresh(), 2000)
    return () => window.clearInterval(timer)
  }, [busy, refresh])

  // Session IDs are native engine identities. A pasted or stale query must not
  // attach this source to an existing matter or another template conversation.
  useEffect(() => {
    setSessionError("")
    setCheckedSession(undefined)
    if (!draft || !session || createdRef.current === session) return
    let live = true
    boundedRequest(matterClient(draft.dir).session.get({ sessionID: session }), 5000, "Template conversation check")
      .then((owner) => {
        if (!live) return
        if (owner.location?.directory !== draft.dir) throw new Error("This conversation belongs to another document. Start a fresh template conversation.")
        setCheckedSession(session)
      })
      .catch((error) => live && setSessionError(humanizeError(error, "Could not reopen this template conversation.")))
    return () => { live = false }
  }, [draft?.dir, session])

  // The library's conversation list links to a draft, not one browser's saved
  // URL. Discover its native chat before mounting a fresh composer.
  useEffect(() => {
    if (!draft || session || createdRef.current || resumeDirectory === draft.dir) return
    let live = true
    boundedRequest(listSessions(matterClient(draft.dir)), 5000, "Template conversations load")
      .then((sessions) => {
        if (!live) return
        const latest = sessions.filter((item) => !item.parentID).sort((a, b) => b.time.updated - a.time.updated)[0]
        if (latest) setParams({ draft: id, session: latest.id }, { replace: true })
        setResumeDirectory(draft.dir)
      })
      .catch((error) => live && setSessionError(humanizeError(error, "Could not find the saved template conversation.")))
    return () => { live = false }
  }, [draft?.dir, session, resumeDirectory, id, setParams])

  const onBusyChange = useCallback((working: boolean) => {
    setBusy(working)
    if (!working) void refresh()
  }, [refresh])
  const sourceBytes = useCallback(() => fetchTemplateDraftBytes(id, "source"), [id])
  const proposalBytes = useCallback(() => fetchTemplateDraftBytes(id, "proposal"), [id])

  async function save() {
    if (!draft || savingRef.current || busy || reviewedRevision !== draft.revision || draft.revision < 1 || draft.status === "published") return
    savingRef.current = true
    setSaving(true)
    setSaveError("")
    try {
      let saved: TemplateDraft
      try {
        saved = await publishTemplateDraft(id, { revision: draft.revision, name: name.trim(), description: description.trim() })
      } catch (error) {
        // Reconcile a lost response or a changed revision before retrying a
        // publication. A successful save stays successful if its reply was lost.
        const current = await refresh()
        if (current?.status === "published") saved = current
        else {
          setSaveError(humanizeError(error, "Could not save this template."))
          return
        }
      }
      setDraft(saved)
      setPreview(undefined)
      toast("success", `Saved template "${saved.publishedName ?? saved.name}".`)
      await onSaved().catch(() => toast("error", "Saved, but the library could not refresh. Reload to see the saved template."))
    } finally {
      savingRef.current = false
      setSaving(false)
    }
  }

  if (!draft) return <div className="template-draft-builder">
    {error ? <p className="msg-error" role="alert">{error} <button type="button" onClick={() => void refresh()}>Retry</button></p>
      : <p role="status">Attaching document…</p>}
  </div>

  const published = draft.status === "published"
  const ready = draft.revision > 0
  const reviewReady = ready && !busy && !checkingDraft && !error
  const openProposal = () => { setSaveError(""); setPreview("proposal") }
  const canOpenSession = session ? createdRef.current === session || checkedSession === session : resumeDirectory === draft.dir
  return <div className="template-draft-builder">
    <section className="template-source" aria-label="Attached source document">
      <div>
        <strong>{draft.source.name}</strong>
        <p>Source copy captured {new Date(draft.createdAt).toLocaleString()}. The original document stays unchanged.</p>
      </div>
      <button type="button" onClick={() => setPreview("source")}>View source</button>
    </section>
    {error && <p className="msg-error" role="alert">{error} <button type="button" onClick={() => void refresh()}>Retry</button></p>}
    <section className="template-review-status" aria-label="Template review">
      {published ? <>
        <p role="status">Saved to your library as <strong>{draft.publishedName ?? draft.name}</strong>.</p>
        <Link to="/templates">Back to saved templates</Link>
      </> : <>
        <p>{ready ? "Your proposed template is ready to review. Check the document and changes before saving it to the library." : "Describe the intended use and position. Template Builder will prepare a reusable copy for your review."}</p>
        <button type="button" className="primary" disabled={!reviewReady} onClick={openProposal}>Review proposed template</button>
      </>}
    </section>
    {sessionError ? <p className="msg-error" role="alert">{sessionError} <button type="button" onClick={() => {
      setSessionError("")
      createdRef.current = undefined
      setResumeDirectory(draft.dir)
      setParams({ draft: id }, { replace: true })
    }}>Start a fresh conversation</button></p>
      : canOpenSession ? <ChatPanel
        key={session ?? "new"}
        directory={draft.dir}
        sessionID={session}
        created={createdRef.current === session}
        initialPrompt={`Create a reusable template from the attached document "${draft.source.name}". Identify variable details and provisions that need my decision, preserve its wording and formatting, and prepare a proposed template for me to review before saving.`}
        agent={TEMPLATE_BUILDER.name}
        available={new Set([TEMPLATE_BUILDER.name])}
        pinned={TEMPLATE_BUILDER}
        title="Create reusable template"
        emptyHint="The attached source copy is the starting point. Tell Template Builder how the template will be used."
        starters={[]}
        composerPlaceholder="Describe the template you need…"
        onAgentChange={() => {}}
        onSessionPersisted={(sid) => {
          createdRef.current = sid
          const url = new URL(window.location.href)
          url.searchParams.set("session", sid)
          window.history.replaceState(window.history.state, "", url)
        }}
        onSessionCreated={(sid) => {
          createdRef.current = sid
          setParams({ draft: id, session: sid }, { replace: true })
        }}
        onViewDocument={() => setPreview("proposal")}
        onDocumentsChanged={refresh}
        onBusyChange={onBusyChange}
        transcriptFooter={reviewReady && !published ? <section className="template-proposal-card" aria-label="Proposed template ready">
          <div className="template-proposal-icon" aria-hidden="true">
            <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <path d="M14 2v6h6M8 13h8M8 17h5" />
            </svg>
          </div>
          <div className="template-proposal-copy">
            <p className="template-proposal-eyebrow">Ready for your review</p>
            <h3>{draft.name || "Proposed reusable template"}</h3>
            <p>Review the document and prepared changes before saving it to your library.</p>
            <button type="button" className="primary" disabled={!reviewReady} onClick={openProposal}>Review proposed template</button>
          </div>
        </section> : undefined}
      /> : <p role="status">Opening template conversation…</p>}

    <Exit when={preview}>{(view) => <TemplateViewer
      key={`${view}:${draft.revision}`}
      name={view === "source" ? draft.source.name : "Proposed reusable template"}
      placeholders={[]}
      loadBytes={view === "source" ? sourceBytes : proposalBytes}
      onRendered={view === "proposal" ? () => setReviewedRevision(draft.revision) : undefined}
      onClose={() => { if (!saving) setPreview(undefined) }}
      sidebar={view === "source" ? <>
        <div className="changes-head">Source document</div>
        <p>This is the exact source copy captured when you started this template conversation.</p>
        <p>Changes are prepared in a separate copy.</p>
      </> : <div className="template-draft-review">
        <div className="changes-head">Review proposed template</div>
        {draft.notes.length > 0 && <ul className="template-review-notes">{draft.notes.map((note, index) => <li key={index}>{note}</li>)}</ul>}
        <h3>Prepared changes</h3>
        <ol className="template-draft-changes">{draft.changes.map((change, index) => <li key={index}>
          <div className="redline-old">{change.before_quote || "(No previous text)"}</div>
          <div className="redline-new">{change.replacement || "(Removed)"}</div>
          {(change.paragraph_before || change.paragraph_after) && <details><summary>Paragraph context</summary>
            {change.paragraph_before && <p>{change.paragraph_before}</p>}
            {change.paragraph_after && <p>{change.paragraph_after}</p>}
          </details>}
        </li>)}</ol>
        {!draft.changes.length && <p>No text changes have been prepared.</p>}
      </div>}
      footer={view === "proposal" ? <section className="template-save-bar" aria-label="Template details and saving">
        <div className="template-save-fields">
          <label>Template name<input value={name} disabled={saving || published} onChange={(event) => { editedName.current = true; setName(event.target.value) }} /></label>
          <label>Description<textarea value={description} disabled={saving || published} rows={2} onChange={(event) => { editedDescription.current = true; setDescription(event.target.value) }} /></label>
        </div>
        {saveError && <p className="msg-error" role="alert">{saveError}</p>}
        {published ? <p role="status">Saved to your library.</p> : <div className="template-save-confirmation">
          <p>Check the document for any remaining client details and confirm that the wording suits its intended use.</p>
          <div className="template-review-actions">
            <button type="button" disabled={saving} onClick={() => setPreview(undefined)}>Keep editing</button>
            <button type="button" className="primary" disabled={saving || busy || !ready || reviewedRevision !== draft.revision || !name.trim() || !description.trim()} onClick={() => void save()}>{saving ? "Saving…" : "Save to library"}</button>
          </div>
        </div>}
      </section> : undefined}
    />}</Exit>
  </div>
}
