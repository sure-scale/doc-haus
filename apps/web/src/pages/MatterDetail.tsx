import { useCallback, useEffect, useRef, useState } from "react"
import { useNavigate, useParams, useSearchParams } from "react-router-dom"
import type { CompareParams } from "../api/ingest/compare"
import { humanizeError } from "../api/ingest/errors"
import { getPreferences } from "../api/ingest/preferences"
import { getMatter, listJurisdictions, listPlaybooks, MatterNotFound, renameMatter, type Jurisdiction, type MatterDetail as Detail, type Playbook } from "../api/ingest/matters"
import { listWorkflows, type CustomWorkflow } from "../api/ingest/library"
import { createTemplateDraft } from "../api/ingest/template-drafts"
import { disposeInstance, listAgents } from "../api/opencode/sessions"
import { matterClient } from "../api/opencode/client"
import { AUTO, WORKFLOWS } from "../agents"
import DocumentsDrawer from "../components/matter/DocumentsDrawer"
import { setDocumentsCount, setDocumentsDrawer, useDocumentsDrawer } from "../components/matter/documents-drawer"
import DocumentViewer from "../components/DocumentViewer"
import ConflictsDialog from "../components/matter/ConflictsDialog"
import CompareDialog from "../components/CompareDialog"
import BlacklineViewer from "../components/BlacklineViewer"
import ChatPanel from "../components/ChatPanel"
import ReviewGrid from "../components/ReviewGrid"
import MatterHeader, { type MatterFields } from "../components/matter/MatterHeader"
import MatterRail from "../components/matter/MatterRail"
import useDraftingStatus from "../components/drafting/useDraftingStatus"
import useNeedsYouTitle from "../components/drafting/useNeedsYouTitle"
import { useToast } from "../components/Toast"
import useExit, { Exit } from "../components/useExit"
import NotFound from "./NotFound"

type Agent = { name: string; description?: string; mode?: string; hidden?: boolean }

// One matter conversation with a documents drawer and a review-table modal.
// The view query opens the review over the current conversation; closing it
// preserves that session. Workflows run as seeded conversations.
export default function MatterDetail({ onSessionsChanged }: { onSessionsChanged: () => void }) {
  const { id } = useParams<{ id: string }>()
  const [params, setParams] = useSearchParams()
  const navigate = useNavigate()
  // Documents is a drawer beside the rail, not a surface; the old
  // `view=documents` links open it over the chat.
  const view = params.get("view") === "review" ? "review" : "chat"
  const drawer = useDocumentsDrawer()
  const [drawerShown, drawerClosing] = useExit(drawer.open)
  // A drawer already open as the matter loads is restored, not opened: it
  // appears in place instead of sliding in.
  const [drawerRestored, setDrawerRestored] = useState(drawer.open)
  if (drawerRestored && !drawer.open) setDrawerRestored(false)
  // Files dropped on the composer, waiting for the drawer's upload list.
  const [incoming, setIncoming] = useState<File[]>()

  useEffect(() => {
    if (params.get("view") !== "documents") return
    setDocumentsDrawer(true)
    const next = new URLSearchParams(params)
    next.set("view", "chat")
    setParams(next, { replace: true })
  }, [params])
  const session = params.get("session") ?? undefined
  const [matter, setMatter] = useState<Detail>()
  // Why the matter could not be loaded: "missing" when ingest has no such matter
  // (deleted, or a mistyped link), otherwise the readable error.
  const [loadError, setLoadError] = useState<string>()
  const toast = useToast()
  // Playbooks and agents stay undefined until they first load, so the composer
  // can hold their controls' places rather than popping them in late.
  const [playbooks, setPlaybooks] = useState<Playbook[]>()
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [custom, setCustom] = useState<CustomWorkflow[]>([])
  const [agents, setAgents] = useState<Agent[]>()
  const [agent, setAgent] = useState(AUTO)
  const [viewing, setViewing] = useState<string>()
  const [creatingTemplate, setCreatingTemplate] = useState(false)
  const creatingTemplateRef = useRef(false)
  // The redline proposal to scroll to when the viewer opens, set when the user
  // follows a chat redline preview's "View in document" link (undefined otherwise).
  const [focusRedline, setFocusRedline] = useState<number>()
  // The Word file being compared with another version (the dialog), and the
  // comparison whose blackline is open. A compare started from the viewer
  // returns to it when the dialog is cancelled.
  const [recordingConflicts, setRecordingConflicts] = useState(false)
  const [attorney, setAttorney] = useState<string>()
  const [comparing, setComparing] = useState<{ name: string; fromViewer?: boolean }>()
  const [blackline, setBlackline] = useState<CompareParams>()
  // The change to show when the blackline opens, set when the user follows a
  // chat answer's "Show in blackline" link (undefined otherwise).
  const [focusChange, setFocusChange] = useState<string>()
  // The title-area fields are inline-click-editable: clicking one swaps its label
  // for an input (or, for jurisdictions, the shared multi-select) seeded with the
  // current value. Each commit persists the whole matter via the ingest service
  // and updates in place. Reference, client and title each track their own edit flag/draft.
  const [renaming, setRenaming] = useState(false)
  const [draftTitle, setDraftTitle] = useState("")
  const [editingRef, setEditingRef] = useState(false)
  const [draftRef, setDraftRef] = useState("")
  const [editingClient, setEditingClient] = useState(false)
  const [draftClient, setDraftClient] = useState("")
  // The session id the composer just minted. The first send flips the `session`
  // param, remounting ChatPanel under the new id; this lets that remount know it
  // is the same composer (not a rail reopen) so it keeps the cursor seated.
  const createdRef = useRef<string | undefined>(undefined)
  const drafting = useDraftingStatus(id)
  useNeedsYouTitle(drafting.status)
  // The chat's answer in progress, which the status bar shows while it runs.
  const [chat, setChat] = useState<{ activity?: string }>()
  const onChatBusy = useCallback((busy: boolean, activity?: string) => setChat(busy ? { activity } : undefined), [])
  const runStatus = drafting.status?.status

  useEffect(() => {
    listPlaybooks()
      .then(setPlaybooks)
      .catch(() => setPlaybooks([]))
    listJurisdictions().then(setJurisdictions)
  }, [])

  function refresh() {
    if (!id) return
    getMatter(id)
      .then((next) => {
        setMatter(next)
        setLoadError(undefined)
      })
      .catch((e) => setLoadError(e instanceof MatterNotFound ? "missing" : humanizeError(e, "Could not load this matter.")))
  }

  useEffect(() => {
    setLoadError(undefined)
    refresh()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id])

  // A run that completes registers its Word files in the matter folder; reload
  // the matter so they appear in Documents without a page refresh.
  useEffect(() => {
    if (runStatus === "completed") refresh()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [runStatus])

  // The rail's Documents badge follows this matter's live document list.
  const documentCount = matter?.documents.length
  useEffect(() => {
    setDocumentsCount(documentCount)
  }, [documentCount])
  useEffect(() => () => setDocumentsCount(undefined), [])

  useEffect(() => {
    if (!matter) return
    const dir = matter.dir
    let cancelled = false
    // Right after navigating back into a matter the engine can be briefly busy and
    // app.agents() rejects. The assistant and workflow pickers hide when they have
    // no agents to offer (see ModelSelector/WorkflowLauncher), so a single failed
    // call would silently drop the composer controls for the whole mount. Retry
    // until it answers rather than leaving them missing.
    async function load() {
      // Fetch custom workflows here so the self-heal check below has the registry.
      const wfRes = await listWorkflows().catch(() => ({ workflows: [] as CustomWorkflow[] }))
      if (cancelled) return
      setCustom(wfRes.workflows)
      const customNames = new Set(wfRes.workflows.map((w) => w.name))
      // A warm instance only rescans the agent dir after dispose; edits to an
      // existing workflow's .md don't change names, so they're picked up on the
      // instance's natural next rebuild. Only missing names need a dispose.
      let healed = false
      while (!cancelled) {
        const list = await listAgents(matterClient(dir)).catch(() => undefined)
        if (cancelled) return
        if (list) {
          const agentNames = new Set((list as Agent[]).map((a) => a.name))
          // If a workflow's agent is missing from the live instance, the instance
          // predates the workflow — dispose once so it rescans on next wake.
          if (!healed && [...customNames].some((n) => !agentNames.has(n))) {
            healed = true
            await disposeInstance(matterClient(dir))
            continue
          }
          setAgents(list as Agent[])
          // A new chat opens on Auto (a pseudo-assistant always offered, so no
          // availability check); an existing session restores the agent it last
          // used (see ChatPanel's load effect), so don't force Auto here. When a
          // surface deep-links in with an `agent` param (e.g. a library
          // template's Use action), honour it if it is a real assistant — this must
          // live in this effect or the unconditional setAgent(AUTO) clobbers it.
          if (!session) {
            const pinned = params.get("agent")
            setAgent(pinned && (list as Agent[]).some((a) => a.name === pinned) ? pinned : AUTO)
          }
          return
        }
        await new Promise((resolve) => setTimeout(resolve, 1000))
      }
    }
    load()
    return () => {
      cancelled = true
    }
    // Re-run when the deep-linked agent param changes (e.g. using a template
    // pins draft mode while already in the matter) so the pin is honoured even
    // though `matter` is unchanged.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matter, params.get("agent")])

  if (loadError === "missing")
    return <NotFound title="Matter not found" detail="This matter may have been deleted, or the link may be wrong." />

  if (!matter && loadError)
    return (
      <div className="not-found">
        <h1>Could not open this matter</h1>
        <p>{loadError}</p>
        <button type="button" className="primary" onClick={refresh}>
          Try again
        </button>
      </div>
    )

  if (!matter)
    return (
      <div className="skeleton-list">
        <div className="skeleton-line w-40" />
        <div className="skeleton-row" />
      </div>
    )

  const available = new Set((agents ?? []).map((a) => a.name))
  const workflows = [...WORKFLOWS, ...custom]

  // One persist path for every title-area field: merge the changed fields over the
  // current matter, write the whole record, and reflect what came back. Sessions
  // refresh too so the sidebar/picker pick up a renamed or re-referenced matter.
  async function patch(fields: MatterFields) {
    if (!matter) return
    const next = { ...matter, ...fields }
    // The client and the Posture go only with their own edit, so another field's save never reverts a newer one.
    const updated = await renameMatter(matter.id, next.title, next.reference, next.jurisdictions, next.playbook,
      "client" in fields ? next.client : undefined, "posture" in fields ? next.posture : undefined).catch((e) => {
      toast("error", humanizeError(e, "Could not save the matter. Try again."))
      return undefined
    })
    if (!updated) return
    setMatter((m) => m && { ...m, title: updated.title, reference: updated.reference, jurisdictions: updated.jurisdictions, playbook: updated.playbook, client: updated.client, posture: updated.posture })
    onSessionsChanged()
  }

  async function commitRename() {
    setRenaming(false)
    const next = draftTitle.trim()
    if (!matter || !next || next === matter.title) return
    await patch({ title: next })
  }

  async function commitRef() {
    setEditingRef(false)
    const next = draftRef.trim()
    if (!matter || next === (matter.reference ?? "")) return
    await patch({ reference: next || undefined })
  }

  // "" clears the client; undefined would keep the stored one.
  async function commitClient() {
    setEditingClient(false)
    const next = draftClient.trim()
    if (!matter || next === (matter.client ?? "")) return
    await patch({ client: next })
  }

  function changePlaybook(next?: string) {
    return patch({ playbook: next })
  }

  return (
    <>
      <MatterHeader
        matter={matter}
        jurisdictions={jurisdictions}
        editingRef={editingRef}
        setEditingRef={setEditingRef}
        draftRef={draftRef}
        setDraftRef={setDraftRef}
        commitRef={commitRef}
        editingClient={editingClient}
        setEditingClient={setEditingClient}
        draftClient={draftClient}
        setDraftClient={setDraftClient}
        commitClient={commitClient}
        renaming={renaming}
        setRenaming={setRenaming}
        draftTitle={draftTitle}
        setDraftTitle={setDraftTitle}
        commitRename={commitRename}
        patch={patch}
        onEditConflicts={() => {
          // The attorney name pre-fills "by"; the dialog opens without it if preferences cannot load.
          getPreferences().then((p) => setAttorney(p.attorney.trim() || undefined), () => {}).finally(() => setRecordingConflicts(true))
        }}
      />

      <div className="matter-body">
        <div className="workspace">
          <ChatPanel
            // A ?prompt= deep link (a workflow, template or review cell) prefills a fresh chat.
            key={session ?? params.get("prompt") ?? "new"}
            directory={matter.dir}
            matterId={matter.id}
            sessionID={session}
            created={createdRef.current === session}
            initialPrompt={params.get("prompt") ?? undefined}
            agent={agent}
            available={available}
            workflows={workflows}
            playbooks={playbooks ?? []}
            playbook={matter.playbook}
            toolsLoading={!agents || !playbooks}
            documents={matter.documents}
            onAddFiles={(files) => {
              setIncoming(files)
              setDocumentsDrawer(true)
            }}
            onPlaybookChange={changePlaybook}
            onAgentChange={setAgent}
            onSessionCreated={(sid) => {
              createdRef.current = sid
              setParams({ view: "chat", session: sid }, { replace: true })
            }}
            onSessionPersisted={(sid) => {
              const url = new URL(window.location.href)
              url.search = new URLSearchParams({ view: "chat", session: sid }).toString()
              window.history.replaceState(window.history.state, "", url)
            }}
            onSessionStarted={onSessionsChanged}
            onViewDocument={(name, redlineId) => {
              setFocusRedline(redlineId)
              setViewing(name)
            }}
            onOpenBlackline={(params, change) => {
              setFocusChange(change)
              setBlackline(params)
            }}
            onDocumentsChanged={refresh}
            onBusyChange={onChatBusy}
            feed={drafting}
          />
        </div>
      </div>

      <Exit when={view === "review"}>{() => (
        <ReviewGrid matterId={matter.id} title={matter.title} directory={matter.dir} documents={matter.documents} onClose={() => {
          const next = new URLSearchParams(params)
          next.delete("view")
          setParams(next, { replace: true })
        }} onAsk={(prompt) => setParams({ view: "chat", prompt })} />
      )}</Exit>

      {drawerShown && (
        <DocumentsDrawer
          closing={drawerClosing}
          restored={drawerRestored}
          matterId={matter.id}
          documents={matter.documents}
          drafting={drafting.status}
          incoming={incoming}
          onIncoming={() => setIncoming(undefined)}
          onUploaded={refresh}
          onView={(name) => {
            setFocusRedline(undefined)
            setViewing(name)
          }}
          onCompare={(name) => setComparing({ name })}
        />
      )}

      <MatterRail matter={matter} session={session} drafting={drafting.status} chat={chat} showStatusbar />

      <Exit when={viewing}>{(viewing) => (
        <DocumentViewer
          key={viewing}
          matterId={matter.id}
          name={viewing}
          focusId={focusRedline}
          onClose={() => setViewing(undefined)}
          onChanged={refresh}
          onConverted={(name) => setViewing(name)}
          onCompare={
            viewing.toLowerCase().endsWith(".docx")
              ? () => {
                  setComparing({ name: viewing, fromViewer: true })
                  setViewing(undefined)
                }
              : undefined
          }
          onSaveTemplate={
            viewing.toLowerCase().endsWith(".docx")
              ? async () => {
                  if (creatingTemplateRef.current) return
                  creatingTemplateRef.current = true
                  setCreatingTemplate(true)
                  try {
                    const draft = await createTemplateDraft(matter.id, viewing)
                    navigate(`/templates?draft=${encodeURIComponent(draft.id)}`)
                  } catch (error) {
                    toast("error", humanizeError(error, "Could not start Template Builder."))
                  } finally {
                    creatingTemplateRef.current = false
                    setCreatingTemplate(false)
                  }
                }
              : undefined
          }
          creatingTemplate={creatingTemplate}
        />
      )}</Exit>

      <Exit when={recordingConflicts}>{() => (
        <ConflictsDialog
          matterId={matter.id}
          current={matter.conflicts}
          attorney={attorney}
          onSaved={(saved) => setMatter((m) => m && { ...m, conflicts: saved.conflicts })}
          onClose={() => setRecordingConflicts(false)}
        />
      )}</Exit>

      <Exit when={comparing}>{(comparing) => (
        <CompareDialog
          matterId={matter.id}
          name={comparing.name}
          documents={matter.documents}
          onUploaded={refresh}
          onClose={() => {
            setComparing(undefined)
            if (comparing.fromViewer) setViewing(comparing.name)
          }}
          onCompare={(params) => {
            setComparing(undefined)
            setFocusChange(undefined)
            setBlackline(params)
          }}
        />
      )}</Exit>

      <Exit when={blackline}>{(blackline) => <BlacklineViewer matterId={matter.id} params={blackline} focus={focusChange} onClose={() => setBlackline(undefined)} />}</Exit>
    </>
  )
}
