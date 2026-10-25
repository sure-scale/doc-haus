import { draftingUpdateAdmission } from "../../../../dochaus/lib/drafting-update-admission"
import { furnitureLegend } from "../../../../dochaus/lib/furniture-legend"
import { Fragment, useEffect, useMemo, useRef, useState, type ReactNode } from "react"
import type { CompareParams } from "../api/ingest/compare"
import type { Document } from "../api/ingest/documents"
import type { LeanDraftingStatus } from "../api/ingest/lean"
import { matterClient, type Client } from "../api/opencode/client"
import { CLIENT_SUMMARY, EXPLICIT_ONLY, QUICK_DRAFT, isAuto, type AssistantMeta, type WorkflowMeta } from "../agents"
import { QUICK_BANNER, quickDraftAnswer } from "./chat/quick-draft"
import { getPreferences } from "../api/ingest/preferences"
import { useToast } from "./Toast"
import CitationView from "./CitationView"
import IssuesView, { parseIssues, stripIssues } from "./IssuesView"
import TodoView from "./TodoView"
import Markdown from "./markdown/Markdown"
import { isEmptyAnswer, readTurn, type VersionCompare } from "./chat/turn-model"
import { agentLabel, basename } from "./chat/step-labels"
import { IconPencil, IconRetry } from "./chat/icons"
import ChatSkeleton from "./chat/ChatSkeleton"
import StepsPanel from "./chat/StepsPanel"
import PermissionCard from "./chat/PermissionCard"
import RedlineView from "./chat/RedlineView"
import DraftView from "./chat/DraftView"
import CompareView from "./chat/CompareView"
import CompareAnswer, { chatCompare } from "./chat/CompareAnswer"
import Composer from "./chat/Composer"
import DocumentsDock from "./chat/DocumentsDock"
import { setDocumentsDrawer } from "./matter/documents-drawer"
import ChatEmpty from "./chat/ChatEmpty"
import UserText from "./chat/UserText"
import { useChatState } from "./chat/use-chat-state"
import { useTurnSettlement } from "./chat/use-turn-settlement"
import { useSessionEvents } from "./chat/use-session-events"
import { useChatActions } from "./chat/use-chat-actions"
import { useStickToBottom } from "./chat/use-stick-to-bottom"
import DraftingNarration from "./narration/DraftingNarration"
import { useDraftingNarration } from "./narration/use-drafting-narration"
import { cleanTitle, leanRunTakesUpdates, openQuestions, queuedNote, updateProgressLabel } from "./drafting/model"
import SessionQuestions from "./chat/SessionQuestions"
import QuestionDock from "./drafting/QuestionDock"
import RunHeader from "./drafting/RunHeader"
import useRunActions from "./drafting/useRunActions"
import type { DraftingFeed } from "./drafting/useDraftingStatus"

// Quiet starter prompts so a fresh matter is not a blank box and the user has
// ready questions about the documents in scope.
const STARTERS = [
  "What are the key points in these materials?",
  "What deadlines or required actions should I know about?",
  "Flag ambiguities, inconsistencies, or items needing review.",
]

export default function ChatPanel({
  directory,
  matterId,
  sessionID,
  created,
  initialPrompt,
  onPromptSeated,
  agent,
  available,
  pinned,
  title,
  emptyHint = "Ask a question about this matter's documents. Every answer cites the source section.",
  starters = STARTERS,
  composerPlaceholder = "e.g. What deadlines or required actions are here?",
  documents,
  onAddFiles,
  toolsLoading,
  workflows = [],
  playbooks,
  playbook,
  onPlaybookChange,
  onAgentChange,
  onSessionCreated,
  onSessionPersisted,
  onSessionStarted,
  onViewDocument,
  onOpenBlackline,
  onDocumentsChanged,
  onBusyChange,
  transcriptFooter,
  feed,
}: {
  directory: string
  // Present on the matter chat. It lets Draft start the matter's drafting run,
  // which this chat then shows; pinned library assistants omit it.
  matterId?: string
  sessionID?: string
  // This mount is the just-minted session, not a chat reopened from the rail.
  // The first send mints a session, which flips the URL and remounts the panel
  // under the new id (see onSessionCreated). The user was typing here a beat ago,
  // so seat the cursor like a fresh chat rather than treating it as a reopen.
  created?: boolean
  // A prompt to seat the composer with on a fresh chat (no session yet).
  // Prefilled, not auto-sent: the lawyer reviews and sends it.
  initialPrompt?: string
  // The surface clears its prompt link once seated. A prompt that arrives while
  // this chat is open is then seated here too, without a remount, when the
  // composer is empty; text the lawyer typed is never overwritten.
  onPromptSeated?: () => void
  agent: string
  available: Set<string>
  // A single fixed assistant for this surface (e.g. the Templates page's
  // template-builder). Drops the composer tools row entirely — no assistant
  // picker, no workflow launcher, nothing to choose when one agent owns the room.
  pinned?: AssistantMeta
  // Surface copy — defaults read for a matter chat, which has no heading of its
  // own (the page head names the matter); other surfaces (the template library)
  // pass their heading, empty-state hint, starters, and placeholder.
  title?: string
  emptyHint?: string
  starters?: string[]
  composerPlaceholder?: string
  // The matter's documents. The composer gets a Documents control (the matter's
  // drawer) and takes dropped files, handed to onAddFiles. None (zero) swaps the
  // starters for an upload-first prompt; surfaces without documents (the
  // library builders) omit both.
  documents?: Document[]
  onAddFiles?: (files: File[]) => void
  // The assistant, workflow and playbook choices are still loading; the
  // composer holds their places so nothing shifts when they arrive.
  toolsLoading?: boolean
  // The full launchable workflow list (built-in + custom). Passed to WorkflowLauncher.
  // Surfaces that don't offer workflows omit it; the launcher's own options guard hides it.
  workflows?: WorkflowMeta[]
  // The firm playbooks available to bind to this matter, plus the currently-bound
  // one and its setter — surfaced as a composer control alongside the assistant
  // and workflow pickers. All optional so a fixed surface (the template library,
  // which passes `pinned`) renders no playbook control; the selector shows only
  // when onPlaybookChange is provided.
  playbooks?: { name: string; description: string }[]
  playbook?: string
  onPlaybookChange?: (name?: string) => void
  onAgentChange: (name: string) => void
  // Open the redline viewer on a document, optionally scrolled to a specific
  // proposal — fired by the in-chat redline preview's "View in document" link.
  onViewDocument: (name: string, redlineId?: number) => void
  // Open the blackline of a comparison the compare-versions tool ran,
  // optionally at one of its changes (a C- or U-id).
  onOpenBlackline?: (params: CompareParams, change?: string) => void
  // Called once the first message lazily mints a session, so the parent can put
  // its id in the URL — a reload then restores the conversation instead of a
  // blank chat. Fired at the first safe idle boundary, before any automatic
  // checkpoint recovery starts another provider turn.
  onSessionCreated?: (id: string) => void
  // Persist the newly minted session identity in the address bar without
  // remounting this component, so a reload during the first provider turn can
  // still adopt the in-flight session.
  onSessionPersisted?: (id: string) => void
  // Fired the instant a session is minted, so the conversation rail can list the
  // new chat immediately rather than waiting for the turn to settle. Unlike
  // onSessionCreated this must not change the URL — that would remount the panel
  // and drop the in-flight stream — it only signals a re-list.
  onSessionStarted?: () => void
  // Fired when a tool changed the matter's document set mid-turn (a draft was
  // created), so the documents rail refreshes without waiting for a reload.
  onDocumentsChanged?: () => void
  // Whether a turn is working, and what it is doing now when that has a plain
  // name; the matter page shows it in its status bar.
  onBusyChange?: (busy: boolean, activity?: string) => void
  // A surface-owned result belongs in the scrolling conversation after its
  // settled turns. Keep it hidden while native history or a turn is in flight.
  transcriptFooter?: ReactNode
  // The matter page's poll of its drafting run, shared with the documents rail;
  // the run header's actions write their result back to it.
  feed?: DraftingFeed
}) {
  const client = useMemo<Client>(() => matterClient(directory), [directory])
  const matterName = basename(directory)

  const pendingFormsRef = useRef(0)
  const formsRefreshRef = useRef<() => void>(() => {})
  const state = useChatState(sessionID)
  const settlement = useTurnSettlement({ state, client, matterName, matterId, directory, onSessionCreated })
  // A drafting run started here: its narration answers under the request turn.
  const [drafting, setDrafting] = useState<{ owner: string; turn: number }>()
  const narration = useDraftingNarration({ client, matterId, matterName, startedOwner: drafting?.owner, sessionID })
  const run = narration.visible ? latest(feed?.status, narration.status) : undefined
  const actions = useRunActions(matterId, run, (next) => feed?.setStatus(next))
  useSessionEvents({
    state,
    pendingFormsRef,
    formsRefreshRef,
    settlement,
    client,
    sessionID,
    matterId,
    matterName,
    agent,
    onAgentChange,
    directory,
    onDocumentsChanged,
    onForeignPart: narration.onForeignPart,
    onForeignSession: narration.acceptsSession,
  })
  // The request a pending quick-draft offer holds: nothing has run for it yet.
  const [pendingRoute, setPendingRoute] = useState<string>()
  const {
    onSend,
    onStop,
    runWorkflow,
    resendFrom,
    onPermissionReply,
    chooseDraft,
    dismissDraftOffer,
    upgradeQuickDraft,
    quickLimit,
    dismissQuickLimit,
    runFullInstead,
  } = useChatActions({
    state,
    settlement,
    client,
    agent,
    matterId,
    directory,
    workflows,
    onSessionPersisted,
    onSessionStarted,
    onDraftingStarted: (owner, turn) => setDrafting({ owner, turn }),
    sourceCount: documents?.filter((document) => document.role !== "deliverable").length,
  })
  const {
    turns,
    input,
    setInput,
    busy,
    loading,
    loadError,
    permissions,
    permissionError,
    editing,
    setEditing,
    resolvedAgent,
    setResolvedAgent,
    sessionRef,
    partsRef,
    rolesRef,
    logRef,
    stickRef,
    inputRef,
    turnRef,
  } = state

  const scroll = useStickToBottom(logRef, stickRef)

  // A fresh chat (no session yet) seats the cursor in the composer on mount so
  // the lawyer can type immediately, as does the remount that follows the first
  // send minting a session (`created`) — same composer, so keep the cursor in it.
  // A conversation reopened from the rail skips this; it focuses once a turn
  // settles (see below).
  useEffect(() => {
    if (!sessionID || created) inputRef.current?.focus({ preventScroll: true })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // A chat deep-linked with a prompt (a workflow, a template, a review-table
  // answer) seats it in the composer for the lawyer to review and send —
  // prefill, never auto-send — and never over text the lawyer has typed.
  const toast = useToast()
  useEffect(() => {
    if (!initialPrompt || (sessionID && !onPromptSeated)) return
    if (input.trim() && input !== initialPrompt) toast("info", "Your unsent message was kept, so the suggested question was not added. Send or clear it, then try again.")
    else {
      setInput(initialPrompt)
      inputRef.current?.focus({ preventScroll: true })
    }
    onPromptSeated?.()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialPrompt])

  // Refocus the composer the moment a turn settles (busy → false). Done in an
  // effect, not inline in finalize(), so it runs after React commits the
  // settled turns: calling focus() inside finalize fires before that re-render,
  // landing on a node React then reconciles away — which is why the refocus only
  // worked intermittently.
  const wasBusy = useRef(false)
  useEffect(() => {
    if (wasBusy.current && !busy) inputRef.current?.focus({ preventScroll: true })
    wasBusy.current = busy
  }, [busy])

  const liveAgent = isAuto(agent) && resolvedAgent ? resolvedAgent : agent
  const live = readTurn(partsRef.current, rolesRef.current, matterName, sessionRef.current, liveAgent)
  const activity = busy ? live.activity : undefined
  // The firm's legends, read once a quick draft or client summary shows, as a
  // rendered document reads them: each carries its form's legend from the
  // firm's settings, never one the model wrote.
  const [legends, setLegends] = useState<{ agreementLegend: string; adviceLegend: string }>()
  // Settings that could not be read leave the answer without its legend; say so
  // rather than show nothing.
  const [legendsFailed, setLegendsFailed] = useState(false)
  const showsLegend = turns.some((t) => t.agent && EXPLICIT_ONLY.has(t.agent)) || (busy && EXPLICIT_ONLY.has(liveAgent))
  useEffect(() => {
    if (!showsLegend || legends) return
    let current = true
    getPreferences()
      .then((prefs) => current && setLegends({ agreementLegend: prefs.agreementLegend ?? "", adviceLegend: prefs.adviceLegend ?? "" }))
      .catch(() => current && setLegendsFailed(true))
    return () => {
      current = false
    }
  }, [showsLegend, legends])
  const legendMissing = legendsFailed && <p className="quick-draft-legend-missing" role="note">Firm legend could not be loaded.</p>
  const stopping = busy && turnRef.current.stopped
  useEffect(() => {
    onBusyChange?.(busy, activity)
  }, [busy, activity, onBusyChange])
  // The narration answers under the request turn this chat sent; anywhere else
  // (a reload, the run's own conversation) it shows the run's recorded request.
  const anchor =
    drafting && narration.status?.owner_session_id === drafting.owner && drafting.turn < turns.length
      ? drafting.turn
      : undefined
  // The run's own conversation (or the one that started it) carries its story
  // and its questions; any other conversation gets one line leading there.
  const own = run !== undefined && (run.owner_session_id === sessionID || run.owner_session_id === drafting?.owner)
  // A completed or stopped run takes no updates: the composer starts a new request.
  const updatingRun = Boolean(run && leanRunTakesUpdates(run) && matterId)
  const updateAdmission = run ? draftingUpdateAdmission(run) : "closed"
  const updateTooLate = updateAdmission === "closed"
  const clarifyingRun = updateAdmission === "clarification"
  async function sendRunUpdate() {
    const text = input.trim()
    if (!text || actions.busy || busy || updateTooLate) return
    if (await actions.update(text)) {
      setInput((current) => (current.trim() === text ? "" : current))
      if (inputRef.current) inputRef.current.style.height = "auto"
    }
  }
  const asked = run ? openQuestions(run) : []
  // Only the latest quick draft can be upgraded, and only while no drafting
  // run is taking updates on this matter.
  const latestQuick = turns.findLastIndex((t) => t.role === "assistant" && t.agent === QUICK_DRAFT)
  const canUpgrade = Boolean(matterId) && !busy && !updatingRun && !pendingRoute
  const quickBanner = (index?: number) => (
    <p className="quick-draft-banner" role="note">
      {QUICK_BANNER}
      {index !== undefined && index === latestQuick && canUpgrade && (
        <>
          {" "}
          <button type="button" className="linklike" onClick={() => void upgradeQuickDraft(index)}>
            Run full drafting cycle
          </button>
        </>
      )}
    </p>
  )
  // After a reload the offer's request is not in the stored history yet.
  const showPendingRoute = pendingRoute !== undefined && !(turns.at(-1)?.role === "user" && turns.at(-1)?.text === pendingRoute)
  // A comparison's answer draws each change it names as a redline.
  const answer = (text: string, compares?: VersionCompare[]) => {
    const compare = chatCompare(compares)
    return compare ? <CompareAnswer text={text} compare={compare} onOpen={onOpenBlackline} /> : <Markdown>{text}</Markdown>
  }
  // A quick draft shows the firm's legend for its form, its text without the
  // form line and, once settled, the placeholders left in the draft as open
  // items, which code finds rather than trusting the answer's own list.
  const quickAnswer = (text: string, settled: boolean) => {
    const quick = quickDraftAnswer(text)
    const legend = quick.form && legends ? furnitureLegend(quick.form, legends).trim() : ""
    return (
      <>
        {legend && <p className="quick-draft-legend">{legend}</p>}
        {quick.form && legendMissing}
        {answer(stripIssues(quick.text))}
        {settled && !quick.listed && quick.placeholders.length > 0 && (
          <p className="quick-draft-open" role="note">
            Open items in the draft: {quick.placeholders.join(", ")}
          </p>
        )}
      </>
    )
  }

  // A client summary is advice to the client, so it carries the firm's advice
  // legend; the model writes none.
  const clientAnswer = (text: string) => {
    const legend = legends ? furnitureLegend("advice", legends).trim() : ""
    return (
      <>
        {legend && <p className="quick-draft-legend">{legend}</p>}
        {legendMissing}
        {answer(stripIssues(text))}
      </>
    )
  }
  const agentAnswer = (agentName: string | undefined, text: string, settled: boolean, compares?: VersionCompare[]) =>
    agentName === QUICK_DRAFT ? quickAnswer(text, settled) : agentName === CLIENT_SUMMARY ? clientAnswer(text) : answer(stripIssues(text), compares)

  const narrationView = own && matterId && (
    <DraftingNarration
      narration={narration}
      matterId={matterId}
      showRequest={anchor === undefined}
      onViewDocument={onViewDocument}
    />
  )

  const composer = (documentsControl?: ReactNode) => (
    <Composer
      documentsControl={documentsControl}
      draftingUpdate={updatingRun}
      draftingClarification={clarifyingRun}
      pinned={pinned}
      available={available}
      agent={agent}
      resolvedAgent={resolvedAgent}
      setResolvedAgent={setResolvedAgent}
      onAgentChange={onAgentChange}
      inputRef={inputRef}
      workflows={workflows}
      runWorkflow={runWorkflow}
      playbooks={playbooks}
      playbook={playbook}
      onPlaybookChange={onPlaybookChange}
      toolsLoading={toolsLoading}
      composerPlaceholder={clarifyingRun ? "Clarify the source or applicability for the held checklist…" : updatingRun ? "Add an instruction or consideration for this draft…" : composerPlaceholder}
      input={input}
      setInput={setInput}
      onSend={updatingRun ? sendRunUpdate : onSend}
      onStop={busy ? onStop : undefined}
      stopping={stopping}
      busy={busy || pendingRoute !== undefined || (updatingRun && (actions.busy || updateTooLate))}
    />
  )

  return (
    <div className="card chat-card design-conversation">
      {!own && (updatingRun || title) && (
        <div className="chat-heading">
          <h2>{updatingRun ? "Update current draft" : title}</h2>
        </div>
      )}
      {run && (
        <RunHeader
          status={run}
          actions={actions}
          own={own}
          statusError={feed?.error}
          checkedAt={feed?.checkedAt}
          onOpen={(name) => onViewDocument(name)}
          onShowDrafts={() => {
            // Leave the bottom first, so a render during the smooth scroll
            // cannot re-pin the log and cancel it.
            stickRef.current = false
            const still = window.matchMedia("(prefers-reduced-motion: reduce)").matches
            document.getElementById("run-result")?.scrollIntoView({ block: "start", behavior: still ? "auto" : "smooth" })
          }}
        />
      )}
      <div className="chat-log-wrap">
        <div
          className={`chat-log${scroll.top ? " fade-top" : ""}${scroll.bottom ? " fade-bottom" : ""}`}
          ref={logRef}
          {...scroll.bind}
        >
          {loadError && (
            <div className="msg-error" role="alert">
              {loadError}
            </div>
          )}
          {loading && turns.length === 0 && <ChatSkeleton />}
          {!loading && !loadError && turns.length === 0 && !busy && !own && !updatingRun && pendingRoute === undefined && (
            <ChatEmpty
              hint={emptyHint}
              starters={starters}
              documentCount={documents?.length}
              onPick={setInput}
              onAddDocuments={documents && onAddFiles ? () => setDocumentsDrawer(true) : undefined}
            />
          )}
          {turns.map((t, i) => (
            <Fragment key={i}>
              {t.role === "user" ? (
                <div className="msg-user-wrap">
                  <div className="msg user">
                    {editing?.index === i ? (
                      <div className="msg-edit">
                        <textarea
                          autoFocus
                          aria-label="Edit your message"
                          value={editing.draft}
                          onChange={(e) => setEditing({ index: i, draft: e.target.value })}
                          onKeyDown={(e) => {
                            if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) resendFrom(t, i, editing.draft)
                            if (e.key === "Escape") setEditing(null)
                          }}
                        />
                        <div className="msg-edit-actions">
                          <button className="icon-btn" onClick={() => setEditing(null)}>
                            Cancel
                          </button>
                          <button className="primary" onClick={() => resendFrom(t, i, editing.draft)}>
                            Save & send
                          </button>
                        </div>
                      </div>
                    ) : (
                      <UserText text={t.text} />
                    )}
                  </div>
                  {editing?.index !== i && t.id && !busy && !updatingRun && (
                    <div className="msg-actions">
                      <button
                        className="icon-btn"
                        title="Edit message"
                        aria-label="Edit message"
                        onClick={() => setEditing({ index: i, draft: t.text })}
                      >
                        <IconPencil />
                      </button>
                      <button className="icon-btn" title="Retry" aria-label="Retry message" onClick={() => resendFrom(t, i, t.text)}>
                        <IconRetry />
                      </button>
                    </div>
                  )}
                </div>
              ) : (
                <div className="msg assistant">
                  {t.agent && <div className="msg-agent">{agentLabel(t.agent)}</div>}
                  {t.agent === QUICK_DRAFT && !t.error && !isEmptyAnswer(t.text) && quickBanner(i)}
                  <StepsPanel steps={t.steps} busy={false} answered={!isEmptyAnswer(t.text)} />
                  {t.text &&
                    !isEmptyAnswer(t.text) &&
                    agentAnswer(t.agent, t.text, true, t.compares)}
                  <IssuesView issues={parseIssues(t.text)} />
                  {t.error && <div className="msg-error">{t.error}</div>}
                  {t.stopped && !t.superseded && (
                    <p className="msg-stopped" role="note">
                      {isEmptyAnswer(t.text) ? "Stopped before an answer." : "Stopped. This answer is incomplete."}
                    </p>
                  )}
                  {!t.superseded &&
                    !t.stopped &&
                    !t.error &&
                    isEmptyAnswer(t.text) &&
                    t.redlines.length === 0 &&
                    t.drafts.length === 0 &&
                    !t.compares?.length && (
                      <div className="msg-error">
                        No answer was produced for this question.{" "}
                        <button
                          className="linklike"
                          onClick={() => {
                            const prev = turns[i - 1]
                            if (prev) resendFrom(prev, i - 1, prev.text)
                          }}
                        >
                          Retry
                        </button>
                      </div>
                    )}
                  <TodoView todos={t.todos} />
                  <CitationView citations={t.citations} matterId={matterId} documents={documents} />
                  <RedlineView redlines={t.redlines} onView={onViewDocument} />
                  <DraftView drafts={t.drafts} onView={onViewDocument} />
                  <CompareView compares={t.compares} onOpen={onOpenBlackline} />
                </div>
              )}
              {/* Auto's acknowledgement opens the reply to this message: display only. */}
              {t.role === "user" && t.ack && (
                <p className="msg-ack">
                  {t.ack}
                </p>
              )}
              {anchor === i && narrationView}
            </Fragment>
          ))}
          {anchor === undefined && narrationView}
          {showPendingRoute && (
            <div className="msg-user-wrap">
              <div className="msg user">
                <UserText text={pendingRoute} />
              </div>
            </div>
          )}
          {(own || updatingRun) && (run?.updates?.length ?? 0) > 0 && (
            <div className="run-updates">
              {run!.updates!.map((update) => (
                <div key={update.id}>
                  <div className="msg user">{update.text}</div>
                  <p className="muted" role="status">
                    {updateProgressLabel(update)}
                  </p>
                </div>
              ))}
            </div>
          )}
          {busy && (
            <div className={`msg assistant${live.text ? " streaming" : ""}`}>
              <div className="msg-agent">{agentLabel(liveAgent)}</div>
              {liveAgent === QUICK_DRAFT && quickBanner()}
              <StepsPanel steps={live.steps} busy answered={Boolean(live.text)} />
              {live.text && agentAnswer(liveAgent, live.text, false, live.compares)}
              {/* The answer is still coming: before its first step, during each step and between them. */}
              <p className="msg-working" role="status">
                <span className="working" aria-hidden>
                  <span className="working-dot" />
                  <span className="working-dot" />
                  <span className="working-dot" />
                </span>
                {stopping ? "Stopping" : (activity ?? "Working")}
              </p>
              <IssuesView issues={parseIssues(live.text)} />
              <TodoView todos={live.todos} />
              <CitationView citations={live.citations} matterId={matterId} documents={documents} />
              {live.text && <RedlineView redlines={live.redlines} onView={onViewDocument} />}
              <DraftView drafts={live.drafts} onView={onViewDocument} />
              <CompareView compares={live.compares} onOpen={onOpenBlackline} />
            </div>
          )}
          {!loading && !loadError && !busy && transcriptFooter}
        </div>
        {scroll.away && (
          <button type="button" className="chat-jump" onClick={scroll.jump}>
            Jump to latest
            <svg
              width="12"
              height="12"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden
            >
              <path d="M12 5v14" />
              <path d="m19 12-7 7-7-7" />
            </svg>
          </button>
        )}
      </div>
      {permissionError && <div className="msg-error">{permissionError}</div>}
      {permissions.map((p) => (
        <PermissionCard key={p.id} request={p} onReply={onPermissionReply} />
      ))}
      {own && run && queuedNote(run) && <p className="run-note" role="status">{queuedNote(run)}</p>}
      {own && run && matterId && (
        <QuestionDock
          matterId={matterId}
          ownerSessionID={run.owner_session_id}
          questions={asked}
          answered={run.questions.length - asked.length}
          answeredQuestions={run.stopped ? [] : run.questions.filter(question => question.status === "answered")}
          queuedAnswerIDs={run.queued_answer_ids}
          titles={new Map((run.plan?.documents ?? []).map((row) => [row.id, cleanTitle(row.title)]))}
          busy={actions.busy}
          error={actions.error}
          errorFor={actions.errorFor}
          onSend={(question, answer) => actions.answer(question.id, answer)}
          onChoose={(question, optionID) => actions.choose(question.id, optionID)}
          onDecide={(question) => actions.decide(question.id)}
        />
      )}
      <SessionQuestions
        matterId={matterId}
        client={client}
        sessionRef={sessionRef}
        pendingRef={pendingFormsRef}
        refreshRef={formsRefreshRef}
        directory={directory}
        onRoute={(request, choice) => void chooseDraft(request, choice)}
        onRouteDismissed={dismissDraftOffer}
        onRoutePending={setPendingRoute}
      />
      {quickLimit && (
        <div className="question-card route-offer" role="status">
          <p className="question-text">
            {quickLimit.reason} A quick draft covers one short document from a few of the matter's documents, with no
            outside law, firm template or legal effect.
          </p>
          <div className="question-actions">
            <button type="button" className="icon-btn" onClick={dismissQuickLimit}>
              Dismiss
            </button>
            <button type="button" className="primary" disabled={busy} onClick={() => void runFullInstead()}>
              Run full drafting cycle
            </button>
          </div>
        </div>
      )}
      {updatingRun && (
        <p className="run-note" role="status">
          {updateTooLate
            ? "Applying or finishing this draft. Wait for work to settle before clarifying a held checklist, or keep changes for a later amendment. You can still answer specific questions."
            : clarifyingRun
              ? "Clarify the source or applicability for checklist inspection. This preserves the draft and proposed amendments. A supported wording correction can then be submitted with Continue drafting; clarification does not approve a proposal or edit wording."
              : "Messages here update the current draft after the current step. Earlier work may be repeated to reflect your instructions."}
        </p>
      )}
      {updatingRun && !own && actions.error && (
        <p className="msg-error" role="alert">
          {actions.error}
        </p>
      )}
      {documents && onAddFiles ? (
        <DocumentsDock count={documents.length} onAddFiles={onAddFiles}>
          {composer}
        </DocumentsDock>
      ) : (
        composer()
      )}
    </div>
  )
}

// The matter page and the narration each poll the same run; show the newer read.
function latest(a: LeanDraftingStatus | undefined, b: LeanDraftingStatus | undefined) {
  if (!a || !b) return a ?? b
  return Date.parse(b.updated_at) > Date.parse(a.updated_at) ? b : a
}
