import { useEffect, useRef, useState } from "react"
import type {
  Chronology,
  ChronologyAccount,
  ChronologyDecision,
  ChronologyEvent,
  ChronologyEvidence,
  ChronologyProposal,
} from "../../api/ingest/chronology"
import JurisdictionSelect from "../JurisdictionSelect"
import { dateLabel, eventEvidence, eventQualifications, simpleAddition } from "./model"
import "./proposal-review.css"

const purposes: Record<ChronologyProposal["kind"], { title: string; explanation: string; short: string }> = {
  addition: {
    title: "Add an event",
    short: "New event",
    explanation: "Check the new event and its supporting sources before adding it.",
  },
  amendment: {
    title: "Update this event",
    short: "Update",
    explanation: "Check the changed information and its sources.",
  },
  conflict: {
    title: "Review a competing account",
    short: "Competing account",
    explanation:
      "Accepting keeps the existing accounts and the disagreement. It does not decide which account is true.",
  },
  removal: {
    title: "Hide this event",
    short: "Hide event",
    explanation: "The event and its history are kept. You can restore it using the hidden-events filter.",
  },
  merge: {
    title: "Combine these events",
    short: "Combine events",
    explanation:
      "Check that these entries describe the same occurrence. Source accounts and links to the previous records are kept.",
  },
  split: {
    title: "Separate this event",
    short: "Separate events",
    explanation: "Check the separate occurrences. Source accounts and links to the previous entry are kept.",
  },
}

/** Exact structured equality: object key order and absent optional properties are
 * immaterial, while source identity, passage, date precision and qualifications count. */
function sameValue(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (!left || !right || typeof left !== "object" || typeof right !== "object") return false
  if (Array.isArray(left) || Array.isArray(right))
    return (
      Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((value, at) => sameValue(value, right[at]))
    )
  const a = left as Record<string, unknown>,
    b = right as Record<string, unknown>
  const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])]
  return keys.every((key) => sameValue(a[key], b[key]))
}
const sameEvidence = (left: ChronologyEvidence, right: ChronologyEvidence) =>
  sameValue({ ...left, id: undefined }, { ...right, id: undefined })
const sameAccount = (left: ChronologyAccount, right: ChronologyAccount) =>
  sameValue(
    { ...left, id: undefined, evidence: left.evidence.map(({ id, ...source }) => source) },
    { ...right, id: undefined, evidence: right.evidence.map(({ id, ...source }) => source) },
  )

// Amendments replace sourced content, while the decision path keeps these lawyer choices.
function reviewEvent(event: ChronologyEvent, saved: ChronologyEvent[], kind: ChronologyProposal["kind"]) {
  const current = saved[0]
  if (saved.length > 0 && (kind === "merge" || kind === "split"))
    return {
      ...event,
      key: saved.some((target) => target.key),
      hidden: saved.every((target) => target.hidden),
      note: saved
        .map((target) => target.note)
        .filter(Boolean)
        .join("\n\n"),
    }
  return current && (kind === "amendment" || kind === "conflict")
    ? {
        ...event,
        origin: current.origin,
        hidden: current.hidden,
        key: current.key,
        note: current.note,
        lawyerEdited: current.lawyerEdited,
        disputed: kind === "conflict" || current.disputed,
        dispute: current.dispute,
      }
    : event
}

export default function ProposalReview({
  chronology,
  busy,
  initialProposal,
  onDecision,
  onBatch,
  onDiscard,
  onClose,
  onEvidence,
  onAsk,
}: {
  chronology: Chronology
  busy: boolean
  initialProposal?: string
  onDecision: (proposal: ChronologyProposal, decision: ChronologyDecision) => Promise<boolean>
  onBatch: (ids: string[]) => Promise<boolean>
  onDiscard: () => Promise<boolean>
  onClose: () => void
  onEvidence: (evidence: ChronologyEvidence) => void
  onAsk: (event: ChronologyEvent, proposal: ChronologyProposal) => void
}) {
  const proposals = chronology.proposals.filter(
    (proposal) => proposal.status === "pending" || proposal.status === "deferred",
  )
  const [selected, setSelected] = useState(initialProposal ?? proposals[0]?.id)
  const [batch, setBatch] = useState(false)
  const [discard, setDiscard] = useState(false)
  const [batchIds, setBatchIds] = useState<string[]>([])
  const heading = useRef<HTMLHeadingElement>(null),
    scroll = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!proposals.some((proposal) => proposal.id === selected))
      setSelected(proposals.find((proposal) => proposal.status === "pending")?.id ?? proposals[0]?.id)
    setBatchIds((current) =>
      current.filter((id) => proposals.some((proposal) => proposal.id === id && simpleAddition(proposal))),
    )
    // A stored revision changes the available decisions; local navigation does not.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chronology.revision])
  const proposal = proposals.find((item) => item.id === selected),
    index = proposals.findIndex((item) => item.id === selected)
  const additions = proposals.filter(simpleAddition),
    refreshSuggestions = proposals.filter((item) => item.origin === "refresh"),
    later = proposals.filter((item) => item.status === "deferred").length
  const saved = proposal
    ? proposal.targetIds.flatMap((id) => {
        const event = chronology.events.find((event) => event.id === id)
        return event ? [event] : []
      })
    : []
  const eventTitle = (item: ChronologyProposal) =>
    item.events[0]?.title ??
    chronology.events.find((event) => item.targetIds.includes(event.id))?.title ??
    "Previous event unavailable"
  useEffect(() => {
    if (!batch && !discard) {
      if (scroll.current) scroll.current.scrollTop = 0
      heading.current?.focus({ preventScroll: true })
    }
  }, [proposal?.id, batch, discard])
  const askEvent = proposal?.events[0] ?? saved[0]
  return (
    <section className="chronology-proposal-review update-review" aria-label="Review chronology updates">
      <div ref={scroll} className="update-review-scroll">
        <div className="update-review-intro">
          <p className="update-review-progress" role="status">
            {proposals.length
              ? `${proposals.length} ${proposals.length === 1 ? "change awaiting" : "changes awaiting"} your decision${later ? ` · ${later} saved for later` : ""}`
              : "No changes awaiting a decision"}
          </p>
          <p>Saved events stay unchanged until you accept.</p>
        </div>
        {proposals.length > 0 && (
          <div className="update-review-navigation">
            <div className="update-review-picker">
              <JurisdictionSelect
                jurisdictions={proposals.map((item, at) => ({
                  code: item.id,
                  name: `${at + 1}. ${purposes[item.kind].short} — ${eventTitle(item)}${item.status === "deferred" ? " · For later" : ""}`,
                }))}
                selected={selected ? [selected] : []}
                multiple={false}
                showCode={false}
                disabled={busy}
                label="Update to review"
                searchLabel="Search updates"
                emptyLabel="Choose an update"
                noResultsLabel="No updates match your search."
                trigger={
                  <>
                    <span>
                      <small>
                        {index + 1} of {proposals.length}
                      </small>
                      <strong>{proposal ? eventTitle(proposal) : "Choose an update"}</strong>
                    </span>
                    <span aria-hidden="true">⌄</span>
                  </>
                }
                onChange={(ids) => {
                  if (ids[0]) {
                    setSelected(ids[0])
                    setBatch(false)
                    setDiscard(false)
                  }
                }}
              />
            </div>
            <div className="update-review-step-buttons">
              <button
                type="button"
                aria-label="Previous update"
                title="Previous update"
                disabled={busy || index <= 0}
                onClick={() => {
                  setSelected(proposals[index - 1]?.id)
                  setBatch(false)
                  setDiscard(false)
                }}
              >
                ‹
              </button>
              <button
                type="button"
                aria-label="Next update"
                title="Next update"
                disabled={busy || index < 0 || index >= proposals.length - 1}
                onClick={() => {
                  setSelected(proposals[index + 1]?.id)
                  setBatch(false)
                  setDiscard(false)
                }}
              >
                ›
              </button>
            </div>
            {(additions.length > 0 || refreshSuggestions.length > 0) && (
              <details
                className="update-review-more"
                key={`more-${proposal?.id}`}
                onKeyDown={(event) => {
                  if (event.key === "Escape" && event.currentTarget.open) {
                    event.preventDefault()
                    event.stopPropagation()
                    event.currentTarget.open = false
                    event.currentTarget.querySelector("summary")?.focus()
                  }
                }}
                onBlur={(event) => {
                  if (!event.currentTarget.contains(event.relatedTarget as Node | null))
                    event.currentTarget.open = false
                }}
              >
                <summary aria-label="More update actions" title="More update actions">
                  •••
                </summary>
                <div>
                  {additions.length > 0 && (
                    <button
                      type="button"
                      disabled={busy}
                      onClick={(event) => {
                        event.currentTarget.closest("details")?.removeAttribute("open")
                        setBatch((value) => !value)
                        setBatchIds(additions.map((item) => item.id))
                        setDiscard(false)
                      }}
                    >
                      <strong>
                        Review {additions.length} new {additions.length === 1 ? "event" : "events"} together…
                      </strong>
                      <small>Choose which to add as unreviewed.</small>
                    </button>
                  )}
                  {refreshSuggestions.length > 0 && (
                    <button
                      type="button"
                      disabled={busy}
                      onClick={(event) => {
                        event.currentTarget.closest("details")?.removeAttribute("open")
                        setDiscard((value) => !value)
                        setBatch(false)
                      }}
                    >
                      <strong>Reject refresh suggestions…</strong>
                      <small>Keep your saved chronology unchanged.</small>
                    </button>
                  )}
                </div>
              </details>
            )}
          </div>
        )}
        {discard && (
          <div className="update-review-notice" role="note">
            <strong>Reject the refresh suggestions?</strong>
            <p>
              {refreshSuggestions.length} refresh {refreshSuggestions.length === 1 ? "suggestion" : "suggestions"} will
              be rejected. Your saved chronology and chat suggestions are kept.
            </p>
            <div className="update-review-inline">
              <button type="button" onClick={() => setDiscard(false)}>
                Keep suggestions
              </button>
              <button
                type="button"
                disabled={busy}
                onClick={async () => {
                  if (await onDiscard()) setDiscard(false)
                }}
              >
                Reject suggestions
              </button>
            </div>
          </div>
        )}
        {batch && (
          <div className="chronology-preview update-review-batch" aria-label="Batch addition preview">
            <h3>Choose new events to add</h3>
            <p>
              Only straightforward new events are included. They remain <strong>unreviewed</strong>. Changes to existing
              events, conflicts and hiding decisions need individual review.
            </p>
            <div className="update-review-batch-list">
              {additions.map((item) => (
                <label className="chronology-check" key={item.id}>
                  <input
                    type="checkbox"
                    disabled={busy}
                    checked={batchIds.includes(item.id)}
                    onChange={(change) =>
                      setBatchIds((current) =>
                        change.target.checked ? [...current, item.id] : current.filter((id) => id !== item.id),
                      )
                    }
                  />
                  <span>{item.events.map((event) => `${dateLabel(event.date)} · ${event.title}`).join("; ")}</span>
                </label>
              ))}
            </div>
            <div className="update-review-inline">
              <button type="button" onClick={() => setBatch(false)}>
                Cancel batch
              </button>
              <button
                type="button"
                className="primary"
                disabled={busy || batchIds.length === 0}
                onClick={async () => {
                  if (await onBatch(batchIds)) setBatch(false)
                }}
              >
                Add {batchIds.length} {batchIds.length === 1 ? "event" : "events"} as unreviewed
              </button>
            </div>
          </div>
        )}
        {proposal && !batch && !discard ? (
          <article key={proposal.id} className="chronology-proposal update-review-change">
            <header className="update-review-purpose">
              <h3 ref={heading} tabIndex={-1}>
                {purposes[proposal.kind].title}
              </h3>
              {proposal.status === "deferred" && <span className="update-review-later">For later</span>}
              {!["amendment", "addition"].includes(proposal.kind) && <p>{purposes[proposal.kind].explanation}</p>}
            </header>
            {saved.length !== proposal.targetIds.length && (
              <p className="chronology-warning">A previous entry is unavailable. Reload before deciding.</p>
            )}
            <ChangeOverview proposal={proposal} saved={saved} onEvidence={onEvidence} />
            <details className="update-review-full-comparison">
              <summary>
                {proposal.kind === "addition" ? "See all event and source details" : "See full comparison"}
              </summary>
              <div
                className={`update-review-comparison${proposal.kind === "addition" || proposal.kind === "removal" ? " is-single" : ""}`}
              >
                {proposal.kind !== "addition" && (
                  <section className="update-review-saved" aria-label="Current saved entry">
                    <h4>{saved.length > 1 ? "Saved entries now" : "Saved entry now"}</h4>
                    {saved.map((event) => (
                      <Preview key={event.id} event={event} onEvidence={onEvidence} />
                    ))}
                  </section>
                )}
                {proposal.kind !== "removal" && (
                  <section className="update-review-proposed" aria-label="Proposed result">
                    <h4>{proposal.kind === "addition" ? "New event to add" : "Proposed content"}</h4>
                    {["amendment", "conflict", "merge"].includes(proposal.kind) && (
                      <p className="update-review-retained-note">
                        Existing source accounts are retained alongside proposed content.
                      </p>
                    )}
                    {proposal.events.map((event) => (
                      <Preview
                        key={event.id}
                        event={reviewEvent(event, saved, proposal.kind)}
                        proposed
                        onEvidence={onEvidence}
                      />
                    ))}
                  </section>
                )}
              </div>
            </details>
            <details className="update-review-reason">
              <summary>Why this is suggested</summary>
              <p>{proposal.reason}</p>
              <h4>What to consider</h4>
              <p>{proposal.recommendation}</p>
              <p className="update-review-evidence-note">
                A verified quotation confirms a text match, not that an account is true.
              </p>
            </details>
          </article>
        ) : !proposal ? (
          <div className="update-review-empty">
            <span aria-hidden="true">✓</span>
            <h3 ref={heading} tabIndex={-1}>
              No pending updates
            </h3>
            <p>Your saved chronology is ready to return to.</p>
            <button type="button" onClick={onClose}>
              Back to saved events
            </button>
          </div>
        ) : null}
      </div>
      {proposal && !batch && !discard && (
        <footer className="update-review-decision-area">
          <div className="update-review-footer-top">
            <p>
              {proposal.kind === "removal"
                ? "Accept hides the event and keeps its history."
                : "Accept saves this change as reviewed."}
            </p>
            {askEvent && (
              <button
                type="button"
                className="update-review-ask"
                disabled={busy}
                onClick={() => onAsk(askEvent, proposal)}
              >
                Ask doc.haus
              </button>
            )}
          </div>
          <div className="update-review-decisions">
            <button
              type="button"
              className="primary"
              disabled={busy}
              onClick={() => void onDecision(proposal, "accept")}
            >
              Accept this change
            </button>
            <button
              type="button"
              disabled={busy}
              title="Reject this suggestion and keep the saved chronology unchanged"
              onClick={() => void onDecision(proposal, "keep")}
            >
              Keep current
            </button>
            <button
              type="button"
              disabled={busy}
              title="Keep this suggestion for a later decision"
              onClick={async () => {
                if (await onDecision(proposal, "defer")) {
                  const next = proposals.find((item) => item.id !== proposal.id && item.status === "pending")
                  if (next) setSelected(next.id)
                }
              }}
            >
              Skip for later
            </button>
          </div>
        </footer>
      )}
    </section>
  )
}

function ChangeOverview({
  proposal,
  saved,
  onEvidence,
}: {
  proposal: ChronologyProposal
  saved: ChronologyEvent[]
  onEvidence: (evidence: ChronologyEvidence) => void
}) {
  const current = saved.length === 1 ? saved[0] : undefined,
    next = proposal.events.length === 1 ? proposal.events[0] : undefined
  const singleAmendment = proposal.kind === "amendment" && current && next
  const previousEvidence = saved.flatMap(eventEvidence)
  const changedAccounts = proposal.events
    .flatMap((event) => event.accounts)
    .filter((account) => !saved.some((event) => event.accounts.some((old) => sameAccount(old, account))))
  const newEvidence = changedAccounts
    .flatMap((account) => account.evidence)
    .filter((source) => !previousEvidence.some((old) => sameEvidence(old, source)))
    .filter((source, at, all) => all.findIndex((other) => sameEvidence(other, source)) === at)
  const removedEvidence = previousEvidence.filter(
    (source) => !proposal.events.flatMap(eventEvidence).some((next) => sameEvidence(source, next)),
  )
  const changedFields = singleAmendment
    ? [
        { name: "Event", before: current.title, after: next.title },
        { name: "Date", before: current.date, after: next.date },
        { name: "Description", before: current.detail, after: next.detail },
      ].filter((field) => !sameValue(field.before, field.after))
    : []
  const events =
    proposal.kind === "removal" ? saved : proposal.events.map((event) => reviewEvent(event, saved, proposal.kind))
  return (
    <section className="update-review-delta" aria-label="What changes">
      {events.map((event) => (
        <div className="update-review-overview-event" key={event.id}>
          <strong className="update-review-event-title">{event.title}</strong>
          <p className="update-review-event-date">{dateLabel(event.date)}</p>
          {!singleAmendment && event.detail && <p>{event.detail}</p>}
          <Qualifications event={event} />
        </div>
      ))}
      {singleAmendment && (
        <>
          {changedFields.length === 0 ? (
            <p className="update-review-unchanged">
              <span aria-hidden="true">✓</span> No change to event wording or date
            </p>
          ) : (
            <div className="update-review-field-changes">
              {changedFields.map((field) => (
                <div key={field.name}>
                  <h4>{field.name}</h4>
                  <div>
                    <span>Saved</span>
                    <FieldValue name={field.name} value={field.before} />
                  </div>
                  <div>
                    <span>Proposed</span>
                    <FieldValue name={field.name} value={field.after} />
                  </div>
                </div>
              ))}
            </div>
          )}
          {newEvidence.length > 0 && (
            <h4 className="update-review-delta-label">
              <span aria-hidden="true">＋</span>{" "}
              {newEvidence.some((source) => previousEvidence.some((old) => old.id === source.id))
                ? "Added or changed supporting evidence"
                : "Added supporting evidence"}{" "}
              <span>({newEvidence.length})</span>
            </h4>
          )}
          {newEvidence.map((source, at) => (
            <Source key={`${source.id}:${at}`} source={source} onEvidence={onEvidence} />
          ))}
          {changedAccounts.length > 0 && (
            <div className="update-review-account-summary">
              <p>
                {changedAccounts.length}{" "}
                {changedAccounts.length === 1 ? "new or changed source account" : "new or changed source accounts"} ·{" "}
                {current.accounts.length} saved · {next.accounts.length} proposed
              </p>
              <details>
                <summary>See account changes</summary>
                {changedAccounts.map((account) => (
                  <Account key={account.id} account={account} />
                ))}
              </details>
            </div>
          )}
          {removedEvidence.length > 0 && (
            <p className="chronology-warning">
              {removedEvidence.length} saved supporting {removedEvidence.length === 1 ? "passage is" : "passages are"}{" "}
              absent or changed in the proposed content. Existing source accounts are retained; check the full
              comparison before accepting.
            </p>
          )}
          {!newEvidence.length && !changedAccounts.length && !changedFields.length && (
            <p className="chronology-muted">Check the full comparison and suggestion details before deciding.</p>
          )}
        </>
      )}
      {events.some((event) => eventEvidence(event).some((source) => source.provenance === "supplied-chronology")) && (
        <p className="chronology-warning update-review-provenance-note">
          Includes supplied chronology evidence; its underlying original source is not independently established.
        </p>
      )}
      {proposal.kind === "addition" &&
        events
          .flatMap(eventEvidence)
          .map((source, at) => <Source key={`${source.id}:${at}`} source={source} onEvidence={onEvidence} />)}
      {proposal.kind === "conflict" && (
        <div className="update-review-conflict-note">
          <strong>Competing accounts remain unresolved</strong>
          <div className="update-review-conflict-accounts">
            <div>
              <h4>Saved accounts</h4>
              {saved
                .flatMap((event) => event.accounts)
                .map((account, at) => (
                  <AccountSummary key={`${account.id}:${at}`} account={account} />
                ))}
            </div>
            <div>
              <h4>New or changed accounts</h4>
              {changedAccounts.map((account, at) => (
                <AccountSummary key={`${account.id}:${at}`} account={account} />
              ))}
            </div>
          </div>
          {newEvidence.map((source, at) => (
            <Source key={`${source.id}:${at}`} source={source} onEvidence={onEvidence} />
          ))}
        </div>
      )}
      {proposal.kind === "removal" && (
        <p className="update-review-hide-note">
          After accepting: hidden, with its history kept. This does not establish that the event did not happen.
        </p>
      )}
      {(proposal.kind === "merge" || proposal.kind === "split") && (
        <p className="update-review-structure-note">
          {saved.length} saved {saved.length === 1 ? "entry" : "entries"} → {proposal.events.length} proposed{" "}
          {proposal.events.length === 1 ? "entry" : "entries"}. Review the full comparison before accepting.
        </p>
      )}
      {saved
        .filter(
          (event) =>
            !events.some(
              (next) =>
                eventQualifications(event).every((qualification) =>
                  eventQualifications(next).includes(qualification),
                ) &&
                (!event.lawyerEdited || next.lawyerEdited),
            ),
        )
        .map((event) => (
          <Qualifications key={event.id} event={event} />
        ))}
    </section>
  )
}

function FieldValue({ name, value }: { name: string; value: unknown }) {
  if (name === "Date") {
    const date = value as ChronologyEvent["date"]
    return (
      <div className="update-review-date-change">
        <p>{dateLabel(date)}</p>
        <p>
          Precision: {date.precision}
          {date.start
            ? ` · Timeline date: ${date.start}${date.end ? ` – ${date.end}` : ""}`
            : " · Not placed on timeline"}
          {date.rangeKind ? ` · ${date.rangeKind === "duration" ? "Known period" : "Uncertain interval"}` : ""}
        </p>
        {date.relative && (
          <p>
            Relative to {date.relative.anchorDocument}
            {date.relative.anchorDate ? ` (${date.relative.anchorDate})` : " (date unresolved)"}
            {date.relative.calculation ? ` · ${date.relative.calculation}` : ""}
            {date.relative.interpretation === "legal-interpretation" ? " · Legal interpretation" : ""}
          </p>
        )}
        {date.correction && <p>Lawyer-corrected date</p>}
        {date.anchorIssue && <p className="chronology-warning">{date.anchorIssue}</p>}
      </div>
    )
  }
  return <p>{typeof value === "string" && value ? value : "Not supplied"}</p>
}

function Qualifications({ event }: { event: ChronologyEvent }) {
  const qualifications = eventQualifications(event)
  return (
    <>
      {qualifications.length > 0 && <p className="update-review-qualifications">{qualifications.join(" · ")}</p>}
      {event.hidden && (
        <p className="chronology-warning">
          This entry is hidden. Accepting an update does not restore it automatically.
        </p>
      )}
      {event.lawyerEdited && <p className="chronology-muted update-review-lawyer-edit">Includes lawyer corrections.</p>}
    </>
  )
}

function Preview({
  event,
  proposed = false,
  onEvidence,
}: {
  event: ChronologyEvent
  proposed?: boolean
  onEvidence: (evidence: ChronologyEvidence) => void
}) {
  const evidence = eventEvidence(event)
  return (
    <article className="chronology-preview-event update-review-entry">
      <strong className="update-review-event-title">{event.title}</strong>
      <p className="update-review-event-date">
        {dateLabel(event.date)}
        {!proposed ? ` · ${event.reviewed ? "Reviewed" : "Unreviewed"}` : ""}
      </p>
      <p>{event.detail}</p>
      <Qualifications event={event} />
      <details className="update-review-accounts">
        <summary>Source accounts ({event.accounts.length})</summary>
        {event.accounts.map((account) => (
          <Account key={account.id} account={account} />
        ))}
      </details>
      <details className="update-review-evidence">
        <summary>Supporting evidence ({evidence.length})</summary>
        {evidence.length ? (
          evidence.map((source, index) => (
            <Source key={`${source.id}:${index}`} source={source} onEvidence={onEvidence} />
          ))
        ) : (
          <p className="chronology-warning">No supporting source was supplied.</p>
        )}
      </details>
    </article>
  )
}

function AccountSummary({ account }: { account: ChronologyAccount }) {
  return (
    <div className="update-review-account-summary-item">
      <strong>{dateLabel(account.date)}</strong>
      <p>{account.title}</p>
      <p>
        {account.status === "reported"
          ? "Reported occurrence"
          : account.status === "required"
            ? "Required event"
            : "Scheduled event"}
        {account.attribution ? ` · ${account.attribution}` : ""}
      </p>
      {(account.privilege === "uncertain" || account.privilege === "privileged") && (
        <p className="chronology-warning">Privilege: lawyer review needed</p>
      )}
    </div>
  )
}

function Account({ account }: { account: ChronologyAccount }) {
  return (
    <div className="update-review-account">
      <strong>{dateLabel(account.date)}</strong>
      <p>{account.title}</p>
      {account.detail && <p>{account.detail}</p>}
      {account.attribution && <p>Reported by {account.attribution}</p>}
      <p>
        {account.status === "reported"
          ? "Reported occurrence"
          : account.status === "required"
            ? "Required event"
            : "Scheduled event"}{" "}
        · Privilege: {account.privilege} · Significance: {account.significance}
      </p>
      {account.people.length > 0 && <p>People: {account.people.join(", ")}</p>}
      {account.gaps.map((gap, at) => (
        <p className="chronology-warning" key={at}>
          {gap}
        </p>
      ))}
    </div>
  )
}

function Source({
  source,
  onEvidence,
}: {
  source: ChronologyEvidence
  onEvidence: (evidence: ChronologyEvidence) => void
}) {
  return (
    <div className="update-review-source">
      <strong>
        {source.documentName}
        {source.section ? ` · ${source.section}` : ""}
      </strong>
      <blockquote>{source.excerpt || "Source passage unavailable."}</blockquote>
      <div className="update-review-source-status">
        <p>
          {source.verified ? "Quotation verified" : "Quotation unverified"} ·{" "}
          {source.freshness === "current"
            ? "Source unchanged"
            : source.freshness === "stale"
              ? "Source changed — recheck"
              : source.freshness === "unavailable"
                ? "Source unavailable"
                : "Source not checked"}
        </p>
        <button type="button" onClick={() => onEvidence(source)}>
          Inspect source
        </button>
      </div>
      {source.provenance === "supplied-chronology" && (
        <p className="chronology-warning">
          Supplied chronology — the underlying original source is not independently established by this passage.
        </p>
      )}
      {source.issue && <p className="chronology-warning">{source.issue}</p>}
    </div>
  )
}
