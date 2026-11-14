import type { Grounding } from "./chat/grounding"
import "./GroundingBadge.css"

const LABEL: Record<Exclude<Grounding["state"], "chat">, string> = {
  grounded: "Grounded",
  limited: "Limited support",
  abstain: "No verified support",
}

// The state a Q&A answer earned from code's structural check: a grounded chip
// when every substantive statement names a quotation the legal plugin verified
// word for word; limited and no-verified-support chips otherwise, with what was
// and was not linked behind the disclosure. Chat answers (no claim) show nothing.
export default function GroundingBadge({ grounding }: { grounding?: Grounding }) {
  if (!grounding || grounding.state === "chat") return null
  const label = LABEL[grounding.state]
  // Verified quotations with no marker anywhere: the gap is the missing map,
  // not any one paragraph, so listing every block adds nothing.
  const unmapped = grounding.citations > 0 && grounding.supported === 0
  return (
    <details className={`grounding grounding-${grounding.state}`}>
      <summary>
        <span className="grounding-chip">
          <Glyph state={grounding.state} />
          {label}
        </span>
        <span className="grounding-count">
          {grounding.supported} statement{grounding.supported === 1 ? "" : "s"} linked
        </span>
      </summary>
      <div className="grounding-body">
        <p>{EXPLAIN[grounding.state]}</p>
        {unmapped && <p>This answer cites verified sources but does not link them to its statements.</p>}
        {!unmapped && grounding.unlinked.length > 0 && (
          <ul>
            {grounding.unlinked.slice(0, 3).map((text, index) => <li key={index}>{text}</li>)}
            {grounding.unlinked.length > 3 && <li>{grounding.unlinked.length - 3} more not linked to a verified quotation.</li>}
          </ul>
        )}
        <ul>
          {grounding.unbound > 0 && <li>{grounding.unbound} citation{grounding.unbound === 1 ? "" : "s"} named no quotation verified against the document.</li>}
          {grounding.rejected > 0 && <li>{grounding.rejected} quotation{grounding.rejected === 1 ? "" : "s"} could not be verified against the document and {grounding.rejected === 1 ? "was" : "were"} not used.</li>}
          {grounding.unverifiedQuotes > 0 && <li>{grounding.unverifiedQuotes} quoted passage{grounding.unverifiedQuotes === 1 ? "" : "s"} in the answer match{grounding.unverifiedQuotes === 1 ? "es" : ""} no verified quotation.</li>}
          {grounding.citations > 0 && <li>{grounding.citations} verified quotation{grounding.citations === 1 ? "" : "s"} in this answer.</li>}
        </ul>
        <p className="grounding-limit">The check confirms citation wording against the documents. It does not judge whether a passage supports the conclusion drawn from it.</p>
      </div>
    </details>
  )
}

const EXPLAIN: Record<Exclude<Grounding["state"], "chat">, string> = {
  grounded: "Every substantive statement in this answer is linked to a quotation verified against the source document.",
  limited: "Part of this answer is not linked to a verified quotation. Read the unlinked text against the documents before relying on it.",
  abstain: "No verified quotation supports this answer. Treat it as unsupported and check the documents before relying on it.",
}

function Glyph({ state }: { state: Grounding["state"] }) {
  if (state === "grounded")
    return (
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
        <path d="m5 12 4 4L19 6" />
      </svg>
    )
  return (
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <path d="M12 3 2.5 20h19L12 3Z" />
      <path d="M12 9v5" />
      <path d="M12 17.5v.01" />
    </svg>
  )
}
