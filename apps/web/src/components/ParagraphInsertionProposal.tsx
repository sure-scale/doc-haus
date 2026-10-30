import type { Redline } from "../api/ingest/redlines"
import "./ParagraphInsertionProposal.css"

type Proposal = Pick<Redline, "scope" | "paragraph_insert" | "paragraph_spacing" | "table_row" | "target_json" | "stale">
type Insertion = NonNullable<Redline["paragraph_insert"]>
type Location = Insertion["location"]
const KIND = "paragraph_insert"
const record = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === "object" && !Array.isArray(value))
const sourceText = (value: unknown): value is string => typeof value === "string" && value.isWellFormed() && !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value)

function insertionLocation(value: unknown): Location | undefined {
  if (!record(value) || Object.keys(value).length !== 4 || typeof value.body_paragraph_number !== "number" || !Number.isSafeInteger(value.body_paragraph_number) || value.body_paragraph_number < 1 ||
      !Array.isArray(value.headings) || !value.headings.every(sourceText) ||
      value.previous_paragraph !== null && !sourceText(value.previous_paragraph) || value.next_paragraph !== null && !sourceText(value.next_paragraph)) return
  return { body_paragraph_number: value.body_paragraph_number, headings: value.headings, previous_paragraph: value.previous_paragraph, next_paragraph: value.next_paragraph }
}

function sameLocation(first: Location, second: Location): boolean {
  return first.body_paragraph_number === second.body_paragraph_number && first.previous_paragraph === second.previous_paragraph && first.next_paragraph === second.next_paragraph &&
    first.headings.length === second.headings.length && first.headings.every((heading, index) => heading === second.headings[index])
}

export function paragraphInsertionMarker(proposal: Proposal): boolean {
  if (proposal.scope === "paragraph_insert" || proposal.paragraph_insert !== undefined) return true
  try {
    const value: unknown = JSON.parse(proposal.target_json ?? "null")
    return record(value) && (value.version === 5 || value.kind === KIND)
  } catch { return false }
}

function insertionValue(proposal: Proposal): Insertion | undefined {
  if (proposal.scope !== "paragraph_insert" || proposal.paragraph_spacing !== undefined || proposal.table_row !== undefined) return
  const view: unknown = proposal.paragraph_insert
  if (!record(view) || Object.keys(view).length !== 5 || view.kind !== KIND || view.position !== "before" && view.position !== "after" ||
      typeof view.reference_text !== "string" || !view.reference_text.length || view.reference_text.length > 8192 || !view.reference_text.isWellFormed() || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(view.reference_text) ||
      typeof view.text !== "string" || !view.text.length || view.text.length > 4000 || !view.text.isWellFormed() || view.text !== view.text.trim() || /[\u0000-\u001f\u007f]/.test(view.text)) return
  const location = insertionLocation(view.location)
  if (!location) return
  if (proposal.target_json != null) {
    try {
      const binding: unknown = JSON.parse(proposal.target_json)
      if (!record(binding) || binding.version !== 5 || binding.kind !== KIND || binding.reference_text !== view.reference_text || binding.position !== view.position || binding.text !== view.text) return
      const boundLocation = insertionLocation(binding.location)
      if (!boundLocation || !sameLocation(location, boundLocation)) return
    } catch { return }
  }
  return { kind: KIND, reference_text: view.reference_text, position: view.position, text: view.text, location }
}

export function reviewableParagraphInsertion(proposal: Proposal): boolean {
  return !proposal.stale && insertionValue(proposal) !== undefined
}

export function ParagraphInsertionProposal({ proposal }: { proposal: Redline }) {
  const view = insertionValue(proposal)
  if (!view) return <p className="paragraph-insertion-unavailable">Paragraph insertion preview unavailable. Reject this proposal and prepare a fresh change list.</p>
  const stale = Boolean(proposal.stale)
  const title = `Insert a paragraph ${view.position} the ${stale ? "original " : ""}reference`
  const blocks = view.position === "before" ? ["new", "reference"] as const : ["reference", "new"] as const
  return <div className="paragraph-insertion-proposal">
    <p className="paragraph-insertion-title">{title}</p>
    {stale && <p className="paragraph-insertion-unavailable">Reference location needs refreshing.</p>}
    <details className="paragraph-insertion-location" open>
      <summary>{stale ? "Original body paragraph" : "Body paragraph"} {view.location.body_paragraph_number}</summary>
      <dl>
        <div><dt>{stale ? "Original heading path" : "Heading path"}</dt>
          <dd>{view.location.headings.length ? <ol>{view.location.headings.map((heading, index) => <li key={index}><SourceText text={heading} /></li>)}</ol> : <span className="paragraph-insertion-context-note">No heading</span>}</dd>
        </div>
        {(["previous", "next"] as const).map(side => {
          const text = view.location[`${side}_paragraph`]
          return <div key={side}>
            <dt>{stale ? `Original ${side} paragraph` : side === "previous" ? "Previous paragraph" : "Next paragraph"}</dt>
            <dd>{text === null ? <span className="paragraph-insertion-context-note">No adjacent body paragraph</span> : <SourceText text={text} />}</dd>
          </div>
        })}
      </dl>
    </details>
    {blocks.map(kind => <section key={kind} className={`paragraph-insertion-block paragraph-insertion-${kind}`} aria-label={kind === "new" ? "New paragraph" : stale ? "Original reference paragraph" : "Reference paragraph, unchanged"}>
      <h4>{kind === "new" ? "New paragraph" : stale ? "Original reference paragraph" : "Reference paragraph · unchanged"}</h4>
      <p className="paragraph-insertion-text">{kind === "new" ? view.text : view.reference_text}</p>
      {kind === "reference" && !view.reference_text.trim() && <p className="paragraph-insertion-note">Whitespace only</p>}
    </section>)}
  </div>
}

function SourceText({ text }: { text: string }) {
  return <><span className="paragraph-insertion-source-text">{text}</span>{!text.trim() && <span className="paragraph-insertion-note">{text ? "Whitespace only" : "Empty paragraph"}</span>}</>
}
