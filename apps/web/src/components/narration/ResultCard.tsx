import { documentContentUrl } from "../../api/ingest/documents"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import DraftedList, { useCombinedDownload } from "../drafting/DraftedList"
import { checklistFindingsCount, draftedGroups, findingsTitle, OUTPUT_ROLE, type DraftedOutput } from "../drafting/model"
import { resultChecks } from "./outcome"

/** A completed run in brief: what the checks established, what remains for the
 * lawyer, and the files it produced, each one click from the viewer. Every file
 * to review is featured: a lone draft as a page, several drafts as rows, and
 * the drafting notes as a row beneath them. Redlines stay with all the files.
 * A draft delivered as one combined file with its memo downloads as that file. */
export default function ResultCard({
  status,
  matterId,
  onViewDocument,
}: {
  status: LeanDraftingStatus
  matterId: string
  onViewDocument: (name: string) => void
}) {
  const checks = resultChecks(status)
  // A review draft delivered with unresolved checklist findings is not a clean completion.
  const findings = checklistFindingsCount(status)
  const groups = draftedGroups(status)
  const outputs = groups.flatMap((group) => [group.main, ...group.extras])
  const drafts = outputs.filter((output) => output.role === "draft")
  const notes = outputs.filter((output) => output.role === "notes")
  const lone = drafts.length === 1 ? drafts[0] : undefined
  const rows = [...(lone ? [] : drafts), ...notes]
  const combinedOf = (output: DraftedOutput) => groups.find((group) => group.main === output)?.combined
  const loneCombined = lone && combinedOf(lone)
  return (
    <section id="run-result" className="narration-result" aria-labelledby="run-result-title">
      <h3 id="run-result-title" className={`narration-result-title${findings === undefined ? "" : " is-findings"}`}>
        <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden>
          <circle cx="8" cy="8" r="7" fill="currentColor" />
          {findings === undefined ? (
            <path d="m5 8.2 2 2 4-4.2" fill="none" stroke="var(--text-on-accent)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
          ) : (
            <path d="M8 4.5v4.2M8 11.4h.01" fill="none" stroke="var(--text-on-accent)" strokeWidth="1.6" strokeLinecap="round" />
          )}
        </svg>
        {findings === undefined ? "Ready for review" : findingsTitle(findings)}
      </h3>
      {checks.length > 0 && (
        <ul className="result-chips">
          {checks.map((check) => (
            <li key={check.text} className={`result-chip is-${check.tone}`}>
              {check.tone === "pass" && (
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                  <path d="M20 6 9 17l-5-5" />
                </svg>
              )}
              {check.tone === "warn" && (
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" aria-hidden>
                  <path d="M12 7v6M12 17h.01" />
                </svg>
              )}
              {check.text}
            </li>
          ))}
        </ul>
      )}
      {lone && (
        <div className="result-document">
          <div className="result-paper" aria-hidden="true">
            <span className="result-paper-label">Draft</span>
            <span className="result-paper-name">{lone.name}</span>
            <span className="result-paper-lines" />
          </div>
          <div className="result-document-info">
            <p className="result-document-kind">Created by Doc.Haus</p>
            <h4>{lone.name}</h4>
            <div className="result-document-actions">
              <button type="button" className="primary result-review" onClick={() => onViewDocument(lone.name)}>
                Review draft
              </button>
              {loneCombined ? (
                <CombinedDownload matterId={matterId} name={loneCombined} />
              ) : (
                <a href={documentContentUrl(matterId, lone.name)} download={lone.name}>
                  Download
                </a>
              )}
            </div>
          </div>
        </div>
      )}
      {rows.length > 0 && (
        <ul className="result-stack" aria-label={lone ? "Also for review" : "Files for review"}>
          {rows.map((output) => (
            <li key={output.name}>
              <ReviewRow output={output} matterId={matterId} onView={onViewDocument} combined={combinedOf(output)} />
            </li>
          ))}
        </ul>
      )}
      {groups.length > 0 && (
        <details className="result-files" open={drafts.length + notes.length === 0 ? true : undefined}>
          <summary>
            All generated files <span className="result-file-count">{outputs.length + groups.filter((group) => group.combined).length}</span>
          </summary>
          <DraftedList matterId={matterId} groups={groups} onView={onViewDocument} />
        </details>
      )}
    </section>
  )
}

function ReviewRow({
  output,
  matterId,
  onView,
  combined,
}: {
  output: DraftedOutput
  matterId: string
  onView: (name: string) => void
  /** The one file this draft downloads as with the requested memo. */
  combined?: string
}) {
  const role = OUTPUT_ROLE[output.role]
  return (
    <div className={`result-row is-${output.role}`}>
      <span className="result-row-paper" aria-hidden="true">
        <span className="result-paper-label">{role}</span>
        <span className="result-paper-lines" />
      </span>
      <div className="result-row-info">
        <p className="result-document-kind">
          {role} · Created by Doc.Haus
        </p>
        <h4>{output.name}</h4>
      </div>
      <div className="result-document-actions">
        <button type="button" className="result-review" onClick={() => onView(output.name)} aria-label={`Review ${output.name}`}>
          Review
        </button>
        {combined ? (
          <CombinedDownload matterId={matterId} name={combined} />
        ) : (
          <a href={documentContentUrl(matterId, output.name)} download={output.name} aria-label={`Download ${output.name}`}>
            Download
          </a>
        )}
      </div>
    </div>
  )
}

/** A combined document's one file, downloaded in place of its draft. */
function CombinedDownload({ matterId, name }: { matterId: string; name: string }) {
  const { busy, download } = useCombinedDownload(matterId)
  return (
    <button
      type="button"
      className="linklike result-download"
      disabled={busy}
      onClick={() => download(name)}
      aria-label={`Download ${name}`}
      title="Download the memo and the draft as one .docx"
    >
      Download
    </button>
  )
}
