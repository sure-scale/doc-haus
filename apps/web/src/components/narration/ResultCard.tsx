import { documentContentUrl } from "../../api/ingest/documents"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import DraftedList from "../drafting/DraftedList"
import { draftedGroups } from "../drafting/model"
import { resultChecks } from "./outcome"

/** A completed run in brief: what the checks established, what remains for the
 * lawyer, and the files it produced, each one click from the viewer. */
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
  const groups = draftedGroups(status)
  const featured = groups.find((group) => group.main.role === "draft")
  const count = groups.reduce((total, group) => total + 1 + group.extras.length, 0)
  return (
    <section id="run-result" className="narration-result" aria-labelledby="run-result-title">
      <h3 id="run-result-title" className="narration-result-title">
        <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden>
          <circle cx="8" cy="8" r="7" fill="currentColor" />
          <path d="m5 8.2 2 2 4-4.2" fill="none" stroke="var(--text-on-accent)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
        Ready for review
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
      {featured && (
        <div className="result-document">
          <div className="result-paper" aria-hidden="true">
            <span className="result-paper-label">Draft</span>
            <span className="result-paper-name">{featured.main.name}</span>
            <span className="result-paper-lines" />
          </div>
          <div className="result-document-info">
            <p className="result-document-kind">Created by Doc.Haus</p>
            <h4>{featured.main.name}</h4>
            <div className="result-document-actions">
              <button type="button" className="primary result-review" onClick={() => onViewDocument(featured.main.name)}>
                Review draft
              </button>
              <a href={documentContentUrl(matterId, featured.main.name)} download={featured.main.name}>
                Download
              </a>
            </div>
          </div>
        </div>
      )}
      {groups.length > 0 && (
        <details className="result-files" open={!featured || groups.filter((group) => group.main.role === "draft").length > 1 ? true : undefined}>
          <summary>
            All generated files <span className="result-file-count">{count}</span>
          </summary>
          <DraftedList matterId={matterId} groups={groups} onView={onViewDocument} />
        </details>
      )}
    </section>
  )
}
