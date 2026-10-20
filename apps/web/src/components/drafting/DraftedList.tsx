import type { ReactNode } from "react"
import { documentContentUrl } from "../../api/ingest/documents"
import { CompareIcon } from "../CompareDialog"
import { OUTPUT_ROLE, type DraftedGroup, type DraftedOutput } from "./model"

// The Word files a drafting run produced: each draft with its redline (and, for
// a one-document run, the drafting notes) nested beneath it. The name opens the
// file in the viewer; the icon button downloads the .docx. `fresh` marks rows the
// run produced while the lawyer watched, for a one-time highlight. A combined
// document's one file leads its group, with the draft it holds beneath it; a
// combined file from an earlier run says so on its row.
export default function DraftedList({
  matterId,
  groups,
  onView,
  onCompare,
  pending,
  fresh,
  renderLabel,
}: {
  matterId: string
  groups: DraftedGroup[]
  onView?: (name: string) => void
  /** Compare a Word file with another version; the row offers it when given. */
  onCompare?: (name: string) => void
  /** Pending redline changes per file, shown as a badge on its row. */
  pending?: Map<string, number>
  fresh?: boolean
  renderLabel?: (output: DraftedOutput) => ReactNode
}) {
  return (
    <ul className={`drafted-list${fresh ? " is-fresh" : ""}`}>
      {groups.map((group) => {
        const beneath = group.combined ? [group.main, ...group.extras] : group.extras
        return (
          <li key={group.main.name} className="drafted-group">
            {group.combined ? (
              <CombinedRow matterId={matterId} name={group.combined} />
            ) : (
              <OutputRow matterId={matterId} output={group.main} onView={onView} onCompare={onCompare} pending={pending?.get(group.main.name)} label={renderLabel?.(group.main)} main />
            )}
            {beneath.length > 0 && (
              <ul className="drafted-extras">
                {beneath.map((output) => (
                  <li key={output.name}>
                    <OutputRow matterId={matterId} output={output} onView={onView} onCompare={onCompare} pending={pending?.get(output.name)} label={renderLabel?.(output)} />
                  </li>
                ))}
              </ul>
            )}
          </li>
        )
      })}
    </ul>
  )
}

/** The combined file downloads as the file saved in the matter, as every
 * other row does. */
function CombinedRow({ matterId, name }: { matterId: string; name: string }) {
  return (
    <div className="drafted-row is-main">
      <FileIcon />
      <div className="drafted-file-details">
        <span className="drafted-name" title={name}>
          <span className="drafted-name-text">{name}</span>
        </span>
        <div className="drafted-file-meta">
          <span className="document-origin is-created">Created by Doc.Haus</span>
          <span className="drafted-tag is-combined">Memo and draft</span>
        </div>
      </div>
      <a
        className="icon-btn drafted-download"
        href={documentContentUrl(matterId, name)}
        download={name}
        aria-label={`Download ${name}`}
        title="Download the memo and the draft as one .docx"
      >
        <DownloadIcon />
      </a>
    </div>
  )
}

function OutputRow({
  matterId,
  output,
  onView,
  onCompare,
  pending,
  main,
  label,
}: {
  matterId: string
  output: DraftedOutput
  onView?: (name: string) => void
  onCompare?: (name: string) => void
  pending?: number
  main?: boolean
  label?: ReactNode
}) {
  return (
    <div className={`drafted-row${main ? " is-main" : ""}`}>
      <FileIcon />
      <div className="drafted-file-details">
        {onView ? (
          <button type="button" className="linklike drafted-name" onClick={() => onView(output.name)} title={`Open ${output.name}`}>
            <span className="drafted-name-text">{output.name}</span>
          </button>
        ) : (
          <span className="drafted-name" title={output.name}>
            <span className="drafted-name-text">{output.name}</span>
          </span>
        )}
        <div className="drafted-file-meta">
          <span className="document-origin is-created">Created by Doc.Haus{output.role === "earlier_combined" ? " in an earlier run" : ""}</span>
          {label ?? <span className={`drafted-tag is-${output.role}`}>{OUTPUT_ROLE[output.role]}</span>}
        </div>
      </div>
      {pending ? (
        <span className="redline-badge" title={`${pending} pending change${pending === 1 ? "" : "s"} to review`}>
          {pending}
        </span>
      ) : null}
      {onCompare && /\.docx$/i.test(output.name) && (
        <button
          type="button"
          className="icon-btn drafted-compare"
          onClick={() => onCompare(output.name)}
          aria-label={`Compare ${output.name} with another version`}
          title="Compare with another version"
        >
          <CompareIcon size={14} />
        </button>
      )}
      <a
        className="icon-btn drafted-download"
        href={documentContentUrl(matterId, output.name)}
        download={output.name}
        aria-label={`Download ${output.name}`}
        title="Download .docx"
      >
        <DownloadIcon />
      </a>
    </div>
  )
}

function FileIcon() {
  return (
    <span className="drafted-file-icon" aria-hidden="true">
      <svg width="18" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
        <path d="M14 2v6h6M8 13h8M8 17h5" />
      </svg>
    </span>
  )
}

function DownloadIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
      <path d="M12 4v11" />
      <path d="m7 10 5 5 5-5" />
      <path d="M5 20h14" />
    </svg>
  )
}
