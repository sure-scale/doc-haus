// In-chat card for each document a turn drafted: the new file's name with a link
// straight into the document viewer, so the lawyer opens the draft without
// hunting for it in the documents rail.
export default function DraftView({ drafts, onView }: { drafts: string[]; onView: (name: string) => void }) {
  if (drafts.length === 0) return null
  return (
    <div className="drafts">
      {drafts.map((name, i) => (
        <div className="draft-card" key={i}>
          <svg className="draft-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <path d="M14 2v6h6" />
          </svg>
          <span className="draft-doc">{name}</span>
          <button className="linklike" onClick={() => onView(name)}>
            Open document
          </button>
        </div>
      ))}
    </div>
  )
}
