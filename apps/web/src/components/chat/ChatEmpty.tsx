// The chat's empty state. A matter with no documents yet has nothing for the
// starter questions to ask about, so it leads with adding documents instead;
// surfaces without documents (the library builders) pass no count and always
// get their starters.
export default function ChatEmpty({
  hint,
  starters,
  documentCount,
  onPick,
  onAddDocuments,
}: {
  hint: string
  starters: string[]
  documentCount?: number
  onPick: (starter: string) => void
  onAddDocuments?: () => void
}) {
  if (documentCount === 0)
    return (
      <div className="chat-empty is-upload-first">
        <IconUpload />
        <p className="chat-empty-title">Start by adding this matter's documents</p>
        <p className="muted">
          Drop Word, PDF, email or Excel files onto the message box below, or add them in Documents. Once they are read
          you can ask about them here, and every answer cites the source section.
        </p>
        {onAddDocuments && (
          <button type="button" className="primary chat-empty-action" onClick={onAddDocuments}>
            Add documents
          </button>
        )}
      </div>
    )

  return (
    <div className="chat-empty">
      <p className="chat-empty-intro">{hint}</p>
      <div className="starters">
        {starters.map((s) => (
          <button key={s} className="starter" onClick={() => onPick(s)}>
            <span className="starter-text">{s}</span>
            <svg className="starter-arrow" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
              <path d="M5 12h14" />
              <path d="m12 5 7 7-7 7" />
            </svg>
          </button>
        ))}
      </div>
    </div>
  )
}

function IconUpload() {
  return (
    <svg className="chat-empty-icon" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <path d="m17 8-5-5-5 5" />
      <path d="M12 3v12" />
    </svg>
  )
}
