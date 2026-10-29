/** A source format icon is descriptive, with no credibility or privilege meaning. */
export default function SourceIcon({ name = "", kind = "" }: { name?: string; kind?: string }) {
  const email = /\.(eml|msg)$/i.test(name) || /\b(email|e-mail|correspondence)\b/i.test(kind)
  const sheet = /\.xlsx$/i.test(name)
  return (
    <svg
      className="chronology-detail-source-icon"
      width="17"
      height="17"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.3"
      aria-hidden="true"
    >
      {email ? (
        <>
          <rect x="2" y="4" width="16" height="12" rx="1.5" />
          <path d="m2 5 8 6 8-6" />
        </>
      ) : sheet ? (
        <>
          <rect x="3" y="2" width="14" height="16" rx="1" />
          <path d="M3 7h14M3 12h14M8 2v16" />
        </>
      ) : (
        <path d="M5 2.5h6l4 4v11H5zM11 2.5v4h4M8 10h4M8 13h4" />
      )}
    </svg>
  )
}
