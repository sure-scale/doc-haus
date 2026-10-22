import { useNavigate } from "react-router-dom"

// Any unknown path, and a matter link whose matter no longer exists.
export default function NotFound({ title = "Page not found", detail }: { title?: string; detail?: string }) {
  const navigate = useNavigate()
  return (
    <div className="not-found">
      <span className="not-found-symbol" aria-hidden="true">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6M9 13h6M9 17h3" /></svg>
      </span>
      <h1>{title}</h1>
      <p>{detail ?? "This page doesn't exist. It may have been moved, or the link may be mistyped."}</p>
      <button type="button" className="primary" onClick={() => navigate("/")}>
        Back to matters
      </button>
    </div>
  )
}
