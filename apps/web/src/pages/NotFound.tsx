import { useNavigate } from "react-router-dom"

// Any unknown path, and a matter link whose matter no longer exists.
export default function NotFound({ title = "Page not found", detail }: { title?: string; detail?: string }) {
  const navigate = useNavigate()
  return (
    <div className="not-found">
      <h1>{title}</h1>
      <p>{detail ?? "This page doesn't exist. It may have been moved, or the link may be mistyped."}</p>
      <button type="button" className="primary" onClick={() => navigate("/")}>
        Back to matters
      </button>
    </div>
  )
}
