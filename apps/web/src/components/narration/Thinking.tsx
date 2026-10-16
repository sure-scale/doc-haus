import { useEffect, useState } from "react"
import { elapsed } from "./timeline"

// The live "work is happening" pill. It sits at the foot of the narration and
// sticks to the bottom of the chat log, so it stays in view while the run works.
// It names what the run is doing right now and how long the stage has run.
export default function Thinking({
  label,
  since,
  summarising,
}: {
  label: string
  since: number
  summarising: boolean
}) {
  const [now, setNow] = useState(Date.now())
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 1_000)
    return () => clearInterval(timer)
  }, [])
  const text = summarising ? "Summarising progress" : label
  return (
    <div className="narration-thinking" role="status" aria-live="polite" title={text}>
      <span className="narration-pulse" aria-hidden />
      <span className="narration-thinking-label">
        <span className="narration-thinking-text">{text}</span>
        <span className="narration-dots" aria-hidden>
          <span>.</span>
          <span>.</span>
          <span>.</span>
        </span>
      </span>
      <span className="narration-meta" aria-hidden>
        {elapsed(now - since)}
      </span>
    </div>
  )
}
