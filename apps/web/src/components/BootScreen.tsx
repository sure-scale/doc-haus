import { useEffect, useState } from "react"

// Both backends (the engine and the document service) are separate processes
// that may still be starting when the web app first loads. Until both answer,
// the list pages would fetch a dead socket and fall through to a misleading
// "nothing here yet" empty state — so we hold the whole app behind this screen
// and surface which service is still coming up.

// Shown one at a time while the backends come up. Whimsical or factually true;
// nothing that implies real legal work is happening (same rule as onboarding).
const BOOT_PHRASES = [
  "Sharpening pencils...",
  "Unlocking the filing cabinet...",
  "Warming up the engine...",
  "Stacking the case files...",
  "Aligning the margins...",
  "Untangling the red tape...",
]

export default function BootScreen({ engineReady, ingestReady, error, onRetry }: { engineReady: boolean; ingestReady: boolean; error?: string; onRetry?: () => void }) {
  const [phrase, setPhrase] = useState(0)
  useEffect(() => {
    const t = setInterval(() => setPhrase((p) => (p + 1) % BOOT_PHRASES.length), 2200)
    return () => clearInterval(t)
  }, [])

  const steps = [
    { label: "Starting the engine", ready: engineReady },
    { label: "Opening the document service", ready: ingestReady },
  ]

  return (
    <div className="boot-screen">
      <div className="boot-loader">
        <div className="boot-house" aria-hidden="true">
          <span className="boot-house-line is-base-left" />
          <span className="boot-house-line is-wall-left" />
          <span className="boot-house-line is-roof-left" />
          <span className="boot-house-line is-roof-right" />
          <span className="boot-house-line is-wall-right" />
          <span className="boot-house-line is-base-right" />
        </div>
        {/* Keyed on the index so each phrase replays the fade-in. */}
        <p key={phrase} className="onboard-phrase">
          {BOOT_PHRASES[phrase]}
        </p>
      </div>
      {error && <div role="alert" className="boot-error"><p>{error}</p><button onClick={onRetry}>Retry connection</button></div>}
      <ul className="boot-steps">
        {steps.map((s) => (
          <li key={s.label} className={s.ready ? "boot-step done" : "boot-step"}>
            <span className="boot-dot" />
            {s.label}
          </li>
        ))}
      </ul>
    </div>
  )
}
