import "./ai-setup.css"

export type ModelTestResult = {
  role: "primary" | "fast"
  model: string
  status: "checking" | "ready" | "auth" | "model" | "unavailable"
  message?: string
  shared?: boolean
}

export type TestResultsProps = {
  results: ModelTestResult[]
  connectionStatus?: string
}

const outcomes = {
  checking: { label: "Checking…", message: "Waiting for the selected model to respond." },
  ready: { label: "Ready", message: "The selected model responded." },
  auth: { label: "Connection failed", message: "Check the connection and sign-in, then test again." },
  model: { label: "Model unavailable", message: "Check the model ID and your account's access to this model." },
  unavailable: { label: "Could not test", message: "The test could not complete. Check the connection, then try again." },
} as const

export default function TestResults({ results, connectionStatus }: TestResultsProps) {
  return <div className="ai-setup-test-results" role="status" aria-live="polite" aria-atomic="true">
    {connectionStatus && <p className="settings-hint ai-setup-connection-status">{connectionStatus}</p>}
    {results.length > 0 && <ul className="ai-setup-result-list" aria-label="Model test results">
      {results.map(result => <li key={result.role} className={`ai-setup-result ai-setup-result--${result.status}`}>
        <div className="ai-setup-result-heading">
          <strong>{result.role === "primary" ? "Primary model" : "Fast model"}</strong>
          <span className="ai-setup-result-status">{outcomes[result.status].label}</span>
        </div>
        <code className="ai-setup-result-model">{result.model}</code>
        <p className="settings-hint">{result.message || outcomes[result.status].message}</p>
        {result.shared && <p className="settings-hint muted">Uses the primary model test.</p>}
      </li>)}
    </ul>}
  </div>
}
