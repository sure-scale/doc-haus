import type { Provider } from "./types"

export default function HostSignInStatus({ providers, verified }: { providers: Provider[]; verified: Set<string> }) {
  const detected = providers.filter(provider => provider.activation !== "disabled")
  return <section className="settings-section">
    <h3>Connection status</h3>
    {detected.map(provider => <div key={provider.id}>
      <h4>{provider.id === "amazon-bedrock" ? "Amazon Bedrock" : provider.id === "google-vertex" ? "Google Vertex AI" : provider.name}</h4>
      <p>{provider.credentialSources.length ? `Detected credentials: ${provider.credentialSources.join(", ")}.` : provider.nativeAvailable ? "Available models; model access has not been tested here." : provider.host?.sources.length ? "Cloud sign-in source detected; provider setup is incomplete." : "No existing connection detected. Sign in below."}</p>
      <p className="settings-hint muted">Detection does not run a model test.</p>
      {!!provider.host?.sources.length && <p>Detected on server: {provider.host.sources.join(", ")}.</p>}
      {provider.host?.type === "aws-chain" ? <p className="settings-hint muted">
        {provider.host.options.profile && <>AWS profile: {provider.host.options.profile}. </>}
        {provider.host.options.region ? <>Region: {provider.host.options.region}.</> : "Choose an AWS region below."}
      </p> : provider.host && <p className="settings-hint muted">
        {provider.host.options.project ? <>GCP project: {provider.host.options.project}. </> : "Choose a GCP project below. ADC alone does not select a project. "}
        {provider.host.options.location && <>Location: {provider.host.options.location}.</>}
      </p>}
      {provider.host && <>
        <p className="settings-hint muted">{verified.has(provider.id) ? "Model access was explicitly verified for this browser." : "Test model access below before choosing cloud models."}</p>
      </>}
    </div>)}
  </section>
}
