import { INGEST_URL } from "../../config"

// GCP projects / AWS profiles visible to the host the engine runs on, for the
// provider-setup comboboxes in Settings. Empty when the host has no such
// sign-in (or the ingest service is down — the fields then take plain typing).
export async function listGcpProjects(): Promise<string[]> {
  const res = await fetch(`${INGEST_URL}/host/gcp-projects`, { signal: AbortSignal.timeout(10000) }).catch(() => null)
  if (!res?.ok) return []
  return ((await res.json()) as { projects: string[] }).projects
}

export async function listAwsProfiles(): Promise<string[]> {
  const res = await fetch(`${INGEST_URL}/host/aws-profiles`, { signal: AbortSignal.timeout(10000) }).catch(() => null)
  if (!res?.ok) return []
  return ((await res.json()) as { profiles: string[] }).profiles
}

// Verify a Vertex project/location with one real call made by the ingest
// service on the host — the same ADC the engine will use — instead of
// prompting through the engine, whose first-touch cold boot can outlast any
// probe timeout. Failures carry Google's error message verbatim.
export async function probeVertexHost(input: {
  project: string
  location: string
  publisher: "google" | "anthropic"
  model: string
}): Promise<{ ok: boolean; error?: string }> {
  const res = await fetch(`${INGEST_URL}/host/probe-vertex`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
    signal: AbortSignal.timeout(35000),
  }).catch(() => null)
  if (!res) return { ok: false, error: "Ingest service unreachable." }
  return res.json()
}
