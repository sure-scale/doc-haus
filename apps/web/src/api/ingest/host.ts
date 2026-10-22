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
