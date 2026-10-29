import { INGEST_URL } from "../../config"

export async function dashboardDirectory(signal?: AbortSignal): Promise<string> {
  const res = await fetch(`${INGEST_URL}/dashboard/directory`, { signal })
  if (!res.ok) throw new Error("Could not prepare the dashboard. Try again.")
  const value: unknown = await res.json()
  if (!value || typeof value !== "object" || !("dir" in value) || typeof value.dir !== "string" || !value.dir)
    throw new Error("Could not prepare the dashboard directory.")
  return value.dir
}
