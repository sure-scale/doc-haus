import path from "node:path"
import { readRegularFile } from "./regular-file"

/** Model capabilities are scoped from their native session, never a model's
 * matter ID. Only reading and staging are exposed; decisions belong to the UI. */
export function chronologyMatter(directory: string) {
  const bytes = readRegularFile(path.join(directory, "matter.json"))
  if (!bytes) throw new Error("Open an existing matter before using its chronology.")
  const matter: unknown = JSON.parse(bytes.toString("utf8"))
  if (
    !matter ||
    typeof matter !== "object" ||
    !("id" in matter) ||
    typeof matter.id !== "string" ||
    !/^[a-zA-Z0-9_-]{1,160}$/.test(matter.id)
  )
    throw new Error("This session has no valid matter identity.")
  return matter.id
}

export async function chronologyRequest(directory: string, signal: AbortSignal, body?: unknown) {
  const matterId = chronologyMatter(directory)
  const base = (process.env.INGEST_URL ?? "http://127.0.0.1:4500").replace(/\/$/, "")
  const response = await fetch(
    `${base}/matters/${encodeURIComponent(matterId)}/chronology${body === undefined ? "" : "/proposals"}`,
    {
      method: body === undefined ? "GET" : "POST",
      headers: { "content-type": "application/json", "x-dochaus-actor": "local agent session" },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      signal,
    },
  )
  const result = (await response.json()) as Record<string, unknown>
  if (!response.ok)
    throw new Error(
      typeof result.error === "string"
        ? result.error
        : `Chronology request failed (${response.status}). Reload the saved chronology before retrying.`,
    )
  return { matterId, result }
}
