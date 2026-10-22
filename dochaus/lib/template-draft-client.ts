import path from "node:path"

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export type TemplateDraftInfo = {
  id: string
  dir: string
  source: { name: string; matterId: string }
  revision: number
  name: string
  description: string
  status: "draft" | "published"
  notes: string[]
  changes: { before_quote: string; replacement: string }[]
}

// The model never chooses a matter, source path or draft ID. The native session's
// directory selects the handoff; ingest confirms that it owns that directory.
export function isTemplateDraftDirectory(directory: string) {
  return path.basename(path.dirname(path.resolve(directory))) === ".template-drafts"
}

export async function attachedTemplateDraft(directory: string, signal?: AbortSignal): Promise<TemplateDraftInfo> {
  const resolved = path.resolve(directory)
  const id = path.basename(resolved)
  if (!isTemplateDraftDirectory(resolved) || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id))
    throw new Error("No source document is attached to this conversation. Start Create reusable template from a document.")
  const draft = await templateDraftRequest<TemplateDraftInfo>(id, "", { signal })
  if (path.resolve(draft.dir) !== resolved || draft.id !== id)
    throw new Error("The attached template document does not belong to this conversation.")
  return draft
}

export async function templateDraftRequest<T>(id: string, route: string, options?: RequestInit): Promise<T> {
  const response = await fetch(`${ingestUrl}/template-drafts/${encodeURIComponent(id)}${route}`, options)
  if (!response.ok) {
    const body = await response.json().catch(() => undefined) as { error?: string } | undefined
    throw new Error(body?.error ?? `The attached template document could not be loaded (${response.status}).`)
  }
  return response.json() as Promise<T>
}

export const TEMPLATE_DRAFT_REVIEW_MESSAGE =
  "This conversation prepares the attached document as a reusable template. Use read-template-source and prepare-template, then ask the user to open Review proposed template and save it. Library changes cannot bypass that review."
