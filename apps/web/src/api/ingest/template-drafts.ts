import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

export type TemplateDraftChange = {
  before_quote: string
  replacement: string
  paragraph_before?: string
  paragraph_after?: string
}

export type TemplateDraft = {
  id: string
  dir: string
  source: { name: string; matterId: string }
  createdAt: number
  revision: number
  name: string
  description: string
  status: "draft" | "published"
  publishedName?: string
  changes: TemplateDraftChange[]
  notes: string[]
}

export async function listTemplateDrafts(): Promise<TemplateDraft[]> {
  const response = await fetch(`${INGEST_URL}/template-drafts`, { cache: "no-store" })
  if (!response.ok) throw new Error(await ingestError(response, "Could not load template conversations"))
  return ((await response.json()) as { drafts: TemplateDraft[] }).drafts
}

export async function createTemplateDraft(matterId: string, name: string): Promise<TemplateDraft> {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matterId)}/template-drafts`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name }),
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not attach this document to Template Builder"))
  return response.json()
}

export async function getTemplateDraft(id: string): Promise<TemplateDraft> {
  const response = await fetch(`${INGEST_URL}/template-drafts/${encodeURIComponent(id)}`, { cache: "no-store" })
  if (!response.ok) throw new Error(await ingestError(response, "Could not load this template conversation"))
  return response.json()
}

export async function fetchTemplateDraftBytes(id: string, view: "source" | "proposal"): Promise<Uint8Array> {
  const response = await fetch(`${INGEST_URL}/template-drafts/${encodeURIComponent(id)}/content?view=${view}`, { cache: "no-store" })
  if (!response.ok) throw new Error(await ingestError(response, "Could not load this document preview"))
  return new Uint8Array(await response.arrayBuffer())
}

export async function publishTemplateDraft(id: string, input: { revision: number; name: string; description: string }): Promise<TemplateDraft> {
  const response = await fetch(`${INGEST_URL}/template-drafts/${encodeURIComponent(id)}/publish`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not save this template"))
  return response.json()
}
