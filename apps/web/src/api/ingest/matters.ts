import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import type { Document } from "./documents"
import type { Posture } from "./preferences"

// Matters and documents live in the ingest service: it owns matter directories
// under WORKSPACE_ROOT and turns uploaded DOCX into the per-matter embedding DB.
// OpenCode has no upload endpoint, so all document I/O goes through here.

export type Matter = {
  id: string
  title: string
  reference?: string
  jurisdictions?: string[]
  playbook?: string
  // Whom the firm acts for on the matter; drafting takes that side when the
  // request does not say whom the lawyer acts for.
  client?: string
  // How hard drafts lean toward that client on this matter; absent means the firm's.
  posture?: Posture
  // The fictional sample matter, labelled as a sample in the matters list.
  sample?: boolean
  dir: string
  created_at: number
}
// A jurisdiction pack the matter can be steered by (dochaus/jurisdiction/<code>).
export type Jurisdiction = { code: string; name: string; citationStyle: string }
// A playbook the matter can be steered by (dochaus/playbook-<name>). Names start
// with "playbook-"; unlike jurisdictions a matter binds at most one.
export type Playbook = {
  name: string
  description: string
  content: string
  version: string
  last_reviewed: string
  builtin: boolean
  updated_at: number
  matters: string[]
}

export type MatterDetail = Matter & { documents: Document[] }

export async function listMatters(): Promise<Matter[]> {
  const res = await fetch(`${INGEST_URL}/matters`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load matters"))
  return res.json()
}

export async function listJurisdictions(): Promise<Jurisdiction[]> {
  const res = await fetch(`${INGEST_URL}/jurisdictions`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load jurisdictions"))
  return res.json()
}

export async function listPlaybooks(): Promise<Playbook[]> {
  const res = await fetch(`${INGEST_URL}/playbooks`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load playbooks"))
  return res.json()
}

export async function createMatter(
  title: string,
  reference?: string,
  jurisdictions?: string[],
  playbook?: string,
  client?: string,
  posture?: Posture,
): Promise<Matter> {
  const res = await fetch(`${INGEST_URL}/matters`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, reference, jurisdictions, playbook, client, posture }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not create the matter"))
  return res.json()
}

// An undefined client or posture drops out of the JSON body, which keeps the
// stored one; "" clears it. Callers that do not edit them therefore never wipe them.
export async function renameMatter(
  id: string,
  title: string,
  reference?: string,
  jurisdictions?: string[],
  playbook?: string,
  client?: string,
  posture?: Posture | "",
): Promise<Matter> {
  const res = await fetch(`${INGEST_URL}/matters/${id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, reference, jurisdictions, playbook, client, posture }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not save the matter"))
  return res.json()
}

// Adds the fictional sample matter (an engagement letter with cited Q&A and a
// legal review to explore). Idempotent: an existing sample is returned as is.
export async function addSampleMatter(): Promise<Matter> {
  const res = await fetch(`${INGEST_URL}/sample-matter`, { method: "POST" })
  if (!res.ok) throw new Error(await ingestError(res, "Could not add the sample matter"))
  return ((await res.json()) as { matter: Matter }).matter
}

export async function deleteMatter(id: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/matters/${id}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, "Could not delete the matter"))
}

export async function getMatter(id: string): Promise<MatterDetail> {
  const res = await fetch(`${INGEST_URL}/matters/${id}`)
  if (res.status === 404) throw new MatterNotFound(id)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load the matter"))
  return res.json()
}

export async function deleteMatterSession(matterId: string, sessionID: string): Promise<void> {
  const res = await fetch(
    `${INGEST_URL}/matters/${matterId}/sessions/${encodeURIComponent(sessionID)}`,
    { method: "DELETE" },
  )
  if (!res.ok) throw new Error(await ingestError(res, "Failed to delete conversation"))
}

// Thrown by getMatter when the id names no matter, so the page can say so
// instead of showing a loading skeleton forever.
export class MatterNotFound extends Error {
  constructor(id: string) {
    super(`No matter with id ${id}`)
    this.name = "MatterNotFound"
  }
}
