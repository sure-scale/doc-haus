import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

// Clients added on their own, before or without a matter. A client named on a
// matter needs no record here; lib/clients.ts joins the two.
export type StoredClient = { id: number; name: string; reference?: string; created_at: number }

export async function listClients(): Promise<StoredClient[]> {
  const res = await fetch(`${INGEST_URL}/clients`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load clients"))
  return res.json()
}

// A name that is already a client returns that client with created false.
export async function createClient(name: string, reference?: string): Promise<{ client: StoredClient; created: boolean }> {
  const res = await fetch(`${INGEST_URL}/clients`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, reference }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not add the client"))
  return { client: await res.json(), created: res.status === 201 }
}

export async function deleteClient(id: number): Promise<void> {
  const res = await fetch(`${INGEST_URL}/clients/${id}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, "Could not delete the client"))
}
