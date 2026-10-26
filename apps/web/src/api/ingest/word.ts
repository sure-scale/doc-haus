import type { WordOpenDocument } from "../../../../../dochaus/lib/word-addin-contract"
import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

export type { WordOpenDocument }
export type WordConfig = { available: boolean; origin: string | null; manifestUrl: string | null; reason?: string }

const wordUrl = (matterId: string, name?: string) => `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/word${name === undefined ? "" : `?name=${encodeURIComponent(name)}`}`

export async function fetchWordConfig(): Promise<WordConfig> {
  const response = await fetch(`${INGEST_URL}/word/config`)
  if (!response.ok) throw new Error(await ingestError(response, "Could not check Word setup"))
  return response.json()
}

export async function fetchWordDocuments(matterId: string): Promise<WordOpenDocument[]> {
  const response = await fetch(wordUrl(matterId))
  if (!response.ok) throw new Error(await ingestError(response, "Could not check which documents are open in Word"))
  return ((await response.json()) as { documents: WordOpenDocument[] }).documents
}

// The normal app requests a linked copy; its write-only capability stays in
// that copy. The browser never reads or stores the capability.
export async function linkedWordDocument(matterId: string, name: string, origin: string): Promise<Blob> {
  const response = await fetch(wordUrl(matterId, name), {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ origin }),
  })
  if (!response.ok) throw new Error(await ingestError(response, `Could not open ${name} in Word`))
  return response.blob()
}

export async function revokeWordDocument(matterId: string, name: string): Promise<void> {
  const response = await fetch(wordUrl(matterId, name), { method: "DELETE" })
  if (!response.ok) throw new Error(await ingestError(response, `Could not disconnect ${name} from Word`))
}
