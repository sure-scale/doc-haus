import { INGEST_URL } from "../../config"
import { VersionUnavailableError } from "./versions"
import { documentContentUrl } from "./documents"

export function documentExportUrl(matterId: string, source: { versionId: string } | { name: string; proposals?: boolean }) {
  const query = "versionId" in source ? `versionId=${encodeURIComponent(source.versionId)}` : `name=${encodeURIComponent(source.name)}${source.proposals ? "&proposals=1" : ""}`
  return `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/export?${query}`
}
export function documentDownloadUrl(matterId: string, name: string) {
  return /\.docx$/i.test(name) ? documentExportUrl(matterId, { name }) : documentContentUrl(matterId, name)
}

/** User-facing sharing copies. Preview/source/archival readers continue using
 * their exact content routes; exporting never records a send or a version. */
export async function fetchDocumentExportBytes(matterId: string,
  source: { versionId: string } | { name: string; proposals?: boolean },
  notice?: (message: string | undefined) => void,
): Promise<Uint8Array> {
  const response = await fetch(documentExportUrl(matterId, source))
  if (!response.ok) {
    const detail = await response.json().catch(() => undefined) as { error?: unknown; code?: unknown } | undefined
    const message = typeof detail?.error === "string" ? detail.error : `Could not prepare this Word copy (${response.status})`
    if ("versionId" in source && [404, 409, 410].includes(response.status))
      throw new VersionUnavailableError(response.status, typeof detail?.code === "string" ? detail.code : undefined, message)
    throw new Error(message)
  }
  const encoded = response.headers.get("X-DocHaus-Revision-Notice")
  notice?.(encoded ? decodeURIComponent(encoded) : undefined)
  return new Uint8Array(await response.arrayBuffer())
}
