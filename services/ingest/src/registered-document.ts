import { deliverableDocumentName } from "./deliverable"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"
import { matterDocumentName } from "./ingest"

// How routes that read one matter document by name resolve it: a source or
// deliverable basename with a supported extension, never a path, then the
// matter file that owns that name. A symbolic-link alias resolves to nothing.
export function registeredDocumentName(requested: string) {
  return matterDocumentName(requested) ?? deliverableDocumentName(requested)
}

export function readDocumentTarget(directory: string, name: string) {
  try {
    return resolveMatterDocumentTarget(directory, name)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) return
    throw error
  }
}
