import { DeliverableConflictError, writeDeliverable } from "./deliverable"

// Lean drafting's deliverable port: finished Word files become matter documents
// with role "deliverable" through the same registry write as every other
// generated deliverable. A refused name (a source, a lawyer's accepted revision,
// or a file whose bytes changed) is reported as undefined so the run tries the
// next name; any other failure propagates. Each registration is the system's own
// write: its stored bytes are recorded (system-docx.ts), so a saved combined file
// stays the system's to rebuild until anything else writes it.
export function registerLeanDeliverable(directory: string, fileName: string, bytes: Uint8Array, priorSha256: string | null) {
  return writeDeliverable(directory, fileName, Buffer.from(bytes), {
    precondition: priorSha256 ? { kind: "replace", prior_sha256: priorSha256 } : { kind: "create" },
    system: true,
  }).then(
    written => written.name,
    (error: unknown) => {
      if (error instanceof DeliverableConflictError) return undefined
      throw error
    },
  )
}
