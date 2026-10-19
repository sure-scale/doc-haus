import { existsSync, readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { hashStageInputs, leanFile, readLeanState } from "../workflows/lean/state"
import { LEAN_PROMPT_VERSION, stageInputs } from "../workflows/lean/stages"
import type { DocID } from "../workflows/lean/plan"

/** Current artifacts follow the accepted apply's initial-draft ancestry, never timestamps.
 * Applied/redline files no current apply accepted remain on disk as history. */
export function currentDocumentArtifacts(directory: string, document: DocID): { body?: string; redline?: string; applied: boolean; warning?: string } {
  const step = `${document}.apply` as const
  const accepted = readLeanState(directory)?.stages[step]
  const ancestry = appliedDraftMatches(directory, document)
  const current = ancestry ?? (accepted?.status === "done" && accepted.input_sha === hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
  const applied = `${document}/body-applied.docx`, draft = `${document}/body.docx`, redline = `${document}/redline.docx`
  if (current && existsSync(leanFile(directory, applied))) return {
    body: applied, applied: true, ...(existsSync(leanFile(directory, redline)) ? { redline } : {}),
  }
  return {
    ...(existsSync(leanFile(directory, draft)) ? { body: draft } : {}), applied: false,
    ...(existsSync(leanFile(directory, applied)) || existsSync(leanFile(directory, redline))
      ? { warning: "Saved applied/redline artifacts are not a current accepted apply result. Current evidence uses the initial draft; prior artifacts remain historical and must not supply current amendment targets." } : {}),
  }
}

/** Snapshot revision order is the existing sole writer's stage order. Pending
 * changes may replace the apply inputs without replacing the accepted Word. */
function appliedDraftMatches(directory: string, document: DocID): boolean | undefined {
  const root = leanFile(directory, "snapshots")
  if (!existsSync(root)) return
  const snapshots = readdirSync(root).flatMap(name => {
    const match = /^(\d+)-(.+)$/.exec(name)
    return match ? [{ name, revision: Number(match[1]), step: match[2] }] : []
  }).sort((a, b) => b.revision - a.revision)
  const applied = leanFile(directory, `${document}/body-applied.docx`)
  if (!existsSync(applied)) return false
  const apply = snapshots.find(row => row.step === `${document}.apply` && sameBytes(path.join(root, row.name, "body-applied.docx"), applied))
  if (!apply) return
  const draft = snapshots.find(row => row.step === `${document}.draft` && row.revision < apply.revision)
  if (!draft) return
  return sameBytes(path.join(root, draft.name, "body.docx"), leanFile(directory, `${document}/body.docx`))
}

function sameBytes(a: string, b: string) {
  return existsSync(a) && existsSync(b) && readFileSync(a).equals(readFileSync(b))
}
