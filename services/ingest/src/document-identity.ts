import { lstatSync, readdirSync, realpathSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../../../dochaus/capabilities/artifacts/mutation-lock"

export class MatterDocumentIdentityError extends Error {}

// Resolve one validated matter-document basename to the physical spelling that
// owns its artifact lock. This keeps the filesystem, registry, and mutation
// boundary on the same identity even on case-insensitive filesystems. Symlinked
// documents are rejected because an alias (or an out-of-matter target) must not
// acquire an independent source/deliverable role.
export function resolveMatterDocumentTarget(matterDir: string, fileName: string) {
  const root = realpathSync.native(matterDir)
  const requested = path.join(root, fileName)
  const requestedStat = lstatOrUndefined(requested)
  if (requestedStat?.isSymbolicLink())
    throw new MatterDocumentIdentityError(`${fileName} is a symbolic-link alias, not a matter document.`)
  const requestedIdentity = canonicalArtifactIdentity(root, fileName)
  const entries = readdirSync(root)
  const exact = entries.find((entry) => entry === fileName)
  const existing = exact ?? entries.find(
    (entry) => !lstatSync(path.join(root, entry)).isSymbolicLink() && canonicalArtifactIdentity(root, entry) === requestedIdentity,
  )
  const name = existing ?? fileName
  const docPath = path.join(root, name)
  if (existing && lstatSync(docPath).isSymbolicLink())
    throw new MatterDocumentIdentityError(`${fileName} is a symbolic-link alias, not a matter document.`)
  return { name, docPath, identity: canonicalArtifactIdentity(root, name) }
}

export function documentRegistryIdentity(docPath: string) {
  return canonicalArtifactIdentity(path.dirname(docPath), path.basename(docPath))
}

function lstatOrUndefined(file: string) {
  try {
    return lstatSync(file)
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") return
    throw error
  }
}
