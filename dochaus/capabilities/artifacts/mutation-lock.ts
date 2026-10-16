import { createHash } from "node:crypto"
import { realpathSync } from "node:fs"
import path from "node:path"
import { withReentrantDirectoryLock, withReentrantDirectoryLockAsync } from "../../lib/directory-lock"

export function withArtifactMutationLock<T>(directory: string, artifact: string, fn: () => T) {
  return withArtifactMutationLocks(directory, [artifact], fn)
}

export function withArtifactMutationLocks<T>(directory: string, artifacts: Iterable<string>, fn: () => T) {
  return withArtifactMutationLockPaths(artifactMutationLockPaths(directory, artifacts), fn)
}

export function withArtifactMutationLockAsync<T>(directory: string, artifact: string, fn: () => Promise<T>) {
  return withArtifactMutationLocksAsync(directory, [artifact], fn)
}

export async function withArtifactMutationLocksAsync<T>(
  directory: string,
  artifacts: Iterable<string>,
  fn: () => Promise<T>,
) {
  return withArtifactMutationLockPathsAsync(artifactMutationLockPaths(directory, artifacts), fn)
}

function withArtifactMutationLockPaths<T>(locks: readonly { lock: string; artifact: string }[], fn: () => T): T {
  const next = locks[0]
  if (!next) return fn()
  return withReentrantDirectoryLock(
    next.lock,
    `${next.artifact} is already being updated. Retry after the active mutation finishes.`,
    () => withArtifactMutationLockPaths(locks.slice(1), fn),
  )
}

async function withArtifactMutationLockPathsAsync<T>(
  locks: readonly { lock: string; artifact: string }[],
  fn: () => Promise<T>,
): Promise<T> {
  const next = locks[0]
  if (!next) return fn()
  return withReentrantDirectoryLockAsync(
    next.lock,
    `${next.artifact} is already being updated. Retry after the active mutation finishes.`,
    () => withArtifactMutationLockPathsAsync(locks.slice(1), fn),
  )
}

function artifactMutationLockPaths(directory: string, artifacts: Iterable<string>) {
  const root = canonicalArtifactIdentity(directory, ".")
  const lockRoot = canonicalArtifactIdentity(directory, ".dochaus/artifact-mutation-locks")
  if (!lockRoot.startsWith(`${root}${path.sep}`))
    throw new Error("The artifact mutation lock root resolves outside the matter.")
  const identities = [...new Set([...artifacts].map((artifact) => canonicalArtifactIdentity(directory, artifact)))]
  if (identities.some((identity) => identity !== root && !identity.startsWith(`${root}${path.sep}`)))
    throw new Error("The artifact mutation target resolves outside the matter.")
  return identities.toSorted().map((identity) => ({
    lock: path.join(lockRoot, createHash("sha256").update(identity).digest("hex")),
    artifact: path.basename(identity),
  }))
}

export function canonicalArtifactIdentity(directory: string, artifact: string) {
  const target = path.resolve(directory, artifact)
  const canonical = realpathOrFallback(target)
  return process.platform === "darwin" || process.platform === "win32" ? canonical.toLowerCase() : canonical
}

function realpathOrFallback(target: string): string {
  try {
    return realpathSync.native(target)
  } catch {
    const parent = path.dirname(target)
    if (parent === target) return target
    return path.join(realpathOrFallback(parent), path.basename(target))
  }
}
